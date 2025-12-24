#!/usr/bin/env python3
"""
PDF図抽出Mixinモジュール

PDFToMarkdownConverterクラスの図抽出機能を提供します。
このモジュールはMixinクラスとして設計されており、メインクラスから継承されます。

機能:
- ベクター図形と埋め込み画像の統合抽出
- 図のクラスタリングとカラム判定
- 図内テキストの抽出
- 画像のレンダリングと保存
"""

import os
import re
from typing import List, Dict, Any, Tuple, Optional, Set

try:
    import fitz
except ImportError as e:
    raise ImportError(
        "PyMuPDFライブラリが必要です: pip install PyMuPDF または uv sync を実行してください"
    ) from e

try:
    from PIL import Image
except ImportError as e:
    raise ImportError(
        "Pillowライブラリが必要です: pip install pillow または uv sync を実行してください"
    ) from e


def debug_print(*args, **kwargs):
    """デバッグ出力（pdf2mdモジュールに委譲）"""
    try:
        from pdf2md import debug_print as _dp
        _dp(*args, **kwargs)
    except ImportError:
        pass


class _FiguresMixin:
    """図抽出機能を提供するMixinクラス
    
    このクラスはPDFToMarkdownConverterに継承され、
    図抽出、クラスタリング、レンダリング機能を提供します。
    
    注意: このクラスは単独では使用できません。
    PDFToMarkdownConverterクラスと組み合わせて使用してください。
    """

    def _extract_all_figures(
        self, page, page_num: int, header_footer_patterns: Set[str] = None
    ) -> List[Dict[str, Any]]:
        """ベクター図形と埋め込み画像を統合して図を抽出
        
        ベクター描画と埋め込み画像を統合してクラスタリングし、
        クラスタリング後にカラム判定を行う（先にクラスタリング、後でカラム判定）。
        ヘッダー/フッター領域内の図クラスタは除外する。
        
        Args:
            page: PyMuPDFのページオブジェクト
            page_num: ページ番号
            header_footer_patterns: ヘッダ・フッタパターンのセット
            
        Returns:
            抽出された図の情報リスト
        """
        if header_footer_patterns is None:
            header_footer_patterns = set()
        
        figures = []
        page_width = page.rect.width
        page_height = page.rect.height
        gutter_x = page_width / 2
        gutter_margin = 10.0  # ガター跨ぎ判定のマージン
        
        # ヘッダー/フッター領域のY座標境界を動的に計算
        header_y_max = None
        footer_y_min = None
        
        if header_footer_patterns:
            try:
                text_dict = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)
                header_lines_y = []
                footer_lines_y = []
                
                for block in text_dict.get("blocks", []):
                    if block.get("type") != 0:
                        continue
                    
                    for line in block.get("lines", []):
                        line_text = ""
                        for span in line.get("spans", []):
                            line_text += span.get("text", "")
                        
                        line_text = line_text.strip()
                        if not line_text:
                            continue
                        
                        line_bbox = line.get("bbox", (0, 0, 0, 0))
                        y_center = (line_bbox[1] + line_bbox[3]) / 2
                        
                        # 既存の_is_header_footer関数でヘッダー/フッター判定
                        if self._is_header_footer(
                            line_text, header_footer_patterns,
                            y_pos=line_bbox[1], page_height=page_height
                        ):
                            # ページ中央より上ならヘッダー、下ならフッター
                            if y_center < page_height / 2:
                                header_lines_y.append(line_bbox[3])
                            else:
                                footer_lines_y.append(line_bbox[1])
                
                # ヘッダー/フッター領域の境界を決定
                # ヘッダー: テキスト下端+マージン、またはページ上端から8%の大きい方
                # フッター: テキスト上端-マージン、またはページ下端から8%の小さい方
                if header_lines_y:
                    text_based_y = max(header_lines_y) + 15.0
                    position_based_y = page_height * 0.08
                    header_y_max = max(text_based_y, position_based_y)
                    debug_print(f"[DEBUG] page={page_num+1}: ヘッダー領域検出 y_max={header_y_max:.1f} (text={text_based_y:.1f}, pos={position_based_y:.1f})")
                if footer_lines_y:
                    text_based_y = min(footer_lines_y) - 15.0
                    position_based_y = page_height * 0.92
                    footer_y_min = min(text_based_y, position_based_y)
                    debug_print(f"[DEBUG] page={page_num+1}: フッター領域検出 y_min={footer_y_min:.1f} (text={text_based_y:.1f}, pos={position_based_y:.1f})")
                    
            except Exception as e:
                debug_print(f"[DEBUG] ヘッダー/フッター領域計算エラー: {e}")
        
        # ページ個別でヘッダー検出が失敗した場合、doc-wideの値をフォールバックとして使用
        if header_y_max is None and self._doc_header_y_max is not None:
            header_y_max = self._doc_header_y_max
            debug_print(f"[DEBUG] page={page_num+1}: doc-wideヘッダー領域をフォールバック適用 y_max={header_y_max:.1f}")
        if footer_y_min is None and self._doc_footer_y_min is not None:
            footer_y_min = self._doc_footer_y_min
            debug_print(f"[DEBUG] page={page_num+1}: doc-wideフッター領域をフォールバック適用 y_min={footer_y_min:.1f}")
        
        # 外れ値ガード: ページ個別のヘッダー/フッター領域がdoc-wide値から大きく外れている場合、
        # doc-wide値にフォールバックする（目次ページなどで異常に大きな領域が検出されるのを防ぐ）
        # 閾値: doc-wide値の3倍を超える場合は外れ値とみなす
        if header_y_max is not None and self._doc_header_y_max is not None:
            if header_y_max > self._doc_header_y_max * 3.0:
                debug_print(f"[DEBUG] page={page_num+1}: ヘッダー領域が外れ値 ({header_y_max:.1f} > {self._doc_header_y_max * 3.0:.1f})、doc-wide値にフォールバック")
                header_y_max = self._doc_header_y_max
        if footer_y_min is not None and self._doc_footer_y_min is not None:
            # フッターは下端からの距離で比較（page_height - footer_y_min）
            page_footer_dist = page_height - footer_y_min
            doc_footer_dist = page_height - self._doc_footer_y_min
            if page_footer_dist > doc_footer_dist * 3.0:
                debug_print(f"[DEBUG] page={page_num+1}: フッター領域が外れ値、doc-wide値にフォールバック")
                footer_y_min = self._doc_footer_y_min
        
        # 罫線ベースの表を検出（図抽出から除外するため）
        line_based_table_bboxes = []
        try:
            tables = page.find_tables()
            if tables.tables:
                for table in tables.tables:
                    bbox = table.bbox
                    rows = table.extract()
                    if rows and len(rows) >= 2 and len(rows[0]) >= 2:
                        line_based_table_bboxes.append(bbox)
                        debug_print(f"[DEBUG] page={page_num+1}: 罫線ベース表検出 bbox=({bbox[0]:.1f}, {bbox[1]:.1f}, {bbox[2]:.1f}, {bbox[3]:.1f})")
        except Exception as e:
            debug_print(f"[DEBUG] find_tables()エラー: {e}")
        
        # 要素タイプ（drawing/image）を保持するリスト
        all_elements = []
        
        try:
            drawings = page.get_drawings()
            for d in drawings:
                rect = d.get("rect")
                if rect:
                    bbox = (rect.x0, rect.y0, rect.x1, rect.y1)
                    area = (bbox[2] - bbox[0]) * (bbox[3] - bbox[1])
                    if area >= 200:
                        all_elements.append({"bbox": bbox, "type": "drawing"})
        except Exception as e:
            debug_print(f"[DEBUG] 描画取得エラー: {e}")
        
        try:
            image_list = page.get_images(full=True)
            for img_info in image_list:
                xref = img_info[0]
                # 同じ画像が複数の場所に配置されている場合があるため、すべてのrectを取得
                for img_rect in page.get_image_rects(xref):
                    bbox = (img_rect.x0, img_rect.y0, img_rect.x1, img_rect.y1)
                    area = (bbox[2] - bbox[0]) * (bbox[3] - bbox[1])
                    if area >= 100:
                        all_elements.append({"bbox": bbox, "type": "image"})
        except Exception as e:
            debug_print(f"[DEBUG] 画像取得エラー: {e}")
        
        # 互換性のためにall_bboxesも作成
        all_bboxes = [e["bbox"] for e in all_elements]
        
        # 要素がない場合は早期リターン
        if len(all_bboxes) == 0:
            return []
        
        # 単独の大きな画像も図として抽出するため、要素が1個の場合は特別処理
        # （クラスタリングは2個以上必要だが、単独の大きな画像は図として有効）
        single_image_candidate = None
        if len(all_bboxes) == 1:
            bbox = all_bboxes[0]
            area = (bbox[2] - bbox[0]) * (bbox[3] - bbox[1])
            # 単独要素は面積が十分大きい場合のみ図として扱う（最小10000ピクセル^2）
            if area >= 10000:
                # ヘッダー/フッター領域内の場合は除外
                in_header = header_y_max is not None and bbox[3] < header_y_max
                in_footer = footer_y_min is not None and bbox[1] > footer_y_min
                if not in_header and not in_footer:
                    # 単独画像を候補として保存（後で通常の処理フローに乗せる）
                    col = "full" if (bbox[2] - bbox[0]) > page_width * 0.5 else ("left" if (bbox[0] + bbox[2]) / 2 < page_width / 2 else "right")
                    single_image_candidate = {
                        "union_bbox": bbox,
                        "raw_union_bbox": bbox,
                        "column": col,
                        "cluster_size": 1,
                        "is_embedded": all_elements[0]["type"] == "image",
                    }
                    debug_print(f"[DEBUG] page={page_num+1}: 単独画像を候補として追加 bbox={bbox}, area={area:.1f}")
                else:
                    if in_header:
                        debug_print(f"[DEBUG] page={page_num+1}: 単独画像がヘッダー領域内のため除外")
                    if in_footer:
                        debug_print(f"[DEBUG] page={page_num+1}: 単独画像がフッター領域内のため除外")
            else:
                debug_print(f"[DEBUG] page={page_num+1}: 単独画像が小さすぎるため除外 area={area:.1f}")
        
        def get_column_for_union_bbox(bbox):
            """クラスタのunion_bboxからカラムを判定（ガター跨ぎを優先）"""
            x0, y0, x1, y1 = bbox
            width = x1 - x0
            center_x = (x0 + x1) / 2
            
            # ガター跨ぎ判定（左端がガター左側、右端がガター右側）
            crosses_gutter = x0 < gutter_x - gutter_margin and x1 > gutter_x + gutter_margin
            
            # 全幅判定
            is_full_width = width > page_width * 0.5
            
            if crosses_gutter or is_full_width:
                return "full"
            elif center_x < gutter_x:
                return "left"
            else:
                return "right"
        
        def get_bbox_column(bbox):
            """個々のbboxのカラムを判定"""
            x0, y0, x1, y1 = bbox
            center_x = (x0 + x1) / 2
            # ガター跨ぎ判定
            crosses = x0 < gutter_x - gutter_margin and x1 > gutter_x + gutter_margin
            if crosses:
                return "full"
            elif center_x < gutter_x:
                return "left"
            else:
                return "right"
        
        # 各bboxのカラムを事前計算
        bbox_columns = [get_bbox_column(b) for b in all_bboxes]
        
        # 図キャプションの位置を検出（クラスタ分離に使用）
        figure_caption_lines = []
        try:
            import re as re_mod
            caption_pattern = re_mod.compile(r'^図\s*\d+[\.\:．：]')
            text_dict = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)
            for block in text_dict.get("blocks", []):
                if block.get("type") != 0:
                    continue
                for line in block.get("lines", []):
                    line_text = ""
                    for span in line.get("spans", []):
                        line_text += span.get("text", "")
                    line_text = line_text.strip()
                    if caption_pattern.match(line_text):
                        line_bbox = line.get("bbox", (0, 0, 0, 0))
                        figure_caption_lines.append({
                            "text": line_text,
                            "bbox": line_bbox,
                            "y_center": (line_bbox[1] + line_bbox[3]) / 2
                        })
        except Exception as e:
            debug_print(f"[DEBUG] 図キャプション検出エラー: {e}")
        
        def has_caption_between(bbox1, bbox2):
            """2つのbbox間に図キャプションがあるかどうかを判定"""
            y1_bottom = bbox1[3]
            y2_top = bbox2[1]
            y1_top = bbox1[1]
            y2_bottom = bbox2[3]
            
            # bbox1が上、bbox2が下の場合
            if y1_bottom < y2_top:
                gap_top = y1_bottom
                gap_bottom = y2_top
            # bbox2が上、bbox1が下の場合
            elif y2_bottom < y1_top:
                gap_top = y2_bottom
                gap_bottom = y1_top
            else:
                # 重なっている場合はキャプションなし
                return False
            
            # キャプションがギャップ内にあるか確認
            for cap in figure_caption_lines:
                cap_y = cap["y_center"]
                if gap_top <= cap_y <= gap_bottom:
                    return True
            return False
        
        def cluster_with_gutter_constraint(bboxes, x_threshold=100.0, y_threshold=40.0):
            """ガター制約付きクラスタリング（左右カラム同士は繋げない）"""
            if not bboxes:
                return []
            
            n = len(bboxes)
            visited = [False] * n
            clusters = []
            
            def boxes_close(idx1, idx2):
                b1, b2 = bboxes[idx1], bboxes[idx2]
                col1, col2 = bbox_columns[idx1], bbox_columns[idx2]
                
                # ガター制約: 左右カラム同士で、どちらもガター跨ぎでない場合は繋げない
                if col1 != col2 and col1 != "full" and col2 != "full":
                    return False
                
                # キャプション制約: 2つのbbox間に図キャプションがある場合は繋げない
                if has_caption_between(b1, b2):
                    debug_print(f"[DEBUG] キャプション制約: bbox1=({b1[1]:.1f}-{b1[3]:.1f}), bbox2=({b2[1]:.1f}-{b2[3]:.1f}) 間にキャプションあり")
                    return False
                
                x0_1, y0_1, x1_1, y1_1 = b1
                x0_2, y0_2, x1_2, y1_2 = b2
                
                x_gap = max(0, max(x0_1, x0_2) - min(x1_1, x1_2))
                y_gap = max(0, max(y0_1, y0_2) - min(y1_1, y1_2))
                
                return x_gap <= x_threshold and y_gap <= y_threshold
            
            for i in range(n):
                if visited[i]:
                    continue
                
                cluster = [i]
                visited[i] = True
                queue = [i]
                
                while queue:
                    current = queue.pop(0)
                    
                    for j in range(n):
                        if visited[j]:
                            continue
                        
                        if boxes_close(current, j):
                            cluster.append(j)
                            visited[j] = True
                            queue.append(j)
                
                clusters.append(cluster)
            
            return clusters
        
        # ガター制約付きクラスタリング
        clusters = cluster_with_gutter_constraint(all_bboxes)
        
        all_figure_candidates = []
        for cluster in clusters:
            if len(cluster) < 2:
                bbox = all_bboxes[cluster[0]]
                area = (bbox[2] - bbox[0]) * (bbox[3] - bbox[1])
                if area < 2000:
                    continue
            
            cluster_bboxes = [all_bboxes[i] for i in cluster]
            cluster_types = [all_elements[i]["type"] for i in cluster]
            
            # raw_union_bbox（パディングなし）を計算
            raw_x0 = min(b[0] for b in cluster_bboxes)
            raw_y0 = min(b[1] for b in cluster_bboxes)
            raw_x1 = max(b[2] for b in cluster_bboxes)
            raw_y1 = max(b[3] for b in cluster_bboxes)
            raw_union_bbox = (raw_x0, raw_y0, raw_x1, raw_y1)
            raw_area = (raw_x1 - raw_x0) * (raw_y1 - raw_y0)
            
            # ヘッダー/フッター領域内のクラスタを除外
            # ヘッダー: クラスタの上端がヘッダー領域内、または大部分がヘッダー領域内
            # フッター: クラスタの下端がフッター領域内、または大部分がフッター領域内
            cluster_height = raw_y1 - raw_y0
            is_in_header = False
            is_in_footer = False
            if header_y_max is not None:
                overlap_with_header = max(0, min(raw_y1, header_y_max) - raw_y0)
                is_in_header = overlap_with_header > cluster_height * 0.5 or raw_y0 < header_y_max * 0.5
            if footer_y_min is not None:
                overlap_with_footer = max(0, raw_y1 - max(raw_y0, footer_y_min))
                is_in_footer = overlap_with_footer > cluster_height * 0.5 or raw_y1 > footer_y_min + (page_height - footer_y_min) * 0.5
            if is_in_header or is_in_footer:
                region = "ヘッダー" if is_in_header else "フッター"
                debug_print(f"[DEBUG] page={page_num+1}: {region}領域内のクラスタを除外 bbox=({raw_x0:.1f}, {raw_y0:.1f}, {raw_x1:.1f}, {raw_y1:.1f})")
                continue
            
            # 埋め込み画像判定: クラスタがimageのみで構成され、小面積の場合
            is_image_only = all(t == "image" for t in cluster_types)
            is_embedded = is_image_only and len(cluster) <= 2 and raw_area < 10000
            
            # パディングを決定（埋め込み画像は0、通常図形は20pt）
            padding = 0.0 if is_embedded else 20.0
            x0 = raw_x0 - padding
            y0 = raw_y0 - padding
            x1 = raw_x1 + padding
            y1 = raw_y1 + padding
            union_bbox = (max(0, x0), max(0, y0), min(page_width, x1), min(page.rect.height, y1))
            
            area = (union_bbox[2] - union_bbox[0]) * (union_bbox[3] - union_bbox[1])
            if area < 1000:
                continue
            
            # クラスタリング後にカラム判定（raw_union_bboxを使用）
            column = get_column_for_union_bbox(raw_union_bbox)
            
            # 画像の数をカウント（左右マージの条件に使用）
            image_count = sum(1 for t in cluster_types if t == "image")
            
            all_figure_candidates.append({
                "union_bbox": union_bbox,
                "raw_union_bbox": raw_union_bbox,
                "cluster_size": len(cluster),
                "column": column,
                "is_embedded": is_embedded,
                "image_count": image_count
            })
        
        # 包含除去フィルタ: 大きいbboxが小さいbboxを包含している場合、小さい方を除去
        def bbox_contains(outer, inner, margin=10.0):
            """outerがinnerを包含しているか判定"""
            return (outer[0] - margin <= inner[0] and 
                    outer[1] - margin <= inner[1] and 
                    outer[2] + margin >= inner[2] and 
                    outer[3] + margin >= inner[3])
        
        filtered_candidates = []
        for i, cand in enumerate(all_figure_candidates):
            is_contained = False
            for j, other in enumerate(all_figure_candidates):
                if i != j:
                    if bbox_contains(other["union_bbox"], cand["union_bbox"]):
                        is_contained = True
                        break
            if not is_contained:
                filtered_candidates.append(cand)
        
        all_figure_candidates = filtered_candidates
        
        # 単独画像候補がある場合は追加（要素が1個の場合の特別処理）
        # ただし、既存の候補と重複している場合は追加しない
        if single_image_candidate is not None:
            single_bbox = single_image_candidate["union_bbox"]
            is_duplicate = False
            for cand in all_figure_candidates:
                cand_bbox = cand["union_bbox"]
                # 重複判定: 一方が他方を包含、または高いIoU
                if bbox_contains(cand_bbox, single_bbox) or bbox_contains(single_bbox, cand_bbox):
                    is_duplicate = True
                    debug_print(f"[DEBUG] page={page_num+1}: 単独画像候補が既存候補と重複のため追加しない")
                    break
            if not is_duplicate:
                all_figure_candidates.append(single_image_candidate)
                debug_print(f"[DEBUG] page={page_num+1}: 単独画像候補をall_figure_candidatesに追加")
        
        # 表領域を検出して図候補から除外
        # 表は罫線（ベクター描画）として認識されるため、図として出力されてしまう問題を防ぐ
        def detect_table_bboxes_from_text(text_lines, page_width):
            """テキスト行の配置パターンから表領域を検出"""
            import re as re_mod
            table_bboxes = []
            gutter = page_width / 2
            
            # キャプションパターン（図X、表Xで始まる行は表検出から除外）
            caption_pattern = re_mod.compile(r'^[図表]\s*\d+')
            
            # 左右カラムごとに処理
            for is_left in [True, False]:
                col_lines = []
                for line in text_lines:
                    # キャプション行は除外
                    if caption_pattern.match(line.get("text", "")):
                        continue
                    line_bbox = line["bbox"]
                    center_x = (line_bbox[0] + line_bbox[2]) / 2
                    if is_left and center_x < gutter:
                        col_lines.append(line)
                    elif not is_left and center_x >= gutter:
                        col_lines.append(line)
                
                if not col_lines:
                    continue
                
                # Y座標でグループ化（同じ行にある要素を検出）
                y_tolerance = 5
                y_groups = {}
                for line in col_lines:
                    y_key = round(line["bbox"][1] / y_tolerance) * y_tolerance
                    if y_key not in y_groups:
                        y_groups[y_key] = []
                    y_groups[y_key].append(line)
                
                # 複数セルがある行を検出（2セル以上で表として認識）
                # ただし、短いテキスト（3文字以下）だけで構成される行は除外（図内ラベルの誤検出防止）
                multi_cell_rows = []
                for y_key in sorted(y_groups.keys()):
                    cells = y_groups[y_key]
                    x_positions = sorted(set(round(c["bbox"][0] / 20) * 20 for c in cells))
                    if len(x_positions) >= 2:
                        # 短いテキストだけで構成される行は除外
                        texts = [c.get("text", "") for c in cells]
                        has_long_text = any(len(t) > 3 for t in texts)
                        if not has_long_text:
                            continue
                        all_bboxes = [c["bbox"] for c in cells]
                        row_bbox = (
                            min(b[0] for b in all_bboxes),
                            min(b[1] for b in all_bboxes),
                            max(b[2] for b in all_bboxes),
                            max(b[3] for b in all_bboxes)
                        )
                        multi_cell_rows.append({"y": y_key, "bbox": row_bbox})
                
                # 連続する複数セル行を表領域としてグループ化（5行以上で表として認識）
                # 2列表（表2など）を検出しつつ、図内ラベルの誤検出を防ぐ
                if len(multi_cell_rows) < 5:
                    continue
                
                current_region = {
                    "y_start": multi_cell_rows[0]["y"],
                    "y_end": multi_cell_rows[0]["y"] + 20,
                    "x0": multi_cell_rows[0]["bbox"][0],
                    "x1": multi_cell_rows[0]["bbox"][2],
                    "rows": [multi_cell_rows[0]]
                }
                
                for i in range(1, len(multi_cell_rows)):
                    row = multi_cell_rows[i]
                    prev_row = multi_cell_rows[i - 1]
                    
                    if row["y"] - prev_row["y"] < 50:
                        current_region["rows"].append(row)
                        current_region["y_end"] = row["y"] + 20
                        current_region["x0"] = min(current_region["x0"], row["bbox"][0])
                        current_region["x1"] = max(current_region["x1"], row["bbox"][2])
                    else:
                        # 連続する行が5行以上で表領域として確定
                        if len(current_region["rows"]) >= 5:
                            table_bboxes.append((
                                current_region["x0"] - 5,
                                current_region["y_start"] - 10,
                                current_region["x1"] + 5,
                                current_region["y_end"] + 5
                            ))
                        current_region = {
                            "y_start": row["y"],
                            "y_end": row["y"] + 20,
                            "x0": row["bbox"][0],
                            "x1": row["bbox"][2],
                            "rows": [row]
                        }
                
                # 連続する行が5行以上で表領域として確定
                if len(current_region["rows"]) >= 5:
                    table_bboxes.append((
                        current_region["x0"] - 5,
                        current_region["y_start"] - 10,
                        current_region["x1"] + 5,
                        current_region["y_end"] + 5
                    ))
            
            return table_bboxes
        
        def bbox_overlap_ratio(bbox1, bbox2):
            """2つのbboxの重なり率を計算（小さい方の面積に対する比率）"""
            x0 = max(bbox1[0], bbox2[0])
            y0 = max(bbox1[1], bbox2[1])
            x1 = min(bbox1[2], bbox2[2])
            y1 = min(bbox1[3], bbox2[3])
            
            if x0 >= x1 or y0 >= y1:
                return 0.0
            
            inter_area = (x1 - x0) * (y1 - y0)
            area1 = (bbox1[2] - bbox1[0]) * (bbox1[3] - bbox1[1])
            area2 = (bbox2[2] - bbox2[0]) * (bbox2[3] - bbox2[1])
            min_area = min(area1, area2)
            
            if min_area <= 0:
                return 0.0
            return inter_area / min_area
        
        # テキスト行を取得（表検出用）
        text_dict = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)
        page_text_lines = []
        for block in text_dict.get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                line_bbox = line.get("bbox", (0, 0, 0, 0))
                line_text = ""
                for span in line.get("spans", []):
                    line_text += span.get("text", "")
                page_text_lines.append({"text": line_text.strip(), "bbox": line_bbox})
        
        # 段組みレイアウトを検出
        column_count = self._detect_column_layout(text_dict)
        
        # 表領域を検出
        table_bboxes = detect_table_bboxes_from_text(page_text_lines, page_width)
        
        if table_bboxes:
            debug_print(f"[DEBUG] page={page_num+1}: 表領域を{len(table_bboxes)}個検出")
            for i, tb in enumerate(table_bboxes):
                debug_print(f"[DEBUG]   表{i+1}: y={tb[1]:.1f}-{tb[3]:.1f}, x={tb[0]:.1f}-{tb[2]:.1f}")
        
        # 表領域がMarkdownテーブルとして出力可能かを判定
        def can_output_as_markdown_table(table_bbox, text_lines, is_two_column_layout):
            """表領域がMarkdownテーブルとして出力可能かを判定
            
            条件:
            - 3行以上の行がある
            - 80%以上の行が同じ列数
            - 段組みページでも、列数の一貫性がある表はMarkdownテーブルとして出力可能
            """
            # 表領域内のテキスト行を取得
            table_lines = []
            for line in text_lines:
                line_bbox = line["bbox"]
                line_center_y = (line_bbox[1] + line_bbox[3]) / 2
                line_center_x = (line_bbox[0] + line_bbox[2]) / 2
                if (table_bbox[1] <= line_center_y <= table_bbox[3] and
                    table_bbox[0] <= line_center_x <= table_bbox[2]):
                    table_lines.append(line)
            
            if len(table_lines) < 3:
                debug_print(f"[DEBUG] 表行数不足: {len(table_lines)} < 3")
                return False
            
            # Y座標でグループ化
            y_groups = {}
            for line in table_lines:
                y_key = round(line["bbox"][1] / 5) * 5
                if y_key not in y_groups:
                    y_groups[y_key] = []
                y_groups[y_key].append(line)
            
            # 各行の列数をカウント
            col_counts = []
            for y_key in sorted(y_groups.keys()):
                cells = y_groups[y_key]
                x_positions = sorted(set(round(c["bbox"][0] / 20) * 20 for c in cells))
                if len(x_positions) >= 2:
                    col_counts.append(len(x_positions))
            
            if len(col_counts) < 3:
                debug_print(f"[DEBUG] 複数列行数不足: {len(col_counts)} < 3")
                return False
            
            # 列数の一貫性をチェック（80%以上が同じ列数）
            most_common = max(set(col_counts), key=col_counts.count)
            consistent = sum(1 for c in col_counts if c == most_common)
            consistency_ratio = consistent / len(col_counts)
            
            debug_print(f"[DEBUG] 表の列数一貫性: {consistent}/{len(col_counts)} = {consistency_ratio:.2f}, 最頻列数={most_common}")
            
            if consistency_ratio < 0.8:
                return False
            
            # 3列以上の表のみMarkdownテーブルとして出力可能
            # _detect_table_regionsは3セル以上で表として認識するため
            # 2列の表（表2など）はMarkdownテーブルとして出力されない
            if most_common < 3:
                debug_print(f"[DEBUG] 2列表のためMarkdownテーブルとして出力不可")
                return False
            
            return True
        
        # 表領域と重なる図候補を分類（Markdown表可能なら除外、不可能なら表画像として出力）
        debug_print(f"[DEBUG] page={page_num+1}: 表フィルタ前の候補数={len(all_figure_candidates)}")
        table_filtered = []
        table_image_candidates = []  # 表画像として出力する候補
        
        for cand in all_figure_candidates:
            is_table = False
            cand_bbox = cand["union_bbox"]
            matched_table_bbox = None
            
            for table_bbox in table_bboxes:
                overlap = bbox_overlap_ratio(cand_bbox, table_bbox)
                if overlap >= 0.7:
                    matched_table_bbox = table_bbox
                    is_table = True
                    break
            
            if is_table and matched_table_bbox:
                # Markdownテーブルとして出力可能かを判定
                is_two_col = (column_count >= 2)
                if can_output_as_markdown_table(matched_table_bbox, page_text_lines, is_two_col):
                    debug_print(f"[DEBUG] page={page_num+1}: 図候補をMarkdown表として除外: overlap={overlap:.2f}")
                else:
                    # Markdownテーブルとして出力不可能な場合は表画像として出力
                    debug_print(f"[DEBUG] page={page_num+1}: 図候補を表画像として出力: overlap={overlap:.2f}")
                    # 表画像用のclip_bboxを計算
                    # 上端: キャプションを除外（union_bboxの上端を基準に探索）
                    # 下端: union_bbox（描画要素のbbox）の下端を使用
                    union = cand_bbox
                    
                    # 上端: union_bboxの上端を基準にキャプション行を検出
                    union_top = union[1]
                    clip_y0 = union_top - 2  # デフォルトはunion_bboxの上端
                    
                    # キャプション行を検出（「表N」パターン）
                    # union_bboxの上端付近にあるキャプション行を探す
                    import re
                    for line in page_text_lines:
                        line_bbox = line["bbox"]
                        line_text = line["text"]
                        # union_bboxの上端付近（上端から20pt以内）にあるキャプション行を探す
                        if line_bbox[1] >= union_top - 5 and line_bbox[3] <= union_top + 20:
                            if re.match(r'^表\s*\d', line_text):
                                # キャプション行の下端を上端にする
                                clip_y0 = line_bbox[3] + 2
                                debug_print(f"[DEBUG] 表画像: キャプション検出 '{line_text[:20]}' y={line_bbox[3]:.1f}")
                                break
                    
                    # 下端: union_bboxの下端を使用（表の罫線全体を含める）
                    # ただし、下端付近にキャプションがある場合は除外
                    clip_x0 = union[0] - 2
                    clip_x1 = union[2] + 2
                    clip_y1 = union[3] + 5  # デフォルトは少しのマージン
                    
                    # 下端付近のキャプション行を検出（「図N」「表N」パターン）
                    union_bottom = union[3]
                    for line in page_text_lines:
                        line_bbox = line["bbox"]
                        line_text = line["text"]
                        # union_bboxの下端付近（下端から20pt以内）にあるキャプション行を探す
                        if line_bbox[1] >= union_bottom - 20 and line_bbox[1] <= union_bottom + 30:
                            if re.match(r'^(図|表)\s*\d', line_text):
                                # キャプション行の上端を下端にする
                                clip_y1 = line_bbox[1] - 2
                                debug_print(f"[DEBUG] 表画像: 下キャプション検出 '{line_text[:20]}' y={line_bbox[1]:.1f}")
                                break
                    
                    debug_print(f"[DEBUG] 表画像: clip_bbox=({clip_x0:.1f}, {clip_y0:.1f}, {clip_x1:.1f}, {clip_y1:.1f})")
                    
                    table_clip_bbox = (clip_x0, clip_y0, clip_x1, clip_y1)
                    cand["clip_bbox"] = table_clip_bbox
                    cand["is_table_image"] = True
                    table_image_candidates.append(cand)
            else:
                table_filtered.append(cand)
        
        debug_print(f"[DEBUG] page={page_num+1}: 表フィルタ後の候補数={len(table_filtered)}, 表画像候補数={len(table_image_candidates)}")
        all_figure_candidates = table_filtered + table_image_candidates
        
        # 第2段: 同一カラム内のクラスタを安全にマージ
        # 本文バリアがない場合のみマージする
        def is_body_text_line(text, line_width, col_width):
            """本文らしい行かどうかを判定"""
            if len(text) < 15:
                return False
            if "。" in text:
                return True
            particles = ["が", "を", "に", "で", "は", "の", "と", "も", "や"]
            if any(p in text for p in particles) and line_width > col_width * 0.5:
                return True
            return False
        
        def has_body_barrier(bbox1, bbox2, text_lines, col_width):
            """2つのbbox間に本文バリアがあるか判定"""
            import re as re_mod
            y_min = min(bbox1[3], bbox2[3])
            y_max = max(bbox1[1], bbox2[1])
            x_overlap_start = max(bbox1[0], bbox2[0])
            x_overlap_end = min(bbox1[2], bbox2[2])
            
            # 図キャプションパターン（「図X.X:」形式）
            caption_pattern = re_mod.compile(r'^図\s*\d+[\.\:．：]')
            
            # bbox1とbbox2の間のギャップを計算
            if bbox1[3] < bbox2[1]:
                gap_top = bbox1[3]
                gap_bottom = bbox2[1]
            elif bbox2[3] < bbox1[1]:
                gap_top = bbox2[3]
                gap_bottom = bbox1[1]
            else:
                gap_top = y_min
                gap_bottom = y_max
            
            for line in text_lines:
                line_bbox = line["bbox"]
                line_center_y = (line_bbox[1] + line_bbox[3]) / 2
                line_center_x = (line_bbox[0] + line_bbox[2]) / 2
                line_width = line_bbox[2] - line_bbox[0]
                
                # ギャップ内またはギャップ近傍（±30pt）にキャプションがあるか確認
                if caption_pattern.match(line["text"]):
                    if gap_top - 30 <= line_center_y <= gap_bottom + 30:
                        if x_overlap_start - 20 < line_center_x < x_overlap_end + 20:
                            debug_print(f"[DEBUG] 図キャプションバリア検出: '{line['text'][:30]}...' y={line_center_y:.1f}")
                            return True
                
                # 本文バリアの判定（従来のロジック）
                if y_min < line_center_y < y_max:
                    if x_overlap_start - 20 < line_center_x < x_overlap_end + 20:
                        if is_body_text_line(line["text"], line_width, col_width):
                            return True
            return False
        
        def get_x_overlap_ratio(bbox1, bbox2):
            """x方向の重なり率を計算"""
            x_overlap = max(0, min(bbox1[2], bbox2[2]) - max(bbox1[0], bbox2[0]))
            min_width = min(bbox1[2] - bbox1[0], bbox2[2] - bbox2[0])
            if min_width <= 0:
                return 0
            return x_overlap / min_width
        
        # page_text_linesは表検出で既に取得済み
        col_width = page_width / 2
        
        # 同一カラム内でマージ可能なクラスタをマージ
        debug_print(f"[DEBUG] クラスタ再マージ開始: {len(all_figure_candidates)}個のクラスタ")
        merged = True
        merge_iteration = 0
        while merged:
            merge_iteration += 1
            merged = False
            new_candidates = []
            used = set()
            
            for i, cand1 in enumerate(all_figure_candidates):
                if i in used:
                    continue
                
                best_merge = None
                best_y_gap = float('inf')
                
                for j, cand2 in enumerate(all_figure_candidates):
                    if i >= j or j in used:
                        continue
                    
                    # 同一カラムのみマージ
                    if cand1["column"] != cand2["column"]:
                        continue
                    
                    bbox1 = cand1["union_bbox"]
                    bbox2 = cand2["union_bbox"]
                    
                    # x方向の重なりが十分あるか
                    x_overlap_ratio = get_x_overlap_ratio(bbox1, bbox2)
                    if x_overlap_ratio < 0.3:
                        debug_print(f"[DEBUG] クラスタ{i}と{j}: x重なり不足 {x_overlap_ratio:.2f}")
                        continue
                    
                    # y方向のギャップを計算
                    y_gap = max(0, max(bbox1[1], bbox2[1]) - min(bbox1[3], bbox2[3]))
                    
                    # y_gapが80pt以内で、本文バリアがない場合のみマージ
                    if y_gap <= 80:
                        debug_print(f"[DEBUG] クラスタ{i}と{j}: bbox1=({bbox1[1]:.1f}-{bbox1[3]:.1f}), bbox2=({bbox2[1]:.1f}-{bbox2[3]:.1f}), y_gap={y_gap:.1f}")
                        if not has_body_barrier(bbox1, bbox2, page_text_lines, col_width):
                            debug_print(f"[DEBUG] クラスタ{i}と{j}: マージ候補 y_gap={y_gap:.1f}")
                            if y_gap < best_y_gap:
                                best_y_gap = y_gap
                                best_merge = j
                        else:
                            debug_print(f"[DEBUG] クラスタ{i}と{j}: 本文バリアあり")
                    else:
                        debug_print(f"[DEBUG] クラスタ{i}と{j}: y_gap超過 {y_gap:.1f}")
                
                if best_merge is not None:
                    cand2 = all_figure_candidates[best_merge]
                    bbox1 = cand1["union_bbox"]
                    bbox2 = cand2["union_bbox"]
                    
                    # マージしたbboxを作成
                    merged_bbox = (
                        min(bbox1[0], bbox2[0]),
                        min(bbox1[1], bbox2[1]),
                        max(bbox1[2], bbox2[2]),
                        max(bbox1[3], bbox2[3])
                    )
                    
                    debug_print(f"[DEBUG] クラスタ{i}と{best_merge}をマージ")
                    new_candidates.append({
                        "union_bbox": merged_bbox,
                        "cluster_size": cand1["cluster_size"] + cand2["cluster_size"],
                        "column": cand1["column"],
                        "image_count": cand1.get("image_count", 0) + cand2.get("image_count", 0)
                    })
                    used.add(i)
                    used.add(best_merge)
                    merged = True
                else:
                    new_candidates.append(cand1)
                    used.add(i)
            
            all_figure_candidates = new_candidates
        
        debug_print(f"[DEBUG] クラスタ再マージ完了: {len(all_figure_candidates)}個のクラスタ")
        
        # 左右クラスタのマージ処理
        # Y方向の重なりが大きく、ガター近傍で隣接し、本文バリアがない場合にマージ
        def get_y_overlap_ratio(bbox1, bbox2):
            """Y方向の重なり率を計算"""
            y_overlap = max(0, min(bbox1[3], bbox2[3]) - max(bbox1[1], bbox2[1]))
            min_height = min(bbox1[3] - bbox1[1], bbox2[3] - bbox2[1])
            if min_height <= 0:
                return 0
            return y_overlap / min_height
        
        debug_print(f"[DEBUG] 左右クラスタマージ開始: {len(all_figure_candidates)}個のクラスタ")
        lr_merged = True
        while lr_merged:
            lr_merged = False
            new_candidates = []
            used = set()
            
            # 左右クラスタのペアを探す
            left_candidates = [(i, c) for i, c in enumerate(all_figure_candidates) 
                               if c["column"] == "left" and i not in used]
            right_candidates = [(i, c) for i, c in enumerate(all_figure_candidates) 
                                if c["column"] == "right" and i not in used]
            
            for left_idx, left_cand in left_candidates:
                if left_idx in used:
                    continue
                
                best_right_idx = None
                best_y_overlap = 0
                
                for right_idx, right_cand in right_candidates:
                    if right_idx in used:
                        continue
                    
                    # 両クラスタに画像が含まれている場合のみマージを許可
                    # （段組ドキュメントで左右に別々の図がある場合の誤マージを防止）
                    left_image_count = left_cand.get("image_count", 0)
                    right_image_count = right_cand.get("image_count", 0)
                    if left_image_count == 0 or right_image_count == 0:
                        debug_print(f"[DEBUG] 左右マージ候補{left_idx},{right_idx}: 画像なし (left={left_image_count}, right={right_image_count})")
                        continue
                    
                    left_bbox = left_cand["union_bbox"]
                    right_bbox = right_cand["union_bbox"]
                    
                    # Y方向の重なり率を計算（90%以上必要）
                    y_overlap_ratio = get_y_overlap_ratio(left_bbox, right_bbox)
                    if y_overlap_ratio < 0.9:
                        debug_print(f"[DEBUG] 左右マージ候補{left_idx},{right_idx}: Y重なり不足 {y_overlap_ratio:.2f}")
                        continue
                    
                    # 左クラスタの右端と右クラスタの左端がガター近傍にあるか
                    left_right_edge = left_bbox[2]
                    right_left_edge = right_bbox[0]
                    x_gap = right_left_edge - left_right_edge
                    
                    # ガター近傍（ガターから±50pt以内）にあるか
                    if not (gutter_x - 50 < left_right_edge < gutter_x + 50 and
                            gutter_x - 50 < right_left_edge < gutter_x + 50):
                        debug_print(f"[DEBUG] 左右マージ候補{left_idx},{right_idx}: ガター近傍でない")
                        continue
                    
                    # X方向のギャップが小さいか（100pt以内）
                    if x_gap > 100:
                        debug_print(f"[DEBUG] 左右マージ候補{left_idx},{right_idx}: X gap超過 {x_gap:.1f}")
                        continue
                    
                    # 本文バリアがないか
                    if has_body_barrier(left_bbox, right_bbox, page_text_lines, col_width):
                        debug_print(f"[DEBUG] 左右マージ候補{left_idx},{right_idx}: 本文バリアあり")
                        continue
                    
                    # 最もY重なりが大きいペアを選択
                    if y_overlap_ratio > best_y_overlap:
                        best_y_overlap = y_overlap_ratio
                        best_right_idx = right_idx
                
                if best_right_idx is not None:
                    right_cand = all_figure_candidates[best_right_idx]
                    left_bbox = left_cand["union_bbox"]
                    right_bbox = right_cand["union_bbox"]
                    
                    # マージしたbboxを作成
                    merged_bbox = (
                        min(left_bbox[0], right_bbox[0]),
                        min(left_bbox[1], right_bbox[1]),
                        max(left_bbox[2], right_bbox[2]),
                        max(left_bbox[3], right_bbox[3])
                    )
                    
                    debug_print(f"[DEBUG] 左右クラスタ{left_idx}と{best_right_idx}をマージ (Y重なり={best_y_overlap:.2f})")
                    new_candidates.append({
                        "union_bbox": merged_bbox,
                        "cluster_size": left_cand["cluster_size"] + right_cand["cluster_size"],
                        "column": "full",
                        "image_count": left_cand.get("image_count", 0) + right_cand.get("image_count", 0)
                    })
                    used.add(left_idx)
                    used.add(best_right_idx)
                    lr_merged = True
                else:
                    new_candidates.append(left_cand)
                    used.add(left_idx)
            
            # マージされなかった右クラスタを追加
            for right_idx, right_cand in right_candidates:
                if right_idx not in used:
                    new_candidates.append(right_cand)
                    used.add(right_idx)
            
            # fullクラスタを追加
            for i, cand in enumerate(all_figure_candidates):
                if cand["column"] == "full" and i not in used:
                    new_candidates.append(cand)
            
            all_figure_candidates = new_candidates
        
        debug_print(f"[DEBUG] 左右クラスタマージ完了: {len(all_figure_candidates)}個のクラスタ")
        
        if not all_figure_candidates:
            return []
        
        debug_print(f"[DEBUG] ページ {page_num + 1}: {len(all_bboxes)}個の要素を{len(all_figure_candidates)}個の図にグループ化")
        
        # 切り出し範囲制御: キャプションの上、本文の下でトリム
        import re as re_module
        
        def find_caption_below(graphics_bbox, text_lines):
            """図の下部付近にあるキャプション行を探す
            
            graphics_bboxの内側にあるキャプションも検出する（境界トリム用）
            """
            caption_pattern = re_module.compile(r'^(図|表)\s*\d+')
            best_caption = None
            best_y = float('inf')
            
            # graphics_bboxの下半分以降にあるキャプションを探す
            search_y_start = (graphics_bbox[1] + graphics_bbox[3]) / 2
            
            for line in text_lines:
                line_bbox = line["bbox"]
                line_text = line["text"].strip()
                
                # キャプションパターンにマッチするか
                if not caption_pattern.match(line_text):
                    continue
                
                # 図の下半分以降にあるか
                if line_bbox[1] < search_y_start:
                    continue
                
                # x方向で図と重なりがあるか
                x_overlap = max(0, min(graphics_bbox[2], line_bbox[2]) - max(graphics_bbox[0], line_bbox[0]))
                if x_overlap < 20:
                    continue
                
                # 最も近いキャプションを選択
                if line_bbox[1] < best_y:
                    best_y = line_bbox[1]
                    best_caption = line
            
            return best_caption
        
        def find_body_text_above(graphics_bbox, text_lines, col_width):
            """図の上部付近にある本文行を探す
            
            graphics_bboxの内側にある本文も検出する（境界トリム用）
            """
            best_body = None
            best_y = 0
            
            # graphics_bboxの上半分以前にある本文を探す
            search_y_end = (graphics_bbox[1] + graphics_bbox[3]) / 2
            
            for line in text_lines:
                line_bbox = line["bbox"]
                line_text = line["text"].strip()
                line_width = line_bbox[2] - line_bbox[0]
                
                # 図の上半分以前にあるか
                if line_bbox[3] > search_y_end:
                    continue
                
                # x方向で図と重なりがあるか
                x_overlap = max(0, min(graphics_bbox[2], line_bbox[2]) - max(graphics_bbox[0], line_bbox[0]))
                if x_overlap < 20:
                    continue
                
                # 本文らしい行か（短いラベルは除外）
                if is_body_text_line(line_text, line_width, col_width):
                    if line_bbox[3] > best_y:
                        best_y = line_bbox[3]
                        best_body = line
            
            return best_body
        
        def compute_clip_bbox(graphics_bbox, text_lines, col_width, page_height, column, is_embedded_image=False,
                               header_y_max_val=None, footer_y_min_val=None):
            """graphics_bboxからclip_bboxを計算（トリム処理）
            
            本文の下〜キャプションの上でトリムする。
            graphics_bboxを侵食してでも、境界を正しく設定する。
            カラム境界でクランプして他段のテキストが混入しないようにする。
            埋め込み画像の場合は最小限のpadding（1pt）を使用。
            ヘッダー/フッター領域でクリップして、ヘッダー/フッター部分が図に含まれないようにする。
            """
            # 埋め込み画像の場合は最小限のpadding（1pt）を使用
            padding = 1.0 if is_embedded_image else 20.0
            clip_x0 = max(0, graphics_bbox[0] - padding)
            clip_y0 = max(0, graphics_bbox[1] - padding)
            clip_x1 = min(page_width, graphics_bbox[2] + padding)
            clip_y1 = min(page_height, graphics_bbox[3] + padding)
            
            # ヘッダー/フッター領域でクリップ
            if header_y_max_val is not None and clip_y0 < header_y_max_val:
                clip_y0 = header_y_max_val
                debug_print(f"[DEBUG] ヘッダー領域クリップ: clip_y0を{clip_y0:.1f}に設定")
            if footer_y_min_val is not None and clip_y1 > footer_y_min_val:
                clip_y1 = footer_y_min_val
                debug_print(f"[DEBUG] フッター領域クリップ: clip_y1を{clip_y1:.1f}に設定")
            
            # カラム境界でクランプ（他段のテキスト混入防止）
            if column == "left":
                clip_x1 = min(clip_x1, gutter_x - 5)
                debug_print(f"[DEBUG] 左カラム: clip_x1を{clip_x1:.1f}にクランプ (gutter_x={gutter_x:.1f})")
            elif column == "right":
                old_clip_x0 = clip_x0
                clip_x0 = max(clip_x0, gutter_x + 5)
                debug_print(f"[DEBUG] 右カラム: clip_x0を{old_clip_x0:.1f}→{clip_x0:.1f}にクランプ (gutter_x={gutter_x:.1f})")
            
            # 埋め込み画像の場合は上下トリムをスキップ
            if is_embedded_image:
                debug_print(f"[DEBUG] 埋め込み画像: 上下トリムをスキップ")
                return (clip_x0, clip_y0, clip_x1, clip_y1)
            
            # 下側: キャプションの上までトリム（常に適用）
            caption = find_caption_below(graphics_bbox, text_lines)
            if caption:
                caption_y0 = caption["bbox"][1]
                new_clip_y1 = caption_y0 - 5.0
                clip_y1 = min(clip_y1, new_clip_y1)
                debug_print(f"[DEBUG] キャプション検出: clip_y1を{clip_y1:.1f}にトリム")
            
            # 上側: 本文の下までトリム（常に適用）
            body_above = find_body_text_above(graphics_bbox, text_lines, col_width)
            if body_above:
                body_y1 = body_above["bbox"][3]
                new_clip_y0 = body_y1 + 5.0
                clip_y0 = max(clip_y0, new_clip_y0)
                debug_print(f"[DEBUG] 本文検出: clip_y0を{clip_y0:.1f}にトリム")
            
            # 健全性チェック: clip_y0 < clip_y1を保証（最小高さ50pt）
            if clip_y1 - clip_y0 < 50:
                center_y = (clip_y0 + clip_y1) / 2
                clip_y0 = center_y - 25
                clip_y1 = center_y + 25
                debug_print(f"[DEBUG] 最小高さ確保: clip_y0={clip_y0:.1f}, clip_y1={clip_y1:.1f}")
            
            return (clip_x0, clip_y0, clip_x1, clip_y1)
        
        for fig_info in all_figure_candidates:
            try:
                # 埋め込み画像かどうかを判定（クラスタ構築時に判定済み）
                is_embedded_image = fig_info.get("is_embedded", False)
                # 表画像かどうかを判定（表フィルタ時に判定済み）
                is_table_image = fig_info.get("is_table_image", False)
                
                # 埋め込み画像の場合はraw_union_bbox（パディングなし）を使用
                if is_embedded_image:
                    graphics_bbox = fig_info.get("raw_union_bbox", fig_info["union_bbox"])
                else:
                    graphics_bbox = fig_info["union_bbox"]
                
                column = fig_info["column"]
                union_bbox = graphics_bbox
                
                # 表画像の場合は既に設定されたclip_bboxを使用
                if is_table_image and "clip_bbox" in fig_info:
                    clip_bbox = fig_info["clip_bbox"]
                    debug_print(f"[DEBUG] 表画像: 既存のclip_bboxを使用")
                else:
                    # clip_bboxを計算（トリム処理、ヘッダー/フッター領域でクリップ）
                    clip_bbox = compute_clip_bbox(
                        graphics_bbox, page_text_lines, col_width, page.rect.height, column, 
                        is_embedded_image, header_y_max, footer_y_min
                    )
                
                self.image_counter += 1
                image_filename = f"{self.base_name}_fig_{page_num + 1:03d}_{self.image_counter:03d}"
                
                # デバッグ: 候補のbbox情報を出力
                debug_print(f"[DEBUG] 図候補出力: page={page_num+1}, union_bbox=({graphics_bbox[0]:.1f}, {graphics_bbox[1]:.1f}, {graphics_bbox[2]:.1f}, {graphics_bbox[3]:.1f}), clip_bbox=({clip_bbox[0]:.1f}, {clip_bbox[1]:.1f}, {clip_bbox[2]:.1f}, {clip_bbox[3]:.1f}), column={column}")
                
                clip_rect = fitz.Rect(clip_bbox)
                matrix = fitz.Matrix(2.0, 2.0)
                pix = page.get_pixmap(matrix=matrix, clip=clip_rect)
                
                if self.output_format == 'svg':
                    image_path = os.path.join(self.images_dir, f"{image_filename}.svg")
                    temp_png = os.path.join(self.images_dir, f"temp_fig_{self.image_counter}.png")
                    pix.save(temp_png)
                    self._convert_png_to_svg(temp_png, image_path)
                    if os.path.exists(temp_png):
                        os.remove(temp_png)
                else:
                    image_path = os.path.join(self.images_dir, f"{image_filename}.png")
                    pix.save(image_path)
                
                # 図形内テキスト抽出にはclip_bboxを使用（本文の巻き込みを防ぐ）
                # clip_bboxは本文の下〜キャプションの上でトリムされている
                # 罫線ベースの表領域内のテキストは除外する
                figure_texts, expanded_bbox = self._extract_text_in_bbox(
                    page, clip_bbox, expand_for_labels=True, column=column, gutter_x=gutter_x,
                    exclude_table_bboxes=line_based_table_bboxes
                )
                
                figures.append({
                    "path": image_path,
                    "filename": os.path.basename(image_path),
                    "bbox": expanded_bbox,
                    "y_position": union_bbox[1],
                    "texts": figure_texts,
                    "column": column
                })
                
                debug_print(f"[DEBUG] 図を抽出: {image_path} ({fig_info['cluster_size']}要素, {len(figure_texts)}テキスト, {column})")
                
            except Exception as e:
                debug_print(f"[DEBUG] 図抽出エラー: {e}")
                continue
        
        figures.sort(key=lambda x: x["y_position"])
        return figures

    def _extract_vector_figures(
        self, page, page_num: int, header_footer_patterns: Set[str] = None
    ) -> List[Dict[str, Any]]:
        """ベクタ描画（図）を抽出（統合版を使用）
        
        Args:
            page: PyMuPDFのページオブジェクト
            page_num: ページ番号
            header_footer_patterns: ヘッダ・フッタパターンのセット
            
        Returns:
            抽出された図の情報リスト
        """
        return self._extract_all_figures(page, page_num, header_footer_patterns)

    def _extract_text_in_bbox(
        self, page, bbox: Tuple[float, float, float, float],
        expand_for_labels: bool = True,
        column: str = None,
        gutter_x: float = None,
        exclude_table_bboxes: List[Tuple[float, float, float, float]] = None
    ) -> Tuple[List[str], Tuple[float, float, float, float]]:
        """指定されたbbox内のテキストを抽出
        
        図のラベルテキストを含めるため、bboxを近傍の短いテキストで拡張する。
        カラム境界を考慮して、隣のカラムのテキストを取り込まないようにする。
        罫線ベースの表領域内のテキストは除外する。
        
        Args:
            page: PyMuPDFのページオブジェクト
            bbox: バウンディングボックス (x0, y0, x1, y1)
            expand_for_labels: ラベルテキストを含めるためにbboxを拡張するか
            column: 図のカラム ("left", "right", "full")
            gutter_x: カラム境界のX座標
            exclude_table_bboxes: 除外する罫線ベースの表領域のリスト
            
        Returns:
            (抽出されたテキストのリスト, 拡張後のbbox)
        """
        import re
        texts = []
        expanded_bbox = bbox
        
        if gutter_x is None:
            gutter_x = page.rect.width / 2
        
        if exclude_table_bboxes is None:
            exclude_table_bboxes = []
        
        def is_in_table_area(line_bbox):
            """行が罫線ベースの表領域内にあるかどうかを判定"""
            line_center_x = (line_bbox[0] + line_bbox[2]) / 2
            line_center_y = (line_bbox[1] + line_bbox[3]) / 2
            for table_bbox in exclude_table_bboxes:
                if (table_bbox[0] - 5 <= line_center_x <= table_bbox[2] + 5 and
                    table_bbox[1] - 5 <= line_center_y <= table_bbox[3] + 5):
                    return True
            return False
        
        try:
            text_dict = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)
            
            label_margin = 50.0
            
            def is_in_same_column(line_center_x):
                if column == "full":
                    return True
                elif column == "left":
                    return line_center_x < gutter_x
                elif column == "right":
                    return line_center_x >= gutter_x
                return True
            
            if expand_for_labels:
                for block in text_dict.get("blocks", []):
                    if block.get("type") != 0:
                        continue
                    
                    for line in block.get("lines", []):
                        line_bbox = line.get("bbox", (0, 0, 0, 0))
                        line_center_x = (line_bbox[0] + line_bbox[2]) / 2
                        line_center_y = (line_bbox[1] + line_bbox[3]) / 2
                        
                        if not is_in_same_column(line_center_x):
                            continue
                        
                        near_bbox = (
                            bbox[0] - label_margin <= line_center_x <= bbox[2] + label_margin and
                            bbox[1] - label_margin <= line_center_y <= bbox[3] + label_margin
                        )
                        
                        if not near_bbox:
                            continue
                        
                        line_text = ""
                        for span in line.get("spans", []):
                            line_text += span.get("text", "")
                        line_text = line_text.strip()
                        
                        if re.match(r'^図\d+', line_text) or re.match(r'^表\d+', line_text):
                            continue
                        
                        is_label = (
                            line_text and
                            len(line_text) <= 30 and
                            "。" not in line_text and
                            not any(c in line_text for c in ["、", "が", "を", "に", "で", "は"])
                        )
                        
                        if is_label:
                            new_x0 = min(expanded_bbox[0], line_bbox[0])
                            new_x1 = max(expanded_bbox[2], line_bbox[2])
                            
                            if column == "left" and new_x1 > gutter_x - 10:
                                continue
                            if column == "right" and new_x0 < gutter_x + 10:
                                continue
                            
                            expanded_bbox = (
                                new_x0,
                                min(expanded_bbox[1], line_bbox[1]),
                                new_x1,
                                max(expanded_bbox[3], line_bbox[3])
                            )
            
            # テキストとY座標を一緒に収集（後でソートするため）
            text_with_positions = []
            
            for block in text_dict.get("blocks", []):
                if block.get("type") != 0:
                    continue
                
                for line in block.get("lines", []):
                    line_bbox = line.get("bbox", (0, 0, 0, 0))
                    
                    line_center_x = (line_bbox[0] + line_bbox[2]) / 2
                    line_center_y = (line_bbox[1] + line_bbox[3]) / 2
                    
                    if not is_in_same_column(line_center_x):
                        continue
                    
                    # 罫線ベースの表領域内のテキストは除外
                    if is_in_table_area(line_bbox):
                        continue
                    
                    if (expanded_bbox[0] <= line_center_x <= expanded_bbox[2] and
                        expanded_bbox[1] <= line_center_y <= expanded_bbox[3]):
                        
                        line_text = ""
                        for span in line.get("spans", []):
                            line_text += span.get("text", "")
                        
                        line_text_stripped = line_text.strip()
                        
                        # キャプションパターンをフィルタリング（図形内テキストから除外）
                        if re.match(r'^図\s*\d+', line_text_stripped) or re.match(r'^表\s*\d+', line_text_stripped):
                            continue
                        
                        if line_text_stripped:
                            # Y座標（上端）とX座標（左端）を記録
                            text_with_positions.append((line_bbox[1], line_bbox[0], line_text_stripped))
            
            # Y座標（上から下）、同じY座標ならX座標（左から右）でソート
            text_with_positions.sort(key=lambda item: (item[0], item[1]))
            texts = [item[2] for item in text_with_positions]
        
        except Exception as e:
            debug_print(f"[DEBUG] bbox内テキスト抽出エラー: {e}")
        
        return texts, expanded_bbox

    def _format_figure_texts_as_details(self, texts: List[str]) -> str:
        """図内テキストを<details>タグ形式に整形
        
        x2md_graphics.pyと同様の形式で出力する。
        
        Args:
            texts: 図内テキストのリスト
            
        Returns:
            整形されたテキスト
        """
        if not texts:
            return ""
        
        quoted_texts = [f'"{t}"' for t in texts]
        texts_line = ', '.join(quoted_texts)
        
        lines = []
        lines.append("<details>")
        lines.append("<summary>図形内テキスト</summary>")
        lines.append("")
        lines.append(texts_line)
        lines.append("")
        lines.append("</details>")
        
        return '\n'.join(lines)

    def _cluster_image_bboxes(
        self, bboxes: List[Tuple[float, float, float, float]], 
        distance_threshold: float = 15.0
    ) -> List[List[int]]:
        """画像のbboxをクラスタリング
        
        近接する画像要素をグループ化する。
        
        Args:
            bboxes: bboxのリスト [(x0, y0, x1, y1), ...]
            distance_threshold: クラスタリングの距離閾値（ピクセル）
            
        Returns:
            クラスタのリスト（各クラスタはbboxのインデックスリスト）
        """
        if not bboxes:
            return []
        
        n = len(bboxes)
        visited = [False] * n
        clusters = []
        
        def boxes_overlap_or_close(b1, b2, threshold):
            """2つのbboxが重なるか、近接しているかを判定"""
            x0_1, y0_1, x1_1, y1_1 = b1
            x0_2, y0_2, x1_2, y1_2 = b2
            
            # 重なりチェック
            if not (x1_1 < x0_2 - threshold or x1_2 < x0_1 - threshold or
                    y1_1 < y0_2 - threshold or y1_2 < y0_1 - threshold):
                return True
            return False
        
        for i in range(n):
            if visited[i]:
                continue
            
            # 新しいクラスタを開始
            cluster = [i]
            visited[i] = True
            queue = [i]
            
            while queue:
                current = queue.pop(0)
                current_bbox = bboxes[current]
                
                for j in range(n):
                    if visited[j]:
                        continue
                    
                    if boxes_overlap_or_close(current_bbox, bboxes[j], distance_threshold):
                        cluster.append(j)
                        visited[j] = True
                        queue.append(j)
            
            clusters.append(cluster)
        
        return clusters

    def _get_cluster_union_bbox(
        self, bboxes: List[Tuple[float, float, float, float]], 
        indices: List[int],
        margin: float = 5.0
    ) -> Tuple[float, float, float, float]:
        """クラスタ内のbboxの和集合を計算
        
        Args:
            bboxes: 全bboxのリスト
            indices: クラスタに含まれるbboxのインデックス
            margin: 周囲に追加するマージン（ピクセル）
            
        Returns:
            和集合のbbox (x0, y0, x1, y1)
        """
        cluster_bboxes = [bboxes[i] for i in indices]
        x0 = min(b[0] for b in cluster_bboxes) - margin
        y0 = min(b[1] for b in cluster_bboxes) - margin
        x1 = max(b[2] for b in cluster_bboxes) + margin
        y1 = max(b[3] for b in cluster_bboxes) + margin
        return (max(0, x0), max(0, y0), x1, y1)

    def _extract_embedded_images(self, page, page_num: int) -> List[Dict[str, Any]]:
        """PDFページから埋め込み画像を抽出
        
        複数の図形要素をクラスタリングして1つの図としてまとめる。
        
        Args:
            page: PyMuPDFのページオブジェクト
            page_num: ページ番号（0始まり）
            
        Returns:
            抽出された画像情報のリスト
        """
        images = []
        
        # 画像のbboxを収集
        image_bboxes = []
        image_xrefs = []
        
        try:
            image_list = page.get_images(full=True)
        except Exception as e:
            debug_print(f"[DEBUG] 画像リスト取得エラー: {e}")
            return []
        
        for img_info in image_list:
            try:
                xref = img_info[0]
                for img_rect in page.get_image_rects(xref):
                    bbox = (img_rect.x0, img_rect.y0, img_rect.x1, img_rect.y1)
                    # 小さすぎる画像は除外（面積が100平方ピクセル未満）
                    area = (bbox[2] - bbox[0]) * (bbox[3] - bbox[1])
                    if area >= 100:
                        image_bboxes.append(bbox)
                        image_xrefs.append(xref)
                    break
            except Exception as e:
                debug_print(f"[DEBUG] 画像bbox取得エラー: {e}")
                continue
        
        if not image_bboxes:
            return []
        
        # 画像が少ない場合はクラスタリングせずに個別に出力
        if len(image_bboxes) <= 3:
            return self._extract_individual_images(page, page_num, image_bboxes, image_xrefs)
        
        # 画像をクラスタリング
        clusters = self._cluster_image_bboxes(image_bboxes, distance_threshold=20.0)
        
        debug_print(f"[DEBUG] ページ {page_num + 1}: {len(image_bboxes)}個の画像要素を{len(clusters)}個のクラスタにグループ化")
        
        for cluster_idx, cluster in enumerate(clusters):
            try:
                # クラスタが1つの画像のみの場合
                if len(cluster) == 1:
                    bbox = image_bboxes[cluster[0]]
                    # 十分な大きさがある場合のみ出力
                    area = (bbox[2] - bbox[0]) * (bbox[3] - bbox[1])
                    if area < 500:  # 小さすぎるクラスタはスキップ
                        continue
                
                # クラスタの和集合bboxを計算
                union_bbox = self._get_cluster_union_bbox(image_bboxes, cluster, margin=3.0)
                
                # クラスタ領域をレンダリング
                self.image_counter += 1
                image_filename = f"{self.base_name}_fig_{page_num + 1:03d}_{self.image_counter:03d}"
                
                # クリップ領域を指定してレンダリング
                clip_rect = fitz.Rect(union_bbox)
                matrix = fitz.Matrix(2.0, 2.0)  # 2倍の解像度
                pix = page.get_pixmap(matrix=matrix, clip=clip_rect)
                
                if self.output_format == 'svg':
                    image_path = os.path.join(self.images_dir, f"{image_filename}.svg")
                    temp_png = os.path.join(self.images_dir, f"temp_cluster_{self.image_counter}.png")
                    pix.save(temp_png)
                    self._convert_png_to_svg(temp_png, image_path)
                    if os.path.exists(temp_png):
                        os.remove(temp_png)
                else:
                    image_path = os.path.join(self.images_dir, f"{image_filename}.png")
                    pix.save(image_path)
                
                images.append({
                    "path": image_path,
                    "filename": os.path.basename(image_path),
                    "bbox": union_bbox,
                    "y_position": union_bbox[1]
                })
                
                debug_print(f"[DEBUG] クラスタ画像を抽出: {image_path} ({len(cluster)}要素)")
                
            except Exception as e:
                debug_print(f"[DEBUG] クラスタ画像抽出エラー: {e}")
                continue
        
        # Y座標でソート（上から順に）
        images.sort(key=lambda x: x["y_position"])
        
        return images

    def _extract_individual_images(
        self, page, page_num: int, 
        bboxes: List[Tuple[float, float, float, float]],
        xrefs: List[int]
    ) -> List[Dict[str, Any]]:
        """個別の画像を抽出（クラスタリングなし）
        
        Args:
            page: PyMuPDFのページオブジェクト
            page_num: ページ番号
            bboxes: 画像のbboxリスト
            xrefs: 画像のxrefリスト
            
        Returns:
            抽出された画像情報のリスト
        """
        images = []
        doc = page.parent
        
        for bbox, xref in zip(bboxes, xrefs):
            try:
                base_image = doc.extract_image(xref)
                if not base_image:
                    continue
                
                image_bytes = base_image.get("image")
                image_ext = base_image.get("ext", "png")
                
                if not image_bytes:
                    continue
                
                self.image_counter += 1
                image_filename = f"{self.base_name}_img_{page_num + 1:03d}_{self.image_counter:03d}"
                
                if self.output_format == 'svg':
                    image_path = os.path.join(self.images_dir, f"{image_filename}.svg")
                    temp_png = os.path.join(self.images_dir, f"temp_{self.image_counter}.png")
                    with open(temp_png, 'wb') as f:
                        f.write(image_bytes)
                    
                    try:
                        with Image.open(temp_png) as img:
                            png_path = temp_png
                            if image_ext.lower() not in ('png',):
                                png_path = temp_png.replace('.png', '_conv.png')
                                img.save(png_path, 'PNG')
                        
                        self._convert_png_to_svg(png_path, image_path)
                        
                        if os.path.exists(temp_png):
                            os.remove(temp_png)
                        if png_path != temp_png and os.path.exists(png_path):
                            os.remove(png_path)
                    except Exception as e:
                        debug_print(f"[DEBUG] SVG変換エラー: {e}")
                        image_path = os.path.join(self.images_dir, f"{image_filename}.png")
                        with open(image_path, 'wb') as f:
                            f.write(image_bytes)
                else:
                    image_path = os.path.join(self.images_dir, f"{image_filename}.png")
                    with open(image_path, 'wb') as f:
                        f.write(image_bytes)
                
                images.append({
                    "path": image_path,
                    "filename": os.path.basename(image_path),
                    "bbox": bbox,
                    "y_position": bbox[1] if bbox else 0
                })
                
                debug_print(f"[DEBUG] 埋め込み画像を抽出: {image_path}")
                
            except Exception as e:
                debug_print(f"[DEBUG] 画像抽出エラー: {e}")
                continue
        
        images.sort(key=lambda x: x["y_position"])
        return images
