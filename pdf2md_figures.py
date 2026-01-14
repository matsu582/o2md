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

リファクタリング:
- _extract_all_figuresをオーケストレータ化
- 処理フェーズごとにサブメソッドに分割
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
    
    リファクタリング後の構造:
    - _extract_all_figures: オーケストレータ（各フェーズを呼び出す）
    - _fig_compute_header_footer_bounds: ヘッダー/フッター領域計算
    - _fig_collect_graphics_elements: 描画要素と画像の収集
    - _fig_detect_figure_captions: 図キャプション検出
    - _fig_cluster_elements: クラスタリングと候補生成
    - _fig_filter_table_regions: 表領域フィルタリング
    - _fig_merge_same_column: 同一カラム内マージ
    - _fig_merge_left_right: 左右クラスタマージ
    - _fig_render_candidates: 画像レンダリング
    """

    # ユーティリティメソッド（staticmethod化）
    @staticmethod
    def _fig_bbox_contains(outer: Tuple, inner: Tuple, margin: float = 10.0) -> bool:
        """outerがinnerを包含しているか判定"""
        return (outer[0] - margin <= inner[0] and 
                outer[1] - margin <= inner[1] and 
                outer[2] + margin >= inner[2] and 
                outer[3] + margin >= inner[3])

    @staticmethod
    def _fig_bbox_overlap_ratio(bbox1: Tuple, bbox2: Tuple) -> float:
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

    @staticmethod
    def _fig_get_x_overlap_ratio(bbox1: Tuple, bbox2: Tuple) -> float:
        """x方向の重なり率を計算"""
        x_overlap = max(0, min(bbox1[2], bbox2[2]) - max(bbox1[0], bbox2[0]))
        min_width = min(bbox1[2] - bbox1[0], bbox2[2] - bbox2[0])
        if min_width <= 0:
            return 0
        return x_overlap / min_width

    @staticmethod
    def _fig_get_y_overlap_ratio(bbox1: Tuple, bbox2: Tuple) -> float:
        """Y方向の重なり率を計算"""
        y_overlap = max(0, min(bbox1[3], bbox2[3]) - max(bbox1[1], bbox2[1]))
        min_height = min(bbox1[3] - bbox1[1], bbox2[3] - bbox2[1])
        if min_height <= 0:
            return 0
        return y_overlap / min_height

    @staticmethod
    def _fig_is_body_text_line(text: str, line_width: float, col_width: float) -> bool:
        """本文らしい行かどうかを判定"""
        if len(text) < 15:
            return False
        if "。" in text:
            return True
        particles = ["が", "を", "に", "で", "は", "の", "と", "も", "や"]
        if any(p in text for p in particles) and line_width > col_width * 0.5:
            return True
        return False

    def _fig_is_text_box_candidate(
        self, candidate: Dict, text_lines: List[Dict], col_width: float
    ) -> bool:
        """図形候補が囲み記事（テキストボックス）かどうかを判定
        
        囲み記事の特徴:
        - image要素を含まない（image_count == 0）
        - bbox内に段落っぽい本文行が多数存在する
        
        Args:
            candidate: 図形候補の辞書
            text_lines: ページ内のテキスト行リスト
            col_width: カラム幅
            
        Returns:
            囲み記事と判定された場合True
        """
        # image要素を含む場合は囲み記事ではない
        image_count = candidate.get("image_count", 0)
        if image_count > 0:
            return False
        
        # bboxを取得
        bbox = candidate.get("raw_union_bbox", candidate.get("union_bbox"))
        if not bbox:
            return False
        
        # bbox内のテキスト行を収集
        body_line_count = 0
        total_line_count = 0
        
        for line in text_lines:
            line_bbox = line.get("bbox", (0, 0, 0, 0))
            line_text = line.get("text", "").strip()
            
            # bbox内のテキスト行かどうかを判定
            line_center_y = (line_bbox[1] + line_bbox[3]) / 2
            line_center_x = (line_bbox[0] + line_bbox[2]) / 2
            
            if not (bbox[0] <= line_center_x <= bbox[2] and 
                    bbox[1] <= line_center_y <= bbox[3]):
                continue
            
            total_line_count += 1
            line_width = line_bbox[2] - line_bbox[0]
            
            # 本文らしい行かどうかを判定
            if self._fig_is_body_text_line(line_text, line_width, col_width):
                body_line_count += 1
        
        # 本文行が10行以上ある場合は囲み記事と判定
        # （短い行や空白行が多い場合でも、本文行が十分にあれば囲み記事）
        if body_line_count >= 10:
            debug_print(f"[DEBUG] 囲み記事検出: {body_line_count}本文行/{total_line_count}行")
            return True
        
        return False

    def _fig_compute_header_footer_bounds(
        self, page, page_num: int, page_height: float, 
        header_footer_patterns: Set[str]
    ) -> Tuple[Optional[float], Optional[float]]:
        """ヘッダー/フッター領域のY座標境界を計算
        
        Args:
            page: PyMuPDFのページオブジェクト
            page_num: ページ番号
            page_height: ページの高さ
            header_footer_patterns: ヘッダ・フッタパターンのセット
            
        Returns:
            (header_y_max, footer_y_min) のタプル
        """
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
                        
                        if self._is_header_footer(
                            line_text, header_footer_patterns,
                            y_pos=line_bbox[1], page_height=page_height
                        ):
                            if y_center < page_height / 2:
                                header_lines_y.append(line_bbox[3])
                            else:
                                footer_lines_y.append(line_bbox[1])
                
                if header_lines_y:
                    text_based_y = max(header_lines_y) + 15.0
                    position_based_y = page_height * 0.08
                    header_y_max = max(text_based_y, position_based_y)
                    debug_print(f"[DEBUG] page={page_num+1}: ヘッダー領域検出 y_max={header_y_max:.1f}")
                if footer_lines_y:
                    text_based_y = min(footer_lines_y) - 15.0
                    position_based_y = page_height * 0.92
                    footer_y_min = min(text_based_y, position_based_y)
                    debug_print(f"[DEBUG] page={page_num+1}: フッター領域検出 y_min={footer_y_min:.1f}")
                    
            except Exception as e:
                debug_print(f"[DEBUG] ヘッダー/フッター領域計算エラー: {e}")
        
        # doc-wideの値をフォールバックとして使用
        if header_y_max is None and self._doc_header_y_max is not None:
            header_y_max = self._doc_header_y_max
            debug_print(f"[DEBUG] page={page_num+1}: doc-wideヘッダー領域をフォールバック適用")
        if footer_y_min is None and self._doc_footer_y_min is not None:
            footer_y_min = self._doc_footer_y_min
            debug_print(f"[DEBUG] page={page_num+1}: doc-wideフッター領域をフォールバック適用")
        
        # 外れ値ガード
        if header_y_max is not None and self._doc_header_y_max is not None:
            if header_y_max > self._doc_header_y_max * 3.0:
                debug_print(f"[DEBUG] page={page_num+1}: ヘッダー領域が外れ値、doc-wide値にフォールバック")
                header_y_max = self._doc_header_y_max
        if footer_y_min is not None and self._doc_footer_y_min is not None:
            page_footer_dist = page_height - footer_y_min
            doc_footer_dist = page_height - self._doc_footer_y_min
            if page_footer_dist > doc_footer_dist * 3.0:
                debug_print(f"[DEBUG] page={page_num+1}: フッター領域が外れ値、doc-wide値にフォールバック")
                footer_y_min = self._doc_footer_y_min
        
        return header_y_max, footer_y_min

    def _fig_collect_graphics_elements(
        self, page, page_num: int
    ) -> Tuple[List[Dict], List[Tuple]]:
        """描画要素と画像を収集
        
        Args:
            page: PyMuPDFのページオブジェクト
            page_num: ページ番号
            
        Returns:
            (all_elements, all_bboxes) のタプル
        """
        all_elements = []
        page_area = page.rect.width * page.rect.height
        
        try:
            drawings = page.get_drawings()
            for d in drawings:
                rect = d.get("rect")
                if rect:
                    bbox = (rect.x0, rect.y0, rect.x1, rect.y1)
                    area = (bbox[2] - bbox[0]) * (bbox[3] - bbox[1])
                    if area >= 200:
                        # PPT由来の背景矩形を除外（ページ面積の90%以上を覆うdrawing）
                        if area >= page_area * 0.9:
                            debug_print(f"[DEBUG] page={page_num+1}: 背景矩形を除外（面積比={area/page_area:.2f}）")
                            continue
                        all_elements.append({"bbox": bbox, "type": "drawing"})
        except Exception as e:
            debug_print(f"[DEBUG] 描画取得エラー: {e}")
        
        # 生テキストが存在するかどうかを確認（背景画像除外の判定に使用）
        raw_text = page.get_text().strip()
        has_raw_text = len(raw_text) > 0
        
        # 繰り返し出現するヘッダ装飾画像のxrefセットを取得
        repeated_xrefs = getattr(self, '_repeated_image_xrefs', set())
        
        try:
            image_list = page.get_images(full=True)
            for img_info in image_list:
                xref = img_info[0]
                # ヘッダ装飾画像を除外（繰り返し出現する小さな上部画像）
                if xref in repeated_xrefs:
                    debug_print(f"[DEBUG] page={page_num+1}: ヘッダ装飾画像を除外（xref={xref}）")
                    continue
                for img_rect in page.get_image_rects(xref):
                    bbox = (img_rect.x0, img_rect.y0, img_rect.x1, img_rect.y1)
                    area = (bbox[2] - bbox[0]) * (bbox[3] - bbox[1])
                    if area >= 100:
                        # 生テキストが存在するページでは、全面画像を背景として除外
                        if has_raw_text and area >= page_area * 0.9:
                            debug_print(f"[DEBUG] page={page_num+1}: 背景画像を除外（面積比={area/page_area:.2f}）")
                            continue
                        all_elements.append({"bbox": bbox, "type": "image"})
        except Exception as e:
            debug_print(f"[DEBUG] 画像取得エラー: {e}")
        
        all_bboxes = [e["bbox"] for e in all_elements]
        return all_elements, all_bboxes

    def _fig_detect_figure_captions(self, page) -> List[Dict]:
        """図キャプションの位置を検出
        
        Args:
            page: PyMuPDFのページオブジェクト
            
        Returns:
            図キャプション情報のリスト
        """
        figure_caption_lines = []
        try:
            caption_pattern = re.compile(r'^図\s*\d+[\.\:．：]')
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
        return figure_caption_lines

    def _fig_has_caption_between(
        self, bbox1: Tuple, bbox2: Tuple, figure_caption_lines: List[Dict]
    ) -> bool:
        """2つのbbox間に図キャプションがあるかどうかを判定"""
        y1_bottom = bbox1[3]
        y2_top = bbox2[1]
        y1_top = bbox1[1]
        y2_bottom = bbox2[3]
        
        if y1_bottom < y2_top:
            gap_top = y1_bottom
            gap_bottom = y2_top
        elif y2_bottom < y1_top:
            gap_top = y2_bottom
            gap_bottom = y1_top
        else:
            return False
        
        for cap in figure_caption_lines:
            cap_y = cap["y_center"]
            if gap_top <= cap_y <= gap_bottom:
                return True
        return False

    def _fig_cluster_elements(
        self, all_bboxes: List[Tuple], all_elements: List[Dict],
        bbox_columns: List[str], figure_caption_lines: List[Dict],
        page_width: float, page_height: float, page_num: int,
        header_y_max: Optional[float], footer_y_min: Optional[float],
        gutter_x: float, gutter_margin: float,
        is_slide_document: bool = False
    ) -> Tuple[List[Dict], Optional[Dict]]:
        """ガター制約付きクラスタリングと候補生成
        
        Args:
            all_bboxes: 全bboxのリスト
            all_elements: 全要素のリスト
            bbox_columns: 各bboxのカラム情報
            figure_caption_lines: 図キャプション情報
            page_width: ページ幅
            page_height: ページ高さ
            page_num: ページ番号
            header_y_max: ヘッダー領域の下端Y座標
            footer_y_min: フッター領域の上端Y座標
            gutter_x: ガターのX座標
            gutter_margin: ガターマージン
            is_slide_document: スライド文書フラグ
            
        Returns:
            (all_figure_candidates, single_image_candidate) のタプル
        """
        # 単独画像候補の処理
        single_image_candidate = None
        if len(all_bboxes) == 1:
            bbox = all_bboxes[0]
            area = (bbox[2] - bbox[0]) * (bbox[3] - bbox[1])
            if area >= 10000:
                in_header = header_y_max is not None and bbox[3] < header_y_max
                in_footer = footer_y_min is not None and bbox[1] > footer_y_min
                if not in_header and not in_footer:
                    col = "full" if (bbox[2] - bbox[0]) > page_width * 0.5 else (
                        "left" if (bbox[0] + bbox[2]) / 2 < page_width / 2 else "right"
                    )
                    # 単独画像候補にdrawing_countとimage_bboxesを追加
                    elem_type = all_elements[0]["type"]
                    is_image = elem_type == "image"
                    # 有意な画像かどうかを判定（ページ面積の5%以上）
                    img_area = (bbox[2] - bbox[0]) * (bbox[3] - bbox[1])
                    page_area_for_single = page_width * page_height
                    is_significant = is_image and img_area >= page_area_for_single * 0.05
                    single_image_candidate = {
                        "union_bbox": bbox,
                        "raw_union_bbox": bbox,
                        "column": col,
                        "cluster_size": 1,
                        "is_embedded": is_image,
                        "image_count": 1 if is_image else 0,
                        "drawing_count": 0 if is_image else 1,
                        "significant_image_count": 1 if is_significant else 0,
                        "image_bboxes": [bbox] if is_image else [],
                    }
                    debug_print(f"[DEBUG] page={page_num+1}: 単独画像を候補として追加")
                else:
                    if in_header:
                        debug_print(f"[DEBUG] page={page_num+1}: 単独画像がヘッダー領域内のため除外")
                    if in_footer:
                        debug_print(f"[DEBUG] page={page_num+1}: 単独画像がフッター領域内のため除外")
            else:
                debug_print(f"[DEBUG] page={page_num+1}: 単独画像が小さすぎるため除外")

        def boxes_close(idx1, idx2):
            b1, b2 = all_bboxes[idx1], all_bboxes[idx2]
            col1, col2 = bbox_columns[idx1], bbox_columns[idx2]
            
            # スライド文書の場合はカラム制約を無効化（図形が分割されるのを防ぐ）
            if not is_slide_document:
                if col1 != col2 and col1 != "full" and col2 != "full":
                    return False
            
            if self._fig_has_caption_between(b1, b2, figure_caption_lines):
                debug_print(f"[DEBUG] キャプション制約: bbox間にキャプションあり")
                return False
            
            x0_1, y0_1, x1_1, y1_1 = b1
            x0_2, y0_2, x1_2, y1_2 = b2
            
            x_gap = max(0, max(x0_1, x0_2) - min(x1_1, x1_2))
            y_gap = max(0, max(y0_1, y0_2) - min(y1_1, y1_2))
            
            # スライド文書の場合はy_gap閾値を緩和（図形が分割されるのを防ぐ）
            y_gap_threshold = 150.0 if is_slide_document else 40.0
            return x_gap <= 100.0 and y_gap <= y_gap_threshold

        # クラスタリング
        n = len(all_bboxes)
        visited = [False] * n
        clusters = []
        
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

        # 混在クラスタの分割処理
        # drawing要素とimage要素が混在し、y方向に離れている場合は分割
        # スライド文書の場合は分割を無効化（図形が分割されるのを防ぐ）
        if not is_slide_document:
            split_clusters = []
            for cluster in clusters:
                if len(cluster) < 2:
                    split_clusters.append(cluster)
                    continue
                
                # クラスタ内の要素タイプを確認
                drawing_indices = [i for i in cluster if all_elements[i]["type"] == "drawing"]
                image_indices = [i for i in cluster if all_elements[i]["type"] == "image"]
                
                # drawing要素とimage要素が両方存在する場合のみ分割を検討
                if not drawing_indices or not image_indices:
                    split_clusters.append(cluster)
                    continue
                
                # drawing要素のy座標範囲を取得
                drawing_y_max = max(all_bboxes[i][3] for i in drawing_indices)
                drawing_y_min = min(all_bboxes[i][1] for i in drawing_indices)
                
                # image要素のy座標範囲を取得
                image_y_min = min(all_bboxes[i][1] for i in image_indices)
                image_y_max = max(all_bboxes[i][3] for i in image_indices)
                
                # drawing要素とimage要素がy方向に離れている場合（20px以上）は分割
                y_gap = image_y_min - drawing_y_max
                if y_gap > 20.0:
                    # drawingクラスタとimageクラスタに分割
                    split_clusters.append(drawing_indices)
                    split_clusters.append(image_indices)
                    debug_print(f"[DEBUG] page={page_num+1}: 混在クラスタを分割（y_gap={y_gap:.1f}）")
                else:
                    split_clusters.append(cluster)
            
            clusters = split_clusters

        # 候補生成
        all_figure_candidates = []
        for cluster in clusters:
            if len(cluster) < 2:
                bbox = all_bboxes[cluster[0]]
                area = (bbox[2] - bbox[0]) * (bbox[3] - bbox[1])
                if area < 2000:
                    continue
            
            cluster_bboxes = [all_bboxes[i] for i in cluster]
            cluster_types = [all_elements[i]["type"] for i in cluster]
            
            raw_x0 = min(b[0] for b in cluster_bboxes)
            raw_y0 = min(b[1] for b in cluster_bboxes)
            raw_x1 = max(b[2] for b in cluster_bboxes)
            raw_y1 = max(b[3] for b in cluster_bboxes)
            raw_union_bbox = (raw_x0, raw_y0, raw_x1, raw_y1)
            raw_area = (raw_x1 - raw_x0) * (raw_y1 - raw_y0)
            
            cluster_height = raw_y1 - raw_y0
            is_in_header = False
            is_in_footer = False
            # 巨大クラスタ（ページの50%以上の高さ）は緩和判定
            is_large_cluster = cluster_height > page_height * 0.5
            if header_y_max is not None:
                overlap_with_header = max(0, min(raw_y1, header_y_max) - raw_y0)
                if is_large_cluster:
                    # 巨大クラスタはoverlap率のみで判定（緩和）
                    # raw_y0の条件は適用しない（表画像などがヘッダー領域から始まる場合がある）
                    is_in_header = overlap_with_header > cluster_height * 0.8
                else:
                    is_in_header = overlap_with_header > cluster_height * 0.5 or raw_y0 < header_y_max * 0.5
            if footer_y_min is not None:
                overlap_with_footer = max(0, raw_y1 - max(raw_y0, footer_y_min))
                if is_large_cluster:
                    # 巨大クラスタはoverlap率のみで判定（緩和）
                    is_in_footer = overlap_with_footer > cluster_height * 0.8
                else:
                    is_in_footer = overlap_with_footer > cluster_height * 0.5 or raw_y1 > footer_y_min + (page_height - footer_y_min) * 0.5
            if is_in_header or is_in_footer:
                region = "ヘッダー" if is_in_header else "フッター"
                debug_print(f"[DEBUG] page={page_num+1}: {region}領域内のクラスタを除外")
                continue
            
            is_image_only = all(t == "image" for t in cluster_types)
            is_embedded = is_image_only and len(cluster) <= 2 and raw_area < 10000
            
            padding = 0.0 if is_embedded else 20.0
            x0 = raw_x0 - padding
            y0 = raw_y0 - padding
            x1 = raw_x1 + padding
            y1 = raw_y1 + padding
            union_bbox = (max(0, x0), max(0, y0), min(page_width, x1), min(page_height, y1))
            
            area = (union_bbox[2] - union_bbox[0]) * (union_bbox[3] - union_bbox[1])
            if area < 1000:
                continue
            
            # カラム判定
            width = raw_x1 - raw_x0
            center_x = (raw_x0 + raw_x1) / 2
            crosses_gutter = raw_x0 < gutter_x - gutter_margin and raw_x1 > gutter_x + gutter_margin
            is_full_width = width > page_width * 0.5
            if crosses_gutter or is_full_width:
                column = "full"
            elif center_x < gutter_x:
                column = "left"
            else:
                column = "right"
            
            image_count = sum(1 for t in cluster_types if t == "image")
            drawing_count = sum(1 for t in cluster_types if t == "drawing")
            
            # 有意な画像数を計算（ページ面積の5%以上の画像のみカウント）
            # 小さなロゴ等は「有意な画像」としてカウントしない
            # また、埋め込み画像のbboxリストを収集（本文フィルタリング用）
            page_area = page_width * page_height
            significant_image_count = 0
            image_bboxes = []
            for idx in cluster:
                if all_elements[idx]["type"] == "image":
                    img_bbox = all_bboxes[idx]
                    image_bboxes.append(img_bbox)
                    img_area = (img_bbox[2] - img_bbox[0]) * (img_bbox[3] - img_bbox[1])
                    if img_area >= page_area * 0.05:
                        significant_image_count += 1
            
            all_figure_candidates.append({
                "union_bbox": union_bbox,
                "raw_union_bbox": raw_union_bbox,
                "cluster_size": len(cluster),
                "column": column,
                "is_embedded": is_embedded,
                "image_count": image_count,
                "drawing_count": drawing_count,
                "significant_image_count": significant_image_count,
                "image_bboxes": image_bboxes
            })

        # 包含除去フィルタ
        filtered_candidates = []
        for i, cand in enumerate(all_figure_candidates):
            is_contained = False
            for j, other in enumerate(all_figure_candidates):
                if i != j:
                    if self._fig_bbox_contains(other["union_bbox"], cand["union_bbox"]):
                        is_contained = True
                        break
            if not is_contained:
                filtered_candidates.append(cand)
        
        all_figure_candidates = filtered_candidates
        
        # 単独画像候補の追加
        if single_image_candidate is not None:
            single_bbox = single_image_candidate["union_bbox"]
            is_duplicate = False
            for cand in all_figure_candidates:
                cand_bbox = cand["union_bbox"]
                if self._fig_bbox_contains(cand_bbox, single_bbox) or self._fig_bbox_contains(single_bbox, cand_bbox):
                    is_duplicate = True
                    debug_print(f"[DEBUG] page={page_num+1}: 単独画像候補が既存候補と重複のため追加しない")
                    break
            if not is_duplicate:
                all_figure_candidates.append(single_image_candidate)
                debug_print(f"[DEBUG] page={page_num+1}: 単独画像候補をall_figure_candidatesに追加")
        
        return all_figure_candidates, single_image_candidate

    def _fig_detect_table_bboxes_from_text(
        self, text_lines: List[Dict], page_width: float
    ) -> List[Tuple]:
        """テキスト行の配置パターンから表領域を検出"""
        table_bboxes = []
        gutter = page_width / 2
        
        caption_pattern = re.compile(r'^[図表]\s*\d+')
        
        for is_left in [True, False]:
            col_lines = []
            for line in text_lines:
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
            
            y_tolerance = 5
            y_groups = {}
            for line in col_lines:
                y_key = round(line["bbox"][1] / y_tolerance) * y_tolerance
                if y_key not in y_groups:
                    y_groups[y_key] = []
                y_groups[y_key].append(line)
            
            multi_cell_rows = []
            for y_key in sorted(y_groups.keys()):
                cells = y_groups[y_key]
                x_positions = sorted(set(round(c["bbox"][0] / 20) * 20 for c in cells))
                if len(x_positions) >= 2:
                    texts = [c.get("text", "") for c in cells]
                    has_long_text = any(len(t) > 3 for t in texts)
                    if not has_long_text:
                        continue
                    all_bboxes_row = [c["bbox"] for c in cells]
                    row_bbox = (
                        min(b[0] for b in all_bboxes_row),
                        min(b[1] for b in all_bboxes_row),
                        max(b[2] for b in all_bboxes_row),
                        max(b[3] for b in all_bboxes_row)
                    )
                    multi_cell_rows.append({"y": y_key, "bbox": row_bbox})
            
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
            
            if len(current_region["rows"]) >= 5:
                table_bboxes.append((
                    current_region["x0"] - 5,
                    current_region["y_start"] - 10,
                    current_region["x1"] + 5,
                    current_region["y_end"] + 5
                ))
        
        return table_bboxes

    def _fig_can_output_as_markdown_table(
        self, table_bbox: Tuple, text_lines: List[Dict], is_two_column_layout: bool
    ) -> bool:
        """表領域がMarkdownテーブルとして出力可能かを判定"""
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
        
        y_groups = {}
        for line in table_lines:
            y_key = round(line["bbox"][1] / 5) * 5
            if y_key not in y_groups:
                y_groups[y_key] = []
            y_groups[y_key].append(line)
        
        col_counts = []
        for y_key in sorted(y_groups.keys()):
            cells = y_groups[y_key]
            x_positions = sorted(set(round(c["bbox"][0] / 20) * 20 for c in cells))
            if len(x_positions) >= 2:
                col_counts.append(len(x_positions))
        
        if len(col_counts) < 3:
            debug_print(f"[DEBUG] 複数列行数不足: {len(col_counts)} < 3")
            return False
        
        most_common = max(set(col_counts), key=col_counts.count)
        consistent = sum(1 for c in col_counts if c == most_common)
        consistency_ratio = consistent / len(col_counts)
        
        debug_print(f"[DEBUG] 表の列数一貫性: {consistent}/{len(col_counts)} = {consistency_ratio:.2f}")
        
        if consistency_ratio < 0.8:
            return False
        
        if most_common < 3:
            debug_print(f"[DEBUG] 2列表のためMarkdownテーブルとして出力不可")
            return False
        
        return True

    def _fig_filter_table_regions(
        self, all_figure_candidates: List[Dict], table_bboxes: List[Tuple],
        page_text_lines: List[Dict], column_count: int, page_num: int
    ) -> List[Dict]:
        """表領域と重なる図候補をフィルタリング"""
        debug_print(f"[DEBUG] page={page_num+1}: 表フィルタ前の候補数={len(all_figure_candidates)}")
        table_filtered = []
        table_image_candidates = []
        
        for cand in all_figure_candidates:
            is_table = False
            cand_bbox = cand["union_bbox"]
            matched_table_bbox = None
            
            for table_bbox in table_bboxes:
                overlap = self._fig_bbox_overlap_ratio(cand_bbox, table_bbox)
                if overlap >= 0.7:
                    matched_table_bbox = table_bbox
                    is_table = True
                    break
            
            if is_table and matched_table_bbox:
                is_two_col = (column_count >= 2)
                if self._fig_can_output_as_markdown_table(matched_table_bbox, page_text_lines, is_two_col):
                    debug_print(f"[DEBUG] page={page_num+1}: 図候補をMarkdown表として除外")
                else:
                    debug_print(f"[DEBUG] page={page_num+1}: 図候補を表画像として出力")
                    union = cand_bbox
                    
                    union_top = union[1]
                    clip_y0 = union_top - 2
                    
                    for line in page_text_lines:
                        line_bbox = line["bbox"]
                        line_text = line["text"]
                        if line_bbox[1] >= union_top - 5 and line_bbox[3] <= union_top + 20:
                            if re.match(r'^表\s*\d', line_text):
                                clip_y0 = line_bbox[3] + 2
                                debug_print(f"[DEBUG] 表画像: キャプション検出 '{line_text[:20]}'")
                                break
                    
                    clip_x0 = union[0] - 2
                    clip_x1 = union[2] + 2
                    clip_y1 = union[3] + 5
                    
                    union_bottom = union[3]
                    for line in page_text_lines:
                        line_bbox = line["bbox"]
                        line_text = line["text"]
                        if line_bbox[1] >= union_bottom - 20 and line_bbox[1] <= union_bottom + 30:
                            if re.match(r'^(図|表)\s*\d', line_text):
                                clip_y1 = line_bbox[1] - 2
                                debug_print(f"[DEBUG] 表画像: 下キャプション検出 '{line_text[:20]}'")
                                break
                    
                    # 注釈テキスト（※で始まるテキスト）を検出してclip_y1を調整
                    for line in page_text_lines:
                        line_bbox = line["bbox"]
                        line_text = line["text"]
                        # union_bottom付近（-30〜+50）にある注釈テキストを検出
                        if line_bbox[1] >= union_bottom - 30 and line_bbox[1] <= union_bottom + 50:
                            if re.match(r'^※\d', line_text):
                                # 注釈テキストの上端をclip_y1として使用
                                new_clip_y1 = line_bbox[1] - 2
                                if new_clip_y1 < clip_y1:
                                    clip_y1 = new_clip_y1
                                    debug_print(f"[DEBUG] 表画像: 注釈テキスト検出 '{line_text[:30]}', clip_y1={clip_y1:.1f}")
                                break
                    
                    debug_print(f"[DEBUG] 表画像: clip_bbox=({clip_x0:.1f}, {clip_y0:.1f}, {clip_x1:.1f}, {clip_y1:.1f})")
                    
                    table_clip_bbox = (clip_x0, clip_y0, clip_x1, clip_y1)
                    cand["clip_bbox"] = table_clip_bbox
                    cand["is_table_image"] = True
                    table_image_candidates.append(cand)
            else:
                table_filtered.append(cand)
        
        debug_print(f"[DEBUG] page={page_num+1}: 表フィルタ後の候補数={len(table_filtered)}, 表画像候補数={len(table_image_candidates)}")
        return table_filtered + table_image_candidates

    def _fig_has_body_barrier(
        self, bbox1: Tuple, bbox2: Tuple, text_lines: List[Dict], col_width: float
    ) -> bool:
        """2つのbbox間に本文バリアがあるか判定"""
        y_min = min(bbox1[3], bbox2[3])
        y_max = max(bbox1[1], bbox2[1])
        x_overlap_start = max(bbox1[0], bbox2[0])
        x_overlap_end = min(bbox1[2], bbox2[2])
        
        caption_pattern = re.compile(r'^図\s*\d+[\.\:．：]')
        
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
            
            if caption_pattern.match(line["text"]):
                if gap_top - 30 <= line_center_y <= gap_bottom + 30:
                    if x_overlap_start - 20 < line_center_x < x_overlap_end + 20:
                        debug_print(f"[DEBUG] 図キャプションバリア検出: '{line['text'][:30]}...'")
                        return True
            
            if y_min < line_center_y < y_max:
                if x_overlap_start - 20 < line_center_x < x_overlap_end + 20:
                    if self._fig_is_body_text_line(line["text"], line_width, col_width):
                        return True
        return False

    def _fig_merge_same_column(
        self, all_figure_candidates: List[Dict], page_text_lines: List[Dict],
        col_width: float
    ) -> List[Dict]:
        """同一カラム内のクラスタを安全にマージ"""
        debug_print(f"[DEBUG] クラスタ再マージ開始: {len(all_figure_candidates)}個のクラスタ")
        merged = True
        while merged:
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
                    
                    if cand1["column"] != cand2["column"]:
                        continue
                    
                    bbox1 = cand1["union_bbox"]
                    bbox2 = cand2["union_bbox"]
                    
                    x_overlap_ratio = self._fig_get_x_overlap_ratio(bbox1, bbox2)
                    if x_overlap_ratio < 0.3:
                        debug_print(f"[DEBUG] クラスタ{i}と{j}: x重なり不足 {x_overlap_ratio:.2f}")
                        continue
                    
                    y_gap = max(0, max(bbox1[1], bbox2[1]) - min(bbox1[3], bbox2[3]))
                    
                    if y_gap <= 80:
                        debug_print(f"[DEBUG] クラスタ{i}と{j}: y_gap={y_gap:.1f}")
                        if not self._fig_has_body_barrier(bbox1, bbox2, page_text_lines, col_width):
                            debug_print(f"[DEBUG] クラスタ{i}と{j}: マージ候補")
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
                    
                    merged_bbox = (
                        min(bbox1[0], bbox2[0]),
                        min(bbox1[1], bbox2[1]),
                        max(bbox1[2], bbox2[2]),
                        max(bbox1[3], bbox2[3])
                    )
                    
                    debug_print(f"[DEBUG] クラスタ{i}と{best_merge}をマージ")
                    # マージ時に重要なフィールドを保持
                    merged_image_bboxes = cand1.get("image_bboxes", []) + cand2.get("image_bboxes", [])
                    new_candidates.append({
                        "union_bbox": merged_bbox,
                        "cluster_size": cand1["cluster_size"] + cand2["cluster_size"],
                        "column": cand1["column"],
                        "image_count": cand1.get("image_count", 0) + cand2.get("image_count", 0),
                        "drawing_count": cand1.get("drawing_count", 0) + cand2.get("drawing_count", 0),
                        "significant_image_count": cand1.get("significant_image_count", 0) + cand2.get("significant_image_count", 0),
                        "image_bboxes": merged_image_bboxes
                    })
                    used.add(i)
                    used.add(best_merge)
                    merged = True
                else:
                    new_candidates.append(cand1)
                    used.add(i)
            
            all_figure_candidates = new_candidates
        
        debug_print(f"[DEBUG] クラスタ再マージ完了: {len(all_figure_candidates)}個のクラスタ")
        return all_figure_candidates

    def _fig_merge_left_right(
        self, all_figure_candidates: List[Dict], page_text_lines: List[Dict],
        col_width: float, gutter_x: float
    ) -> List[Dict]:
        """左右クラスタのマージ処理"""
        debug_print(f"[DEBUG] 左右クラスタマージ開始: {len(all_figure_candidates)}個のクラスタ")
        lr_merged = True
        while lr_merged:
            lr_merged = False
            new_candidates = []
            used = set()
            
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
                    
                    # 画像またはdrawing要素が両方に存在する場合のみマージ対象
                    left_image_count = left_cand.get("image_count", 0)
                    right_image_count = right_cand.get("image_count", 0)
                    left_drawing_count = left_cand.get("drawing_count", 0)
                    right_drawing_count = right_cand.get("drawing_count", 0)
                    left_has_elements = left_image_count > 0 or left_drawing_count > 0
                    right_has_elements = right_image_count > 0 or right_drawing_count > 0
                    if not left_has_elements or not right_has_elements:
                        debug_print(f"[DEBUG] 左右マージ候補{left_idx},{right_idx}: 要素なし")
                        continue
                    
                    left_bbox = left_cand["union_bbox"]
                    right_bbox = right_cand["union_bbox"]
                    
                    y_overlap_ratio = self._fig_get_y_overlap_ratio(left_bbox, right_bbox)
                    if y_overlap_ratio < 0.9:
                        debug_print(f"[DEBUG] 左右マージ候補{left_idx},{right_idx}: Y重なり不足 {y_overlap_ratio:.2f}")
                        continue
                    
                    left_right_edge = left_bbox[2]
                    right_left_edge = right_bbox[0]
                    x_gap = right_left_edge - left_right_edge
                    
                    if not (gutter_x - 50 < left_right_edge < gutter_x + 50 and
                            gutter_x - 50 < right_left_edge < gutter_x + 50):
                        debug_print(f"[DEBUG] 左右マージ候補{left_idx},{right_idx}: ガター近傍でない")
                        continue
                    
                    if x_gap > 100:
                        debug_print(f"[DEBUG] 左右マージ候補{left_idx},{right_idx}: X gap超過 {x_gap:.1f}")
                        continue
                    
                    if self._fig_has_body_barrier(left_bbox, right_bbox, page_text_lines, col_width):
                        debug_print(f"[DEBUG] 左右マージ候補{left_idx},{right_idx}: 本文バリアあり")
                        continue
                    
                    if y_overlap_ratio > best_y_overlap:
                        best_y_overlap = y_overlap_ratio
                        best_right_idx = right_idx
                
                if best_right_idx is not None:
                    right_cand = all_figure_candidates[best_right_idx]
                    left_bbox = left_cand["union_bbox"]
                    right_bbox = right_cand["union_bbox"]
                    
                    merged_bbox = (
                        min(left_bbox[0], right_bbox[0]),
                        min(left_bbox[1], right_bbox[1]),
                        max(left_bbox[2], right_bbox[2]),
                        max(left_bbox[3], right_bbox[3])
                    )
                    
                    debug_print(f"[DEBUG] 左右クラスタ{left_idx}と{best_right_idx}をマージ (Y重なり={best_y_overlap:.2f})")
                    # マージ時に重要なフィールドを保持
                    merged_image_bboxes = left_cand.get("image_bboxes", []) + right_cand.get("image_bboxes", [])
                    new_candidates.append({
                        "union_bbox": merged_bbox,
                        "cluster_size": left_cand["cluster_size"] + right_cand["cluster_size"],
                        "column": "full",
                        "image_count": left_cand.get("image_count", 0) + right_cand.get("image_count", 0),
                        "drawing_count": left_cand.get("drawing_count", 0) + right_cand.get("drawing_count", 0),
                        "significant_image_count": left_cand.get("significant_image_count", 0) + right_cand.get("significant_image_count", 0),
                        "image_bboxes": merged_image_bboxes
                    })
                    used.add(left_idx)
                    used.add(best_right_idx)
                    lr_merged = True
                else:
                    new_candidates.append(left_cand)
                    used.add(left_idx)
            
            for right_idx, right_cand in right_candidates:
                if right_idx not in used:
                    new_candidates.append(right_cand)
                    used.add(right_idx)
            
            for i, cand in enumerate(all_figure_candidates):
                if cand["column"] == "full" and i not in used:
                    new_candidates.append(cand)
            
            all_figure_candidates = new_candidates
        
        debug_print(f"[DEBUG] 左右クラスタマージ完了: {len(all_figure_candidates)}個のクラスタ")
        return all_figure_candidates

    def _fig_find_caption_below(
        self, graphics_bbox: Tuple, text_lines: List[Dict]
    ) -> Optional[Dict]:
        """図の下部付近にあるキャプション行を探す"""
        caption_pattern = re.compile(r'^(図|表)\s*\d+')
        best_caption = None
        best_y = float('inf')
        
        search_y_start = (graphics_bbox[1] + graphics_bbox[3]) / 2
        
        for line in text_lines:
            line_bbox = line["bbox"]
            line_text = line["text"].strip()
            
            if not caption_pattern.match(line_text):
                continue
            
            if line_bbox[1] < search_y_start:
                continue
            
            x_overlap = max(0, min(graphics_bbox[2], line_bbox[2]) - max(graphics_bbox[0], line_bbox[0]))
            if x_overlap < 20:
                continue
            
            if line_bbox[1] < best_y:
                best_y = line_bbox[1]
                best_caption = line
        
        return best_caption

    def _fig_find_body_text_above(
        self, graphics_bbox: Tuple, text_lines: List[Dict], col_width: float
    ) -> Optional[Dict]:
        """図の上部付近にある本文行を探す"""
        best_body = None
        best_y = 0
        
        search_y_end = (graphics_bbox[1] + graphics_bbox[3]) / 2
        
        for line in text_lines:
            line_bbox = line["bbox"]
            line_text = line["text"].strip()
            line_width = line_bbox[2] - line_bbox[0]
            
            if line_bbox[3] > search_y_end:
                continue
            
            x_overlap = max(0, min(graphics_bbox[2], line_bbox[2]) - max(graphics_bbox[0], line_bbox[0]))
            if x_overlap < 20:
                continue
            
            if self._fig_is_body_text_line(line_text, line_width, col_width):
                if line_bbox[3] > best_y:
                    best_y = line_bbox[3]
                    best_body = line
        
        return best_body

    def _fig_compute_clip_bbox(
        self, graphics_bbox: Tuple, text_lines: List[Dict], col_width: float,
        page_width: float, page_height: float, column: str, gutter_x: float,
        is_embedded_image: bool = False,
        header_y_max: Optional[float] = None, footer_y_min: Optional[float] = None,
        is_slide_document: bool = False,
        is_two_column: bool = True,
        raw_graphics_bbox: Optional[Tuple] = None
    ) -> Tuple:
        """graphics_bboxからclip_bboxを計算（トリム処理）"""
        # スライド文書では小さなマージン（5px）、通常文書では20px
        if is_slide_document:
            padding = 5.0
        elif is_embedded_image:
            padding = 1.0
        else:
            padding = 20.0
        
        clip_x0 = max(0, graphics_bbox[0] - padding)
        clip_y0 = max(0, graphics_bbox[1] - padding)
        clip_x1 = min(page_width, graphics_bbox[2] + padding)
        clip_y1 = min(page_height, graphics_bbox[3] + padding)
        
        # スライド文書の場合、raw_graphics_bboxを基準にclip_y1を制限
        # これにより、クラスタリングで拡張されたbboxがフッタ領域を含まないようにする
        if is_slide_document and raw_graphics_bbox:
            # raw_graphics_bboxの下端 + パディング + マージン（10px）を上限とする
            max_clip_y1 = raw_graphics_bbox[3] + padding + 10.0
            if clip_y1 > max_clip_y1:
                debug_print(f"[DEBUG] スライド文書: clip_y1を{clip_y1:.1f}→{max_clip_y1:.1f}に制限（フッタ除外）")
                clip_y1 = max_clip_y1
        
        # スライド文書の場合、本文行と交差する場合は上端を詰める
        # raw_graphics_bbox（パディング前の元のbbox）を使用して判定
        if is_slide_document:
            ref_bbox = raw_graphics_bbox if raw_graphics_bbox else graphics_bbox
            for line in text_lines:
                line_bbox = line.get("bbox", (0, 0, 0, 0))
                line_text = line.get("text", "").strip()
                line_width = line_bbox[2] - line_bbox[0]
                # 本文らしい行かどうかを判定
                if not self._fig_is_body_text_line(line_text, line_width, col_width):
                    continue
                # 本文行が図形の上端より上にある場合、clip_y0を詰める
                # ref_bbox[1]（パディング前の図形上端）を基準に判定
                if line_bbox[3] > clip_y0 and line_bbox[1] < ref_bbox[1]:
                    # 本文行の下端より下にclip_y0を設定
                    new_clip_y0 = line_bbox[3] + 2.0
                    if new_clip_y0 < ref_bbox[1]:
                        debug_print(f"[DEBUG] スライド文書: 本文行を避けてclip_y0を{clip_y0:.1f}→{new_clip_y0:.1f}に調整")
                        clip_y0 = new_clip_y0
        
        if header_y_max is not None and clip_y0 < header_y_max:
            clip_y0 = header_y_max
            debug_print(f"[DEBUG] ヘッダー領域クリップ: clip_y0を{clip_y0:.1f}に設定")
        if footer_y_min is not None and clip_y1 > footer_y_min:
            clip_y1 = footer_y_min
            debug_print(f"[DEBUG] フッター領域クリップ: clip_y1を{clip_y1:.1f}に設定")
        
        # 2段組み文書の場合のみカラム制約を適用
        # 単一カラム文書では図形が中央付近に配置されることがあるため、制約を適用しない
        if is_two_column:
            if column == "left":
                clip_x1 = min(clip_x1, gutter_x - 5)
                debug_print(f"[DEBUG] 左カラム: clip_x1を{clip_x1:.1f}にクランプ")
            elif column == "right":
                old_clip_x0 = clip_x0
                clip_x0 = max(clip_x0, gutter_x + 5)
                debug_print(f"[DEBUG] 右カラム: clip_x0を{old_clip_x0:.1f}→{clip_x0:.1f}にクランプ")
        
        # スライド文書または埋め込み画像の場合、上下トリムをスキップ
        if is_embedded_image or is_slide_document:
            debug_print(f"[DEBUG] {'スライド文書' if is_slide_document else '埋め込み画像'}: 上下トリムをスキップ")
            
            # スライド文書の場合、図形内テキスト（本文ではないテキスト）を含めるようにclip_bboxを拡張
            if is_slide_document:
                ref_bbox = raw_graphics_bbox if raw_graphics_bbox else graphics_bbox
                for line in text_lines:
                    line_bbox = line.get("bbox", (0, 0, 0, 0))
                    line_text = line.get("text", "").strip()
                    line_width = line_bbox[2] - line_bbox[0]
                    
                    # 本文行はスキップ（図形に含めない）
                    if self._fig_is_body_text_line(line_text, line_width, col_width):
                        continue
                    
                    # フッタ領域のテキストはスキップ（clip_y1より下）
                    if line_bbox[1] >= clip_y1:
                        continue
                    
                    # ヘッダ領域のテキストはスキップ（clip_y0より上）
                    if line_bbox[3] <= clip_y0:
                        continue
                    
                    # テキストが図形領域内またはその近くにあるかチェック
                    # Y方向: 図形の上端から下端の範囲内（厳密に）
                    if line_bbox[1] < ref_bbox[1] - 10 or line_bbox[3] > ref_bbox[3]:
                        continue
                    
                    # X方向: 図形と重なっているか、図形の右側にある
                    if line_bbox[2] < ref_bbox[0] - 10:
                        continue
                    
                    # 図形内テキストを含めるようにclip_bboxを拡張
                    if line_bbox[2] > clip_x1:
                        debug_print(f"[DEBUG] 図形内テキスト検出: clip_x1を{clip_x1:.1f}→{line_bbox[2] + padding:.1f}に拡張（テキスト: {line_text[:20]}）")
                        clip_x1 = min(page_width, line_bbox[2] + padding)
                    if line_bbox[0] < clip_x0:
                        debug_print(f"[DEBUG] 図形内テキスト検出: clip_x0を{clip_x0:.1f}→{line_bbox[0] - padding:.1f}に拡張（テキスト: {line_text[:20]}）")
                        clip_x0 = max(0, line_bbox[0] - padding)
            
            return (clip_x0, clip_y0, clip_x1, clip_y1)
        
        caption = self._fig_find_caption_below(graphics_bbox, text_lines)
        if caption:
            caption_y0 = caption["bbox"][1]
            new_clip_y1 = caption_y0 - 5.0
            clip_y1 = min(clip_y1, new_clip_y1)
            debug_print(f"[DEBUG] キャプション検出: clip_y1を{clip_y1:.1f}にトリム")
        
        body_above = self._fig_find_body_text_above(graphics_bbox, text_lines, col_width)
        if body_above:
            body_y1 = body_above["bbox"][3]
            new_clip_y0 = body_y1 + 5.0
            clip_y0 = max(clip_y0, new_clip_y0)
            debug_print(f"[DEBUG] 本文検出: clip_y0を{clip_y0:.1f}にトリム")
        
        # 見出しテキストを検出してclip_y0を調整
        # 見出しパターン: "1. ", "2. ", "第1条" など
        heading_pattern = re.compile(r'^(\d+\.\s+|第[一二三四五六七八九十0-9]+)')
        # clip_bboxの上端付近（padding分上）から図形の上端までの範囲で見出しを検索
        search_y_start = clip_y0 - padding - 10  # paddingより少し上から検索
        search_y_end = graphics_bbox[1] + 10
        for line in text_lines:
            line_bbox = line.get("bbox", (0, 0, 0, 0))
            line_text = line.get("text", "").strip()
            # 見出しパターンに一致し、clip_bbox付近にある場合
            if heading_pattern.match(line_text):
                # 見出しがclip_bboxの上端付近から図形の上端の間にあるかチェック
                if line_bbox[1] >= search_y_start and line_bbox[3] <= search_y_end:
                    # X方向のオーバーラップを確認
                    x_overlap = max(0, min(graphics_bbox[2], line_bbox[2]) - max(graphics_bbox[0], line_bbox[0]))
                    if x_overlap > 20:
                        # 見出しの下端より下にclip_y0を設定
                        new_clip_y0 = line_bbox[3] + 5.0
                        clip_y0 = max(clip_y0, new_clip_y0)
                        debug_print(f"[DEBUG] 見出し検出: clip_y0を{clip_y0:.1f}にトリム（見出し: {line_text[:20]}）")
        
        if clip_y1 - clip_y0 < 50:
            center_y = (clip_y0 + clip_y1) / 2
            clip_y0 = center_y - 25
            clip_y1 = center_y + 25
            debug_print(f"[DEBUG] 最小高さ確保: clip_y0={clip_y0:.1f}, clip_y1={clip_y1:.1f}")
        
        return (clip_x0, clip_y0, clip_x1, clip_y1)

    def _fig_render_candidates(
        self, page, page_num: int, all_figure_candidates: List[Dict],
        page_text_lines: List[Dict], col_width: float, page_width: float,
        gutter_x: float, header_y_max: Optional[float], footer_y_min: Optional[float],
        line_based_table_bboxes: List[Tuple], is_slide_document: bool = False,
        is_two_column: bool = True
    ) -> List[Dict]:
        """図候補をレンダリングして図情報リストを生成"""
        figures = []
        page_height = page.rect.height
        page_area = page_width * page_height
        
        # スライド文書用: ページ全体の本文テキスト文字数を事前計算
        total_body_text_chars = 0
        if is_slide_document:
            for line in page_text_lines:
                line_text = line.get("text", "").strip()
                line_bbox = line.get("bbox", (0, 0, 0, 0))
                line_width = line_bbox[2] - line_bbox[0]
                if self._fig_is_body_text_line(line_text, line_width, col_width):
                    total_body_text_chars += len(line_text)
        
        for fig_info in all_figure_candidates:
            try:
                is_embedded_image = fig_info.get("is_embedded", False)
                is_table_image = fig_info.get("is_table_image", False)
                
                if is_embedded_image:
                    graphics_bbox = fig_info.get("raw_union_bbox", fig_info["union_bbox"])
                else:
                    graphics_bbox = fig_info["union_bbox"]
                
                # スライド文書の場合、raw_union_bboxを保持して本文テキスト除外に使用
                raw_graphics_bbox = fig_info.get("raw_union_bbox", graphics_bbox)
                
                column = fig_info["column"]
                union_bbox = graphics_bbox
                
                if is_table_image and "clip_bbox" in fig_info:
                    clip_bbox = fig_info["clip_bbox"]
                    # 表画像でも本文テキストを除外するロジックを適用
                    body_line = self._fig_find_body_text_above(
                        clip_bbox, page_text_lines, col_width
                    )
                    if body_line:
                        body_bottom = body_line["bbox"][3]
                        new_clip_y0 = body_bottom + 5.0
                        if new_clip_y0 > clip_bbox[1] and new_clip_y0 < clip_bbox[3]:
                            debug_print(f"[DEBUG] 表画像: 本文検出によりclip_y0を{clip_bbox[1]:.1f}→{new_clip_y0:.1f}にトリム")
                            clip_bbox = (clip_bbox[0], new_clip_y0, clip_bbox[2], clip_bbox[3])
                    debug_print(f"[DEBUG] 表画像: clip_bboxを使用")
                else:
                    # 画像のみクラスタ（drawing_count=0）の場合の処理
                    drawing_count_for_clip = fig_info.get("drawing_count", 0)
                    image_bboxes_for_clip = fig_info.get("image_bboxes", [])
                    significant_image_count_for_clip = fig_info.get("significant_image_count", 0)
                    image_count_for_clip = fig_info.get("image_count", 0)
                    
                    # スライド文書で、画像のみクラスタかつ有意な画像がない場合
                    # 埋め込み画像を直接抽出して出力（テキストを巻き込まない）
                    # ただし、タイル画像（表示サイズが元画像サイズとほぼ同じ）の場合は
                    # ページレンダリングで図の領域を切り出す
                    if is_slide_document and drawing_count_for_clip == 0:
                        has_significant = significant_image_count_for_clip > 0 or image_count_for_clip >= 3
                        if not has_significant and image_bboxes_for_clip:
                            page_images = page.get_images(full=True)
                            page_img_infos = page.get_image_info()
                            
                            # タイル画像かどうかを判定
                            # 表示サイズが元画像サイズとほぼ同じ場合はタイル画像
                            has_tile_image = False
                            for img_bbox in image_bboxes_for_clip:
                                display_w = img_bbox[2] - img_bbox[0]
                                display_h = img_bbox[3] - img_bbox[1]
                                for pi, pimg in enumerate(page_images):
                                    if pi < len(page_img_infos):
                                        info = page_img_infos[pi]
                                        info_bbox = info.get("bbox", (0, 0, 0, 0))
                                        if (abs(info_bbox[0] - img_bbox[0]) < 1 and
                                            abs(info_bbox[1] - img_bbox[1]) < 1):
                                            orig_w = info.get("width", 0)
                                            orig_h = info.get("height", 0)
                                            # 表示サイズと元サイズの比率を確認
                                            if orig_w > 0 and orig_h > 0:
                                                w_ratio = display_w / orig_w
                                                h_ratio = display_h / orig_h
                                                # 比率が0.8〜1.2の範囲ならタイル画像
                                                if 0.8 <= w_ratio <= 1.2 and 0.8 <= h_ratio <= 1.2:
                                                    has_tile_image = True
                                                    debug_print(f"[DEBUG] page={page_num+1}: タイル画像を検出 ratio=({w_ratio:.2f}, {h_ratio:.2f})")
                                            break
                            
                            # タイル画像がある場合はグレー領域を自動検出してレンダリング
                            # テキストを巻き込まないように、グレー領域の境界を検出
                            if has_tile_image:
                                debug_print(f"[DEBUG] page={page_num+1}: タイル画像のためグレー領域を自動検出")
                                
                                # ページをレンダリングしてグレー領域の境界を検出
                                detect_mat = fitz.Matrix(1.0, 1.0)
                                detect_pix = page.get_pixmap(matrix=detect_mat)
                                detect_w = detect_pix.width
                                detect_h = detect_pix.height
                                detect_n = detect_pix.n
                                detect_samples = detect_pix.samples
                                
                                # 行ごとのグレーピクセル比率を計算
                                def row_gray_ratio(y_pos):
                                    gray_cnt = 0
                                    for x_pos in range(detect_w):
                                        idx = (y_pos * detect_w + x_pos) * detect_n
                                        r_val = detect_samples[idx]
                                        g_val = detect_samples[idx + 1]
                                        b_val = detect_samples[idx + 2]
                                        if abs(r_val - g_val) < 15 and abs(g_val - b_val) < 15 and 100 < r_val < 220:
                                            gray_cnt += 1
                                    return gray_cnt / detect_w
                                
                                # y方向のグレー領域を検出（比率が0.5以上の連続区間）
                                gray_y0, gray_y1 = None, None
                                in_gray_region = False
                                for y_pos in range(0, detect_h, 5):
                                    ratio = row_gray_ratio(y_pos)
                                    if ratio > 0.5 and not in_gray_region:
                                        gray_y0 = y_pos
                                        in_gray_region = True
                                    elif ratio < 0.3 and in_gray_region:
                                        gray_y1 = y_pos
                                        break
                                if in_gray_region and gray_y1 is None:
                                    gray_y1 = detect_h
                                
                                # 列ごとのグレーピクセル比率を計算（帯の範囲内で）
                                def col_gray_ratio_in_band(x_pos, band_y0, band_y1):
                                    gray_cnt = 0
                                    band_height = band_y1 - band_y0
                                    if band_height <= 0:
                                        return 0
                                    for y_pos in range(band_y0, band_y1):
                                        idx = (y_pos * detect_w + x_pos) * detect_n
                                        r_val = detect_samples[idx]
                                        g_val = detect_samples[idx + 1]
                                        b_val = detect_samples[idx + 2]
                                        if abs(r_val - g_val) < 15 and abs(g_val - b_val) < 15 and 100 < r_val < 220:
                                            gray_cnt += 1
                                    return gray_cnt / band_height
                                
                                # x方向のグレー領域を検出（帯の範囲内で、比率が0.5以上の連続区間）
                                gray_x0, gray_x1 = None, None
                                in_gray_region = False
                                if gray_y0 is not None and gray_y1 is not None:
                                    for x_pos in range(0, detect_w, 5):
                                        ratio = col_gray_ratio_in_band(x_pos, gray_y0, gray_y1)
                                        if ratio > 0.5 and not in_gray_region:
                                            gray_x0 = x_pos
                                            in_gray_region = True
                                        elif ratio < 0.3 and in_gray_region:
                                            gray_x1 = x_pos
                                            break
                                    if in_gray_region and gray_x1 is None:
                                        gray_x1 = detect_w
                                
                                # グレー領域が検出された場合のみレンダリング
                                if gray_x0 is not None and gray_y0 is not None:
                                    debug_print(f"[DEBUG] page={page_num+1}: グレー領域検出 ({gray_x0}, {gray_y0}, {gray_x1}, {gray_y1})")
                                    
                                    self.image_counter += 1
                                    image_filename = f"{self.base_name}_fig_{page_num + 1:03d}_{self.image_counter:03d}"
                                    
                                    # グレー領域のみをレンダリング
                                    clip_rect = fitz.Rect(gray_x0, gray_y0, gray_x1, gray_y1)
                                    mat = fitz.Matrix(2.0, 2.0)
                                    pix = page.get_pixmap(matrix=mat, clip=clip_rect)
                                    
                                    if self.output_format == 'svg':
                                        temp_img = os.path.join(self.images_dir, f"temp_fig_{self.image_counter}.png")
                                        pix.save(temp_img)
                                        image_path = os.path.join(self.images_dir, f"{image_filename}.svg")
                                        self._convert_png_to_svg(temp_img, image_path)
                                        if os.path.exists(temp_img):
                                            os.remove(temp_img)
                                    else:
                                        image_path = os.path.join(self.images_dir, f"{image_filename}.png")
                                        pix.save(image_path)
                                    
                                    figures.append({
                                        "path": image_path,
                                        "filename": os.path.basename(image_path),
                                        "bbox": (gray_x0, gray_y0, gray_x1, gray_y1),
                                        "y_position": gray_y0,
                                        "texts": [],
                                        "column": column
                                    })
                                    debug_print(f"[DEBUG] page={page_num+1}: グレー領域をレンダリング {image_filename}")
                                    continue
                            else:
                                # 各埋め込み画像を個別に抽出
                                extracted_any = False
                                for img_bbox in image_bboxes_for_clip:
                                    img_area = (img_bbox[2] - img_bbox[0]) * (img_bbox[3] - img_bbox[1])
                                    if img_area < page_area * 0.005:
                                        debug_print(f"[DEBUG] page={page_num+1}: 小さな画像を除外（面積比={img_area/page_area:.2%}）")
                                        continue
                                    
                                    xref = None
                                    smask = 0
                                    for pi, pimg in enumerate(page_images):
                                        if pi < len(page_img_infos):
                                            info_bbox = page_img_infos[pi].get("bbox", (0, 0, 0, 0))
                                            if (abs(info_bbox[0] - img_bbox[0]) < 1 and
                                                abs(info_bbox[1] - img_bbox[1]) < 1 and
                                                abs(info_bbox[2] - img_bbox[2]) < 1 and
                                                abs(info_bbox[3] - img_bbox[3]) < 1):
                                                xref = pimg[0]
                                                smask = pimg[1] if len(pimg) > 1 else 0
                                                break
                                    
                                    if xref is None:
                                        debug_print(f"[DEBUG] page={page_num+1}: 画像のxrefが見つからない bbox={img_bbox}")
                                        continue
                                    
                                    self.image_counter += 1
                                    image_filename = f"{self.base_name}_fig_{page_num + 1:03d}_{self.image_counter:03d}"
                                    
                                    try:
                                        doc = page.parent
                                        
                                        if smask > 0:
                                            pix = fitz.Pixmap(doc, xref)
                                            mask_pix = fitz.Pixmap(doc, smask)
                                            if pix.alpha == 0:
                                                pix = fitz.Pixmap(pix, 1)
                                            pix.set_alpha(mask_pix.samples)
                                            
                                            if self.output_format == 'svg':
                                                temp_img = os.path.join(self.images_dir, f"temp_fig_{self.image_counter}.png")
                                                pix.save(temp_img)
                                                image_path = os.path.join(self.images_dir, f"{image_filename}.svg")
                                                self._convert_png_to_svg(temp_img, image_path)
                                                if os.path.exists(temp_img):
                                                    os.remove(temp_img)
                                            else:
                                                image_path = os.path.join(self.images_dir, f"{image_filename}.png")
                                                pix.save(image_path)
                                        else:
                                            base_image = doc.extract_image(xref)
                                            img_ext = base_image["ext"]
                                            img_data = base_image["image"]
                                            
                                            if self.output_format == 'svg':
                                                temp_img = os.path.join(self.images_dir, f"temp_fig_{self.image_counter}.{img_ext}")
                                                with open(temp_img, "wb") as f:
                                                    f.write(img_data)
                                                image_path = os.path.join(self.images_dir, f"{image_filename}.svg")
                                                self._convert_png_to_svg(temp_img, image_path)
                                                if os.path.exists(temp_img):
                                                    os.remove(temp_img)
                                            else:
                                                image_path = os.path.join(self.images_dir, f"{image_filename}.{img_ext}")
                                                with open(image_path, "wb") as f:
                                                    f.write(img_data)
                                        
                                        figures.append({
                                            "path": image_path,
                                            "filename": os.path.basename(image_path),
                                            "bbox": img_bbox,
                                            "y_position": img_bbox[1],
                                            "texts": [],
                                            "column": column
                                        })
                                        extracted_any = True
                                        debug_print(f"[DEBUG] page={page_num+1}: 埋め込み画像を直接抽出 {image_filename}")
                                    except Exception as e:
                                        debug_print(f"[DEBUG] page={page_num+1}: 画像抽出エラー xref={xref}: {e}")
                                        continue
                                if extracted_any:
                                    continue
                    
                    clip_bbox = self._fig_compute_clip_bbox(
                        graphics_bbox, page_text_lines, col_width,
                        page_width, page_height, column, gutter_x,
                        is_embedded_image, header_y_max, footer_y_min,
                        is_slide_document, is_two_column,
                        raw_graphics_bbox=raw_graphics_bbox
                    )
                
                # スライド文書: ページ全体を覆う図形（面積比50%以上）かつ本文テキスト比50%以上の場合のみ除外
                # 図中ラベル（短いテキスト）は本文としてカウントしない
                # ただし、埋め込み画像が含まれるクラスタは除外しない
                # 条件1: 有意な埋め込み画像（ページ面積の5%以上）が1個以上
                # 条件2: 埋め込み画像が3個以上（小さな画像でも複数あれば図形として扱う）
                significant_image_count = fig_info.get("significant_image_count", 0)
                image_count = fig_info.get("image_count", 0)
                has_significant_images = significant_image_count > 0 or image_count >= 3
                if is_slide_document and total_body_text_chars > 0 and not has_significant_images:
                    # 図形の面積比を計算
                    page_area = page_width * page_height
                    fig_area = (clip_bbox[2] - clip_bbox[0]) * (clip_bbox[3] - clip_bbox[1])
                    area_ratio = fig_area / page_area if page_area > 0 else 0
                    
                    # 面積比が50%以上の大きな図形のみ本文テキスト比チェック
                    if area_ratio >= 0.5:
                        fig_body_text_chars = 0
                        for line in page_text_lines:
                            line_bbox = line.get("bbox", (0, 0, 0, 0))
                            line_text = line.get("text", "").strip()
                            line_width = line_bbox[2] - line_bbox[0]
                            # 行がclip_bbox内に含まれているかチェック
                            if (line_bbox[0] >= clip_bbox[0] - 5 and line_bbox[2] <= clip_bbox[2] + 5 and
                                line_bbox[1] >= clip_bbox[1] - 5 and line_bbox[3] <= clip_bbox[3] + 5):
                                # 本文らしい行のみカウント
                                if self._fig_is_body_text_line(line_text, line_width, col_width):
                                    fig_body_text_chars += len(line_text)
                        
                        body_text_ratio = fig_body_text_chars / total_body_text_chars
                        debug_print(f"[DEBUG] page={page_num+1}: 大きな図形（面積比={area_ratio:.1%}）の本文テキスト比={body_text_ratio:.1%}")
                        if body_text_ratio >= 0.5:
                            debug_print(f"[DEBUG] page={page_num+1}: ページ全体を覆う図形を除外（面積比={area_ratio:.1%}, 本文テキスト比={body_text_ratio:.1%}）")
                            continue
                
                # 小さすぎる図形を除外（高さが60ピクセル未満）
                # 表のヘッダーラベルなど、図として出力すべきでない小さな要素を除外
                clip_height = clip_bbox[3] - clip_bbox[1]
                clip_width = clip_bbox[2] - clip_bbox[0]
                if clip_height < 60:
                    debug_print(f"[DEBUG] page={page_num+1}: 小さすぎる図形を除外（高さ={clip_height:.1f}px < 60px）")
                    continue
                
                # 幅が狭すぎる図形も除外（幅が100ピクセル未満）
                if clip_width < 100:
                    debug_print(f"[DEBUG] page={page_num+1}: 幅が狭すぎる図形を除外（幅={clip_width:.1f}px < 100px）")
                    continue
                
                self.image_counter += 1
                image_filename = f"{self.base_name}_fig_{page_num + 1:03d}_{self.image_counter:03d}"
                
                debug_print(f"[DEBUG] 図候補出力: page={page_num+1}, column={column}, graphics_bbox={graphics_bbox}, clip_bbox={clip_bbox}")
                
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
                
                # 図形のbboxが罫線ベースの表領域と重なる場合は、表領域除外をしない
                # （図形として出力する表のテキストを抽出するため）
                figure_overlaps_table = False
                for table_bbox in line_based_table_bboxes:
                    # 図形と表の重なりを計算
                    overlap_x0 = max(clip_bbox[0], table_bbox[0])
                    overlap_y0 = max(clip_bbox[1], table_bbox[1])
                    overlap_x1 = min(clip_bbox[2], table_bbox[2])
                    overlap_y1 = min(clip_bbox[3], table_bbox[3])
                    if overlap_x0 < overlap_x1 and overlap_y0 < overlap_y1:
                        overlap_area = (overlap_x1 - overlap_x0) * (overlap_y1 - overlap_y0)
                        table_area = (table_bbox[2] - table_bbox[0]) * (table_bbox[3] - table_bbox[1])
                        # 表の50%以上が図形に含まれている場合
                        if table_area > 0 and overlap_area / table_area > 0.5:
                            figure_overlaps_table = True
                            break
                
                exclude_tables = [] if figure_overlaps_table else line_based_table_bboxes
                # スライド文書の場合はラベル拡張をしない（clip_bbox内のテキストのみ抽出）
                expand_labels = not is_slide_document
                # 埋め込み画像bboxを取得（スライド文書用の本文フィルタリング）
                # 画像のみクラスタ（drawing_count=0）の場合は本文テキストを含めない
                # これにより、ページ27のような画像のみのクラスタで本文が巻き込まれるのを防ぐ
                drawing_count = fig_info.get("drawing_count", 0)
                if drawing_count == 0:
                    # 画像のみクラスタの場合、image_bboxesを空にして本文フィルタリングを無効化
                    image_bboxes = []
                else:
                    image_bboxes = fig_info.get("image_bboxes", [])
                
                # 表画像の場合はテキスト抽出を無効化（テキストは表の一部としてレンダリングされている）
                if is_table_image:
                    figure_texts = []
                    expanded_bbox = clip_bbox
                    debug_print(f"[DEBUG] 表画像: テキスト抽出を無効化")
                else:
                    figure_texts, expanded_bbox = self._extract_text_in_bbox(
                        page, clip_bbox, expand_for_labels=expand_labels, column=column, gutter_x=gutter_x,
                        exclude_table_bboxes=exclude_tables, image_bboxes=image_bboxes,
                        is_slide_document=is_slide_document
                    )
                
                # スライド文書: 本文テキストがない場合は除外しない（図中ラベルのみのページ）
                # 本文テキストがある場合のみ、抽出されたテキスト量をチェック
                # この除外ロジックは無効化（フェーズ7.5とレンダリング前チェックで十分）
                
                figures.append({
                    "path": image_path,
                    "filename": os.path.basename(image_path),
                    "bbox": expanded_bbox,
                    "y_position": union_bbox[1],
                    "texts": figure_texts,
                    "column": column,
                    "is_table_image": is_table_image
                })
                
                debug_print(f"[DEBUG] 図を抽出: {image_path} ({fig_info['cluster_size']}要素, {len(figure_texts)}テキスト, {column})")
                
            except Exception as e:
                debug_print(f"[DEBUG] 図抽出エラー: {e}")
                continue
        
        figures.sort(key=lambda x: x["y_position"])
        return figures

    def _extract_all_figures(
        self, page, page_num: int, header_footer_patterns: Set[str] = None,
        is_slide_document: bool = False
    ) -> List[Dict[str, Any]]:
        """ベクター図形と埋め込み画像を統合して図を抽出（オーケストレータ）
        
        ベクター描画と埋め込み画像を統合してクラスタリングし、
        クラスタリング後にカラム判定を行う（先にクラスタリング、後でカラム判定）。
        ヘッダー/フッター領域内の図クラスタは除外する。
        スライド文書の場合は小さな装飾要素を除外する。
        
        Args:
            page: PyMuPDFのページオブジェクト
            page_num: ページ番号
            header_footer_patterns: ヘッダ・フッタパターンのセット
            is_slide_document: スライド文書フラグ
            
        Returns:
            抽出された図の情報リスト
        """
        if header_footer_patterns is None:
            header_footer_patterns = set()
        
        page_width = page.rect.width
        page_height = page.rect.height
        gutter_x = page_width / 2
        gutter_margin = 10.0
        
        # フェーズ1: ヘッダー/フッター領域計算
        header_y_max, footer_y_min = self._fig_compute_header_footer_bounds(
            page, page_num, page_height, header_footer_patterns
        )
        
        # 罫線ベースの表を検出
        line_based_table_bboxes = []
        try:
            tables = page.find_tables()
            if tables.tables:
                for table in tables.tables:
                    bbox = table.bbox
                    rows = table.extract()
                    if rows and len(rows) >= 2 and len(rows[0]) >= 2:
                        line_based_table_bboxes.append(bbox)
                        debug_print(f"[DEBUG] page={page_num+1}: 罫線ベース表検出")
        except Exception as e:
            debug_print(f"[DEBUG] find_tables()エラー: {e}")
        
        # フェーズ2: 描画要素と画像の収集
        all_elements, all_bboxes = self._fig_collect_graphics_elements(page, page_num)
        
        if len(all_bboxes) == 0:
            return []
        
        # フェーズ3: 図キャプション検出
        figure_caption_lines = self._fig_detect_figure_captions(page)
        
        # 各bboxのカラムを事前計算
        def get_bbox_column(bbox):
            x0, y0, x1, y1 = bbox
            center_x = (x0 + x1) / 2
            crosses = x0 < gutter_x - gutter_margin and x1 > gutter_x + gutter_margin
            if crosses:
                return "full"
            elif center_x < gutter_x:
                return "left"
            else:
                return "right"
        
        bbox_columns = [get_bbox_column(b) for b in all_bboxes]
        
        # フェーズ4: クラスタリングと候補生成
        all_figure_candidates, _ = self._fig_cluster_elements(
            all_bboxes, all_elements, bbox_columns, figure_caption_lines,
            page_width, page_height, page_num, header_y_max, footer_y_min,
            gutter_x, gutter_margin, is_slide_document
        )
        
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
        
        # フェーズ5: 表領域フィルタリング
        table_bboxes = self._fig_detect_table_bboxes_from_text(page_text_lines, page_width)
        
        if table_bboxes:
            debug_print(f"[DEBUG] page={page_num+1}: 表領域を{len(table_bboxes)}個検出")
        
        all_figure_candidates = self._fig_filter_table_regions(
            all_figure_candidates, table_bboxes, page_text_lines, column_count, page_num
        )
        
        # フェーズ5.5: 囲み記事（テキストボックス）のフィルタリング
        # マージ処理の前に囲み記事を除外する（マージで画像と結合されるのを防ぐ）
        col_width = page_width / 2
        filtered_candidates = []
        for cand in all_figure_candidates:
            if self._fig_is_text_box_candidate(cand, page_text_lines, col_width):
                debug_print(f"[DEBUG] page={page_num+1}: 囲み記事候補を除外")
                continue
            filtered_candidates.append(cand)
        all_figure_candidates = filtered_candidates
        
        # フェーズ5.6: スライド文書での小さな装飾要素のフィルタリング
        # スライド文書の場合、ページ面積の5%未満の小さな図形を除外
        # ただし、埋め込み画像のみのクラスタ（drawing_count=0）は除外しない
        if is_slide_document and all_figure_candidates:
            page_area = page_width * page_height
            min_area_threshold = page_area * 0.05
            slide_filtered = []
            for cand in all_figure_candidates:
                bbox = cand.get("union_bbox", (0, 0, 0, 0))
                cand_area = (bbox[2] - bbox[0]) * (bbox[3] - bbox[1])
                drawing_count = cand.get("drawing_count", 0)
                # 埋め込み画像のみのクラスタは除外しない（意図的に配置された画像）
                if drawing_count == 0:
                    slide_filtered.append(cand)
                    continue
                if cand_area < min_area_threshold:
                    debug_print(f"[DEBUG] page={page_num+1}: スライド装飾要素を除外（面積比={cand_area/page_area:.2%}）")
                    continue
                slide_filtered.append(cand)
            all_figure_candidates = slide_filtered
        
        if not all_figure_candidates:
            return []
        
        # フェーズ6: 同一カラム内マージ
        all_figure_candidates = self._fig_merge_same_column(
            all_figure_candidates, page_text_lines, col_width
        )
        
        # フェーズ7: 左右クラスタマージ
        all_figure_candidates = self._fig_merge_left_right(
            all_figure_candidates, page_text_lines, col_width, gutter_x
        )
        
        if not all_figure_candidates:
            return []
        
        # フェーズ7.5: スライド文書での図形内テキスト量による除外
        # マージ後の図形に対して、面積比70%以上かつ「本文テキスト比」50%以上の場合のみ除外
        # 図中ラベル（短いテキスト）は本文としてカウントしない
        if is_slide_document and all_figure_candidates:
            page_area = page_width * page_height
            # 本文テキストのみをカウント（図中ラベルは除外）
            total_body_text_chars = 0
            for line in page_text_lines:
                line_text = line.get("text", "").strip()
                line_bbox = line.get("bbox", (0, 0, 0, 0))
                line_width = line_bbox[2] - line_bbox[0]
                if self._fig_is_body_text_line(line_text, line_width, col_width):
                    total_body_text_chars += len(line_text)
            debug_print(f"[DEBUG] page={page_num+1}: フェーズ7.5開始 - 図形候補数={len(all_figure_candidates)}, 本文テキスト文字数={total_body_text_chars}")
            slide_text_filtered = []
            for cand_idx, cand in enumerate(all_figure_candidates):
                bbox = cand.get("union_bbox", (0, 0, 0, 0))
                fig_area = (bbox[2] - bbox[0]) * (bbox[3] - bbox[1])
                area_ratio = fig_area / page_area if page_area > 0 else 0
                debug_print(f"[DEBUG] page={page_num+1}: 図形候補{cand_idx+1} bbox={bbox}, 面積比={area_ratio:.1%}")
                
                # 面積比が50%未満の図形は除外しない（小さな図形は保持）
                if area_ratio < 0.5:
                    slide_text_filtered.append(cand)
                    continue
                
                # 埋め込み画像が含まれるクラスタは除外しない
                # 条件1: 有意な埋め込み画像（ページ面積の5%以上）が1個以上
                # 条件2: 埋め込み画像が3個以上（小さな画像でも複数あれば図形として扱う）
                significant_image_count = cand.get("significant_image_count", 0)
                image_count = cand.get("image_count", 0)
                if significant_image_count > 0 or image_count >= 3:
                    debug_print(f"[DEBUG] page={page_num+1}: 図形候補{cand_idx+1} 埋め込み画像{image_count}個（有意={significant_image_count}個）を含むため除外しない")
                    slide_text_filtered.append(cand)
                    continue
                
                # 面積比50%以上かつ有意な埋め込み画像なしの大きな図形のみ「本文テキスト比」をチェック
                fig_body_text_chars = 0
                for line in page_text_lines:
                    line_bbox = line.get("bbox", (0, 0, 0, 0))
                    line_text = line.get("text", "").strip()
                    line_width = line_bbox[2] - line_bbox[0]
                    # bbox内のテキスト行かどうかを判定
                    if (line_bbox[0] >= bbox[0] - 5 and line_bbox[2] <= bbox[2] + 5 and
                        line_bbox[1] >= bbox[1] - 5 and line_bbox[3] <= bbox[3] + 5):
                        # 本文らしい行のみカウント
                        if self._fig_is_body_text_line(line_text, line_width, col_width):
                            fig_body_text_chars += len(line_text)
                
                if total_body_text_chars > 0:
                    body_text_ratio = fig_body_text_chars / total_body_text_chars
                    debug_print(f"[DEBUG] page={page_num+1}: 図形候補{cand_idx+1} 本文テキスト比={body_text_ratio:.1%} ({fig_body_text_chars}/{total_body_text_chars})")
                    if body_text_ratio >= 0.5:
                        debug_print(f"[DEBUG] page={page_num+1}: ページ全体を覆う図形を除外（面積比={area_ratio:.1%}, 本文テキスト比={body_text_ratio:.1%}）")
                        continue
                else:
                    # 本文テキストがない場合は除外しない（図中ラベルのみのページ）
                    debug_print(f"[DEBUG] page={page_num+1}: 図形候補{cand_idx+1} 本文テキストなし、除外しない")
                
                slide_text_filtered.append(cand)
            all_figure_candidates = slide_text_filtered
        
        # スライド文書で全ての図形候補が除外された場合、個別の埋め込み画像を抽出
        if not all_figure_candidates and is_slide_document:
            debug_print(f"[DEBUG] page={page_num+1}: 図形候補が全て除外されたため、個別の埋め込み画像を抽出")
            individual_images = self._extract_slide_individual_images(
                page, page_num, page_text_lines, header_y_max, footer_y_min
            )
            if individual_images:
                debug_print(f"[DEBUG] page={page_num+1}: 個別画像を{len(individual_images)}個抽出")
                return individual_images
            return []
        
        if not all_figure_candidates:
            return []
        
        debug_print(f"[DEBUG] ページ {page_num + 1}: {len(all_bboxes)}個の要素を{len(all_figure_candidates)}個の図にグループ化")
        
        # フェーズ8: 画像レンダリング
        # 2段組みかどうかを判定（column_count >= 2 の場合は2段組み）
        is_two_column = column_count >= 2
        figures = self._fig_render_candidates(
            page, page_num, all_figure_candidates, page_text_lines, col_width,
            page_width, gutter_x, header_y_max, footer_y_min, line_based_table_bboxes,
            is_slide_document, is_two_column
        )
        
        return figures

    def _extract_vector_figures(
        self, page, page_num: int, header_footer_patterns: Set[str] = None,
        is_slide_document: bool = False
    ) -> List[Dict[str, Any]]:
        """ベクタ描画（図）を抽出（統合版を使用）
        
        Args:
            page: PyMuPDFのページオブジェクト
            page_num: ページ番号
            header_footer_patterns: ヘッダ・フッタパターンのセット
            is_slide_document: スライド文書フラグ
            
        Returns:
            抽出された図の情報リスト
        """
        return self._extract_all_figures(page, page_num, header_footer_patterns, is_slide_document)

    def _extract_text_in_bbox(
        self, page, bbox: Tuple[float, float, float, float],
        expand_for_labels: bool = True,
        column: str = None,
        gutter_x: float = None,
        exclude_table_bboxes: List[Tuple[float, float, float, float]] = None,
        image_bboxes: List[Tuple[float, float, float, float]] = None,
        is_slide_document: bool = False
    ) -> Tuple[List[str], Tuple[float, float, float, float]]:
        """指定されたbbox内のテキストを抽出
        
        図のラベルテキストを含めるため、bboxを近傍の短いテキストで拡張する。
        カラム境界を考慮して、隣のカラムのテキストを取り込まないようにする。
        罫線ベースの表領域内のテキストは除外する。
        スライド文書の場合、埋め込み画像bbox外の本文テキストを除外する。
        
        Args:
            page: PyMuPDFのページオブジェクト
            bbox: バウンディングボックス (x0, y0, x1, y1)
            expand_for_labels: ラベルテキストを含めるためにbboxを拡張するか
            column: 図のカラム ("left", "right", "full")
            gutter_x: カラム境界のX座標
            exclude_table_bboxes: 除外する罫線ベースの表領域のリスト
            image_bboxes: 埋め込み画像のbboxリスト（スライド文書用）
            is_slide_document: スライド文書フラグ
            
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
                        
                        # 見出しパターンはラベル候補から除外（expanded_bboxに含めない）
                        # これにより、見出しは本文側に残り、図形内テキストには含まれない
                        if re.match(r'^\d+\.\s+', line_text):
                            continue
                        if re.match(r'^第[一二三四五六七八九十0-9]+', line_text):
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
                        
                        # 見出しパターンをフィルタリング（図形内テキストから除外）
                        # 例: "1. 近接した図形群", "4. 重なり合う図形", "第1条"
                        if re.match(r'^\d+\.\s+', line_text_stripped):
                            continue
                        if re.match(r'^第[一二三四五六七八九十0-9]+', line_text_stripped):
                            continue
                        # 長い説明文も除外（図形内ラベルは通常短い）
                        # 20文字以上で句読点を含む場合、または「。」で終わる場合は説明文とみなす
                        if len(line_text_stripped) > 20 and ('。' in line_text_stripped or '、' in line_text_stripped):
                            continue
                        if line_text_stripped.endswith('。'):
                            continue
                        
                        # スライド文書の場合、本文テキストを除外
                        # 図形内テキストには短いラベルのみを含め、本文は含めない
                        if is_slide_document:
                            line_width = line_bbox[2] - line_bbox[0]
                            col_width = page.rect.width * 0.4
                            is_body = self._fig_is_body_text_line(line_text_stripped, line_width, col_width)
                            if is_body:
                                # 埋め込み画像がある場合、画像bboxの近傍にある本文のみ含める
                                if image_bboxes:
                                    # 本文行が埋め込み画像bboxの近傍にあるかチェック
                                    # 上下方向: 画像bboxを20px拡張
                                    # 左右方向: 画像bboxを50px拡張
                                    near_image = False
                                    for img_bbox in image_bboxes:
                                        if (img_bbox[0] - 50 <= line_center_x <= img_bbox[2] + 50 and
                                            img_bbox[1] - 20 <= line_center_y <= img_bbox[3] + 20):
                                            near_image = True
                                            break
                                    if not near_image:
                                        debug_print(f"[DEBUG] スライド文書: 本文行を除外（画像近傍外）: {line_text_stripped[:30]}...")
                                        continue
                                else:
                                    # 埋め込み画像がない場合、本文テキストは全て除外
                                    debug_print(f"[DEBUG] スライド文書: 本文行を除外（画像なし）: {line_text_stripped[:30]}...")
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
        threshold: float = 50.0
    ) -> List[List[int]]:
        """画像bboxをクラスタリング
        
        Args:
            bboxes: bboxのリスト
            threshold: クラスタリング閾値
            
        Returns:
            クラスタのリスト（各クラスタはbboxインデックスのリスト）
        """
        if not bboxes:
            return []
        
        n = len(bboxes)
        visited = [False] * n
        clusters = []
        
        def boxes_overlap_or_close(idx1, idx2):
            b1, b2 = bboxes[idx1], bboxes[idx2]
            x_gap = max(0, max(b1[0], b2[0]) - min(b1[2], b2[2]))
            y_gap = max(0, max(b1[1], b2[1]) - min(b1[3], b2[3]))
            return x_gap <= threshold and y_gap <= threshold
        
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
                    if boxes_overlap_or_close(current, j):
                        cluster.append(j)
                        visited[j] = True
                        queue.append(j)
            
            clusters.append(cluster)
        
        return clusters

    def _get_cluster_union_bbox(
        self, bboxes: List[Tuple[float, float, float, float]],
        cluster: List[int]
    ) -> Tuple[float, float, float, float]:
        """クラスタのunion bboxを計算
        
        Args:
            bboxes: bboxのリスト
            cluster: クラスタ（bboxインデックスのリスト）
            
        Returns:
            union bbox
        """
        cluster_bboxes = [bboxes[i] for i in cluster]
        x0 = min(b[0] for b in cluster_bboxes)
        y0 = min(b[1] for b in cluster_bboxes)
        x1 = max(b[2] for b in cluster_bboxes)
        y1 = max(b[3] for b in cluster_bboxes)
        return (x0, y0, x1, y1)

    def _extract_embedded_images(
        self, page, page_num: int, header_footer_patterns: Set[str] = None
    ) -> List[Dict[str, Any]]:
        """埋め込み画像を抽出（統合版を使用）
        
        Args:
            page: PyMuPDFのページオブジェクト
            page_num: ページ番号
            header_footer_patterns: ヘッダ・フッタパターンのセット
            
        Returns:
            抽出された画像の情報リスト
        """
        return self._extract_all_figures(page, page_num, header_footer_patterns)

    def _extract_slide_individual_images(
        self, page, page_num: int, page_text_lines: List[Dict],
        header_y_max: Optional[float], footer_y_min: Optional[float]
    ) -> List[Dict[str, Any]]:
        """スライド文書用のフォールバック画像抽出
        
        クラスタリングで大きな図形が除外された場合に発動する。
        有意な埋め込み画像（ページ面積の5%以上）がある場合のみ画像を出力。
        枠線のみのページ（テキストが主体）は画像を出力しない。
        
        Args:
            page: PyMuPDFのページオブジェクト
            page_num: ページ番号
            page_text_lines: ページ内のテキスト行リスト
            header_y_max: ヘッダー領域の下端Y座標
            footer_y_min: フッター領域の上端Y座標
            
        Returns:
            抽出された画像の情報リスト
        """
        page_width = page.rect.width
        page_height = page.rect.height
        page_area = page_width * page_height
        
        # 埋め込み画像の総面積を計算
        total_image_area = 0
        image_count = 0
        try:
            images = page.get_images(full=True)
            for img in images:
                xref = img[0]
                img_rects = page.get_image_rects(xref)
                for rect in img_rects:
                    img_area = (rect[2] - rect[0]) * (rect[3] - rect[1])
                    # 小さなロゴ（1%未満）は除外
                    if img_area >= page_area * 0.01:
                        total_image_area += img_area
                        image_count += 1
        except Exception as e:
            debug_print(f"[DEBUG] page={page_num+1}: 画像確認エラー: {e}")
        
        total_image_ratio = total_image_area / page_area if page_area > 0 else 0
        
        # 画像の総面積が3%未満の場合は画像を出力しない
        # 枠線のみのページ（テキストが主体）は画像を出力しない
        # 閾値を3%に下げて、ページ24のような4.5%の画像も出力できるようにする
        if total_image_ratio < 0.03:
            debug_print(f"[DEBUG] page={page_num+1}: フォールバック - 画像総面積{total_image_ratio:.1%}、画像出力スキップ")
            return []
        
        debug_print(f"[DEBUG] page={page_num+1}: フォールバック - 画像{image_count}個（総面積{total_image_ratio:.1%}）を抽出")
        return self._extract_individual_embedded_images(
            page, page_num, header_y_max, footer_y_min
        )
    
    def _render_full_page_image(
        self, page, page_num: int,
        header_y_max: Optional[float], footer_y_min: Optional[float]
    ) -> List[Dict[str, Any]]:
        """ページ全体を1枚の画像としてレンダリング
        
        Args:
            page: PyMuPDFのページオブジェクト
            page_num: ページ番号
            header_y_max: ヘッダー領域の下端Y座標
            footer_y_min: フッター領域の上端Y座標
            
        Returns:
            抽出された画像の情報リスト（1要素）
        """
        page_width = page.rect.width
        page_height = page.rect.height
        
        # クリップ領域を計算（ヘッダー/フッターを除外）
        clip_y0 = header_y_max if header_y_max else 0
        clip_y1 = footer_y_min if footer_y_min else page_height
        
        # マージンを追加（上下5px）
        clip_y0 = max(0, clip_y0 - 5)
        clip_y1 = min(page_height, clip_y1 + 5)
        
        clip_bbox = (0, clip_y0, page_width, clip_y1)
        
        try:
            self.image_counter += 1
            image_filename = f"{self.base_name}_fig_{page_num + 1:03d}_{self.image_counter:03d}"
            
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
            
            debug_print(f"[DEBUG] page={page_num+1}: ページ全体画像を抽出 {image_filename}")
            
            return [{
                "path": image_path,
                "filename": os.path.basename(image_path),
                "bbox": clip_bbox,
                "y_position": clip_y0,
                "texts": [],  # テキストは別途Markdownとして出力
                "column": "full"
            }]
            
        except Exception as e:
            debug_print(f"[DEBUG] page={page_num+1}: ページ全体画像抽出エラー: {e}")
            return []
    
    def _extract_individual_embedded_images(
        self, page, page_num: int,
        header_y_max: Optional[float], footer_y_min: Optional[float]
    ) -> List[Dict[str, Any]]:
        """個別の埋め込み画像を抽出
        
        Args:
            page: PyMuPDFのページオブジェクト
            page_num: ページ番号
            header_y_max: ヘッダー領域の下端Y座標
            footer_y_min: フッター領域の上端Y座標
            
        Returns:
            抽出された画像の情報リスト
        """
        images = []
        page_width = page.rect.width
        page_height = page.rect.height
        page_area = page_width * page_height
        
        # 最小面積閾値（ページ面積の2%以上）
        min_area = page_area * 0.02
        
        # 重複除外用のbboxリスト
        accepted_bboxes = []
        
        try:
            img_infos = page.get_image_info()
            
            for img_info in img_infos:
                bbox = img_info.get('bbox')
                if not bbox:
                    continue
                
                area = (bbox[2] - bbox[0]) * (bbox[3] - bbox[1])
                
                # 小さすぎる画像は除外
                if area < min_area:
                    debug_print(f"[DEBUG] page={page_num+1}: 小さな画像を除外（面積={area:.0f}）")
                    continue
                
                # ヘッダー領域の画像は除外
                if header_y_max and bbox[3] < header_y_max:
                    debug_print(f"[DEBUG] page={page_num+1}: ヘッダー領域の画像を除外")
                    continue
                
                # フッター領域の画像は除外
                if footer_y_min and bbox[1] > footer_y_min:
                    debug_print(f"[DEBUG] page={page_num+1}: フッター領域の画像を除外")
                    continue
                
                # 重複するbboxは除外（重なり率90%以上）
                is_duplicate = False
                for accepted_bbox in accepted_bboxes:
                    overlap_ratio = self._fig_bbox_overlap_ratio(bbox, accepted_bbox)
                    if overlap_ratio >= 0.9:
                        debug_print(f"[DEBUG] page={page_num+1}: 重複画像を除外（重なり率={overlap_ratio:.1%}）")
                        is_duplicate = True
                        break
                if is_duplicate:
                    continue
                
                accepted_bboxes.append(bbox)
                
                self.image_counter += 1
                image_filename = f"{self.base_name}_fig_{page_num + 1:03d}_{self.image_counter:03d}"
                
                clip_rect = fitz.Rect(bbox)
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
                
                # 図形内テキストを抽出（ラベル拡張なし）
                figure_texts, _ = self._extract_text_in_bbox(
                    page, bbox, expand_for_labels=False
                )
                
                images.append({
                    "path": image_path,
                    "filename": os.path.basename(image_path),
                    "bbox": bbox,
                    "y_position": bbox[1],
                    "texts": figure_texts,
                    "column": "full"
                })
                
                debug_print(f"[DEBUG] page={page_num+1}: 個別画像を抽出 {image_filename}")
                
        except Exception as e:
            debug_print(f"[DEBUG] page={page_num+1}: 個別画像抽出エラー: {e}")
        
        images.sort(key=lambda x: x["y_position"])
        return images

    def _extract_individual_images(
        self, page, page_num: int
    ) -> List[Dict[str, Any]]:
        """個別の埋め込み画像を抽出
        
        Args:
            page: PyMuPDFのページオブジェクト
            page_num: ページ番号
            
        Returns:
            抽出された画像の情報リスト
        """
        images = []
        
        try:
            image_list = page.get_images(full=True)
            
            for img_info in image_list:
                xref = img_info[0]
                
                try:
                    for img_rect in page.get_image_rects(xref):
                        bbox = (img_rect.x0, img_rect.y0, img_rect.x1, img_rect.y1)
                        area = (bbox[2] - bbox[0]) * (bbox[3] - bbox[1])
                        
                        if area < 1000:
                            continue
                        
                        self.image_counter += 1
                        image_filename = f"{self.base_name}_img_{page_num + 1:03d}_{self.image_counter:03d}"
                        
                        clip_rect = fitz.Rect(bbox)
                        matrix = fitz.Matrix(2.0, 2.0)
                        pix = page.get_pixmap(matrix=matrix, clip=clip_rect)
                        
                        if self.output_format == 'svg':
                            image_path = os.path.join(self.images_dir, f"{image_filename}.svg")
                            temp_png = os.path.join(self.images_dir, f"temp_img_{self.image_counter}.png")
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
                            "bbox": bbox,
                            "y_position": bbox[1]
                        })
                        
                except Exception as e:
                    debug_print(f"[DEBUG] 画像抽出エラー (xref={xref}): {e}")
                    continue
                    
        except Exception as e:
            debug_print(f"[DEBUG] 画像リスト取得エラー: {e}")
        
        images.sort(key=lambda x: x["y_position"])
        return images
