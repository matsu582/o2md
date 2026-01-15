#!/usr/bin/env python3
"""
PDFテキスト抽出・処理Mixinモジュール

PDFToMarkdownConverterクラスのテキスト抽出・処理機能を提供します。
このモジュールはMixinクラスとして設計されており、メインクラスから継承されます。

機能:
- 構造化テキスト抽出（フォントサイズ、位置情報を使用）
- カラム分割と段落リフロー
- リスト継続行のマージ
- 上付き文字行のマージ
- 書式付きテキスト生成
"""

import re
from typing import List, Dict, Any, Tuple, Set, Optional

try:
    import fitz
except ImportError as e:
    raise ImportError(
        "PyMuPDFライブラリが必要です: pip install PyMuPDF または uv sync を実行してください"
    ) from e


def debug_print(*args, **kwargs):
    """デバッグ出力（pdf2mdモジュールに委譲）"""
    try:
        from pdf2md import debug_print as _dp
        _dp(*args, **kwargs)
    except ImportError:
        pass


class _TextMixin:
    """テキスト抽出・処理機能を提供するMixinクラス
    
    このクラスはPDFToMarkdownConverterに継承され、
    テキスト抽出、段落リフロー、書式処理機能を提供します。
    
    注意: このクラスは単独では使用できません。
    PDFToMarkdownConverterクラスと組み合わせて使用してください。
    """

    def _extract_structured_text(self, page) -> List[Dict[str, Any]]:
        """PDFページから構造化されたテキストブロックを抽出
        
        フォントサイズ、位置情報を使用して見出し、段落、箇条書き、表を判定する。
        
        Args:
            page: PyMuPDFのページオブジェクト
            
        Returns:
            構造化されたテキストブロックのリスト
        """
        blocks = []
        
        try:
            # 詳細なテキスト情報を取得
            text_dict = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)
        except Exception as e:
            debug_print(f"[DEBUG] テキスト抽出エラー: {e}")
            return []
        
        if not text_dict.get("blocks"):
            return []
        
        # フォントサイズの統計を収集（見出し判定用）
        font_sizes = []
        for block in text_dict["blocks"]:
            if block.get("type") == 0:  # テキストブロック
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        if span.get("text", "").strip():
                            font_sizes.append(span.get("size", 12))
        
        if not font_sizes:
            return []
        
        # 基準フォントサイズを計算（最頻値を本文サイズとする）
        from collections import Counter
        size_counts = Counter(round(s, 1) for s in font_sizes)
        base_font_size = size_counts.most_common(1)[0][0] if size_counts else 12
        
        # 表の検出用: 同じY座標に複数のテキストがあるかチェック
        table_rows = self._detect_table_structure(text_dict)
        
        for block in text_dict["blocks"]:
            if block.get("type") != 0:  # テキストブロック以外はスキップ
                continue
            
            block_text_parts = []
            block_font_size = base_font_size
            block_is_bold = False
            block_bbox = block.get("bbox", (0, 0, 0, 0))
            
            for line in block.get("lines", []):
                line_text = ""
                line_font_size = base_font_size
                line_is_bold = False
                
                for span in line.get("spans", []):
                    text = span.get("text", "")
                    if text:
                        line_text += text
                        line_font_size = max(line_font_size, span.get("size", 12))
                        font_name = span.get("font", "").lower()
                        if "bold" in font_name or "heavy" in font_name:
                            line_is_bold = True
                
                if line_text.strip():
                    block_text_parts.append(line_text)
                    block_font_size = max(block_font_size, line_font_size)
                    if line_is_bold:
                        block_is_bold = True
            
            if not block_text_parts:
                continue
            
            full_text = "\n".join(block_text_parts)
            
            # ブロックタイプを判定
            block_type = self._classify_block_type(
                full_text, block_font_size, base_font_size, block_is_bold, block_bbox
            )
            
            blocks.append({
                "type": block_type,
                "text": full_text,
                "font_size": block_font_size,
                "bbox": block_bbox
            })
        
        # 表構造がある場合は表として処理
        if table_rows:
            blocks = self._merge_table_blocks(blocks, table_rows)
        
        return blocks

    def _extract_structured_text_v2(
        self, page, header_footer_patterns: Set[str],
        exclude_bboxes: List[Tuple[float, float, float, float]] = None,
        is_slide_document: bool = False
    ) -> List[Dict[str, Any]]:
        """PDFページから構造化されたテキストブロックを抽出（改良版）
        
        行単位で抽出し、カラム分割と段落リフローを行う。
        ヘッダ・フッタを除外する。
        図領域内のテキストも除外する。
        スライド文書の場合は段落結合をスキップする。
        
        Args:
            page: PyMuPDFのページオブジェクト
            header_footer_patterns: ヘッダ・フッタパターンのセット
            exclude_bboxes: 除外するbboxのリスト（図領域など）
            is_slide_document: スライド文書フラグ（Trueの場合、段落結合をスキップ）
            
        Returns:
            構造化されたテキストブロックのリスト
        """
        if exclude_bboxes is None:
            exclude_bboxes = []
        import re
        from collections import Counter
        
        try:
            text_dict = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)
        except Exception as e:
            debug_print(f"[DEBUG] テキスト抽出エラー: {e}")
            return []
        
        if not text_dict.get("blocks"):
            return []
        
        page_width = text_dict.get("width", 612)
        page_height = text_dict.get("height", 792)
        page_center = page_width / 2
        
        # 段組み判定（1段組みの場合は全行をfullにする）
        column_count = self._detect_column_layout(text_dict)
        is_single_column = (column_count == 1)
        
        # 行単位でテキストを収集（span情報も保持）
        lines_data = []
        font_sizes = []
        
        for block in text_dict.get("blocks", []):
            if block.get("type") != 0:
                continue
            
            for line in block.get("lines", []):
                line_bbox = line.get("bbox", (0, 0, 0, 0))
                
                # span情報を収集
                all_spans = []
                for span in line.get("spans", []):
                    text = span.get("text", "")
                    if text:
                        span_size = span.get("size", 12)
                        span_bbox = span.get("bbox", (0, 0, 0, 0))
                        font_name = span.get("font", "").lower()
                        span_flags = span.get("flags", 0)
                        
                        # 太字・斜体の検出（フォント名とflagsの両方をチェック）
                        # PyMuPDF flags: bit 0 = superscript, bit 1 = italic, bit 2 = serifed, bit 3 = monospaced, bit 4 = bold
                        span_is_bold = (
                            "bold" in font_name or 
                            "heavy" in font_name or
                            (span_flags & (1 << 4)) != 0
                        )
                        span_is_italic = (
                            "italic" in font_name or 
                            "oblique" in font_name or
                            (span_flags & (1 << 1)) != 0
                        )
                        
                        # 上付き・下付きの検出
                        # PyMuPDF flags bit 0 = superscript
                        span_is_superscript = (span_flags & (1 << 0)) != 0
                        # 下付きはflagsでは検出できないため、フォントサイズとY座標で推定
                        span_is_subscript = False
                        
                        # 打消し線の検出（PDFでは直接検出困難、フォント名で推定）
                        span_is_strikethrough = "strikeout" in font_name or "strike" in font_name
                        
                        all_spans.append({
                            "text": text,
                            "size": span_size,
                            "bbox": span_bbox,
                            "is_bold": span_is_bold,
                            "is_italic": span_is_italic,
                            "is_superscript": span_is_superscript,
                            "is_subscript": span_is_subscript,
                            "is_strikethrough": span_is_strikethrough
                        })
                        font_sizes.append(span_size)
                
                if not all_spans:
                    continue
                
                # X座標の距離が大きいspanを別の行として分離
                # 同じPDF行内でも、X座標の距離がフォントサイズの1.5倍以上の場合は別の行として扱う
                # 空白のみのspanはスキップして、実際のテキストspanの間の距離を計算
                span_groups = []
                current_group = [all_spans[0]]
                
                for i in range(1, len(all_spans)):
                    curr_span = all_spans[i]
                    
                    # 空白のみのspanは現在のグループに追加して次へ
                    if curr_span["text"].strip() == "":
                        current_group.append(curr_span)
                        continue
                    
                    # 現在のグループの最後の非空白spanを探す
                    prev_span = None
                    for j in range(len(current_group) - 1, -1, -1):
                        if current_group[j]["text"].strip() != "":
                            prev_span = current_group[j]
                            break
                    
                    if prev_span is None:
                        # グループに非空白spanがない場合は現在のspanを追加
                        current_group.append(curr_span)
                        continue
                    
                    prev_x_end = prev_span["bbox"][2]
                    curr_x_start = curr_span["bbox"][0]
                    x_gap = curr_x_start - prev_x_end
                    font_size = min(prev_span["size"], curr_span["size"])
                    
                    # X座標の距離がフォントサイズの1.5倍以上の場合は別グループ
                    if x_gap > font_size * 1.5:
                        span_groups.append(current_group)
                        current_group = [curr_span]
                    else:
                        current_group.append(curr_span)
                
                span_groups.append(current_group)
                
                # 各グループを別の行として処理
                for line_spans in span_groups:
                    line_text = ""
                    line_font_size = 0
                    line_is_bold = False
                    
                    # グループのbboxを計算（空白のみのspanを除外）
                    # 空白spanを含めるとbboxが広がり、_sort_lines_by_columnで
                    # 再度クラスタリングされて結合されてしまう
                    non_whitespace_spans = [s for s in line_spans if s["text"].strip()]
                    if non_whitespace_spans:
                        group_x0 = min(s["bbox"][0] for s in non_whitespace_spans)
                        group_y0 = min(s["bbox"][1] for s in non_whitespace_spans)
                        group_x1 = max(s["bbox"][2] for s in non_whitespace_spans)
                        group_y1 = max(s["bbox"][3] for s in non_whitespace_spans)
                    else:
                        group_x0 = min(s["bbox"][0] for s in line_spans)
                        group_y0 = min(s["bbox"][1] for s in line_spans)
                        group_x1 = max(s["bbox"][2] for s in line_spans)
                        group_y1 = max(s["bbox"][3] for s in line_spans)
                    group_bbox = (group_x0, group_y0, group_x1, group_y1)
                    
                    for span in line_spans:
                        line_text += span["text"]
                        line_font_size = max(line_font_size, span["size"])
                        if span["is_bold"]:
                            line_is_bold = True
                    
                    if not line_text.strip():
                        continue
                    
                    # ヘッダ・フッタを除外（Y座標とページ高さを渡す）
                    if self._is_header_footer(
                        line_text, header_footer_patterns,
                        y_pos=group_bbox[1], page_height=page_height, font_size=line_font_size
                    ):
                        debug_print(f"[DEBUG] ヘッダ・フッタ除外: {line_text.strip()[:30]}...")
                        continue
                    
                    # 視覚的に見えないテキストを除外
                    # ページ中央付近の単独数字（装飾的な要素）を除外
                    line_text_stripped = line_text.strip()
                    line_center_x = (group_bbox[0] + group_bbox[2]) / 2
                    relative_x = line_center_x / page_width
                    is_centered = 0.4 < relative_x < 0.6
                    is_single_digit = re.match(r'^[0-9０-９]$', line_text_stripped)
                    if is_centered and is_single_digit:
                        debug_print(f"[DEBUG] 装飾的テキスト除外: '{line_text_stripped}' at x={line_center_x:.1f}")
                        continue
                    
                    # 図領域内のテキストを除外（中心点が図領域内にある場合）
                    # ただし、キャプションパターン（図X、表X）は除外しない
                    line_center_x = (group_bbox[0] + group_bbox[2]) / 2
                    line_center_y = (group_bbox[1] + group_bbox[3]) / 2
                    line_text_stripped = line_text.strip()
                    is_caption = re.match(r'^図\s*\d+', line_text_stripped) or re.match(r'^表\s*\d+', line_text_stripped)
                    in_figure = False
                    if not is_caption:
                        for fig_bbox in exclude_bboxes:
                            if (fig_bbox[0] <= line_center_x <= fig_bbox[2] and
                                fig_bbox[1] <= line_center_y <= fig_bbox[3]):
                                in_figure = True
                                break
                    if in_figure:
                        debug_print(f"[DEBUG] 図領域内テキスト除外: {line_text_stripped[:30]}...")
                        continue
                    
                    line_width = group_bbox[2] - group_bbox[0]
                    x_center = (group_bbox[0] + group_bbox[2]) / 2
                    
                    # カラム判定: フル幅、左カラム、右カラム
                    # 1段組みページでは全行をfullにして、誤った段落分割を防ぐ
                    if is_single_column:
                        column = "full"
                    elif line_width > page_width * 0.6:
                        column = "full"
                    elif x_center < page_center:
                        column = "left"
                    else:
                        column = "right"
                    
                    # 行全体の斜体フラグを計算
                    line_is_italic = any(s.get("is_italic", False) for s in line_spans)
                    
                    lines_data.append({
                        "text": line_text,
                        "bbox": group_bbox,
                        "font_size": line_font_size,
                        "is_bold": line_is_bold,
                        "is_italic": line_is_italic,
                        "column": column,
                        "y": group_bbox[1],
                        "x": group_bbox[0],
                        "width": line_width,
                        "spans": line_spans
                    })
        
        if not lines_data:
            return []
        
        # 基準フォントサイズを計算
        size_counts = Counter(round(s, 1) for s in font_sizes)
        base_font_size = size_counts.most_common(1)[0][0] if size_counts else 12
        
        # 傍注（上付き文字）を検出して結合
        lines_data = self._merge_superscript_lines(lines_data, base_font_size)
        
        # カラム内の表を検出（リフロー前に行う）
        table_regions = self._detect_table_regions(lines_data, page_center)
        
        # 罫線ベースの表検出（find_tables()を使用）
        # テキストベースの検出で見逃した表を補完する
        line_based_tables = self._detect_line_based_tables(page, lines_data)
        
        # カラムごとにソート（フル幅→左→右の順、各カラム内はY座標順）
        # スライド文書の場合は行クラスタリングをスキップ
        sorted_lines = self._sort_lines_by_column(lines_data, is_slide_document=is_slide_document)
        
        # 段落リフロー（同一カラム内で近接する行を結合、表領域は除外）
        # exclude_bboxes（図形のbbox）を渡して、図形と重なるテーブルの重複出力を防ぐ
        # ページサイズを渡して横長ページ（PPTスライド等）での改行保持を有効化
        # スライド文書フラグを渡して段落結合をスキップ
        reflowed_blocks = self._reflow_paragraphs_with_tables(
            sorted_lines, base_font_size, table_regions, line_based_tables,
            figure_bboxes=exclude_bboxes,
            page_width=page_width,
            page_height=page_height,
            is_slide_document=is_slide_document
        )
        
        # ブロックタイプを判定（カラム情報を保持）
        blocks = []
        for block_data in reflowed_blocks:
            # 見出しブロックはそのまま（カラム情報も保持）
            if block_data.get("is_heading"):
                level = block_data.get("heading_level", 1)
                block_type = f"heading{level}"
                block = {
                    "type": block_type,
                    "text": block_data["text"],
                    "font_size": block_data["font_size"],
                    "bbox": block_data["bbox"],
                    "is_heading": True
                }
                if "column" in block_data:
                    block["column"] = block_data["column"]
                blocks.append(block)
                continue
            
            # 表ブロックはそのまま（カラム情報も保持）
            if block_data.get("is_table"):
                block = {
                    "type": "table",
                    "text": block_data["text"],
                    "font_size": block_data["font_size"],
                    "bbox": block_data["bbox"]
                }
                if "column" in block_data:
                    block["column"] = block_data["column"]
                blocks.append(block)
                continue
            
            block_type = self._classify_block_type(
                block_data["text"],
                block_data["font_size"],
                base_font_size,
                block_data["is_bold"],
                block_data["bbox"],
                is_slide_document=is_slide_document
            )
            block = {
                "type": block_type,
                "text": block_data["text"],
                "font_size": block_data["font_size"],
                "bbox": block_data["bbox"]
            }
            # リストマーカー種別を設定（_finalize_blockで設定済みの場合はそれを使用）
            marker_kind = block_data.get("list_marker_kind")
            if marker_kind:
                block["list_marker_kind"] = marker_kind
            # カラム情報を保持（2段組みの順序維持に必要）
            if "column" in block_data:
                block["column"] = block_data["column"]
            blocks.append(block)
        
        # 番号付きリストの継続行を結合する後処理
        # スライド文書の場合はスキップ（各行を独立に保持）
        if not is_slide_document:
            blocks = self._merge_list_continuations(blocks)
        
        # リストの途中で見出しになっているブロックをlist_itemに降格
        blocks = self._demote_heading_in_list_context(blocks)
        
        # リスト項目にネストレベルを割り当て
        blocks = self._assign_list_levels(blocks)
        
        return blocks

    def _merge_list_continuations(
        self, blocks: List[Dict[str, Any]]
    ) -> List[Dict[str, Any]]:
        """番号付きリストの継続行を前のリスト項目に結合
        
        list_itemまたはリスト項目開始パターンを含むparagraphの直後に
        paragraphが来て、かつ以下の条件を満たす場合に結合:
        - 縦gap（行間）が15px以下
        - 前のブロックが句点（。）で終わらない
        - 次のブロックが新しいリスト項目開始パターンで始まらない
        - 次のブロックが短い（8文字以下）場合はdelta_x制限を緩和
        
        Args:
            blocks: ブロックのリスト
            
        Returns:
            結合後のブロックのリスト
        """
        import re
        
        # リスト項目開始パターン（すべての形式を網羅）
        list_start_pattern = re.compile(
            r'^[\s]*('
            r'[(（][0-9０-９]+[)）]\s*|'  # (1) （１） など
            r'[0-9０-９]+[.)．）]\s*|'  # 1. 2) など
            r'[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]'  # 丸数字
            r')'
        )
        
        if len(blocks) < 2:
            return blocks
        
        merged = []
        skip_next = False
        
        for i, block in enumerate(blocks):
            if skip_next:
                skip_next = False
                continue
            
            # 最後のブロックはそのまま追加
            if i >= len(blocks) - 1:
                merged.append(block)
                continue
            
            next_block = blocks[i + 1]
            curr_text = block.get("text", "")
            next_text = next_block.get("text", "")
            
            # 見出しブロックは継続行結合の対象外
            if block.get("is_heading"):
                merged.append(block)
                continue
            
            # 現在のブロックがリスト項目かどうか判定
            is_curr_list = (
                block.get("type") == "list_item" or
                bool(list_start_pattern.match(curr_text))
            )
            
            # 次のブロックがparagraphで、現在がリスト項目の場合
            if is_curr_list and next_block.get("type") == "paragraph":
                
                curr_bbox = block.get("bbox", (0, 0, 0, 0))
                next_bbox = next_block.get("bbox", (0, 0, 0, 0))
                
                # 左端差（ハングインデント）を計算
                delta_x = next_bbox[0] - curr_bbox[0]
                # 縦gap（行間）を計算
                gap_y = next_bbox[1] - curr_bbox[3]
                
                # 前のブロックの末尾文字を取得
                ends_with_period = curr_text.rstrip().endswith("。")
                
                # 次のブロックが新しいリスト項目開始パターンで始まるかチェック
                starts_with_list_marker = bool(list_start_pattern.match(next_text))
                
                # 次のブロックの文字数（短い継続行の判定用）
                next_text_len = len(next_text.strip())
                
                # 結合条件を判定
                # 短い継続行（8文字以下）の場合はdelta_x制限を緩和
                # ただし、X座標の距離が大きい場合は結合しない（同じ行の別要素を誤結合しないため）
                # 前のブロックの右端と次のブロックの左端の距離を計算
                x_gap = next_bbox[0] - curr_bbox[2]
                large_x_gap = x_gap > 10  # 10px以上離れている場合は別要素とみなす
                
                if next_text_len <= 8:
                    # 短い継続行: delta_x制限緩和、ただしX座標の距離が大きい場合は除外
                    should_merge = (
                        gap_y <= 15 and
                        not ends_with_period and
                        not starts_with_list_marker and
                        not large_x_gap
                    )
                else:
                    # 通常の継続行: ハングインデント範囲内
                    should_merge = (
                        5 <= delta_x <= 60 and
                        gap_y <= 15 and
                        not ends_with_period and
                        not starts_with_list_marker
                    )
                
                if should_merge:
                    # テキストを結合（日本語なのでスペースなしで連結）
                    merged_text = curr_text.rstrip() + next_text.lstrip()
                    # bboxを拡張
                    merged_bbox = (
                        min(curr_bbox[0], next_bbox[0]),
                        curr_bbox[1],
                        max(curr_bbox[2], next_bbox[2]),
                        next_bbox[3]
                    )
                    # 元のブロックタイプを維持（list_itemまたはparagraph）
                    merged_block = {
                        "type": block.get("type", "paragraph"),
                        "text": merged_text,
                        "font_size": block.get("font_size", 0),
                        "bbox": merged_bbox
                    }
                    if "column" in block:
                        merged_block["column"] = block["column"]
                    # リストマーカー種別を保持
                    if "list_marker_kind" in block:
                        merged_block["list_marker_kind"] = block["list_marker_kind"]
                    merged.append(merged_block)
                    skip_next = True
                    continue
            
            merged.append(block)
        
        # 再帰的に結合（複数の継続行がある場合）
        if len(merged) < len(blocks):
            return self._merge_list_continuations(merged)
        
        return merged

    def _demote_heading_in_list_context(
        self, blocks: List[Dict[str, Any]]
    ) -> List[Dict[str, Any]]:
        """リストの途中で見出しになっているブロックをlist_itemに降格
        
        直前がlist_itemで、y_gapが小さく、x0が近い場合は
        headingをlist_itemに降格する。これにより、リストの途中で
        見出しが挿入される問題を防ぐ。
        
        Args:
            blocks: ブロックのリスト
            
        Returns:
            降格後のブロックのリスト
        """
        import re
        
        if len(blocks) < 2:
            return blocks
        
        # 番号付き見出しパターン（「N. タイトル」形式）
        numbered_heading_pattern = re.compile(
            r'^[\d０-９]+[\.．\)）]\s+'
        )
        
        result = []
        for i, block in enumerate(blocks):
            block_type = block.get("type", "")
            text = block.get("text", "").strip()
            
            # headingで、番号付きパターンに一致する場合のみチェック
            if block_type.startswith("heading") and numbered_heading_pattern.match(text):
                # 直前のブロックを探す
                prev_block = result[-1] if result else None
                
                if prev_block:
                    prev_type = prev_block.get("type", "")
                    prev_bbox = prev_block.get("bbox", (0, 0, 0, 0))
                    curr_bbox = block.get("bbox", (0, 0, 0, 0))
                    
                    # 直前がlist_itemの場合
                    if prev_type == "list_item":
                        # y_gap（縦方向の距離）を計算
                        y_gap = curr_bbox[1] - prev_bbox[3]
                        # x0の差（インデントの差）を計算
                        x_diff = abs(curr_bbox[0] - prev_bbox[0])
                        
                        # y_gapが小さく（30pt以下）、x0が近い（20pt以下）場合は降格
                        if y_gap <= 30 and x_diff <= 20:
                            # headingをlist_itemに降格
                            demoted_block = block.copy()
                            demoted_block["type"] = "list_item"
                            result.append(demoted_block)
                            continue
            
            result.append(block)
        
        return result

    def _assign_list_levels(
        self, blocks: List[Dict[str, Any]]
    ) -> List[Dict[str, Any]]:
        """リスト項目にネストレベルを割り当て
        
        丸数字（circled）で始まるリスト項目で、直前に親リスト
        （fullwidth/arabic）がある場合はネストレベルを1に設定。
        
        Args:
            blocks: ブロックのリスト
            
        Returns:
            list_levelが設定されたブロックのリスト
        """
        if len(blocks) < 2:
            return blocks
        
        # 親リストのマーカー種別
        parent_markers = {"fullwidth", "arabic"}
        # 子リストのマーカー種別
        child_markers = {"circled", "paren"}
        
        result = []
        # 親リストが走っているかどうかのフラグ
        parent_list_active = False
        
        for block in blocks:
            block_type = block.get("type", "")
            marker_kind = block.get("list_marker_kind")
            
            # list_item以外のブロックで親リストをリセット
            if block_type != "list_item":
                if block_type.startswith("heading") or block_type == "paragraph":
                    parent_list_active = False
                result.append(block)
                continue
            
            # list_itemの場合
            if marker_kind in parent_markers:
                # 親リストマーカーの場合、親リストを開始
                parent_list_active = True
                block_copy = block.copy()
                block_copy["list_level"] = 0
                result.append(block_copy)
            elif marker_kind in child_markers and parent_list_active:
                # 子リストマーカーで、親リストが走っている場合はネスト
                block_copy = block.copy()
                block_copy["list_level"] = 1
                result.append(block_copy)
            else:
                # その他の場合はレベル0
                block_copy = block.copy()
                block_copy["list_level"] = 0
                result.append(block_copy)
        
        return result

    def _merge_superscript_lines(
        self, lines_data: List[Dict], base_font_size: float
    ) -> List[Dict]:
        """傍注（上付き文字）を検出して前の行に結合
        
        フォントサイズが本文より小さく、前の行の直後に配置されている
        テキストを<sup>タグで囲んで前の行に結合する。
        
        Args:
            lines_data: 行データのリスト
            base_font_size: 基準フォントサイズ
            
        Returns:
            結合後の行データのリスト
        """
        if len(lines_data) < 2:
            return lines_data
        
        # カラムごとに処理
        result = []
        skip_indices = set()
        
        # カラムごとにグループ化
        column_groups = {}
        for i, line in enumerate(lines_data):
            col = line["column"]
            if col not in column_groups:
                column_groups[col] = []
            column_groups[col].append((i, line))
        
        for col, col_lines in column_groups.items():
            # Y座標が近い行のペアを探す
            for idx1, (orig_idx1, line1) in enumerate(col_lines):
                if orig_idx1 in skip_indices:
                    continue
                
                for idx2, (orig_idx2, line2) in enumerate(col_lines):
                    if idx1 == idx2 or orig_idx2 in skip_indices:
                        continue
                    
                    # Y座標が近い（10ピクセル以内）
                    if abs(line1["y"] - line2["y"]) >= 10:
                        continue
                    
                    # どちらの行が先頭に小さいフォントのspanを持つか確認
                    line1_spans = line1.get("spans", [])
                    line2_spans = line2.get("spans", [])
                    
                    # line2の先頭spanが小さいフォントサイズか確認
                    if line2_spans:
                        first_span = line2_spans[0]
                        first_span_size = first_span.get("size", base_font_size)
                        sup_text = first_span.get("text", "").strip()
                        
                        # フォントサイズが本文の70%以下で、テキストが短い（15文字以下）
                        if (first_span_size < base_font_size * 0.7 and
                            len(sup_text) <= 15 and len(sup_text) > 0):
                            
                            # X座標が連続しているか確認（line1の右端とline2の左端）
                            line1_right = line1["bbox"][2]
                            line2_left = line2["bbox"][0]
                            
                            if abs(line1_right - line2_left) < 10:
                                # 残りのテキストを取得
                                remaining_text = ""
                                for j, span in enumerate(line2_spans):
                                    if j == 0:
                                        continue
                                    remaining_text += span.get("text", "")
                                
                                # 結合したテキストを作成
                                merged_text = line1["text"].rstrip()
                                # 注釈参照の場合はMarkdown脚注形式に変換
                                if self._is_footnote_reference(sup_text):
                                    merged_text += self._format_footnote_ref(sup_text)
                                else:
                                    merged_text += f"<sup>{sup_text}</sup>"
                                if remaining_text.strip():
                                    merged_text += remaining_text
                                
                                # 結合した行を作成
                                merged_line = line1.copy()
                                merged_line["text"] = merged_text
                                merged_line["bbox"] = (
                                    min(line1["bbox"][0], line2["bbox"][0]),
                                    min(line1["bbox"][1], line2["bbox"][1]),
                                    max(line1["bbox"][2], line2["bbox"][2]),
                                    max(line1["bbox"][3], line2["bbox"][3])
                                )
                                
                                # line1を更新、line2をスキップ
                                lines_data[orig_idx1] = merged_line
                                skip_indices.add(orig_idx2)
                                debug_print(f"[DEBUG] 傍注結合: {line1['text'][:20]}... + <sup>{sup_text}</sup>")
                                break
        
        # スキップされていない行を結果に追加
        for i, line in enumerate(lines_data):
            if i not in skip_indices:
                result.append(line)
        
        return result

    def _sort_lines_by_column(
        self, lines_data: List[Dict], is_slide_document: bool = False
    ) -> List[Dict]:
        """行をカラムごとにソート
        
        フル幅要素を基準に、その間の区間で左→右の順に出力する。
        同じy座標（±フォントサイズの範囲内）にある行はx座標順にソートする。
        スライド文書の場合はクラスタリングをスキップして各行を独立に保持する。
        
        Args:
            lines_data: 行データのリスト
            is_slide_document: スライド文書フラグ
            
        Returns:
            ソートされた行データのリスト
        """
        import re
        
        def cluster_and_sort_by_row(lines: List[Dict], skip_clustering: bool = False) -> List[Dict]:
            """同じy座標の行をクラスタリングし、クラスタ内はx座標順にソート
            
            Args:
                lines: 行データのリスト
                skip_clustering: Trueの場合、クラスタリングをスキップしてy座標順のみでソート
            """
            if not lines:
                return []
            
            # y座標でソート
            sorted_by_y = sorted(lines, key=lambda x: x["y"])
            
            # スライド文書の場合はクラスタリングをスキップ
            if skip_clustering:
                debug_print("[DEBUG] スライド文書: 行クラスタリングをスキップ")
                return sorted_by_y
            
            # 同じy座標（±フォントサイズの0.5倍以内）の行をクラスタリング
            clusters = []
            current_cluster = [sorted_by_y[0]]
            
            for line in sorted_by_y[1:]:
                prev_line = current_cluster[-1]
                y_diff = abs(line["y"] - prev_line["y"])
                threshold = min(line.get("font_size", 12), prev_line.get("font_size", 12)) * 0.5
                
                if y_diff <= threshold:
                    current_cluster.append(line)
                else:
                    clusters.append(current_cluster)
                    current_cluster = [line]
            
            clusters.append(current_cluster)
            
            # 各クラスタ内をx座標順にソートし、X座標の距離が大きい場合は分離
            result = []
            for cluster in clusters:
                cluster_sorted = sorted(cluster, key=lambda x: x.get("x", 0))
                # X座標の距離が大きい行を分離
                if len(cluster_sorted) > 1:
                    separated = []
                    current_group = [cluster_sorted[0]]
                    for i in range(1, len(cluster_sorted)):
                        prev = cluster_sorted[i - 1]
                        curr = cluster_sorted[i]
                        # 前の行の右端と現在の行の左端の距離を計算
                        prev_x_end = prev.get("bbox", (0, 0, 0, 0))[2]
                        curr_x_start = curr.get("bbox", (0, 0, 0, 0))[0]
                        x_gap = curr_x_start - prev_x_end
                        font_size = min(prev.get("font_size", 12), curr.get("font_size", 12))
                        # X方向のギャップがフォントサイズの1倍以上の場合は分離
                        if x_gap > font_size * 1.0:
                            separated.extend(current_group)
                            current_group = [curr]
                        else:
                            current_group.append(curr)
                    separated.extend(current_group)
                    result.extend(separated)
                else:
                    result.extend(cluster_sorted)
            
            return result
        
        def merge_number_title_lines(lines: List[Dict]) -> List[Dict]:
            """番号のみの行と直後のタイトル行をマージ
            
            例: 「1.1」と「背景」を「1.1 背景」にマージ
            """
            if not lines:
                return []
            
            # 番号のみパターン（例: 1.1, 2.3.1, 第1章 など）
            number_only_pattern = re.compile(r'^[\d.]+$|^第\s*[\d０-９一二三四五六七八十]+\s*章?$')
            
            result = []
            skip_next = False
            
            for i, line in enumerate(lines):
                if skip_next:
                    skip_next = False
                    continue
                
                text = line.get("text", "").strip()
                
                # 番号のみの行かチェック
                if number_only_pattern.match(text) and i + 1 < len(lines):
                    next_line = lines[i + 1]
                    next_text = next_line.get("text", "").strip()
                    
                    # 次の行が同じy座標（±フォントサイズの0.5倍以内）かチェック
                    y_diff = abs(line["y"] - next_line["y"])
                    threshold = min(line.get("font_size", 12), next_line.get("font_size", 12)) * 0.5
                    
                    # 次の行が番号で始まらない短いテキストならマージ
                    if (y_diff <= threshold and 
                        not number_only_pattern.match(next_text) and
                        len(next_text) < 50):
                        # マージ
                        merged_text = f"{text} {next_text}"
                        merged_line = line.copy()
                        merged_line["text"] = merged_text
                        merged_line["bbox"] = (
                            min(line["bbox"][0], next_line["bbox"][0]),
                            min(line["bbox"][1], next_line["bbox"][1]),
                            max(line["bbox"][2], next_line["bbox"][2]),
                            max(line["bbox"][3], next_line["bbox"][3])
                        )
                        result.append(merged_line)
                        skip_next = True
                        debug_print(f"[DEBUG] 番号+タイトルをマージ: '{text}' + '{next_text}' -> '{merged_text}'")
                        continue
                
                result.append(line)
            
            return result
        
        # フル幅行とカラム行を分離
        full_lines = [l for l in lines_data if l["column"] == "full"]
        left_lines = [l for l in lines_data if l["column"] == "left"]
        right_lines = [l for l in lines_data if l["column"] == "right"]
        
        # 各グループをクラスタリング＆ソート
        # スライド文書の場合はクラスタリングをスキップ
        full_lines = cluster_and_sort_by_row(full_lines, skip_clustering=is_slide_document)
        left_lines = cluster_and_sort_by_row(left_lines, skip_clustering=is_slide_document)
        right_lines = cluster_and_sort_by_row(right_lines, skip_clustering=is_slide_document)
        
        # 番号+タイトルのマージ（fullカラムのみ、2カラムでは誤結合の可能性があるため）
        # スライド文書の場合はマージをスキップ
        if not is_slide_document:
            full_lines = merge_number_title_lines(full_lines)
        
        # スライド文書用: タイトル行の直後に「N/N」形式が来たら結合
        # 例: 「１．エグゼクティブサマリー」+ 「1/3」→「１．エグゼクティブサマリー 1/3」
        if is_slide_document and full_lines:
            page_fraction_pattern = re.compile(r'^\s*\d+\s*/\s*\d+\s*$')
            merged_full_lines = []
            skip_next = False
            for i, line in enumerate(full_lines):
                if skip_next:
                    skip_next = False
                    continue
                text = line.get("text", "").strip()
                # 次の行が「N/N」形式かチェック
                if i + 1 < len(full_lines):
                    next_line = full_lines[i + 1]
                    next_text = next_line.get("text", "").strip()
                    if page_fraction_pattern.match(next_text):
                        # タイトル行（番号付き見出し形式）の場合のみ結合
                        if re.match(r'^[０-９\d]+[．.]\s*', text):
                            merged_text = f"{text} {next_text}"
                            merged_line = line.copy()
                            merged_line["text"] = merged_text
                            merged_line["bbox"] = (
                                min(line["bbox"][0], next_line["bbox"][0]),
                                min(line["bbox"][1], next_line["bbox"][1]),
                                max(line["bbox"][2], next_line["bbox"][2]),
                                max(line["bbox"][3], next_line["bbox"][3])
                            )
                            merged_full_lines.append(merged_line)
                            skip_next = True
                            debug_print(f"[DEBUG] スライドタイトル+ページ番号をマージ: '{text}' + '{next_text}'")
                            continue
                merged_full_lines.append(line)
            full_lines = merged_full_lines
        
        # フル幅行がない場合は単純に左→右
        if not full_lines:
            return left_lines + right_lines
        
        # 左右カラムがない場合（1段組み）は単純にY座標でソート
        # 区間処理をスキップして重複を防ぐ
        if not left_lines and not right_lines:
            return full_lines
        
        # フル幅行を基準に区間を作成
        result = []
        added_indices = set()  # 追加済みの行インデックスを追跡
        full_y_positions = [l["y"] for l in full_lines]
        full_y_positions = [-float('inf')] + full_y_positions + [float('inf')]
        
        for i in range(len(full_y_positions) - 1):
            y_start = full_y_positions[i]
            y_end = full_y_positions[i + 1]
            
            # この区間のフル幅行を追加（重複チェック付き）
            if i > 0:
                for idx, fl in enumerate(full_lines):
                    if idx not in added_indices and abs(fl["y"] - y_start) < 1:
                        result.append(fl)
                        added_indices.add(idx)
            
            # この区間の左カラム行を追加
            for ll in left_lines:
                if y_start < ll["y"] < y_end:
                    result.append(ll)
            
            # この区間の右カラム行を追加
            for rl in right_lines:
                if y_start < rl["y"] < y_end:
                    result.append(rl)
        
        # 最後のフル幅行を追加（まだ追加されていない場合）
        if full_lines:
            last_idx = len(full_lines) - 1
            if last_idx not in added_indices:
                result.append(full_lines[last_idx])
        
        return result

    def _reflow_paragraphs(
        self, lines: List[Dict], base_font_size: float
    ) -> List[Dict]:
        """段落リフロー（近接する行を結合）
        
        同一カラム内で縦方向のギャップが小さい行を結合する。
        番号付き箇条書きの後に続く非番号行は別ブロックとして分離する。
        
        Args:
            lines: ソートされた行データのリスト
            base_font_size: 基準フォントサイズ
            
        Returns:
            結合されたブロックのリスト
        """
        import re
        
        if not lines:
            return []
        
        # 番号付き箇条書きパターン（全角数字も含む）
        # 小数（例: 14.0%）を誤認識しないよう、区切り記号の後に数字が続かないことを確認
        # (N)形式の括弧付き番号も含む（半角・全角両対応）
        numbered_list_pattern = re.compile(
            r'^[\s]*('
            r'[0-9０-９]+[.)．）](?=\s*[^0-9０-９])|'  # 1. 2) など
            r'[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]|'  # 丸数字
            r'[(（][0-9０-９]+[)）]\s+'  # (1) （１） など
            r')'
        )
        
        def is_numbered_list_line(text: str) -> bool:
            """行が番号付き箇条書きで始まるかを判定"""
            return bool(numbered_list_pattern.match(text))
        
        # 行高の推定（フォントサイズの1.2倍程度）
        line_height = base_font_size * 1.2
        # 結合する最大ギャップ
        # 閾値を大きくしすぎると段落が過剰に結合されるため、0.8倍に設定
        gap_threshold = line_height * 0.8
        
        blocks = []
        current_block = {
            "texts": [lines[0]["text"]],
            "bbox": list(lines[0]["bbox"]),
            "font_size": lines[0]["font_size"],
            "is_bold": lines[0]["is_bold"],
            "is_italic": lines[0].get("is_italic", False),
            "column": lines[0]["column"],
            "last_y": lines[0]["bbox"][3],
            "last_x": lines[0]["x"],
            "spans_list": [lines[0].get("spans", [])],
            "is_list": is_numbered_list_line(lines[0]["text"])
        }
        
        for i in range(1, len(lines)):
            line = lines[i]
            prev_line = lines[i - 1]
            
            # 結合条件をチェック
            y_gap = line["y"] - current_block["last_y"]
            same_column = line["column"] == current_block["column"]
            x_aligned = abs(line["x"] - current_block["last_x"]) < 20
            
            # 現在の行が番号付き箇条書きかどうか
            curr_is_numbered = is_numbered_list_line(line["text"])
            
            # 箇条書きブロックの終了判定
            # 現在のブロックが箇条書きで、次の行が番号付きでない場合は分離
            # ただし、吊り下げインデント（右に大きくずれている）場合は継続行として結合
            indent_threshold = base_font_size * 1.5  # 吊り下げインデントの閾値
            is_hanging_indent = (line["x"] - current_block["last_x"]) > indent_threshold
            list_boundary = (
                current_block["is_list"] and 
                not curr_is_numbered and
                not is_hanging_indent  # 吊り下げインデントでない場合は分離
            )
            
            # 段落の区切り条件
            is_new_paragraph = (
                y_gap > gap_threshold or
                not same_column or
                line["is_bold"] != current_block["is_bold"] or
                abs(line["font_size"] - current_block["font_size"]) > 1 or
                list_boundary or
                curr_is_numbered  # 新しい番号付き行は常に新しいブロック
            )
            
            if is_new_paragraph:
                # 現在のブロックを確定
                blocks.append(self._finalize_block(current_block))
                
                # 新しいブロックを開始
                current_block = {
                    "texts": [line["text"]],
                    "bbox": list(line["bbox"]),
                    "font_size": line["font_size"],
                    "is_bold": line["is_bold"],
                    "is_italic": line.get("is_italic", False),
                    "column": line["column"],
                    "last_y": line["bbox"][3],
                    "last_x": line["x"],
                    "spans_list": [line.get("spans", [])],
                    "is_list": curr_is_numbered
                }
            else:
                # 行を結合（日本語はスペースなし、英数字はスペースあり）
                prev_text = current_block["texts"][-1]
                curr_text = line["text"]
                
                # 前の行の末尾と現在の行の先頭をチェック
                if prev_text and curr_text:
                    prev_char = prev_text.rstrip()[-1] if prev_text.rstrip() else ""
                    curr_char = curr_text.lstrip()[0] if curr_text.lstrip() else ""
                    
                    # 英数字同士の場合はスペースを入れる
                    if prev_char.isascii() and curr_char.isascii():
                        if prev_char.isalnum() and curr_char.isalnum():
                            current_block["texts"].append(" " + curr_text)
                        else:
                            current_block["texts"].append(curr_text)
                    else:
                        # 日本語の場合はスペースなしで結合
                        current_block["texts"].append(curr_text)
                else:
                    current_block["texts"].append(curr_text)
                
                # bboxを更新
                current_block["bbox"][2] = max(current_block["bbox"][2], line["bbox"][2])
                current_block["bbox"][3] = line["bbox"][3]
                current_block["last_y"] = line["bbox"][3]
                current_block["font_size"] = max(current_block["font_size"], line["font_size"])
                if line["is_bold"]:
                    current_block["is_bold"] = True
                if line.get("is_italic", False):
                    current_block["is_italic"] = True
                # span情報を追加
                current_block["spans_list"].append(line.get("spans", []))
        
        # 最後のブロックを追加
        blocks.append(self._finalize_block(current_block))
        
        return blocks

    def _finalize_block(self, block_data: Dict) -> Dict:
        """ブロックデータを最終形式に変換
        
        span情報を使って太字・斜体の書式を適用する。
        リストマーカー種別（list_marker_kind）を判定して保存する。
        """
        import re
        
        # span情報を使って書式付きテキストを生成
        spans_list = block_data.get("spans_list", [])
        if spans_list and any(spans_list):
            text = self._apply_text_formatting(spans_list)
        else:
            # span情報がない場合は従来通り
            text = "".join(block_data["texts"]).strip()
        
        # リストマーカー種別を判定（変換前のテキストで判定）
        list_marker_kind = self._detect_list_marker_kind(text)
        
        # 番号付き箇条書きの検出と変換
        text = self._convert_numbered_bullets(text)
        
        result = {
            "text": text,
            "bbox": tuple(block_data["bbox"]),
            "font_size": block_data["font_size"],
            "is_bold": block_data["is_bold"],
            "is_italic": block_data.get("is_italic", False)
        }
        
        # リストマーカー種別を保存
        if list_marker_kind:
            result["list_marker_kind"] = list_marker_kind
        
        # カラム情報を保持（2段組みの順序維持に必要）
        if "column" in block_data:
            result["column"] = block_data["column"]
        
        return result
    
    def _detect_list_marker_kind(self, text: str) -> str:
        """テキストの行頭からリストマーカー種別を判定
        
        Args:
            text: 判定するテキスト
            
        Returns:
            マーカー種別（"circled", "fullwidth", "arabic", "paren", "bullet", None）
        """
        import re
        
        text = text.strip()
        if not text:
            return None
        
        # 丸数字（①②③...）
        if re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]', text):
            return "circled"
        
        # 全角数字+区切り（１．２．など）
        if re.match(r'^[０-９]+[．.)）]', text):
            return "fullwidth"
        
        # 半角数字+区切り（1. 2) など）
        if re.match(r'^[0-9]+[.)）]', text):
            return "arabic"
        
        # 括弧付き番号（(1) （１）など）
        if re.match(r'^[(（][0-9０-９]+[)）]', text):
            return "paren"
        
        # 箇条書きマーカー（•, ・, ○, ● など）
        if re.match(r'^[•・○●◆◇▪▫－―＊\-\*]', text):
            return "bullet"
        
        return None

    def _is_footnote_reference(self, text: str) -> bool:
        """テキストが注釈参照かどうかを判定
        
        用語1、[1]、[1][2]などの注釈参照パターンを検出する。
        10^-9などの数式上付きは注釈ではない。
        
        Args:
            text: 判定するテキスト
            
        Returns:
            注釈参照の場合True
        """
        import re
        text = text.strip()
        # 用語N パターン（用語1、用語2など）
        if re.match(r'^用語\d+$', text):
            return True
        # [N] パターン（[1]、[2]など、複数連続も可）
        if re.match(r'^(\[\d+\])+$', text):
            return True
        return False

    def _format_footnote_ref(self, text: str) -> str:
        """注釈参照をMarkdown形式に変換
        
        用語1 → [^用語1]
        [1] → [^1]
        [1][2] → [^1][^2]
        
        Args:
            text: 注釈参照テキスト
            
        Returns:
            Markdown形式の注釈参照
        """
        import re
        text = text.strip()
        # 用語N パターン
        if re.match(r'^用語\d+$', text):
            return f"[^{text}]"
        # [N] パターン（複数連続対応）
        if re.match(r'^(\[\d+\])+$', text):
            refs = re.findall(r'\[(\d+)\]', text)
            return ''.join(f"[^{r}]" for r in refs)
        return text

    def _apply_text_formatting(self, spans_list: List[List[Dict]]) -> str:
        """span情報を使って書式付きテキストを生成
        
        太字は**text**、斜体は*text*、太字斜体は***text***として出力する。
        番号付き箇条書きを検出し、リストマーカーには書式を適用しない。
        
        Args:
            spans_list: 行ごとのspan情報のリスト
            
        Returns:
            書式付きテキスト
        """
        import re
        
        # まず各行のテキストを生成（書式付き）
        formatted_lines = []
        for line_spans in spans_list:
            if not line_spans:
                continue
            
            line_text = ""
            for span in line_spans:
                text = span.get("text", "")
                if not text:
                    continue
                
                is_bold = span.get("is_bold", False)
                is_italic = span.get("is_italic", False)
                is_superscript = span.get("is_superscript", False)
                is_subscript = span.get("is_subscript", False)
                is_strikethrough = span.get("is_strikethrough", False)
                
                # 書式を適用（優先順位: 上付き/下付き > 打消し線 > 太字/斜体）
                if is_superscript:
                    # 注釈参照の場合はMarkdown脚注形式に変換
                    if self._is_footnote_reference(text):
                        formatted = self._format_footnote_ref(text)
                    else:
                        formatted = f"<sup>{text}</sup>"
                elif is_subscript:
                    formatted = f"<sub>{text}</sub>"
                elif is_strikethrough:
                    formatted = f"~~{text}~~"
                elif is_bold and is_italic:
                    formatted = f"***{text}***"
                elif is_bold:
                    formatted = f"**{text}**"
                elif is_italic:
                    formatted = f"*{text}*"
                else:
                    formatted = text
                
                line_text += formatted
            
            if line_text.strip():
                formatted_lines.append(line_text)
        
        if not formatted_lines:
            return ""
        
        # 番号付き箇条書きパターン（行頭の数字+区切り+空白、全角数字も含む）
        # 例: "1. ", "2) ", "1．", "１．", "①", "②", "(1)", "（１）"
        # 小数（例: 14.0%）を誤認識しないよう、区切り記号の後に数字が続かないことを確認
        # (N)形式の括弧付き番号も含む（半角・全角両対応）
        numbered_pattern = re.compile(
            r'^[\s]*('
            r'[0-9０-９]+[.)．）](?=\s*[^0-9０-９])|'  # 1. 2) など
            r'[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]|'  # 丸数字
            r'[(（][0-9０-９]+[)）]\s+'  # (1) （１） など
            r')'
        )
        
        # 各行が番号付き箇条書きかどうかを判定
        is_list_item = []
        for line in formatted_lines:
            # 書式マーカーを除去してパターンマッチ
            plain_line = re.sub(r'\*+([^*]+)\*+', r'\1', line)
            is_list_item.append(bool(numbered_pattern.match(plain_line)))
        
        # 連続する番号付き箇条書きを検出（2つ以上連続で有効）
        list_ranges = []
        i = 0
        while i < len(is_list_item):
            if is_list_item[i]:
                start = i
                while i < len(is_list_item) and is_list_item[i]:
                    i += 1
                if i - start >= 2:
                    list_ranges.append((start, i))
            else:
                i += 1
        
        # リスト範囲内の行を処理（マーカーから書式を除去）
        for start, end in list_ranges:
            for idx in range(start, end):
                line = formatted_lines[idx]
                # 行頭の書式マーカーを除去（リストマーカー部分のみ）
                # 例: "**1.** text" → "1. text"
                plain_line = re.sub(r'\*+([^*]+)\*+', r'\1', line)
                match = numbered_pattern.match(plain_line)
                if match:
                    marker = match.group(0)
                    rest = plain_line[len(marker):]
                    # 丸数字を番号に変換
                    circled_to_num = {
                        '①': '1. ', '②': '2. ', '③': '3. ', '④': '4. ', '⑤': '5. ',
                        '⑥': '6. ', '⑦': '7. ', '⑧': '8. ', '⑨': '9. ', '⑩': '10. ',
                        '⑪': '11. ', '⑫': '12. ', '⑬': '13. ', '⑭': '14. ', '⑮': '15. ',
                        '⑯': '16. ', '⑰': '17. ', '⑱': '18. ', '⑲': '19. ', '⑳': '20. '
                    }
                    for circled, num in circled_to_num.items():
                        if circled in marker:
                            marker = marker.replace(circled, num)
                            break
                    formatted_lines[idx] = marker + rest
        
        # 結果を構築（リスト範囲の後に空行を挿入）
        result_parts = []
        prev_line_end_char = ""
        in_list = False
        
        for i, line in enumerate(formatted_lines):
            # 現在の行がリスト範囲内かどうか
            current_in_list = any(start <= i < end for start, end in list_ranges)
            
            # リストが終了した場合、空行を挿入
            if in_list and not current_in_list:
                result_parts.append("\n")
            
            # リスト内の行は改行で区切る
            if current_in_list:
                if result_parts and not result_parts[-1].endswith("\n"):
                    result_parts.append("\n")
                result_parts.append(line)
            else:
                # 通常の行間の結合（日本語はスペースなし、英数字はスペースあり）
                if result_parts:
                    curr_char = line.lstrip()[0] if line.lstrip() else ""
                    if prev_line_end_char.isascii() and curr_char.isascii():
                        if prev_line_end_char.isalnum() and curr_char.isalnum():
                            result_parts.append(" ")
                result_parts.append(line)
            
            prev_line_end_char = line.rstrip()[-1] if line.rstrip() else ""
            in_list = current_in_list
        
        return "".join(result_parts).strip()

    def _convert_numbered_bullets(self, text: str) -> str:
        """番号付き箇条書きを検出してMarkdownリスト形式に変換
        
        行頭の丸数字（①②③など）および全角番号（１．２．など）を
        Markdownの番号付きリスト形式に変換する。
        本文中の参照（「前記①により」など）は変換しない。
        小数（例: 14.0%）は変換しない。
        
        Args:
            text: 入力テキスト
            
        Returns:
            変換後のテキスト
        """
        import re
        
        # 丸数字を番号に変換するマッピング
        circled_to_num = {
            '①': '1', '②': '2', '③': '3', '④': '4', '⑤': '5',
            '⑥': '6', '⑦': '7', '⑧': '8', '⑨': '9', '⑩': '10',
            '⑪': '11', '⑫': '12', '⑬': '13', '⑭': '14', '⑮': '15',
            '⑯': '16', '⑰': '17', '⑱': '18', '⑲': '19', '⑳': '20'
        }
        
        # 全角数字を半角数字に変換するマッピング
        fullwidth_to_halfwidth = {
            '０': '0', '１': '1', '２': '2', '３': '3', '４': '4',
            '５': '5', '６': '6', '７': '7', '８': '8', '９': '9'
        }
        
        # 行頭の丸数字パターン（行頭または改行直後の丸数字のみ）
        circled_pattern = re.compile(
            r'^([①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳])(\s*)(.*)$',
            re.MULTILINE
        )
        
        def replace_circled(match):
            """行頭の丸数字を番号付きリストに変換"""
            circled = match.group(1)
            space = match.group(2)
            rest = match.group(3)
            num = circled_to_num.get(circled, circled)
            if not space:
                space = ' '
            return f"{num}.{space}{rest}"
        
        # 行頭の全角番号パターン（例: 「１．借主は...」）
        # 小数を除外するため、区切り記号の後に数字が続かないことを確認
        fullwidth_pattern = re.compile(
            r'^([０-９]+)[．.)）](\s*)([^0-9０-９].*)$',
            re.MULTILINE
        )
        
        def replace_fullwidth(match):
            """行頭の全角番号を番号付きリストに変換"""
            fullwidth_num = match.group(1)
            space = match.group(2)
            rest = match.group(3)
            # 全角数字を半角に変換
            halfwidth_num = ''.join(
                fullwidth_to_halfwidth.get(c, c) for c in fullwidth_num
            )
            if not space:
                space = ' '
            return f"{halfwidth_num}.{space}{rest}"
        
        # 行頭の丸数字を変換
        result = circled_pattern.sub(replace_circled, text)
        # 行頭の全角番号を変換
        result = fullwidth_pattern.sub(replace_fullwidth, result)
        
        return result

    def _reflow_paragraphs_with_tables(
        self, lines: List[Dict], base_font_size: float, table_regions: List[Dict],
        line_based_tables: List[Dict] = None,
        figure_bboxes: List[Tuple[float, float, float, float]] = None,
        page_width: float = None,
        page_height: float = None,
        is_slide_document: bool = False
    ) -> List[Dict]:
        """段落リフロー（表領域を考慮）
        
        同一カラム内で縦方向のギャップが小さい行を結合する。
        表領域内の行は結合せず、Markdownテーブルとして出力する。
        番号付き見出しは単独ブロックとして確定する。
        番号付き箇条書きの後に続く非番号行は別ブロックとして分離する。
        図形として出力された領域と重なるテーブルはMarkdownテーブルとして出力しない。
        横長ページ（PPTスライド等）では改行を保持する。
        スライド文書では段落結合をスキップし、各行を別々のブロックとして出力する。
        
        Args:
            lines: ソートされた行データのリスト
            base_font_size: 基準フォントサイズ
            table_regions: 表領域のリスト
            line_based_tables: 罫線ベースで検出された表のリスト
            figure_bboxes: 図形として出力されたbboxのリスト（重複出力を防ぐため）
            page_width: ページ幅（横長ページ検出用）
            page_height: ページ高さ（横長ページ検出用）
            is_slide_document: スライド文書フラグ（Trueの場合、段落結合をスキップ）
            
        Returns:
            結合されたブロックのリスト
        """
        import re
        
        if line_based_tables is None:
            line_based_tables = []
        
        if figure_bboxes is None:
            figure_bboxes = []
        
        if not lines:
            return []
        
        # 番号付き箇条書きパターン（全角数字も含む）
        # 小数（例: 14.0%）を誤認識しないよう、区切り記号の後に数字が続かないことを確認
        # (N)形式の括弧付き番号も含む（半角・全角両対応）
        numbered_list_pattern = re.compile(
            r'^[\s]*('
            r'[0-9０-９]+[.)．）](?=\s*[^0-9０-９])|'  # 1. 2) など
            r'[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]|'  # 丸数字
            r'[(（][0-9０-９]+[)）]\s+'  # (1) （１） など
            r')'
        )
        
        def is_numbered_list_line(text: str) -> bool:
            """行が番号付き箇条書きで始まるかを判定"""
            return bool(numbered_list_pattern.match(text))
        
        def is_structured_field_line(text: str) -> bool:
            """構造化フィールド行かどうかを判定（ラベル＋多スペース＋値）"""
            t = text.strip()
            if not t:
                return False
            if "。" in t or "、" in t:
                return False
            if not re.search(r'[\u3000 ]{2,}', t):
                return False
            return bool(re.match(r'^\S{1,6}[\u3000 ]{2,}\S+', t))
        
        def is_in_table_region(line: Dict) -> Optional[Dict]:
            """行が表領域内にあるかチェック"""
            for region in table_regions:
                if (line["column"] == region["column"] and
                    region["y_start"] - 10 <= line["y"] <= region["y_end"] + 10):
                    return region
            return None
        
        # 罫線ベースの表をY座標でソート
        sorted_line_tables = sorted(line_based_tables, key=lambda t: t["y_start"])
        processed_line_tables = set()
        
        def get_line_based_table_by_y(line: Dict) -> Optional[Dict]:
            """行のy座標が罫線ベースの表領域内にあるかチェック（テーブル出力トリガ用）
            
            y座標のみで判定し、テーブルの出力タイミングを決定する。
            """
            line_y = line.get("y", 0)
            for table in sorted_line_tables:
                if table["y_start"] - 5 <= line_y <= table["y_end"] + 5:
                    return table
            return None
        
        def should_skip_line_for_table(line: Dict, table: Dict) -> bool:
            """行がテーブルの一部としてスキップすべきかチェック（行除外判定用）
            
            2段組の場合、テーブルのx範囲と行のx座標を比較して、
            同じカラム内の行のみをスキップする。
            """
            line_x = line.get("x", 0)
            table_bbox = table.get("bbox", (0, 0, 0, 0))
            table_x_start = table_bbox[0]
            table_x_end = table_bbox[2]
            # 行のx座標がテーブルのx範囲内にあるかチェック
            # 許容範囲を広めに設定（テーブル幅の20%程度）
            table_width = table_x_end - table_x_start
            margin = max(table_width * 0.2, 30)
            return table_x_start - margin <= line_x <= table_x_end + margin
        
        def table_overlaps_with_figure(table: Dict) -> bool:
            """テーブルが図形と重なるかチェック（重複出力防止用）
            
            テーブルのbboxが図形のbboxと80%以上重なる場合はTrueを返す。
            図形として出力された表はMarkdownテーブルとして出力しない。
            """
            if not figure_bboxes:
                return False
            
            table_bbox = table.get("bbox", (0, 0, 0, 0))
            table_x0, table_y0, table_x1, table_y1 = table_bbox
            table_area = (table_x1 - table_x0) * (table_y1 - table_y0)
            
            if table_area <= 0:
                return False
            
            for fig_bbox in figure_bboxes:
                fig_x0, fig_y0, fig_x1, fig_y1 = fig_bbox
                # 重なり領域を計算
                overlap_x0 = max(table_x0, fig_x0)
                overlap_y0 = max(table_y0, fig_y0)
                overlap_x1 = min(table_x1, fig_x1)
                overlap_y1 = min(table_y1, fig_y1)
                
                if overlap_x0 < overlap_x1 and overlap_y0 < overlap_y1:
                    overlap_area = (overlap_x1 - overlap_x0) * (overlap_y1 - overlap_y0)
                    # テーブルの80%以上が図形に含まれている場合
                    if overlap_area / table_area >= 0.8:
                        debug_print(f"[DEBUG] テーブルが図形と重なるためスキップ: table_bbox={table_bbox}, fig_bbox={fig_bbox}")
                        return True
            
            return False
        
        # スライド文書の場合は段落結合をスキップし、各行を別々のブロックとして出力
        if is_slide_document:
            debug_print("[DEBUG] スライド文書: 段落結合をスキップ")
            blocks = []
            for line in lines:
                text = line["text"].strip()
                bbox = line["bbox"]
                
                # ページ番号のフィルタリング（スライド文書のみ）
                # 条件: 数字のみ（1-3桁）、ページ下端（93%以上）、右端または左端（10%以内）、幅が小さい（5%未満）
                if page_width and page_height and text.isdigit() and len(text) <= 3:
                    x_center = (bbox[0] + bbox[2]) / 2
                    y_top = bbox[1]
                    width = bbox[2] - bbox[0]
                    
                    is_at_bottom = y_top > page_height * 0.93
                    is_at_edge = (x_center > page_width * 0.9) or (x_center < page_width * 0.1)
                    is_narrow = width < page_width * 0.05
                    
                    if is_at_bottom and is_at_edge and is_narrow:
                        debug_print(f"[DEBUG] ページ番号を除外: text=\"{text}\", y={y_top:.1f}, x_center={x_center:.1f}")
                        continue
                
                blocks.append({
                    "text": line["text"],
                    "bbox": line["bbox"],
                    "font_size": line["font_size"],
                    "is_bold": line["is_bold"],
                    "column": line["column"]
                })
            return blocks
        
        # 行高の推定（フォントサイズの1.2倍程度）
        line_height = base_font_size * 1.2
        # 結合する最大ギャップ
        # 閾値を大きくしすぎると段落が過剰に結合されるため、0.8倍に設定
        # 横長ページ（PPTスライド等）では改行を保持するため閾値を小さくする
        is_landscape = (page_width is not None and page_height is not None 
                       and page_width > page_height)
        if is_landscape:
            gap_threshold = line_height * 0.3
            debug_print(f"[DEBUG] 横長ページ検出: 改行保持モード（gap_threshold={gap_threshold:.1f}）")
        else:
            gap_threshold = line_height * 0.8
        
        blocks = []
        current_block = None
        processed_table_regions = set()
        
        i = 0
        while i < len(lines):
            line = lines[i]
            
            # 罫線ベースの表領域をチェック
            line_table = get_line_based_table_by_y(line)
            if line_table:
                table_id = (line_table["y_start"], line_table["y_end"])
                # テーブルがまだ出力されていない場合、出力する（y座標のみで判定）
                if table_id not in processed_line_tables:
                    # 図形と重なるテーブルはMarkdownテーブルとして出力しない
                    if not table_overlaps_with_figure(line_table):
                        # 現在のブロックを確定
                        if current_block:
                            blocks.append(self._finalize_block(current_block))
                            current_block = None
                        
                        # 罫線ベースの表をMarkdownテーブルとして出力
                        blocks.append({
                            "text": line_table["markdown"],
                            "bbox": line_table["bbox"],
                            "font_size": base_font_size,
                            "is_bold": False,
                            "is_table": True,
                            "column": "full"
                        })
                    processed_line_tables.add(table_id)
                
                # 行をスキップするかはx座標も含めて判定
                if should_skip_line_for_table(line, line_table):
                    i += 1
                    continue
                # x座標がテーブル範囲外の場合は、行を通常処理する（continueしない）
            
            table_region = is_in_table_region(line)
            
            if table_region:
                region_id = (table_region["column"], table_region["y_start"])
                if region_id not in processed_table_regions:
                    # 現在のブロックを確定
                    if current_block:
                        blocks.append(self._finalize_block(current_block))
                        current_block = None
                    
                    # 表をMarkdownテーブルとして出力（カラム情報も保持）
                    table_md = self._format_table_region(table_region)
                    if table_md:
                        blocks.append({
                            "text": table_md,
                            "bbox": (0, table_region["y_start"], 300, table_region["y_end"]),
                            "font_size": base_font_size,
                            "is_bold": False,
                            "is_table": True,
                            "column": table_region.get("column", "full")
                        })
                    processed_table_regions.add(region_id)
                i += 1
                continue
            
            # 番号付き見出しの検出
            is_heading, heading_level, heading_text = self._is_numbered_heading(line["text"])
            if is_heading:
                # 連続する番号付きリストかどうかを先読み・後読みで判定
                # 同じインデントで連番（1, 2, 3...）が続く場合はリストとして扱う
                is_consecutive_list = False
                line_number_match = re.match(r'^([0-9０-９]+)[．.\s]', line["text"])
                if line_number_match:
                    current_num = int(line_number_match.group(1).translate(
                        str.maketrans('０１２３４５６７８９', '0123456789')))
                    consecutive_count = 1
                    current_x = line["x"]
                    x_tolerance = base_font_size * 0.5
                    
                    # 前方向の先読み（前の行が連番の場合はリストとして扱う）
                    for j in range(i - 1, max(i - 10, -1), -1):
                        prev_line = lines[j]
                        if prev_line.get("column") != line.get("column"):
                            break
                        prev_match = re.match(r'^([0-9０-９]+)[．.\s]', prev_line["text"])
                        if prev_match:
                            prev_num = int(prev_match.group(1).translate(
                                str.maketrans('０１２３４５６７８９', '0123456789')))
                            # 連番かつ同じx位置
                            if prev_num == current_num - 1 and abs(prev_line["x"] - current_x) <= x_tolerance:
                                is_consecutive_list = True
                                break
                            else:
                                break
                        else:
                            # 継続行（インデントされた行）はスキップ
                            if prev_line["x"] > current_x + x_tolerance:
                                continue
                            break
                    
                    # 後続の行を先読み（前方向で見つからなかった場合）
                    if not is_consecutive_list:
                        for j in range(i + 1, min(i + 10, len(lines))):
                            next_line = lines[j]
                            if next_line.get("column") != line.get("column"):
                                break
                            next_match = re.match(r'^([0-9０-９]+)[．.\s]', next_line["text"])
                            if next_match:
                                next_num = int(next_match.group(1).translate(
                                    str.maketrans('０１２３４５６７８９', '0123456789')))
                                # 連番かつ同じx位置
                                if next_num == current_num + consecutive_count and abs(next_line["x"] - current_x) <= x_tolerance:
                                    consecutive_count += 1
                                    if consecutive_count >= 3:
                                        is_consecutive_list = True
                                        break
                                else:
                                    break
                            else:
                                # 継続行（インデントされた行）はスキップ
                                if next_line["x"] > current_x + x_tolerance:
                                    continue
                                break
                
                if is_consecutive_list:
                    # 連続リストの場合は見出しとして扱わない
                    pass
                else:
                    # 現在のブロックを確定
                    if current_block:
                        blocks.append(self._finalize_block(current_block))
                        current_block = None
                    
                    # 見出しの継続行をマージ
                    # 見出しが列幅いっぱいで折り返している場合、次の行を結合
                    merged_heading_text = heading_text
                    merged_bbox = list(line["bbox"])
                    
                    # 見出し行がカラム幅の大部分を占めているかを判定
                    # 短い見出しは折り返しが発生しないため、マージ不要
                    heading_width = line["bbox"][2] - line["bbox"][0]
                    
                    # 同じカラム内の行からカラム幅を推定
                    column_lines = [l for l in lines if l.get("column") == line.get("column")]
                    if column_lines:
                        column_left = min(l["bbox"][0] for l in column_lines)
                        column_right = max(l["bbox"][2] for l in column_lines)
                        column_width = column_right - column_left
                    else:
                        column_width = heading_width
                    
                    # 見出し行がカラム幅の70%以上を占める場合のみマージを許可
                    # （2段組の場合も列幅が短くなるため、比率で判定）
                    should_merge_continuation = (
                        column_width > 0 and 
                        heading_width / column_width >= 0.7
                    )
                    
                    j = i + 1
                    while should_merge_continuation and j < len(lines):
                        next_line = lines[j]
                        # 継続行の条件:
                        # 1. 同じカラム
                        # 2. Y方向のギャップが小さい（行高の1.5倍以内）
                        # 3. フォントサイズが同等
                        # 4. 次の行が番号付き見出しではない
                        # 5. 次の行が短い（折り返しの続き）
                        # 6. X座標の距離が小さい（同じ行の折り返しのみ）
                        if next_line.get("column") != line.get("column"):
                            break
                        y_gap = next_line["y"] - merged_bbox[3]
                        if y_gap > base_font_size * 1.5:
                            break
                        if abs(next_line["font_size"] - line["font_size"]) > 1:
                            break
                        next_is_heading, _, _ = self._is_numbered_heading(next_line["text"])
                        if next_is_heading:
                            break
                        # 次の行が短い場合のみマージ（長い本文は除外）
                        next_text = next_line["text"].strip()
                        if len(next_text) > 30:
                            break
                        # X座標の距離をチェック（同じY座標でも離れている場合はマージしない）
                        x_gap = next_line["bbox"][0] - merged_bbox[2]
                        if x_gap > base_font_size * 1.0:
                            break
                        # マージ
                        merged_heading_text += next_text
                        merged_bbox[2] = max(merged_bbox[2], next_line["bbox"][2])
                        merged_bbox[3] = next_line["bbox"][3]
                        j += 1
                    
                    # 見出しを単独ブロックとして追加（カラム情報も保持）
                    blocks.append({
                        "text": merged_heading_text,
                        "bbox": tuple(merged_bbox),
                        "font_size": line["font_size"],
                        "is_bold": True,
                        "is_heading": True,
                        "heading_level": heading_level,
                        "column": line["column"]
                    })
                    i = j
                    continue
            
            # 現在の行が番号付き箇条書きかどうか
            curr_is_numbered = is_numbered_list_line(line["text"])
            
            if current_block is None:
                current_block = {
                    "texts": [line["text"]],
                    "bbox": list(line["bbox"]),
                    "font_size": line["font_size"],
                    "is_bold": line["is_bold"],
                    "column": line["column"],
                    "last_y": line["bbox"][3],
                    "last_x": line["x"],
                    "is_list": curr_is_numbered,
                    "list_start_x": line["x"] if curr_is_numbered else None
                }
            else:
                # 結合条件をチェック
                y_gap = line["y"] - current_block["last_y"]
                same_column = line["column"] == current_block["column"]
                
                # X座標の距離をチェック（同じY座標でも離れている場合は別段落）
                # 前の行の右端と現在の行の左端の距離を計算
                prev_x_end = current_block["bbox"][2]
                curr_x_start = line["bbox"][0]
                x_gap = curr_x_start - prev_x_end
                # X方向のギャップがフォントサイズの1倍以上の場合は別段落
                # （同じY座標にある別々のテキスト要素を分離するため）
                x_gap_threshold = base_font_size * 1.0
                large_x_gap = x_gap > x_gap_threshold
                
                # 左端の位置が大きく異なる場合も別段落として扱う
                # （異なるX位置にある行を誤って結合しないため）
                prev_x_start = current_block["bbox"][0]
                x_start_diff = abs(curr_x_start - prev_x_start)
                # 左端の差がフォントサイズの3倍以上の場合は別段落
                # ただし、インデントされた継続行は除外（curr_x_start > prev_x_startの場合）
                large_x_start_diff = (
                    x_start_diff > base_font_size * 3.0 and
                    curr_x_start < prev_x_start  # 現在の行が前の行より左にある場合のみ
                )
                
                # 箇条書きブロックの終了判定
                # 現在のブロックが箇条書きで、次の行が番号付きでない場合
                # 継続行かどうかは「先読み」で判定：
                # - インデントされた行の次の行が左端に戻る場合は新しい段落
                # - インデントされた行の次の行も右寄りなら継続行
                list_boundary = False
                if current_block.get("is_list", False) and not curr_is_numbered:
                    list_start_x = current_block.get("list_start_x", current_block["last_x"])
                    indent_threshold = base_font_size * 0.5
                    left_tolerance = base_font_size * 0.3
                    
                    # 継続行は番号マーカーより右にインデントされている
                    is_indented = line["x"] > list_start_x + indent_threshold
                    
                    if is_indented:
                        # 先読み: 次の行が左端に戻るかを確認
                        next_line_returns_left = False
                        if i + 1 < len(lines):
                            next_line = lines[i + 1]
                            # 同一カラムの次の行のみ確認
                            if next_line.get("column") == line.get("column"):
                                # 次の行が左端（番号マーカーの位置）に戻るか
                                next_x_diff = abs(next_line["x"] - list_start_x)
                                next_line_returns_left = next_x_diff <= left_tolerance
                        
                        # 次の行が左端に戻る場合、現在の行は新しい段落の1行目
                        if next_line_returns_left:
                            list_boundary = True
                        else:
                            # 次の行も右寄りなら継続行として結合
                            list_boundary = False
                    else:
                        # インデントされていない（左端に戻った）場合は分離
                        list_boundary = True
                
                # 図表キャプション（「図N」「表N」）で始まるブロックは次の行と結合しない
                prev_is_caption = False
                if current_block.get("texts"):
                    first_text = current_block["texts"][0].strip()
                    if re.match(r'^(図|表)\s*[0-9０-９]+', first_text):
                        prev_is_caption = True
                
                # 現在の行が図表キャプションで始まる場合は新しいブロックとして分離
                curr_is_caption = bool(re.match(r'^(図|表)\s*[0-9０-９]+', line["text"].strip()))
                
                is_new_paragraph = (
                    y_gap > gap_threshold or
                    not same_column or
                    line["is_bold"] != current_block["is_bold"] or
                    abs(line["font_size"] - current_block["font_size"]) > 1 or
                    list_boundary or
                    curr_is_numbered or  # 新しい番号付き行は常に新しいブロック
                    prev_is_caption or  # 図表キャプションの後は新しいブロック
                    curr_is_caption or  # 図表キャプションは新しいブロックとして開始
                    large_x_gap or  # X方向のギャップが大きい場合は別段落
                    large_x_start_diff  # 左端の位置が大きく異なる場合は別段落
                )
                
                if is_new_paragraph:
                    blocks.append(self._finalize_block(current_block))
                    current_block = {
                        "texts": [line["text"]],
                        "bbox": list(line["bbox"]),
                        "font_size": line["font_size"],
                        "is_bold": line["is_bold"],
                        "column": line["column"],
                        "last_y": line["bbox"][3],
                        "last_x": line["x"],
                        "is_list": curr_is_numbered,
                        "list_start_x": line["x"] if curr_is_numbered else None
                    }
                else:
                    # 行を結合
                    prev_text = current_block["texts"][-1]
                    curr_text = line["text"]
                    
                    # 構造化フィールド行（ラベル＋多スペース＋値）の場合は改行を保持
                    if current_block.get("is_list", False) and is_structured_field_line(curr_text):
                        first_line = current_block["texts"][0]
                        marker_match = re.match(r'^(\s*[0-9０-９]+[.)．）]\s*)', first_line)
                        indent = "   " if marker_match else ""
                        current_block["texts"].append("\n" + indent + curr_text.strip())
                    elif prev_text and curr_text:
                        prev_char = prev_text.rstrip()[-1] if prev_text.rstrip() else ""
                        curr_char = curr_text.lstrip()[0] if curr_text.lstrip() else ""
                        
                        if prev_char.isascii() and curr_char.isascii():
                            if prev_char.isalnum() and curr_char.isalnum():
                                current_block["texts"].append(" " + curr_text)
                            else:
                                current_block["texts"].append(curr_text)
                        else:
                            current_block["texts"].append(curr_text)
                    else:
                        current_block["texts"].append(curr_text)
                    
                    current_block["bbox"][2] = max(current_block["bbox"][2], line["bbox"][2])
                    current_block["bbox"][3] = line["bbox"][3]
                    current_block["last_y"] = line["bbox"][3]
                    current_block["font_size"] = max(current_block["font_size"], line["font_size"])
                    if line["is_bold"]:
                        current_block["is_bold"] = True
            
            i += 1
        
        if current_block:
            blocks.append(self._finalize_block(current_block))
        
        return blocks
