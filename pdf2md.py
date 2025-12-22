#!/usr/bin/env python3
"""
PDF to Markdown Converter
PDFファイルをMarkdown形式に変換するツール

特徴:
- PDFの各ページを画像（PNG/SVG）として出力
- 埋め込みテキストの抽出
- manga-ocrによるOCRフォールバック対応
- o2mdファミリーとして統一されたインターフェース
"""

import os
import sys
import tempfile
import subprocess
from pathlib import Path
from typing import Optional, List, Dict, Any, Tuple, Set

from utils import get_libreoffice_path

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

# 設定
LIBREOFFICE_PATH = get_libreoffice_path()

# DPI設定
DEFAULT_DPI = 300
IMAGE_QUALITY = 95


# グローバルverboseフラグ
_VERBOSE = False

def set_verbose(verbose: bool):
    """verboseモードを設定"""
    global _VERBOSE
    _VERBOSE = verbose

def is_verbose() -> bool:
    """verboseモードかどうかを返す"""
    return _VERBOSE

def debug_print(*args, **kwargs):
    """verboseモード時のみ出力するデバッグ用print"""
    if _VERBOSE:
        print(*args, **kwargs)


class PDFToMarkdownConverter:
    """PDFファイルをMarkdown形式に変換するコンバータクラス
    
    o2mdファミリーとして統一されたインターフェースを提供します。
    """
    
    def __init__(
        self,
        pdf_file_path: str,
        output_dir: Optional[str] = None,
        output_format: str = 'png'
    ):
        """コンバータインスタンスの初期化
        
        Args:
            pdf_file_path: 変換するPDFファイルのパス
            output_dir: 出力ディレクトリ（省略時は./output）
            output_format: 出力画像形式 ('png' または 'svg')
        """
        self.pdf_file = pdf_file_path
        self.base_name = Path(pdf_file_path).stem
        
        if output_dir:
            self.output_dir = output_dir
        else:
            self.output_dir = os.path.join(os.getcwd(), "output")
        
        self.images_dir = os.path.join(self.output_dir, "images")
        
        self.output_format = output_format.lower() if output_format else 'png'
        
        # 出力形式の検証
        if self.output_format not in ('png', 'svg'):
            print(f"[WARNING] 不明な出力形式 '{output_format}'。'png'を使用します。")
            self.output_format = 'png'
        
        # ディレクトリ作成
        os.makedirs(self.output_dir, exist_ok=True)
        os.makedirs(self.images_dir, exist_ok=True)
        
        self.markdown_lines = []
        self.image_counter = 0
        
        # manga-ocrインスタンス（遅延初期化）
        self._ocr = None
        
        print(f"[INFO] 出力画像形式: {self.output_format.upper()}")
    
    def _get_ocr(self):
        """manga-ocrインスタンスを取得（遅延初期化）"""
        if self._ocr is None:
            try:
                from manga_ocr import MangaOcr
                self._ocr = MangaOcr()
                print("[INFO] manga-ocrを初期化しました")
            except ImportError:
                print("[WARNING] manga-ocrがインストールされていません。OCR機能は無効です。")
                print("[WARNING] インストール: pip install manga-ocr または uv sync")
                self._ocr = False
            except Exception as e:
                print(f"[WARNING] manga-ocrの初期化に失敗しました: {e}")
                self._ocr = False
        return self._ocr if self._ocr else None
    
    def convert(self) -> str:
        """メイン変換処理
        
        Returns:
            出力ファイルのパス
        """
        print(f"[INFO] PDF文書変換開始: {self.pdf_file}")
        
        # ドキュメントタイトルを先頭に追加
        self.markdown_lines.append(f"# {self.base_name}")
        self.markdown_lines.append("")
        
        try:
            doc = fitz.open(self.pdf_file)
        except Exception as e:
            print(f"[ERROR] PDFファイルを開けません: {self.pdf_file} - {e}")
            raise
        
        try:
            total_pages = len(doc)
            print(f"[INFO] 総ページ数: {total_pages}")
            
            # ヘッダ・フッタパターンを検出（全ページから収集）
            header_footer_patterns = self._detect_header_footer_patterns(doc)
            if header_footer_patterns:
                debug_print(f"[DEBUG] ヘッダ・フッタパターン検出: {len(header_footer_patterns)}個")
            
            for page_num in range(total_pages):
                print(f"[INFO] ページ {page_num + 1}/{total_pages} を処理中...")
                page = doc[page_num]
                self._convert_page(page, page_num, header_footer_patterns)
        finally:
            doc.close()
        
        # Markdownファイルを書き出し
        markdown_content = "\n".join(self.markdown_lines)
        output_file = os.path.join(self.output_dir, f"{self.base_name}.md")
        
        with open(output_file, "w", encoding="utf-8") as f:
            f.write(markdown_content)
        
        print(f"[SUCCESS] 変換完了: {output_file}")
        return output_file
    
    def _detect_header_footer_patterns(self, doc) -> Set[str]:
        """全ページからヘッダ・フッタパターンを検出
        
        ページ間で繰り返されるテキストをヘッダ・フッタとして検出する。
        数字は正規化して比較する。
        
        Args:
            doc: PyMuPDFのドキュメントオブジェクト
            
        Returns:
            ヘッダ・フッタパターンのセット
        """
        import re
        from collections import defaultdict
        
        # 各ページの上端・下端テキストを収集
        page_texts = defaultdict(list)
        
        for page_num in range(len(doc)):
            page = doc[page_num]
            page_height = page.rect.height
            
            try:
                text_dict = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)
            except Exception:
                continue
            
            for block in text_dict.get("blocks", []):
                if block.get("type") != 0:
                    continue
                
                bbox = block.get("bbox", (0, 0, 0, 0))
                y_top = bbox[1]
                y_bottom = bbox[3]
                
                # 上端10%または下端10%にあるテキスト
                is_header = y_top < page_height * 0.1
                is_footer = y_bottom > page_height * 0.9
                
                if is_header or is_footer:
                    for line in block.get("lines", []):
                        line_text = ""
                        for span in line.get("spans", []):
                            line_text += span.get("text", "")
                        
                        if line_text.strip():
                            # 数字を正規化（ページ番号など）
                            normalized = re.sub(r'\d+', '<NUM>', line_text.strip())
                            # 全角数字も正規化
                            normalized = re.sub(r'[０-９]+', '<NUM>', normalized)
                            page_texts[normalized].append(page_num)
        
        # 2ページ以上で出現するパターンをヘッダ・フッタとして認識
        patterns = set()
        for pattern, pages in page_texts.items():
            if len(set(pages)) >= 2:
                patterns.add(pattern)
        
        return patterns
    
    def _is_header_footer(
        self, text: str, patterns: Set[str], 
        y_pos: float = None, page_height: float = None, font_size: float = None
    ) -> bool:
        """テキストがヘッダ・フッタかどうかを判定
        
        Args:
            text: 判定するテキスト
            patterns: ヘッダ・フッタパターンのセット
            y_pos: 行のY座標（オプション）
            page_height: ページの高さ（オプション）
            font_size: フォントサイズ（オプション）
            
        Returns:
            ヘッダ・フッタの場合True
        """
        import re
        text_stripped = text.strip()
        
        # ページ番号パターンを直接検出（−21−、−21 −、-21-など）
        # マイナス記号（−, -, ー）で囲まれた数字
        if re.match(r'^[−\-ー]\s*\d+\s*[−\-ー]$', text_stripped):
            return True
        
        # 正規化してパターンマッチ
        normalized = re.sub(r'\d+', '<NUM>', text_stripped)
        normalized = re.sub(r'[０-９]+', '<NUM>', normalized)
        # スペースを正規化（複数スペースを1つに、前後のスペースを除去）
        normalized = re.sub(r'\s+', ' ', normalized).strip()
        
        # パターンも同様に正規化して比較
        for pattern in patterns:
            pattern_normalized = re.sub(r'\s+', ' ', pattern).strip()
            if normalized == pattern_normalized:
                return True
        
        # ページ下部（84%以下）のフッタキーワード検出
        footer_keywords = [
            '〒', 'E-mail', 'Accepted', 'received', 'Revisions',
            'Japan', 'INCORPORATED', 'Business Unit', 'Bissiness Unit',
            'Original manuscript', 'pp.', 'Vol.'
        ]
        
        if y_pos is not None and page_height is not None:
            # ページ下部（84%以下）にある行をチェック
            if y_pos > page_height * 0.84:
                # フッタキーワードを含む場合は除外
                for keyword in footer_keywords:
                    if keyword in text_stripped:
                        return True
        
        return False
    
    def _convert_page(self, page, page_num: int, header_footer_patterns: Set[str] = None):
        """PDFページを変換
        
        Args:
            page: PyMuPDFのページオブジェクト
            page_num: ページ番号（0始まり）
            header_footer_patterns: ヘッダ・フッタパターンのセット
        """
        if header_footer_patterns is None:
            header_footer_patterns = set()
        
        # 先にベクタ描画（図）を検出して、その領域を取得
        # 図領域内のテキストは本文から除外するため
        vector_figures = self._extract_vector_figures(page, page_num)
        figure_bboxes = [fig["bbox"] for fig in vector_figures]
        
        # テキストベースのPDFかどうかを判定（図領域を除外）
        text_blocks = self._extract_structured_text_v2(
            page, header_footer_patterns, exclude_bboxes=figure_bboxes
        )
        
        if text_blocks or vector_figures:
            # テキストベースのPDF: 構造化されたMarkdownを出力
            debug_print(f"[DEBUG] ページ {page_num + 1}: テキストベースPDFとして処理")
            
            # 統合版の図抽出を使用（ベクター図形と埋め込み画像を統合）
            # vector_figuresは既に_extract_all_figuresで統合処理済み
            all_images = vector_figures
            
            # 構造化テキストと画像を出力
            self._output_structured_markdown_with_images(text_blocks, all_images)
        else:
            # 画像ベースのPDF: 従来の画像+OCR処理
            debug_print(f"[DEBUG] ページ {page_num + 1}: 画像ベースPDFとして処理")
            image_path = self._render_page_as_image(page, page_num)
            if image_path:
                image_filename = os.path.basename(image_path)
                self.markdown_lines.append(f"![ページ {page_num + 1}](images/{image_filename})")
                self.markdown_lines.append("")
            
            # OCRでテキスト抽出
            ocr_text = self._ocr_page(page)
            if ocr_text and ocr_text.strip() and ocr_text != "(OCR利用不可)":
                self.markdown_lines.append("### 抽出テキスト（OCR）")
                self.markdown_lines.append("")
                for line in ocr_text.strip().split('\n'):
                    if line.strip():
                        self.markdown_lines.append(line.strip())
                self.markdown_lines.append("")
        
        # ページ区切り
        self.markdown_lines.append("---")
        self.markdown_lines.append("")
    
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
        exclude_bboxes: List[Tuple[float, float, float, float]] = None
    ) -> List[Dict[str, Any]]:
        """PDFページから構造化されたテキストブロックを抽出（改良版）
        
        行単位で抽出し、カラム分割と段落リフローを行う。
        ヘッダ・フッタを除外する。
        図領域内のテキストも除外する。
        
        Args:
            page: PyMuPDFのページオブジェクト
            header_footer_patterns: ヘッダ・フッタパターンのセット
            exclude_bboxes: 除外するbboxのリスト（図領域など）
            
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
        
        # 行単位でテキストを収集（span情報も保持）
        lines_data = []
        font_sizes = []
        
        for block in text_dict.get("blocks", []):
            if block.get("type") != 0:
                continue
            
            for line in block.get("lines", []):
                line_bbox = line.get("bbox", (0, 0, 0, 0))
                line_text = ""
                line_font_size = 0
                line_is_bold = False
                line_spans = []
                
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
                        
                        line_spans.append({
                            "text": text,
                            "size": span_size,
                            "bbox": span_bbox,
                            "is_bold": span_is_bold,
                            "is_italic": span_is_italic,
                            "is_superscript": span_is_superscript,
                            "is_subscript": span_is_subscript,
                            "is_strikethrough": span_is_strikethrough
                        })
                        line_text += text
                        line_font_size = max(line_font_size, span_size)
                        font_sizes.append(span_size)
                        if span_is_bold:
                            line_is_bold = True
                
                if not line_text.strip():
                    continue
                
                # ヘッダ・フッタを除外（Y座標とページ高さを渡す）
                if self._is_header_footer(
                    line_text, header_footer_patterns,
                    y_pos=line_bbox[1], page_height=page_height, font_size=line_font_size
                ):
                    debug_print(f"[DEBUG] ヘッダ・フッタ除外: {line_text.strip()[:30]}...")
                    continue
                
                # 図領域内のテキストを除外（中心点が図領域内にある場合）
                # ただし、キャプションパターン（図X、表X）は除外しない
                line_center_x = (line_bbox[0] + line_bbox[2]) / 2
                line_center_y = (line_bbox[1] + line_bbox[3]) / 2
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
                
                line_width = line_bbox[2] - line_bbox[0]
                x_center = (line_bbox[0] + line_bbox[2]) / 2
                
                # カラム判定: フル幅、左カラム、右カラム
                if line_width > page_width * 0.6:
                    column = "full"
                elif x_center < page_center:
                    column = "left"
                else:
                    column = "right"
                
                # 行全体の斜体フラグを計算
                line_is_italic = any(s.get("is_italic", False) for s in line_spans)
                
                lines_data.append({
                    "text": line_text,
                    "bbox": line_bbox,
                    "font_size": line_font_size,
                    "is_bold": line_is_bold,
                    "is_italic": line_is_italic,
                    "column": column,
                    "y": line_bbox[1],
                    "x": line_bbox[0],
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
        
        # カラムごとにソート（フル幅→左→右の順、各カラム内はY座標順）
        sorted_lines = self._sort_lines_by_column(lines_data)
        
        # 段落リフロー（同一カラム内で近接する行を結合、表領域は除外）
        reflowed_blocks = self._reflow_paragraphs_with_tables(
            sorted_lines, base_font_size, table_regions
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
                    "bbox": block_data["bbox"]
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
                block_data["bbox"]
            )
            block = {
                "type": block_type,
                "text": block_data["text"],
                "font_size": block_data["font_size"],
                "bbox": block_data["bbox"]
            }
            # カラム情報を保持（2段組みの順序維持に必要）
            if "column" in block_data:
                block["column"] = block_data["column"]
            blocks.append(block)
        
        return blocks
    
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
    
    def _sort_lines_by_column(self, lines_data: List[Dict]) -> List[Dict]:
        """行をカラムごとにソート
        
        フル幅要素を基準に、その間の区間で左→右の順に出力する。
        
        Args:
            lines_data: 行データのリスト
            
        Returns:
            ソートされた行データのリスト
        """
        # フル幅行とカラム行を分離
        full_lines = [l for l in lines_data if l["column"] == "full"]
        left_lines = [l for l in lines_data if l["column"] == "left"]
        right_lines = [l for l in lines_data if l["column"] == "right"]
        
        # 各グループをY座標でソート
        full_lines.sort(key=lambda x: x["y"])
        left_lines.sort(key=lambda x: x["y"])
        right_lines.sort(key=lambda x: x["y"])
        
        # フル幅行がない場合は単純に左→右
        if not full_lines:
            return left_lines + right_lines
        
        # フル幅行を基準に区間を作成
        result = []
        full_y_positions = [l["y"] for l in full_lines]
        full_y_positions = [-float('inf')] + full_y_positions + [float('inf')]
        
        for i in range(len(full_y_positions) - 1):
            y_start = full_y_positions[i]
            y_end = full_y_positions[i + 1]
            
            # この区間のフル幅行を追加
            if i > 0:
                for fl in full_lines:
                    if abs(fl["y"] - y_start) < 1:
                        result.append(fl)
            
            # この区間の左カラム行を追加
            for ll in left_lines:
                if y_start < ll["y"] < y_end:
                    result.append(ll)
            
            # この区間の右カラム行を追加
            for rl in right_lines:
                if y_start < rl["y"] < y_end:
                    result.append(rl)
        
        # 最後のフル幅行を追加
        if full_lines:
            last_full = full_lines[-1]
            if last_full not in result:
                result.append(last_full)
        
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
        
        # 番号付き箇条書きパターン
        numbered_list_pattern = re.compile(
            r'^[\s]*(\d+[.)．]\s|[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳])'
        )
        
        def is_numbered_list_line(text: str) -> bool:
            """行が番号付き箇条書きで始まるかを判定"""
            return bool(numbered_list_pattern.match(text))
        
        # 行高の推定（フォントサイズの1.2倍程度）
        line_height = base_font_size * 1.2
        gap_threshold = line_height * 0.8  # 結合する最大ギャップ
        
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
        """
        # span情報を使って書式付きテキストを生成
        spans_list = block_data.get("spans_list", [])
        if spans_list and any(spans_list):
            text = self._apply_text_formatting(spans_list)
        else:
            # span情報がない場合は従来通り
            text = "".join(block_data["texts"]).strip()
        
        # 番号付き箇条書きの検出と変換
        text = self._convert_numbered_bullets(text)
        
        result = {
            "text": text,
            "bbox": tuple(block_data["bbox"]),
            "font_size": block_data["font_size"],
            "is_bold": block_data["is_bold"],
            "is_italic": block_data.get("is_italic", False)
        }
        
        # カラム情報を保持（2段組みの順序維持に必要）
        if "column" in block_data:
            result["column"] = block_data["column"]
        
        return result
    
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
        
        # 番号付き箇条書きパターン（行頭の数字+区切り+空白）
        # 例: "1. ", "2) ", "1．", "①", "②"
        numbered_pattern = re.compile(r'^[\s]*(\d+[.)．]\s|[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳])')
        
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
        
        ①②③④などの丸数字や、1. 2. 3.などの番号付きマーカーを検出し、
        Markdownの番号付きリスト形式に変換する。
        
        Args:
            text: 入力テキスト
            
        Returns:
            変換後のテキスト
        """
        import re
        
        # 丸数字のパターン（①〜⑳）
        circled_pattern = r'([①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳])'
        
        # 丸数字が2つ以上含まれているか確認
        circled_matches = re.findall(circled_pattern, text)
        if len(circled_matches) >= 2:
            # 丸数字を番号に変換するマッピング
            circled_to_num = {
                '①': '1', '②': '2', '③': '3', '④': '4', '⑤': '5',
                '⑥': '6', '⑦': '7', '⑧': '8', '⑨': '9', '⑩': '10',
                '⑪': '11', '⑫': '12', '⑬': '13', '⑭': '14', '⑮': '15',
                '⑯': '16', '⑰': '17', '⑱': '18', '⑲': '19', '⑳': '20'
            }
            
            # 丸数字で分割
            parts = re.split(circled_pattern, text)
            
            # 結果を構築
            result_lines = []
            current_num = None
            current_text = ""
            
            for part in parts:
                if part in circled_to_num:
                    # 前の項目を保存
                    if current_num is not None and current_text.strip():
                        result_lines.append(f"{current_num}. {current_text.strip()}")
                    current_num = circled_to_num[part]
                    current_text = ""
                else:
                    current_text += part
            
            # 最後の項目を保存
            if current_num is not None and current_text.strip():
                result_lines.append(f"{current_num}. {current_text.strip()}")
            
            if result_lines:
                return "\n".join(result_lines)
        
        return text
    
    def _detect_table_regions(
        self, lines_data: List[Dict], page_center: float
    ) -> List[Dict]:
        """カラム内の表領域を検出
        
        同じY座標に複数のセルがある行が連続する領域を表として検出する。
        
        Args:
            lines_data: 行データのリスト
            page_center: ページの中央X座標
            
        Returns:
            表領域のリスト（各領域は{y_start, y_end, column, rows}を含む）
        """
        import re as re_mod
        # キャプションパターン（図X、表Xで始まる行は表検出から除外）
        caption_pattern = re_mod.compile(r'^[図表]\s*\d+')
        
        # キャプション行を除外してから左右カラムを分離
        filtered_lines = [l for l in lines_data 
                         if not caption_pattern.match(l.get("text", ""))]
        left_lines = [l for l in filtered_lines if l["column"] == "left"]
        right_lines = [l for l in filtered_lines if l["column"] == "right"]
        
        table_regions = []
        
        for column_lines, column_name in [(left_lines, "left"), (right_lines, "right")]:
            if not column_lines:
                continue
            
            # Y座標でグループ化（同じ行にある要素を検出）
            y_tolerance = 5  # Y座標の許容誤差
            y_groups = {}
            
            for line in column_lines:
                y_key = round(line["y"] / y_tolerance) * y_tolerance
                if y_key not in y_groups:
                    y_groups[y_key] = []
                y_groups[y_key].append(line)
            
            # 複数セルがある行を検出（3セル以上で表として認識）
            # ただし、短いテキスト（3文字以下）だけで構成される行は除外（図内ラベルの誤検出防止）
            multi_cell_rows = []
            all_rows = []  # 全ての行（単一セル含む）
            for y_key in sorted(y_groups.keys()):
                cells = y_groups[y_key]
                # X座標でソートして、異なるX位置にあるセルをカウント
                x_positions = sorted(set(round(c["x"] / 20) * 20 for c in cells))
                # 短いテキストだけで構成される行は除外
                texts = [c.get("text", "") for c in cells]
                has_long_text = any(len(t) > 3 for t in texts)
                is_multi = len(x_positions) >= 3 and has_long_text
                row_data = {
                    "y": y_key,
                    "cells": sorted(cells, key=lambda c: c["x"]),
                    "x_positions": x_positions,
                    "is_multi_cell": is_multi
                }
                all_rows.append(row_data)
                if is_multi:
                    multi_cell_rows.append(row_data)
            
            # 連続する複数セル行を表領域としてグループ化
            if not multi_cell_rows:
                continue
            
            current_region = {
                "y_start": multi_cell_rows[0]["y"],
                "y_end": multi_cell_rows[0]["y"] + 20,
                "column": column_name,
                "rows": [multi_cell_rows[0]],
                "all_rows": all_rows  # 単一セル行も含む全行を保持
            }
            
            for i in range(1, len(multi_cell_rows)):
                row = multi_cell_rows[i]
                prev_row = multi_cell_rows[i - 1]
                
                # 連続している場合（Y座標の差が小さい）
                if row["y"] - prev_row["y"] < 50:  # 許容範囲を広げる
                    current_region["rows"].append(row)
                    current_region["y_end"] = row["y"] + 20
                else:
                    # 2行以上の連続した複数セル行があれば表として認識
                    if len(current_region["rows"]) >= 2:
                        table_regions.append(current_region)
                    
                    current_region = {
                        "y_start": row["y"],
                        "y_end": row["y"] + 20,
                        "column": column_name,
                        "rows": [row],
                        "all_rows": all_rows
                    }
            
            # 最後の領域をチェック
            if len(current_region["rows"]) >= 2:
                table_regions.append(current_region)
        
        return table_regions
    
    def _is_numbered_heading(self, text: str) -> Tuple[bool, int, str]:
        """番号付き見出しかどうかを判定
        
        「1　はじめに」「2.1　概要」などのパターンを検出する。
        
        Args:
            text: 判定するテキスト
            
        Returns:
            (見出しかどうか, 見出しレベル, 見出しテキスト)
        """
        import re
        text = text.strip()
        
        # 長すぎる行は見出しではない
        if len(text) > 50:
            return (False, 0, "")
        
        # 末尾が句点で終わる場合は見出しではない
        if text.endswith("。") or text.endswith("．"):
            return (False, 0, "")
        
        # 番号付き見出しパターン: 「1　はじめに」「2.1　概要」など
        # 数字 + (ドット + 数字)* + 全角/半角スペース + タイトル
        match = re.match(r'^(\d+(?:[\.．]\d+)*)\s*[　 ]+(.{1,40})$', text)
        if match:
            number_part = match.group(1)
            title_part = match.group(2).strip()
            
            # タイトル部分が短すぎる場合は除外
            if len(title_part) < 2:
                return (False, 0, "")
            
            # タイトル部分が日本語の見出しらしい文字で始まる場合のみ見出しとして認識
            # 「年」「月」「日」「倍」などの単位で始まる場合は除外
            first_char = title_part[0]
            excluded_chars = '年月日倍個回分秒時点番号件台人円万億兆'
            if first_char in excluded_chars:
                return (False, 0, "")
            
            # 見出しレベルを決定（ドットの数 + 1）
            level = number_part.count('.') + number_part.count('．') + 1
            
            return (True, level, title_part)
        
        return (False, 0, "")
    
    def _reflow_paragraphs_with_tables(
        self, lines: List[Dict], base_font_size: float, table_regions: List[Dict]
    ) -> List[Dict]:
        """段落リフロー（表領域を考慮）
        
        同一カラム内で縦方向のギャップが小さい行を結合する。
        表領域内の行は結合せず、Markdownテーブルとして出力する。
        番号付き見出しは単独ブロックとして確定する。
        番号付き箇条書きの後に続く非番号行は別ブロックとして分離する。
        
        Args:
            lines: ソートされた行データのリスト
            base_font_size: 基準フォントサイズ
            table_regions: 表領域のリスト
            
        Returns:
            結合されたブロックのリスト
        """
        import re
        
        if not lines:
            return []
        
        # 番号付き箇条書きパターン
        numbered_list_pattern = re.compile(
            r'^[\s]*(\d+[.)．]\s|[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳])'
        )
        
        def is_numbered_list_line(text: str) -> bool:
            """行が番号付き箇条書きで始まるかを判定"""
            return bool(numbered_list_pattern.match(text))
        
        def is_in_table_region(line: Dict) -> Optional[Dict]:
            """行が表領域内にあるかチェック"""
            for region in table_regions:
                if (line["column"] == region["column"] and
                    region["y_start"] - 10 <= line["y"] <= region["y_end"] + 10):
                    return region
            return None
        
        # 行高の推定（フォントサイズの1.2倍程度）
        line_height = base_font_size * 1.2
        gap_threshold = line_height * 0.8
        
        blocks = []
        current_block = None
        processed_table_regions = set()
        
        i = 0
        while i < len(lines):
            line = lines[i]
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
                # 現在のブロックを確定
                if current_block:
                    blocks.append(self._finalize_block(current_block))
                    current_block = None
                
                # 見出しを単独ブロックとして追加（カラム情報も保持）
                blocks.append({
                    "text": heading_text,
                    "bbox": tuple(line["bbox"]),
                    "font_size": line["font_size"],
                    "is_bold": True,
                    "is_heading": True,
                    "heading_level": heading_level,
                    "column": line["column"]
                })
                i += 1
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
                
                is_new_paragraph = (
                    y_gap > gap_threshold or
                    not same_column or
                    line["is_bold"] != current_block["is_bold"] or
                    abs(line["font_size"] - current_block["font_size"]) > 1 or
                    list_boundary or
                    curr_is_numbered  # 新しい番号付き行は常に新しいブロック
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
                    
                    if prev_text and curr_text:
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
    
    def _format_table_region(self, table_region: Dict) -> str:
        """表領域をMarkdownテーブル形式に変換
        
        単一セル行（継続行）も含めて処理し、適切にマージする。
        
        Args:
            table_region: 表領域データ
            
        Returns:
            Markdownテーブル文字列
        """
        rows = table_region.get("rows", [])
        all_rows = table_region.get("all_rows", [])
        if not rows:
            return ""
        
        y_start = table_region.get("y_start", 0)
        y_end = table_region.get("y_end", 0)
        
        # 表領域内の全行を取得（単一セル行も含む）
        table_all_rows = []
        for row in all_rows:
            if y_start - 10 <= row["y"] <= y_end + 10:
                table_all_rows.append(row)
        
        if not table_all_rows:
            table_all_rows = rows
        
        # 複数セル行からヘッダ行を特定（最初の複数セル行）
        header_row = None
        for row in table_all_rows:
            if row.get("is_multi_cell", False):
                header_row = row
                break
        
        if not header_row:
            return ""
        
        # ヘッダ行の列位置を基準にする
        column_positions = sorted(header_row.get("x_positions", []))
        if len(column_positions) < 2:
            return ""
        
        # 各行のセルを列に割り当て
        table_rows = []
        for row in table_all_rows:
            cells = row.get("cells", [])
            row_data = [""] * len(column_positions)
            
            for cell in cells:
                cell_x = round(cell["x"] / 20) * 20
                # 最も近い列に割り当て
                min_dist = float("inf")
                best_idx = 0
                for idx, pos in enumerate(column_positions):
                    dist = abs(cell_x - pos)
                    if dist < min_dist:
                        min_dist = dist
                        best_idx = idx
                
                cell_text = cell["text"].strip()
                if not row_data[best_idx]:
                    row_data[best_idx] = cell_text
                else:
                    row_data[best_idx] += " " + cell_text
            
            table_rows.append({
                "data": row_data,
                "is_multi_cell": row.get("is_multi_cell", False),
                "y": row["y"]
            })
        
        if not table_rows:
            return ""
        
        # 継続行をマージ（単一セル行を前の行にマージ）
        merged_rows = []
        for row in table_rows:
            if row["is_multi_cell"]:
                merged_rows.append(row["data"])
            else:
                # 単一セル行: 前の行にマージ
                if merged_rows:
                    prev_row = merged_rows[-1]
                    for i, cell in enumerate(row["data"]):
                        if cell:
                            if prev_row[i]:
                                prev_row[i] += " " + cell
                            else:
                                prev_row[i] = cell
                else:
                    merged_rows.append(row["data"])
        
        if not merged_rows:
            return ""
        
        # Markdownテーブルを生成
        md_lines = []
        
        # ヘッダ行
        header = merged_rows[0]
        md_lines.append("| " + " | ".join(header) + " |")
        
        # 区切り行
        md_lines.append("| " + " | ".join(["---"] * len(header)) + " |")
        
        # データ行
        for row in merged_rows[1:]:
            md_lines.append("| " + " | ".join(row) + " |")
        
        return "\n".join(md_lines)
    
    def _detect_table_in_fullwidth(
        self, text_dict: Dict, header_footer_patterns: Set[str]
    ) -> List[List[Dict]]:
        """フル幅領域での表検出
        
        段組みページでも、フル幅領域（ページ幅の60%以上）にある
        表構造を検出する。
        
        Args:
            text_dict: PyMuPDFのテキスト辞書
            header_footer_patterns: ヘッダ・フッタパターン
            
        Returns:
            表の行リスト
        """
        page_width = text_dict.get("width", 612)
        
        # フル幅領域の行を収集
        y_groups: Dict[int, List[Dict]] = {}
        
        for block in text_dict.get("blocks", []):
            if block.get("type") != 0:
                continue
            
            block_bbox = block.get("bbox", (0, 0, 0, 0))
            block_width = block_bbox[2] - block_bbox[0]
            
            # フル幅ブロックのみ対象
            if block_width < page_width * 0.6:
                continue
            
            for line in block.get("lines", []):
                bbox = line.get("bbox", (0, 0, 0, 0))
                y_key = round(bbox[1] / 8) * 8  # 8ピクセル単位でグループ化
                
                line_text = ""
                for span in line.get("spans", []):
                    line_text += span.get("text", "")
                
                if line_text.strip():
                    # ヘッダ・フッタを除外
                    if self._is_header_footer(line_text, header_footer_patterns):
                        continue
                    
                    if y_key not in y_groups:
                        y_groups[y_key] = []
                    y_groups[y_key].append({
                        "text": line_text.strip(),
                        "x": bbox[0],
                        "bbox": bbox
                    })
        
        # 複数のセルがある行を表の行として抽出
        table_rows = []
        for y_key in sorted(y_groups.keys()):
            cells = y_groups[y_key]
            if len(cells) >= 2:
                cells_sorted = sorted(cells, key=lambda c: c["x"])
                avg_text_len = sum(len(c["text"]) for c in cells_sorted) / len(cells_sorted)
                if avg_text_len < 50:
                    table_rows.append(cells_sorted)
        
        # 表として認識する条件
        if len(table_rows) >= 3:
            col_counts = [len(row) for row in table_rows]
            most_common_cols = max(set(col_counts), key=col_counts.count)
            consistent_rows = sum(1 for c in col_counts if c == most_common_cols)
            if consistent_rows / len(table_rows) >= 0.8:
                return table_rows
        
        return []
    
    def _extract_all_figures(self, page, page_num: int) -> List[Dict[str, Any]]:
        """ベクター図形と埋め込み画像を統合して図を抽出
        
        ベクター描画と埋め込み画像を統合してクラスタリングし、
        クラスタリング後にカラム判定を行う（先にクラスタリング、後でカラム判定）。
        
        Args:
            page: PyMuPDFのページオブジェクト
            page_num: ページ番号
            
        Returns:
            抽出された図の情報リスト
        """
        figures = []
        page_width = page.rect.width
        gutter_x = page_width / 2
        gutter_margin = 10.0  # ガター跨ぎ判定のマージン
        
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
                for img_rect in page.get_image_rects(xref):
                    bbox = (img_rect.x0, img_rect.y0, img_rect.x1, img_rect.y1)
                    area = (bbox[2] - bbox[0]) * (bbox[3] - bbox[1])
                    if area >= 100:
                        all_elements.append({"bbox": bbox, "type": "image"})
                        break
        except Exception as e:
            debug_print(f"[DEBUG] 画像取得エラー: {e}")
        
        # 互換性のためにall_bboxesも作成
        all_bboxes = [e["bbox"] for e in all_elements]
        
        if len(all_bboxes) < 2:
            return []
        
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
            
            all_figure_candidates.append({
                "union_bbox": union_bbox,
                "raw_union_bbox": raw_union_bbox,
                "cluster_size": len(cluster),
                "column": column,
                "is_embedded": is_embedded
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
        
        # 表領域を検出
        table_bboxes = detect_table_bboxes_from_text(page_text_lines, page_width)
        
        if table_bboxes:
            debug_print(f"[DEBUG] page={page_num+1}: 表領域を{len(table_bboxes)}個検出")
            for i, tb in enumerate(table_bboxes):
                debug_print(f"[DEBUG]   表{i+1}: y={tb[1]:.1f}-{tb[3]:.1f}, x={tb[0]:.1f}-{tb[2]:.1f}")
        
        # 表領域と重なる図候補を除外
        debug_print(f"[DEBUG] page={page_num+1}: 表フィルタ前の候補数={len(all_figure_candidates)}")
        table_filtered = []
        for cand in all_figure_candidates:
            is_table = False
            cand_bbox = cand["union_bbox"]
            for table_bbox in table_bboxes:
                overlap = bbox_overlap_ratio(cand_bbox, table_bbox)
                if overlap >= 0.7:
                    debug_print(f"[DEBUG] page={page_num+1}: 図候補を表として除外: overlap={overlap:.2f}, cand_bbox=({cand_bbox[0]:.1f}, {cand_bbox[1]:.1f}, {cand_bbox[2]:.1f}, {cand_bbox[3]:.1f})")
                    is_table = True
                    break
            if not is_table:
                table_filtered.append(cand)
        
        debug_print(f"[DEBUG] page={page_num+1}: 表フィルタ後の候補数={len(table_filtered)}")
        all_figure_candidates = table_filtered
        
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
            y_min = min(bbox1[3], bbox2[3])
            y_max = max(bbox1[1], bbox2[1])
            x_overlap_start = max(bbox1[0], bbox2[0])
            x_overlap_end = min(bbox1[2], bbox2[2])
            
            for line in text_lines:
                line_bbox = line["bbox"]
                line_center_y = (line_bbox[1] + line_bbox[3]) / 2
                line_center_x = (line_bbox[0] + line_bbox[2]) / 2
                line_width = line_bbox[2] - line_bbox[0]
                
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
                        "column": cand1["column"]
                    })
                    used.add(i)
                    used.add(best_merge)
                    merged = True
                else:
                    new_candidates.append(cand1)
                    used.add(i)
            
            all_figure_candidates = new_candidates
        
        debug_print(f"[DEBUG] クラスタ再マージ完了: {len(all_figure_candidates)}個のクラスタ")
        
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
        
        def compute_clip_bbox(graphics_bbox, text_lines, col_width, page_height, column, is_embedded_image=False):
            """graphics_bboxからclip_bboxを計算（トリム処理）
            
            本文の下〜キャプションの上でトリムする。
            graphics_bboxを侵食してでも、境界を正しく設定する。
            カラム境界でクランプして他段のテキストが混入しないようにする。
            埋め込み画像の場合は最小限のpadding（1pt）を使用。
            """
            # 埋め込み画像の場合は最小限のpadding（1pt）を使用
            padding = 1.0 if is_embedded_image else 20.0
            clip_x0 = max(0, graphics_bbox[0] - padding)
            clip_y0 = max(0, graphics_bbox[1] - padding)
            clip_x1 = min(page_width, graphics_bbox[2] + padding)
            clip_y1 = min(page_height, graphics_bbox[3] + padding)
            
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
                
                # 埋め込み画像の場合はraw_union_bbox（パディングなし）を使用
                if is_embedded_image:
                    graphics_bbox = fig_info.get("raw_union_bbox", fig_info["union_bbox"])
                else:
                    graphics_bbox = fig_info["union_bbox"]
                
                column = fig_info["column"]
                union_bbox = graphics_bbox
                
                # clip_bboxを計算（トリム処理）
                clip_bbox = compute_clip_bbox(graphics_bbox, page_text_lines, col_width, page.rect.height, column, is_embedded_image)
                
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
                
                figure_texts, expanded_bbox = self._extract_text_in_bbox(
                    page, union_bbox, expand_for_labels=True, column=column, gutter_x=gutter_x
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
    
    def _extract_vector_figures(self, page, page_num: int) -> List[Dict[str, Any]]:
        """ベクタ描画（図）を抽出（統合版を使用）
        
        Args:
            page: PyMuPDFのページオブジェクト
            page_num: ページ番号
            
        Returns:
            抽出された図の情報リスト
        """
        return self._extract_all_figures(page, page_num)
    
    def _extract_text_in_bbox(
        self, page, bbox: Tuple[float, float, float, float],
        expand_for_labels: bool = True,
        column: str = None,
        gutter_x: float = None
    ) -> Tuple[List[str], Tuple[float, float, float, float]]:
        """指定されたbbox内のテキストを抽出
        
        図のラベルテキストを含めるため、bboxを近傍の短いテキストで拡張する。
        カラム境界を考慮して、隣のカラムのテキストを取り込まないようにする。
        
        Args:
            page: PyMuPDFのページオブジェクト
            bbox: バウンディングボックス (x0, y0, x1, y1)
            expand_for_labels: ラベルテキストを含めるためにbboxを拡張するか
            column: 図のカラム ("left", "right", "full")
            gutter_x: カラム境界のX座標
            
        Returns:
            (抽出されたテキストのリスト, 拡張後のbbox)
        """
        import re
        texts = []
        expanded_bbox = bbox
        
        if gutter_x is None:
            gutter_x = page.rect.width / 2
        
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
            
            for block in text_dict.get("blocks", []):
                if block.get("type") != 0:
                    continue
                
                for line in block.get("lines", []):
                    line_bbox = line.get("bbox", (0, 0, 0, 0))
                    
                    line_center_x = (line_bbox[0] + line_bbox[2]) / 2
                    line_center_y = (line_bbox[1] + line_bbox[3]) / 2
                    
                    if not is_in_same_column(line_center_x):
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
                            texts.append(line_text_stripped)
        
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
    
    def _classify_block_type(
        self, text: str, font_size: float, base_size: float, 
        is_bold: bool, bbox: Tuple[float, float, float, float]
    ) -> str:
        """テキストブロックのタイプを分類
        
        Args:
            text: テキスト内容
            font_size: フォントサイズ
            base_size: 基準フォントサイズ
            is_bold: 太字かどうか
            bbox: バウンディングボックス
            
        Returns:
            ブロックタイプ ('heading1', 'heading2', 'heading3', 'paragraph', 'list_item')
        """
        text_stripped = text.strip()
        
        # 箇条書きの検出（マーカーの後に空白がある場合のみ）
        import re
        
        # 記号マーカー + 空白のパターン（--で始まる行は除外）
        if not text_stripped.startswith('--'):
            # 空白必須のマーカー（-, *, など）
            if re.match(r'^[\-\*]\s+', text_stripped):
                return "list_item"
            # 空白不要のマーカー（•, ・, ○, ● など）
            bullet_markers = ['•', '・', '○', '●', '◆', '◇', '▪', '▫', '－', '―', '＊']
            for marker in bullet_markers:
                if text_stripped.startswith(marker):
                    return "list_item"
        
        # 番号付きリストの検出（区切り記号の後に空白が必須）
        if re.match(r'^[\d０-９]+[\.．\)）]\s+', text_stripped):
            return "list_item"
        
        # 見出しの検出（フォントサイズと太字に基づく）
        size_ratio = font_size / base_size if base_size > 0 else 1.0
        
        if size_ratio >= 1.8 or (size_ratio >= 1.5 and is_bold):
            return "heading1"
        elif size_ratio >= 1.4 or (size_ratio >= 1.2 and is_bold):
            return "heading2"
        elif size_ratio >= 1.15 or is_bold:
            return "heading3"
        
        return "paragraph"
    
    def _detect_column_layout(self, text_dict: Dict) -> int:
        """段組み（カラム）レイアウトを検出
        
        テキストブロックのX座標分布から段組みを判定する。
        
        Args:
            text_dict: PyMuPDFのテキスト辞書
            
        Returns:
            カラム数（1=単一カラム、2=2段組み）
        """
        # ページ幅を取得
        page_width = text_dict.get("width", 612)
        page_center = page_width / 2
        
        # 各行のX座標を収集
        x_positions = []
        for block in text_dict.get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                bbox = line.get("bbox", (0, 0, 0, 0))
                x_start = bbox[0]
                x_end = bbox[2]
                line_width = x_end - x_start
                # 短すぎる行は除外（見出しや箇条書きマーカーなど）
                if line_width > page_width * 0.2:
                    x_positions.append(x_start)
        
        if len(x_positions) < 5:
            return 1
        
        # 左半分と右半分に分類
        left_count = sum(1 for x in x_positions if x < page_center * 0.6)
        right_count = sum(1 for x in x_positions if x > page_center * 0.8)
        
        # 両方に一定数以上の行があれば2段組みと判定
        if left_count >= 3 and right_count >= 3:
            debug_print(f"[DEBUG] 2段組みレイアウトを検出 (左: {left_count}, 右: {right_count})")
            return 2
        
        return 1
    
    def _detect_table_structure(self, text_dict: Dict) -> List[List[Dict]]:
        """表構造を検出
        
        同じY座標に複数のテキストブロックがある場合、表として検出する。
        段組みレイアウトの場合は表検出を無効化する。
        
        Args:
            text_dict: PyMuPDFのテキスト辞書
            
        Returns:
            表の行リスト（各行はセルのリスト）
        """
        # 段組みレイアウトの場合は表検出をスキップ
        column_count = self._detect_column_layout(text_dict)
        if column_count >= 2:
            debug_print("[DEBUG] 段組みレイアウトのため表検出をスキップ")
            return []
        
        page_width = text_dict.get("width", 612)
        
        # Y座標でグループ化
        y_groups: Dict[int, List[Dict]] = {}
        
        for block in text_dict.get("blocks", []):
            if block.get("type") != 0:
                continue
            
            for line in block.get("lines", []):
                bbox = line.get("bbox", (0, 0, 0, 0))
                y_key = round(bbox[1] / 5) * 5  # 5ピクセル単位でグループ化
                
                line_text = ""
                for span in line.get("spans", []):
                    line_text += span.get("text", "")
                
                if line_text.strip():
                    if y_key not in y_groups:
                        y_groups[y_key] = []
                    y_groups[y_key].append({
                        "text": line_text.strip(),
                        "x": bbox[0],
                        "bbox": bbox
                    })
        
        # 複数のセルがある行を表の行として抽出
        table_rows = []
        for y_key in sorted(y_groups.keys()):
            cells = y_groups[y_key]
            if len(cells) >= 2:  # 2つ以上のセルがある行
                # X座標でソート
                cells_sorted = sorted(cells, key=lambda c: c["x"])
                
                # 表の追加条件: セルの平均文字数が短め（長文は段組みの可能性）
                avg_text_len = sum(len(c["text"]) for c in cells_sorted) / len(cells_sorted)
                if avg_text_len < 50:  # 平均50文字未満
                    table_rows.append(cells_sorted)
        
        # 表として認識する条件を強化
        # 1. 連続する行が3行以上
        # 2. 列数が行ごとに大きくブレない
        if len(table_rows) >= 3:
            # 列数の一貫性をチェック
            col_counts = [len(row) for row in table_rows]
            most_common_cols = max(set(col_counts), key=col_counts.count)
            consistent_rows = sum(1 for c in col_counts if c == most_common_cols)
            
            # 80%以上の行が同じ列数なら表として認識
            if consistent_rows / len(table_rows) >= 0.8:
                return table_rows
        
        return []
    
    def _merge_table_blocks(
        self, blocks: List[Dict], table_rows: List[List[Dict]]
    ) -> List[Dict]:
        """表構造をブロックリストにマージ
        
        Args:
            blocks: 既存のブロックリスト
            table_rows: 検出された表の行
            
        Returns:
            更新されたブロックリスト
        """
        if not table_rows:
            return blocks
        
        # 表のMarkdownを生成
        table_md_lines = []
        
        # ヘッダー行
        header_cells = [cell["text"] for cell in table_rows[0]]
        table_md_lines.append("| " + " | ".join(header_cells) + " |")
        table_md_lines.append("|" + "|".join(["---"] * len(header_cells)) + "|")
        
        # データ行
        for row in table_rows[1:]:
            row_cells = [cell["text"] for cell in row]
            # セル数を揃える
            while len(row_cells) < len(header_cells):
                row_cells.append("")
            table_md_lines.append("| " + " | ".join(row_cells[:len(header_cells)]) + " |")
        
        table_text = "\n".join(table_md_lines)
        
        # 表ブロックを追加（既存ブロックの最後に）
        # 表に含まれるテキストを持つブロックを除外
        table_texts = set()
        for row in table_rows:
            for cell in row:
                table_texts.add(cell["text"])
        
        filtered_blocks = []
        for block in blocks:
            block_text = block["text"].strip()
            # 表のセルテキストと完全一致するブロックは除外
            if block_text not in table_texts:
                filtered_blocks.append(block)
        
        # 表ブロックを追加
        filtered_blocks.append({
            "type": "table",
            "text": table_text,
            "font_size": 12,
            "bbox": (0, 0, 0, 0)
        })
        
        return filtered_blocks
    
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
    
    def _output_structured_markdown_with_images(
        self, blocks: List[Dict[str, Any]], images: List[Dict[str, Any]]
    ):
        """構造化されたテキストブロックと画像をMarkdownとして出力
        
        Args:
            blocks: 構造化されたテキストブロックのリスト
            images: 抽出された画像情報のリスト
        """
        # 画像がない場合は従来の処理
        if not images:
            self._output_structured_markdown(blocks)
            return
        
        # ブロックと画像をカラム・Y座標でマージしてソート
        all_items = []
        
        for block in blocks:
            bbox = block.get("bbox", (0, 0, 0, 0))
            # カラム情報を取得（デフォルトは"full"）
            column = block.get("column", "full")
            # カラムを数値に変換（left=0, full=1, right=2）
            column_order = {"left": 0, "full": 1, "right": 2}.get(column, 1)
            all_items.append({
                "type": "block",
                "data": block,
                "y_position": bbox[1],
                "column": column,
                "column_order": column_order
            })
        
        for img in images:
            bbox = img.get("bbox", (0, 0, 0, 0))
            # 画像のカラムをX座標から判定（ページ中央より左なら左カラム）
            img_center_x = (bbox[0] + bbox[2]) / 2 if bbox else 0
            # ページ幅の半分を基準にカラムを判定（297.64は一般的なA4の半分）
            column = "left" if img_center_x < 297.64 else "right"
            column_order = {"left": 0, "full": 1, "right": 2}.get(column, 1)
            all_items.append({
                "type": "image",
                "data": img,
                "y_position": img["y_position"],
                "column": column,
                "column_order": column_order
            })
        
        # カラム順（左→フル→右）、次にY座標でソート
        all_items.sort(key=lambda x: (x["column_order"], x["y_position"]))
        
        prev_type = None
        list_active = False
        
        for item in all_items:
            if item["type"] == "image":
                # 画像を出力
                if list_active:
                    self.markdown_lines.append("")
                    list_active = False
                
                img_data = item["data"]
                self.markdown_lines.append("")
                self.markdown_lines.append(f"![図](images/{img_data['filename']})")
                self.markdown_lines.append("")
                
                # 図内テキストを<details>タグで出力
                figure_texts = img_data.get("texts", [])
                if figure_texts:
                    details_text = self._format_figure_texts_as_details(figure_texts)
                    self.markdown_lines.append(details_text)
                    self.markdown_lines.append("")
                
                prev_type = "image"
                
            else:
                # テキストブロックを出力
                block = item["data"]
                block_type = block["type"]
                text = block["text"].strip()
                
                if not text:
                    continue
                
                # リストの終了処理
                if list_active and block_type != "list_item":
                    self.markdown_lines.append("")
                    list_active = False
                
                if block_type == "heading1":
                    if prev_type:
                        self.markdown_lines.append("")
                    self.markdown_lines.append(f"# {text}")
                    self.markdown_lines.append("")
                    
                elif block_type == "heading2":
                    if prev_type:
                        self.markdown_lines.append("")
                    self.markdown_lines.append(f"## {text}")
                    self.markdown_lines.append("")
                    
                elif block_type == "heading3":
                    if prev_type:
                        self.markdown_lines.append("")
                    self.markdown_lines.append(f"### {text}")
                    self.markdown_lines.append("")
                    
                elif block_type == "list_item":
                    if not list_active:
                        if prev_type:
                            self.markdown_lines.append("")
                        list_active = True
                    
                    # 箇条書きマーカーを検出して適切な形式で出力
                    import re
                    cleaned_text = text
                    output_marker = "-"
                    
                    # 番号付きリストを検出（元の番号を保持）
                    num_match = re.match(r'^([\d０-９]+)[\.．\)）]\s+', text)
                    if num_match:
                        # 全角数字を半角に変換
                        num_str = num_match.group(1)
                        num_str = num_str.translate(str.maketrans('０１２３４５６７８９', '0123456789'))
                        output_marker = f"{num_str}."
                        cleaned_text = text[num_match.end():]
                    else:
                        # 空白必須のマーカー（-, *）
                        cleaned_text = re.sub(r'^[\-\*]\s+', '', cleaned_text)
                        # 空白不要のマーカー（•, ・, ○, ● など）
                        cleaned_text = re.sub(r'^[•・○●◆◇▪▫－―＊]\s*', '', cleaned_text)
                    
                    self.markdown_lines.append(f"{output_marker} {cleaned_text}")
                    
                elif block_type == "table":
                    if prev_type:
                        self.markdown_lines.append("")
                    self.markdown_lines.append(text)
                    self.markdown_lines.append("")
                    
                else:  # paragraph
                    if prev_type and prev_type != "paragraph":
                        self.markdown_lines.append("")
                    self.markdown_lines.append(text)
                    self.markdown_lines.append("")
                
                prev_type = block_type
        
        # 最後のリストの終了処理
        if list_active:
            self.markdown_lines.append("")

    def _output_structured_markdown(self, blocks: List[Dict[str, Any]]):
        """構造化されたテキストブロックをMarkdownとして出力
        
        Args:
            blocks: 構造化されたテキストブロックのリスト
        """
        prev_type = None
        list_active = False
        
        for block in blocks:
            block_type = block["type"]
            text = block["text"].strip()
            
            if not text:
                continue
            
            # リストの終了処理
            if list_active and block_type != "list_item":
                self.markdown_lines.append("")
                list_active = False
            
            if block_type == "heading1":
                if prev_type:
                    self.markdown_lines.append("")
                self.markdown_lines.append(f"# {text}")
                self.markdown_lines.append("")
                
            elif block_type == "heading2":
                if prev_type:
                    self.markdown_lines.append("")
                self.markdown_lines.append(f"## {text}")
                self.markdown_lines.append("")
                
            elif block_type == "heading3":
                if prev_type:
                    self.markdown_lines.append("")
                self.markdown_lines.append(f"### {text}")
                self.markdown_lines.append("")
                
            elif block_type == "list_item":
                if not list_active:
                    if prev_type:
                        self.markdown_lines.append("")
                    list_active = True
                
                # 箇条書きマーカーを検出して適切な形式で出力
                import re
                cleaned_text = text
                output_marker = "-"
                
                # 番号付きリストを検出（元の番号を保持）
                num_match = re.match(r'^([\d０-９]+)[\.．\)）]\s+', text)
                if num_match:
                    # 全角数字を半角に変換
                    num_str = num_match.group(1)
                    num_str = num_str.translate(str.maketrans('０１２３４５６７８９', '0123456789'))
                    output_marker = f"{num_str}."
                    cleaned_text = text[num_match.end():]
                else:
                    # 空白必須のマーカー（-, *）
                    cleaned_text = re.sub(r'^[\-\*]\s+', '', cleaned_text)
                    # 空白不要のマーカー（•, ・, ○, ● など）
                    cleaned_text = re.sub(r'^[•・○●◆◇▪▫－―＊]\s*', '', cleaned_text)
                
                self.markdown_lines.append(f"{output_marker} {cleaned_text}")
                
            elif block_type == "table":
                if prev_type:
                    self.markdown_lines.append("")
                self.markdown_lines.append(text)
                self.markdown_lines.append("")
                
            else:  # paragraph
                if prev_type and prev_type != "paragraph":
                    self.markdown_lines.append("")
                self.markdown_lines.append(text)
                self.markdown_lines.append("")
            
            prev_type = block_type
        
        # 最後のリストの終了処理
        if list_active:
            self.markdown_lines.append("")

    def _render_page_as_image(self, page, page_num: int) -> Optional[str]:
        """PDFページを画像としてレンダリング
        
        Args:
            page: PyMuPDFのページオブジェクト
            page_num: ページ番号（0始まり）
            
        Returns:
            保存された画像ファイルのパス
        """
        try:
            # 高解像度でレンダリング
            matrix = fitz.Matrix(DEFAULT_DPI / 72, DEFAULT_DPI / 72)
            pix = page.get_pixmap(matrix=matrix)
            
            if self.output_format == 'svg':
                # SVG形式で出力
                image_filename = f"{self.base_name}_page_{page_num + 1:03d}.svg"
                image_path = os.path.join(self.images_dir, image_filename)
                
                # まずPNGとして保存し、SVGに変換
                temp_png = os.path.join(self.images_dir, f"temp_{page_num}.png")
                pix.save(temp_png)
                
                # PNGをSVGに変換（埋め込み形式）
                self._convert_png_to_svg(temp_png, image_path)
                
                # 一時ファイルを削除
                if os.path.exists(temp_png):
                    os.remove(temp_png)
            else:
                # PNG形式で出力
                image_filename = f"{self.base_name}_page_{page_num + 1:03d}.png"
                image_path = os.path.join(self.images_dir, image_filename)
                pix.save(image_path)
            
            self.image_counter += 1
            debug_print(f"[DEBUG] 画像を保存: {image_path}")
            return image_path
            
        except Exception as e:
            print(f"[WARNING] ページ {page_num + 1} の画像変換に失敗: {e}")
            return None
    
    def _convert_png_to_svg(self, png_path: str, svg_path: str):
        """PNGをSVG形式に変換（画像埋め込み形式）
        
        Args:
            png_path: 入力PNGファイルのパス
            svg_path: 出力SVGファイルのパス
        """
        import base64
        
        try:
            with Image.open(png_path) as img:
                width, height = img.size
            
            with open(png_path, 'rb') as f:
                png_data = base64.b64encode(f.read()).decode('utf-8')
            
            svg_content = f'''<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" 
     width="{width}" height="{height}" viewBox="0 0 {width} {height}">
  <image width="{width}" height="{height}" 
         xlink:href="data:image/png;base64,{png_data}"/>
</svg>'''
            
            with open(svg_path, 'w', encoding='utf-8') as f:
                f.write(svg_content)
                
        except Exception as e:
            print(f"[WARNING] SVG変換に失敗: {e}")
            # フォールバック: PNGをそのままコピー
            import shutil
            png_fallback = svg_path.replace('.svg', '.png')
            shutil.copy(png_path, png_fallback)
    
    def _extract_text_from_page(self, page, page_num: int) -> str:
        """PDFページからテキストを抽出
        
        埋め込みテキストを優先的に抽出し、
        テキストが取得できない場合はOCRを使用する。
        
        Args:
            page: PyMuPDFのページオブジェクト
            page_num: ページ番号（0始まり）
            
        Returns:
            抽出されたテキスト
        """
        # まず埋め込みテキストを試す
        text = page.get_text("text").strip()
        
        if text:
            debug_print(f"[DEBUG] ページ {page_num + 1}: 埋め込みテキストを抽出")
            return text
        
        # テキストがない場合はOCRを試す
        debug_print(f"[DEBUG] ページ {page_num + 1}: OCRでテキストを抽出")
        ocr_text = self._ocr_page(page)
        return ocr_text
    
    def _ocr_page(self, page) -> str:
        """manga-ocrを使用してページからテキストを抽出
        
        Args:
            page: PyMuPDFのページオブジェクト
            
        Returns:
            OCRで抽出されたテキスト
        """
        ocr = self._get_ocr()
        if ocr is None:
            return "(OCR利用不可)"
        
        try:
            # ページを画像に変換
            matrix = fitz.Matrix(2.0, 2.0)
            pix = page.get_pixmap(matrix=matrix)
            
            # PILイメージに変換
            import io
            img_data = pix.tobytes("png")
            img = Image.open(io.BytesIO(img_data))
            
            # manga-ocrでテキスト抽出
            text = ocr(img)
            return text.strip() if text else ""
            
        except Exception as e:
            print(f"[WARNING] OCR処理中にエラーが発生: {e}")
            return "(OCRエラー)"


def main():
    """メイン関数"""
    import argparse
    
    parser = argparse.ArgumentParser(description='PDFファイルをMarkdownに変換')
    parser.add_argument('pdf_file', help='変換するPDFファイル')
    parser.add_argument('-o', '--output-dir', type=str,
                       help='出力ディレクトリを指定（デフォルト: ./output）')
    parser.add_argument('--format', choices=['png', 'svg'], default='svg',
                       help='出力画像形式を指定（デフォルト: svg）')
    parser.add_argument('-v', '--verbose', action='store_true',
                       help='デバッグ情報を出力')
    
    args = parser.parse_args()
    
    set_verbose(args.verbose)
    
    if not os.path.exists(args.pdf_file):
        print(f"エラー: ファイル '{args.pdf_file}' が見つかりません。")
        sys.exit(1)
    
    if not args.pdf_file.lower().endswith('.pdf'):
        print("エラー: .pdf形式のファイルを指定してください。")
        sys.exit(1)
    
    try:
        converter = PDFToMarkdownConverter(
            args.pdf_file,
            output_dir=args.output_dir,
            output_format=args.format
        )
        output_file = converter.convert()
        print("\n変換完了!")
        print(f"出力ファイル: {output_file}")
        print(f"画像フォルダ: {converter.images_dir}")
        
    except Exception as e:
        print(f"変換エラー: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
