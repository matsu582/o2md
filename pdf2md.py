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
from typing import Optional, List, Dict, Any, Tuple

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
            
            for page_num in range(total_pages):
                print(f"[INFO] ページ {page_num + 1}/{total_pages} を処理中...")
                page = doc[page_num]
                self._convert_page(page, page_num)
        finally:
            doc.close()
        
        # Markdownファイルを書き出し
        markdown_content = "\n".join(self.markdown_lines)
        output_file = os.path.join(self.output_dir, f"{self.base_name}.md")
        
        with open(output_file, "w", encoding="utf-8") as f:
            f.write(markdown_content)
        
        print(f"[SUCCESS] 変換完了: {output_file}")
        return output_file
    
    def _convert_page(self, page, page_num: int):
        """PDFページを変換
        
        Args:
            page: PyMuPDFのページオブジェクト
            page_num: ページ番号（0始まり）
        """
        # ページ見出し
        self.markdown_lines.append(f"## ページ {page_num + 1}")
        self.markdown_lines.append("")
        
        # テキストベースのPDFかどうかを判定
        text_blocks = self._extract_structured_text(page)
        
        if text_blocks:
            # テキストベースのPDF: 構造化されたMarkdownを出力
            debug_print(f"[DEBUG] ページ {page_num + 1}: テキストベースPDFとして処理")
            self._output_structured_markdown(text_blocks)
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
        
        # 箇条書きの検出
        list_markers = ['•', '・', '-', '－', '―', '*', '＊', '○', '●', '◆', '◇', '▪', '▫']
        for marker in list_markers:
            if text_stripped.startswith(marker):
                return "list_item"
        
        # 番号付きリストの検出
        import re
        if re.match(r'^[\d０-９]+[\.．\)）]\s*', text_stripped):
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
    
    def _detect_table_structure(self, text_dict: Dict) -> List[List[Dict]]:
        """表構造を検出
        
        同じY座標に複数のテキストブロックがある場合、表として検出する。
        
        Args:
            text_dict: PyMuPDFのテキスト辞書
            
        Returns:
            表の行リスト（各行はセルのリスト）
        """
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
                table_rows.append(cells_sorted)
        
        # 連続する表の行が3行以上ある場合のみ表として認識
        if len(table_rows) >= 2:
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
                
                # 箇条書きマーカーを統一
                import re
                cleaned_text = re.sub(
                    r'^[•・\-－―\*＊○●◆◇▪▫]\s*', '', text
                )
                cleaned_text = re.sub(
                    r'^[\d０-９]+[\.．\)）]\s*', '', cleaned_text
                )
                self.markdown_lines.append(f"- {cleaned_text}")
                
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
