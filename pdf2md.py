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
# Mixinクラスのインポート
from pdf2md_figures import _FiguresMixin
from pdf2md_tables import _TablesMixin
from pdf2md_text import _TextMixin

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


class PDFToMarkdownConverter(_FiguresMixin, _TablesMixin, _TextMixin):
    """PDFファイルをMarkdown形式に変換するコンバータクラス
    
    o2mdファミリーとして統一されたインターフェースを提供します。
    
    継承:
        _FiguresMixin: 図抽出機能（pdf2md_figures.py）
        _TablesMixin: 表検出・処理機能（pdf2md_tables.py）
        _TextMixin: テキスト抽出・処理機能（pdf2md_text.py）
    """
    
    def __init__(
        self,
        pdf_file_path: str,
        output_dir: Optional[str] = None,
        output_format: str = 'png',
        ocr_engine: str = 'tesseract',
        tessdata_dir: Optional[str] = None
    ):
        """コンバータインスタンスの初期化
        
        Args:
            pdf_file_path: 変換するPDFファイルのパス
            output_dir: 出力ディレクトリ（省略時は./output）
            output_format: 出力画像形式 ('png' または 'svg')
            ocr_engine: OCRエンジン ('manga-ocr' または 'tesseract')
            tessdata_dir: tessdataディレクトリのパス（tessdata_best使用時に指定）
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
        
        # OCRエンジン設定
        self.ocr_engine = ocr_engine.lower() if ocr_engine else 'tesseract'
        if self.ocr_engine not in ('manga-ocr', 'tesseract'):
            print(f"[WARNING] 不明なOCRエンジン '{ocr_engine}'。'tesseract'を使用します。")
            self.ocr_engine = 'tesseract'
        
        # tessdataディレクトリ（tessdata_best使用時に指定）
        self.tessdata_dir = tessdata_dir
        
        # 脚注番号セット（参考文献ブロックから抽出した番号のみ変換対象）
        self._defined_footnote_nums: Set[str] = set()
        
        # doc-wideのヘッダー/フッター領域（フォールバック用）
        self._doc_header_y_max: Optional[float] = None
        self._doc_footer_y_min: Optional[float] = None
        
        print(f"[INFO] 出力画像形式: {self.output_format.upper()}")
    
    def _get_ocr(self):
        """manga-ocrインスタンスを取得（遅延初期化）"""
        if self._ocr is None:
            try:
                # tokenizersのスレッドプール生成を抑止（終了時ハング対策）
                os.environ.setdefault("TOKENIZERS_PARALLELISM", "false")
                
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
            
            # doc-wideのヘッダー/フッター領域を計算（フォールバック用）
            self._compute_doc_wide_header_footer(doc, header_footer_patterns)
            
            for page_num in range(total_pages):
                print(f"[INFO] ページ {page_num + 1}/{total_pages} を処理中...")
                page = doc[page_num]
                self._convert_page(page, page_num, header_footer_patterns)
        finally:
            doc.close()
        
        # Markdownファイルを書き出し
        markdown_content = "\n".join(self.markdown_lines)
        
        # ページ跨ぎの文章を結合する後処理
        markdown_content = self._merge_across_page_breaks(markdown_content)
        
        # 最終パス: 全文に対して脚注参照変換を適用
        # （出力経路によっては変換が漏れる可能性があるため）
        if self._defined_footnote_nums:
            markdown_content = self._convert_inline_footnote_refs(markdown_content)
        
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
    
    def _compute_doc_wide_header_footer(self, doc, header_footer_patterns: Set[str]) -> None:
        """doc-wideのヘッダー/フッター領域を計算（フォールバック用）
        
        全ページからヘッダー/フッター領域のY座標を収集し、
        中央値を計算してフォールバック値として保持する。
        
        Args:
            doc: PyMuPDFのドキュメントオブジェクト
            header_footer_patterns: ヘッダ・フッタパターンのセット
        """
        if not header_footer_patterns:
            return
        
        header_y_values = []
        footer_y_values = []
        
        for page_num in range(len(doc)):
            page = doc[page_num]
            page_height = page.rect.height
            
            try:
                text_dict = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)
            except Exception:
                continue
            
            page_header_y = []
            page_footer_y = []
            
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
                    
                    # ヘッダー/フッター判定
                    if self._is_header_footer(
                        line_text, header_footer_patterns,
                        y_pos=line_bbox[1], page_height=page_height
                    ):
                        if y_center < page_height / 2:
                            page_header_y.append(line_bbox[3])
                        else:
                            page_footer_y.append(line_bbox[1])
            
            # ページごとのヘッダー/フッター領域を収集
            if page_header_y:
                header_y_values.append(max(page_header_y) + 15.0)
            if page_footer_y:
                footer_y_values.append(min(page_footer_y) - 15.0)
        
        # 中央値を計算してフォールバック値として保持
        if header_y_values:
            header_y_values.sort()
            mid = len(header_y_values) // 2
            self._doc_header_y_max = header_y_values[mid]
            debug_print(f"[DEBUG] doc-wideヘッダー領域: y_max={self._doc_header_y_max:.1f} ({len(header_y_values)}ページから計算)")
        
        if footer_y_values:
            footer_y_values.sort()
            mid = len(footer_y_values) // 2
            self._doc_footer_y_min = footer_y_values[mid]
            debug_print(f"[DEBUG] doc-wideフッター領域: y_min={self._doc_footer_y_min:.1f} ({len(footer_y_values)}ページから計算)")
    
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
        # ただし、y座標がヘッダー/フッター帯にある場合のみ除外
        # （章扉ページの本文中の「第N章」「序論」などを誤除外しないため）
        for pattern in patterns:
            pattern_normalized = re.sub(r'\s+', ' ', pattern).strip()
            if normalized == pattern_normalized:
                # y座標情報がある場合は、ヘッダー/フッター帯にいるかチェック
                if y_pos is not None and page_height is not None:
                    # ヘッダー帯: ページ上端12%以内
                    # フッター帯: ページ下端10%以内（90%以降）
                    in_header_zone = y_pos < page_height * 0.12
                    in_footer_zone = y_pos > page_height * 0.90
                    if in_header_zone or in_footer_zone:
                        return True
                    # ヘッダー/フッター帯外にある場合は除外しない
                    debug_print(f"[DEBUG] パターン一致だがヘッダー/フッター帯外のため除外しない: '{text_stripped[:30]}' y={y_pos:.1f}")
                    continue
                else:
                    # y座標情報がない場合は従来通り除外
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
    
    def _is_scan_page(self, page, text_blocks: list, vector_figures: list) -> bool:
        """スキャンページ（画像ベースPDF）かどうかを判定
        
        以下の条件を満たす場合にスキャンページと判定:
        - テキストブロックが空
        - 図が1つだけ
        - その図がページの大部分（80%以上）を覆っている
        
        Args:
            page: PyMuPDFのページオブジェクト
            text_blocks: 抽出されたテキストブロック
            vector_figures: 抽出された図のリスト
            
        Returns:
            スキャンページの場合True
        """
        # テキストブロックがある場合はスキャンページではない
        if text_blocks:
            return False
        
        # 図が1つだけの場合のみチェック
        if len(vector_figures) != 1:
            return False
        
        # ページサイズを取得
        page_rect = page.rect
        page_width = page_rect.width
        page_height = page_rect.height
        page_area = page_width * page_height
        
        # 図のbboxを取得
        fig_bbox = vector_figures[0].get("bbox", (0, 0, 0, 0))
        fig_width = fig_bbox[2] - fig_bbox[0]
        fig_height = fig_bbox[3] - fig_bbox[1]
        fig_area = fig_width * fig_height
        
        # 図がページの80%以上を覆っている場合はスキャンページ
        coverage_ratio = fig_area / page_area if page_area > 0 else 0
        
        if coverage_ratio >= 0.8:
            debug_print(f"[DEBUG] スキャンページ検出: 図の面積比={coverage_ratio:.2%}")
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
        # ヘッダー/フッター領域内の装飾要素を除外するためにpatternsを渡す
        vector_figures = self._extract_vector_figures(page, page_num, header_footer_patterns)
        figure_bboxes = [fig["bbox"] for fig in vector_figures]
        
        # テキストベースのPDFかどうかを判定（図領域を除外）
        text_blocks = self._extract_structured_text_v2(
            page, header_footer_patterns, exclude_bboxes=figure_bboxes
        )
        
        # スキャンページ（画像ベースPDF）かどうかを判定
        is_scan = self._is_scan_page(page, text_blocks, vector_figures)
        
        if is_scan:
            # スキャンページ: 図の画像を出力し、OCRでテキスト抽出
            debug_print(f"[DEBUG] ページ {page_num + 1}: スキャンページとして処理（OCR実行）")
            
            # 図の画像を出力
            if vector_figures:
                fig = vector_figures[0]
                image_filename = fig.get("filename", "")
                if image_filename:
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
        elif text_blocks or vector_figures:
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
        
        # ページ境界マーカーを挿入（後処理でページ跨ぎの結合に使用）
        # 先頭ブロックのy座標を埋め込む（ヘッダ帯判定用）
        # text_blocks（ヘッダ/フッター除外後）の先頭ブロックを使用
        first_block_y = None
        if text_blocks:
            for block in text_blocks:
                block_y = block.get("y", block.get("bbox", [0, 0, 0, 0])[1] if block.get("bbox") else 0)
                if block_y is not None and block_y > 0:
                    first_block_y = block_y
                    break
        self.markdown_lines.append(f"<!--PAGE_BREAK first_y={first_block_y}-->")
        self.markdown_lines.append("")
    
    
    
    
    
    
    
    
    
    
    def _merge_across_page_breaks(self, content: str) -> str:
        """ページ跨ぎの文章を結合する後処理
        
        ページ境界マーカー（<!--PAGE_BREAK first_y=...-->）の前後の段落を分析し、
        文が途中で切れている場合は結合する。
        ただし、次ページの先頭ブロックがヘッダ帯（y ≤ doc_header_y_max）にある場合は
        結合しない（ヘッダ/メタデータとして扱う）。
        
        Args:
            content: ページ境界マーカーを含むMarkdownコンテンツ
            
        Returns:
            ページ跨ぎの文章を結合したMarkdownコンテンツ
        """
        import re
        
        # マーカーパターン（y座標付き）
        page_marker_pattern = r'<!--PAGE_BREAK first_y=([^>]+)-->'
        
        # マーカーが存在しない場合はそのまま返す
        if '<!--PAGE_BREAK' not in content:
            return content
        
        # マーカーで分割してページごとの内容を取得（y座標も抽出）
        parts = re.split(page_marker_pattern, content)
        # parts = [content0, y0, content1, y1, content2, ...]
        # 奇数インデックスがy座標、偶数インデックスがコンテンツ
        
        if len(parts) <= 1:
            return content
        
        # ページコンテンツとy座標を分離
        page_contents = []
        first_block_y_values = []
        for i, part in enumerate(parts):
            if i % 2 == 0:
                page_contents.append(part)
            else:
                # y座標を解析（"None"の場合はNone）
                try:
                    y_val = float(part) if part != 'None' else None
                except ValueError:
                    y_val = None
                first_block_y_values.append(y_val)
        
        if len(page_contents) <= 1:
            return content
        
        # 結合判定用のパターン（結合しない条件）
        # 見出し、リスト、表、図、脚注定義などで始まる場合は結合しない
        no_merge_start_patterns = [
            r'^#{1,6}\s',  # 見出し
            r'^[\d０-９]+[\.．\)）]\s',  # 番号付きリスト（1. 2) など）
            r'^[(（][0-9０-９]+[)）]\s*',  # 括弧付き番号（(1) （１）など）
            r'^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]',  # 丸数字
            r'^[-\*]\s',  # 箇条書き
            r'^\|',  # 表
            r'^!\[',  # 画像
            r'^\[\^',  # 脚注定義
            r'^>\s',  # 引用
            r'^<details',  # 詳細ブロック
            r'^第[0-9０-９一二三四五六七八九十百]+\s*条',  # 「第N条」形式
            r'^(図|表)\s*[\d０-９]+[\.\:．：]',  # 図・表キャプション（図3.1: 表4.2:など）
        ]
        no_merge_start_re = re.compile('|'.join(no_merge_start_patterns))
        
        # 文末記号（これで終わる場合は結合しない）
        sentence_end_chars = set('。！？」』】）)!?')
        
        merged_pages = [page_contents[0]]
        
        for i in range(1, len(page_contents)):
            prev_content = merged_pages[-1]
            curr_content = page_contents[i]
            
            # 前ページの末尾段落を取得
            prev_lines = prev_content.rstrip().split('\n')
            prev_last_line = ''
            for line in reversed(prev_lines):
                stripped = line.strip()
                if stripped:
                    prev_last_line = stripped
                    break
            
            # 現ページの先頭段落を取得
            curr_lines = curr_content.lstrip().split('\n')
            curr_first_line = ''
            curr_first_idx = 0
            for idx, line in enumerate(curr_lines):
                stripped = line.strip()
                if stripped:
                    curr_first_line = stripped
                    curr_first_idx = idx
                    break
            
            # 結合判定
            should_merge = False
            
            if prev_last_line and curr_first_line:
                # 前ページが文末記号で終わっていない
                ends_with_sentence = prev_last_line[-1] in sentence_end_chars
                
                # 現ページが新しい構造要素で始まらない
                starts_with_structure = bool(no_merge_start_re.match(curr_first_line))
                
                # 次ページの先頭ブロックがヘッダ帯にあるかチェック（汎用的な判定）
                # y座標がヘッダ帯（doc_header_y_max以下）にある場合は結合しない
                # 注意: first_block_y_values[i]は次のページ（curr_content）の先頭ブロックのy座標
                curr_first_y = first_block_y_values[i] if i < len(first_block_y_values) else None
                in_header_area = False
                if curr_first_y is not None and self._doc_header_y_max is not None:
                    if curr_first_y <= self._doc_header_y_max:
                        in_header_area = True
                        debug_print(f"[DEBUG] ヘッダ帯検出: y={curr_first_y:.1f} <= header_y_max={self._doc_header_y_max:.1f}")
                
                # 結合条件: 文末でなく、新構造要素でもなく、ヘッダ帯でもない
                if not ends_with_sentence and not starts_with_structure and not in_header_area:
                    should_merge = True
                    debug_print(f"[DEBUG] ページ跨ぎ結合: '{prev_last_line[-20:]}' + '{curr_first_line[:20]}'")
            
            if should_merge:
                # 前ページの末尾と現ページの先頭を結合
                # 前ページの末尾の空行を削除
                prev_stripped = prev_content.rstrip()
                
                # 現ページの先頭の空行を削除し、最初の行を前ページに結合
                curr_remaining_lines = curr_lines[curr_first_idx + 1:]
                curr_remaining = '\n'.join(curr_remaining_lines)
                
                # 結合（日本語の場合はスペースなし、英数字の場合はスペースあり）
                last_char = prev_last_line[-1] if prev_last_line else ''
                first_char = curr_first_line[0] if curr_first_line else ''
                
                # 両方がASCII英数字の場合のみスペースを挿入
                need_space = (last_char.isascii() and last_char.isalnum() and 
                              first_char.isascii() and first_char.isalnum())
                
                separator = ' ' if need_space else ''
                merged_line = prev_stripped + separator + curr_first_line
                
                # 残りの内容を追加
                if curr_remaining.strip():
                    merged_pages[-1] = merged_line + '\n\n' + curr_remaining
                else:
                    merged_pages[-1] = merged_line + '\n'
            else:
                # 結合しない場合は通常通り追加
                merged_pages.append(curr_content)
        
        # マーカーを削除して結合
        result = ''.join(merged_pages)
        
        # 連続する空行を2つまでに正規化
        result = re.sub(r'\n{3,}', '\n\n', result)
        
        return result
    
    def _fix_broken_urls(self, text: str) -> str:
        """PDF抽出時に混入したURL内のスペースを除去
        
        http://example.com/path/file name.pdf
        ↓
        http://example.com/path/filename.pdf
        
        Args:
            text: 処理対象のテキスト
            
        Returns:
            URL内のスペースを除去したテキスト
        """
        import re
        
        def fix_url(match):
            """URL内のスペースを除去"""
            url = match.group(0)
            return re.sub(r'\s+', '', url)
        
        # スペースを含むURLにマッチ（http(s)://で始まり、区切り文字で終わる）
        # 少なくとも1回はスペースを含むURLのみを対象にする
        pattern = r'https?://[^\s,、)\]>\"]+(?:\s+[^\s,、)\]>\"]+)+'
        return re.sub(pattern, fix_url, text)
    
    def _convert_inline_footnote_refs(self, text: str) -> str:
        """本文中のインライン参照[N]を[^N]に変換
        
        脚注定義が存在する番号のみを変換する。
        
        ...として[4]。読む確率を...
        ↓
        ...として[^4]。読む確率を...
        
        Args:
            text: 処理対象のテキスト
            
        Returns:
            インライン参照を変換したテキスト
        """
        import re
        
        # 脚注番号セットが空の場合は変換しない
        if not self._defined_footnote_nums:
            return text
        
        def replace_func(match):
            prefix = match.group(1)
            num = match.group(3)
            # 脚注定義が存在する番号のみ変換
            if num in self._defined_footnote_nums:
                return f"{prefix}[^{num}]"
            return match.group(0)
        
        # 前の文字を含めてマッチし、図/表の直後でないことを確認
        pattern = r'(^|[^^\[図表])(\[(\d{1,2})\])'
        
        return re.sub(pattern, replace_func, text)
    
    def _format_footnote_definitions(self, text: str) -> str:
        """参考文献または用語の説明ブロックを注釈定義形式に変換
        
        参考文献[1]	 説明...[2]	 説明...
        ↓
        ## 参考文献
        
        [^1]: 説明...
        
        [^2]: 説明...
        
        用語の説明用語1：	説明...
        ↓
        ## 用語の説明
        
        [^用語1]: 説明...
        
        Args:
            text: 参考文献または用語の説明ブロックのテキスト
            
        Returns:
            Markdown注釈定義形式のテキスト
        """
        import re
        
        # URL内のスペースを除去
        text = self._fix_broken_urls(text)
        
        # 用語の説明ブロックの処理
        if text.lstrip().startswith("用語の説明") and re.search(r'用語\s*\d+\s*[:：]', text):
            # 「用語の説明」を除去
            text = re.sub(r'^\s*用語の説明\s*', '', text)
            
            # 用語N: マーカーの位置を全て見つける
            markers = list(re.finditer(r'用語\s*(\d+)\s*[:：]\s*', text))
            if not markers:
                return f"## 用語の説明\n\n{text}"
            
            # 各マーカー間のテキストを切り出して注釈定義を生成
            definitions = []
            for i, match in enumerate(markers):
                num = match.group(1)
                start = match.end()
                if i + 1 < len(markers):
                    end = markers[i + 1].start()
                else:
                    end = len(text)
                
                content = text[start:end].strip()
                if content:
                    definitions.append(f"[^用語{num}]: {content}")
            
            if definitions:
                return "## 用語の説明\n\n" + "\n\n".join(definitions) + "\n"
            else:
                return f"## 用語の説明\n\n{text}"
        
        # 参考文献ブロックの処理
        if not re.match(r'^\s*参考文献\s*\[\d+\]', text):
            return text
        
        # 「参考文献」を除去
        text = re.sub(r'^\s*参考文献\s*', '', text)
        
        # [N] マーカーの位置を全て見つける
        markers = list(re.finditer(r'\[(\d+)\]\s*', text))
        if not markers:
            return f"## 参考文献\n\n{text}"
        
        # 各マーカー間のテキストを切り出して注釈定義を生成
        definitions = []
        for i, match in enumerate(markers):
            num = match.group(1)
            start = match.end()
            # 次のマーカーの開始位置、または文字列の終端まで
            if i + 1 < len(markers):
                end = markers[i + 1].start()
            else:
                end = len(text)
            
            content = text[start:end].strip()
            # 末尾の不要な文字を除去
            content = content.rstrip('.')
            if content:
                definitions.append(f"[^{num}]: {content}")
        
        if definitions:
            return "## 参考文献\n\n" + "\n\n".join(definitions) + "\n"
        else:
            return f"## 参考文献\n\n{text}"
    
    def _is_footnote_definition_block(self, text: str) -> bool:
        """テキストが参考文献または用語の説明ブロックかどうかを判定
        
        Args:
            text: 判定するテキスト
            
        Returns:
            参考文献または用語の説明ブロックの場合True
        """
        import re
        # 参考文献ブロック
        if re.match(r'^\s*参考文献\s*\[\d+\]', text):
            return True
        # 用語の説明ブロック
        if text.lstrip().startswith("用語の説明") and re.search(r'用語\s*\d+\s*[:：]', text):
            return True
        return False
    
    def _extract_footnote_nums_from_blocks(self, blocks: List[Dict[str, Any]]) -> None:
        """blocksから脚注番号を抽出してインスタンス変数に追加
        
        参考文献ブロックから[N]形式の番号を抽出し、
        _defined_footnote_numsに追加する。
        （複数ページにまたがる場合も累積される）
        
        Args:
            blocks: 構造化されたテキストブロックのリスト
        """
        import re
        
        found_new = False
        for block in blocks:
            text = block.get("text", "").strip()
            if not text:
                continue
            
            # 参考文献ブロックから番号を抽出
            if re.match(r'^\s*参考文献\s*\[\d+\]', text):
                nums = re.findall(r'\[(\d+)\]', text)
                if nums:
                    self._defined_footnote_nums.update(nums)
                    found_new = True
        
        if found_new:
            debug_print(f"[DEBUG] 脚注番号セット: {sorted(self._defined_footnote_nums, key=int)}")
    
    
    
    
    
    def _is_numbered_heading(self, text: str) -> Tuple[bool, int, str]:
        """番号付き見出しかどうかを判定
        
        「1　はじめに」「2.1　概要」「4.1.1　詳細」「第1条（借入要項）」などのパターンを検出する。
        文末表現（述語終止形）を含む場合は見出しではないと判定する。
        図表キャプション（「図N」「表N」）は見出しではないと判定する。
        
        Args:
            text: 判定するテキスト
            
        Returns:
            (見出しかどうか, 見出しレベル, 見出しテキスト)
        """
        import re
        text = text.strip()
        
        # 図表キャプション（「図N」「表N」）は見出しではない
        if re.match(r'^(図|表)\s*[0-9０-９]+', text):
            return (False, 0, "")
        
        # 「第N条」形式の見出しパターン（先に判定、本文が続いていても対応）
        article_match = re.match(
            r'^(第[0-9０-９一二三四五六七八九十百]+\s*条[　 ]*[（(][^）)]+[）)])',
            text
        )
        if article_match:
            heading_text = article_match.group(1).strip()
            return (True, 3, heading_text)
        
        # 長すぎる行は見出しではない（「第N条」以外のパターン用）
        if len(text) > 50:
            return (False, 0, "")
        
        # 文末表現（述語終止形）を含む場合は見出しではない
        # 句点（。．）で終わる場合
        if text.endswith("。") or text.endswith("．"):
            return (False, 0, "")
        
        # 半角ピリオドで終わる場合の判定
        if text.endswith("."):
            # 「N.」形式の番号のみは除外しない
            if not re.match(r'^[0-9０-９]+\.$', text):
                return (False, 0, "")
        
        # 文として完結する表現を含む場合は見出しではない
        # 述語終止形: である、ある、する、した、いる、いた、ない、れる、られる、ます、です等
        sentence_endings = (
            r'(である|ある|する|した|いる|いた|った|ない|れる|られる|'
            r'ます|です|ました|でした|ません|ている|ていた|ておく|ておいた|'
            r'なる|なった|できる|できた|おく|おいた|くる|きた|いく|いった)$'
        )
        if re.search(sentence_endings, text):
            return (False, 0, "")
        
        # 階層番号パターン: 「4.1　QRコードの構造」「4.1.1　ファインダパターン」など
        # ドットの数に応じて見出しレベルを決定（4.1 → レベル2、4.1.1 → レベル3）
        hierarchical_match = re.match(
            r'^(\d+(?:\.\d+)+)[　 ]+(.{1,40})$',
            text
        )
        if hierarchical_match:
            number_part = hierarchical_match.group(1)
            title_part = hierarchical_match.group(2).strip()
            
            # タイトル部分が短すぎる場合は除外
            if len(title_part) < 2:
                return (False, 0, "")
            
            # ドットの数でレベルを決定（4.1 → 1ドット → レベル2、4.1.1 → 2ドット → レベル3）
            dot_count = number_part.count('.')
            heading_level = min(dot_count + 1, 3)  # 最大レベル3
            
            return (True, heading_level, text)
        
        # 「N．タイトル」形式の見出しパターン（例: 「１．固定金利型の利率変更」）
        numbered_dot_match = re.match(
            r'^([0-9０-９]+)[．.][　 ]*([^0-9０-９].{1,39})$',
            text
        )
        if numbered_dot_match:
            title_part = numbered_dot_match.group(2).strip()
            first_char = title_part[0] if title_part else ""
            # 単位や助数詞で始まる場合は除外
            excluded_chars = '年月日倍個回分秒時点番号件台人円万億兆つ'
            if first_char not in excluded_chars:
                return (True, 3, text)
        
        # 番号付き見出しパターン: 「1　はじめに」など
        match = re.match(r'^(\d+)\s*[　 ]+(.{1,40})$', text)
        if match:
            title_part = match.group(2).strip()
            
            # タイトル部分が短すぎる場合は除外
            if len(title_part) < 2:
                return (False, 0, "")
            
            # 単位や助数詞で始まる場合は除外
            first_char = title_part[0]
            excluded_chars = '年月日倍個回分秒時点番号件台人円万億兆つ'
            if first_char in excluded_chars:
                return (False, 0, "")
            
            # 「名詞＋助詞＋区切り」パターンは見出しではない（例: 「乙は、」「甲が 」）
            particle_pattern = re.compile(r'^.{1,3}[はがをにでともや][、。\s,.]')
            if particle_pattern.match(title_part):
                return (False, 0, "")
            
            return (True, 2, title_part)
        
        return (False, 0, "")
    
    
    
    
    
    
    
    
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
        
        # 見出しから除外する条件（文らしさフィルタ）
        def is_sentence_like(txt: str) -> bool:
            """文らしいテキストかどうかを判定"""
            # 句点で終わる場合は文として扱う（ただし短い場合は除外）
            if len(txt) > 15 and txt.endswith(('.', '。', '．')):
                return True
            # 図表キャプション形式は見出しではない（「図N」「表N」で始まる）
            if re.match(r'^(図|表)\s*[0-9０-９]+', txt):
                return True
            return False
        
        # 文らしいテキストは見出しにしない
        if is_sentence_like(text_stripped):
            return "paragraph"
        
        # 短いテキストは見出しにしない（継続行の可能性が高い）
        # ただし、以下の場合は除外:
        # - 番号付き見出し形式（「N.N」「第N章」など）
        # - フォントサイズが非常に大きい場合（size_ratio >= 2.0）
        if len(text_stripped) <= 10 and size_ratio < 2.0:
            if not re.match(r'^[\d０-９]+[\.\s]', text_stripped):
                if not re.match(r'^第[\d０-９一二三四五六七八九十]+\s*(章|節|条)', text_stripped):
                    return "paragraph"
        
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
    
    
    
    
    
    
    
    def _output_structured_markdown_with_images(
        self, blocks: List[Dict[str, Any]], images: List[Dict[str, Any]]
    ):
        """構造化されたテキストブロックと画像をMarkdownとして出力
        
        Args:
            blocks: 構造化されたテキストブロックのリスト
            images: 抽出された画像情報のリスト
        """
        import re
        
        # 脚注番号セットを抽出（インライン参照変換に使用）
        self._extract_footnote_nums_from_blocks(blocks)
        
        # 画像がない場合は従来の処理
        if not images:
            self._output_structured_markdown(blocks)
            return
        
        # キャプションパターン（図/表のキャプション）
        caption_pattern = re.compile(r'^(図|表)\s*\d+')
        
        # キャプションと画像を関連付け
        # キャプションブロックを検出し、対応する画像と紐付ける
        caption_to_image = {}  # caption_block_id -> image_index
        image_captions = {}  # image_index -> {"above": caption_block, "below": caption_block}
        used_caption_ids = set()
        
        for img_idx, img in enumerate(images):
            img_bbox = img.get("bbox", (0, 0, 0, 0))
            img_y0 = img_bbox[1]  # 画像の上端
            img_y1 = img_bbox[3]  # 画像の下端
            img_x0 = img_bbox[0]
            img_x1 = img_bbox[2]
            
            best_above_caption = None
            best_above_distance = float('inf')
            best_below_caption = None
            best_below_distance = float('inf')
            
            for block_idx, block in enumerate(blocks):
                if block.get("type") != "heading3":
                    continue
                
                text = block.get("text", "").strip()
                if not caption_pattern.match(text):
                    continue
                
                block_bbox = block.get("bbox", (0, 0, 0, 0))
                block_y0 = block_bbox[1]
                block_y1 = block_bbox[3]
                block_x0 = block_bbox[0]
                block_x1 = block_bbox[2]
                
                # X方向の重なりを確認（少なくとも20pt以上の重なり）
                x_overlap = max(0, min(img_x1, block_x1) - max(img_x0, block_x0))
                if x_overlap < 20:
                    continue
                
                # Y方向の距離を計算
                # キャプションが画像の上にある場合
                if block_y1 <= img_y0 + 30:  # 30ptの許容範囲
                    distance = img_y0 - block_y1
                    if distance < 50 and distance < best_above_distance:  # 50pt以内
                        best_above_distance = distance
                        best_above_caption = (block_idx, block)
                
                # キャプションが画像の下にある場合
                if block_y0 >= img_y1 - 30:  # 30ptの許容範囲
                    distance = block_y0 - img_y1
                    if distance < 50 and distance < best_below_distance:  # 50pt以内
                        best_below_distance = distance
                        best_below_caption = (block_idx, block)
            
            # 画像にキャプションを関連付け
            image_captions[img_idx] = {"above": None, "below": None}
            if best_above_caption:
                block_idx, block = best_above_caption
                image_captions[img_idx]["above"] = block
                used_caption_ids.add(block_idx)
                debug_print(f"[DEBUG] 画像{img_idx}に上キャプションを関連付け: {block.get('text', '')[:30]}")
            if best_below_caption:
                block_idx, block = best_below_caption
                image_captions[img_idx]["below"] = block
                used_caption_ids.add(block_idx)
                debug_print(f"[DEBUG] 画像{img_idx}に下キャプションを関連付け: {block.get('text', '')[:30]}")
        
        # ブロックと画像をカラム・Y座標でマージしてソート
        all_items = []
        
        for block_idx, block in enumerate(blocks):
            # 関連付けられたキャプションはスキップ（画像出力時に出力する）
            if block_idx in used_caption_ids:
                continue
            
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
        
        for img_idx, img in enumerate(images):
            bbox = img.get("bbox", (0, 0, 0, 0))
            # 画像に関連付けられたキャプション情報を追加
            img["_caption_above"] = image_captions.get(img_idx, {}).get("above")
            img["_caption_below"] = image_captions.get(img_idx, {}).get("below")
            
            # 画像のカラムを画像自身のcolumn属性から取得（なければX座標から判定）
            column = img.get("column")
            if not column:
                img_center_x = (bbox[0] + bbox[2]) / 2 if bbox else 0
                column = "left" if img_center_x < 297.64 else "right"
            column_order = {"left": 0, "full": 1, "right": 2}.get(column, 1)
            
            # Y座標は上キャプションがあればそのY座標を使用（順序を正しくするため）
            y_position = img["y_position"]
            caption_above = img.get("_caption_above")
            if caption_above:
                caption_bbox = caption_above.get("bbox", (0, 0, 0, 0))
                y_position = min(y_position, caption_bbox[1])
            
            all_items.append({
                "type": "image",
                "data": img,
                "y_position": y_position,
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
                
                # 上キャプションがあれば先に出力
                caption_above = img_data.get("_caption_above")
                if caption_above:
                    caption_text = caption_above.get("text", "").strip()
                    if caption_text:
                        if prev_type:
                            self.markdown_lines.append("")
                        self.markdown_lines.append(f"### {caption_text}")
                        self.markdown_lines.append("")
                
                self.markdown_lines.append("")
                self.markdown_lines.append(f"![図](images/{img_data['filename']})")
                self.markdown_lines.append("")
                
                # 図内テキストを<details>タグで出力
                figure_texts = img_data.get("texts", [])
                if figure_texts:
                    details_text = self._format_figure_texts_as_details(figure_texts)
                    self.markdown_lines.append(details_text)
                    self.markdown_lines.append("")
                
                # 下キャプションがあれば後に出力
                caption_below = img_data.get("_caption_below")
                if caption_below:
                    caption_text = caption_below.get("text", "").strip()
                    if caption_text:
                        self.markdown_lines.append(f"### {caption_text}")
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
                    # 参考文献ブロックの場合は注釈定義形式に変換
                    if self._is_footnote_definition_block(text):
                        formatted_refs = self._format_footnote_definitions(text)
                        self.markdown_lines.append(formatted_refs)
                    else:
                        # インライン参照[N]を[^N]に変換
                        text = self._convert_inline_footnote_refs(text)
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
                # 参考文献ブロックの場合は注釈定義形式に変換
                if self._is_footnote_definition_block(text):
                    formatted_refs = self._format_footnote_definitions(text)
                    self.markdown_lines.append(formatted_refs)
                else:
                    # インライン参照[N]を[^N]に変換
                    text = self._convert_inline_footnote_refs(text)
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
        """テキスト検出とOCRを使用してページからテキストを抽出
        
        comic-text-detectorでテキスト領域を検出し、
        各領域ごとにmanga-ocrでテキストを抽出します。
        
        Args:
            page: PyMuPDFのページオブジェクト
            
        Returns:
            OCRで抽出されたテキスト
        """
        try:
            # ページを画像に変換（300dpi相当: 300/72 ≈ 4.17）
            scale = 300 / 72
            matrix = fitz.Matrix(scale, scale)
            pix = page.get_pixmap(matrix=matrix)
            
            # numpy配列に変換（BGR形式）
            import numpy as np
            img_array = np.frombuffer(pix.samples, dtype=np.uint8)
            img_array = img_array.reshape(pix.height, pix.width, pix.n)
            
            # RGBの場合はBGRに変換
            if pix.n == 3:
                img_bgr = img_array[:, :, ::-1].copy()
            elif pix.n == 4:
                # RGBAの場合はRGBに変換してからBGRに
                img_bgr = img_array[:, :, :3][:, :, ::-1].copy()
            else:
                img_bgr = img_array
            
            # pdf2md_ocrモジュールを使用してテキスト抽出
            from pdf2md_ocr import process_pdf_page_with_detection, set_verbose as ocr_set_verbose
            ocr_set_verbose(is_verbose())
            
            text = process_pdf_page_with_detection(
                img_bgr, 
                ocr_engine=self.ocr_engine,
                tessdata_dir=self.tessdata_dir
            )
            return text.strip() if text else ""
            
        except ImportError as e:
            # ImportErrorは常に表示（原因特定のため）
            print(f"[WARNING] pdf2md_ocrモジュールが利用できません: {e}")
            print("[WARNING] テキスト領域検出が無効です。フォールバックOCRを使用します。")
            # フォールバック: 従来のmanga-ocr直接呼び出し
            return self._ocr_page_fallback(page)
        except Exception as e:
            print(f"[WARNING] OCR処理中にエラーが発生: {e}")
            return "(OCRエラー)"
    
    def _ocr_page_fallback(self, page) -> str:
        """フォールバック: manga-ocrを直接使用してページからテキストを抽出
        
        Args:
            page: PyMuPDFのページオブジェクト
            
        Returns:
            OCRで抽出されたテキスト
        """
        ocr = self._get_ocr()
        if ocr is None:
            return "(OCR利用不可)"
        
        try:
            # ページを画像に変換（300dpi相当）
            scale = 300 / 72
            matrix = fitz.Matrix(scale, scale)
            pix = page.get_pixmap(matrix=matrix)
            
            # PILイメージに変換
            import io
            img_data = pix.tobytes("png")
            img = Image.open(io.BytesIO(img_data))
            
            # manga-ocrでテキスト抽出
            text = ocr(img)
            return text.strip() if text else ""
            
        except Exception as e:
            print(f"[WARNING] フォールバックOCR処理中にエラーが発生: {e}")
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
    parser.add_argument('--ocr-engine', choices=['manga-ocr', 'tesseract'], 
                       default='tesseract',
                       help='OCRエンジンを指定（デフォルト: tesseract）')
    parser.add_argument('--tessdata-dir', type=str,
                       help='tessdataディレクトリを指定（tessdata_best使用時）')
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
            output_format=args.format,
            ocr_engine=args.ocr_engine,
            tessdata_dir=args.tessdata_dir
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
