#!/usr/bin/env python3
"""
Office to Markdown Converter (o2md)
Excel、Word、PowerPoint、PDFファイルを自動判定してMarkdownに変換する統合ツール

機能:
- ファイル拡張子に基づいて自動的に適切な変換クラスを選択
- Excel (.xlsx, .xls) → x2md.ExcelToMarkdownConverter
- Word (.docx, .doc) → d2md.WordToMarkdownConverter
- PowerPoint (.pptx, .ppt) → p2md.PowerPointToMarkdownConverter
- PDF (.pdf) → pdf2md.PDFToMarkdownConverter
- 一太郎 (.jtd, .jtt) → jtd2md.JtdToMarkdownConverter
- 古い形式（.xls, .doc, .ppt）は自動的に新形式に変換してから処理
- フォルダ指定時はサブフォルダを含む全対象ファイルを再帰的に一括変換

対応ファイル形式:
- Excel: .xlsx, .xls
- Word: .docx, .doc
- PowerPoint: .pptx, .ppt
- PDF: .pdf
- 一太郎: .jtd, .jtt

使用例:
    # 基本的な使用方法
    python o2md.py input_files/data.xlsx
    python o2md.py input_files/document.docx
    python o2md.py input_files/presentation.pptx
    
    # 出力ディレクトリを指定
    python o2md.py input_files/data.xlsx -o custom_output
    
    # フォルダ内のファイルを一括変換
    python o2md.py input_files/
    python o2md.py input_files/ -r -o output_all  # サブフォルダも再帰的に処理
    
    # Word文書で見出しテキストをリンクに使用
    python o2md.py input_files/document.docx --use-heading-text
    
    # 古い形式のファイルも変換可能
    python o2md.py input_files/old_file.xls
    python o2md.py input_files/old_doc.doc
    python o2md.py input_files/old_presentation.ppt

出力:
- デフォルトの出力ディレクトリ: ./output/
- Markdownファイル: ./output/[元のファイル名].md
- 画像ファイル: ./output/images/
- フォルダ指定時: ./output/[相対パス]/[ファイル名].md

必要な依存関係:
- x2md.py, d2md.py, p2md.py
- openpyxl, python-docx, python-pptx
- Pillow (PIL)
- LibreOffice (古い形式の変換、図形レンダリングに必要、オプショナル)
"""

import logging
import os
import sys
import gc
import argparse
from pathlib import Path

from o2md.i18n import _, setup_i18n

# 各変換クラスをインポート
try:
    from o2md.x2md import ExcelToMarkdownConverter, convert_xls_to_xlsx
    from o2md import x2md
except ImportError as e:
    raise ImportError(
        "x2md.pyのインポートに失敗しました。必要な依存関係をインストールしてください: uv sync"
    ) from e

try:
    from o2md.d2md import WordToMarkdownConverter, convert_doc_to_docx
    from o2md import d2md
except ImportError as e:
    raise ImportError(
        "d2md.pyのインポートに失敗しました。必要な依存関係をインストールしてください: uv sync"
    ) from e

try:
    from o2md.p2md import PowerPointToMarkdownConverter
    from o2md import p2md
except ImportError as e:
    raise ImportError(
        "p2md.pyのインポートに失敗しました。必要な依存関係をインストールしてください: uv sync"
    ) from e

try:
    from o2md.pdf2md import PDFToMarkdownConverter
    from o2md import pdf2md
except ImportError as e:
    raise ImportError(
        "pdf2md.pyのインポートに失敗しました。必要な依存関係をインストールしてください: uv sync"
    ) from e

try:
    from o2md.jtd2md import JtdToMarkdownConverter
    from o2md import jtd2md
except ImportError as e:
    raise ImportError(
        "jtd2md.pyのインポートに失敗しました。必要な依存関係をインストールしてください: pip install olefile"
    ) from e

try:
    from o2md.img2md import ImageToMarkdownConverter
    from o2md import img2md
except ImportError as e:
    raise ImportError(
        "img2md.pyのインポートに失敗しました。必要な依存関係をインストールしてください: uv sync"
    ) from e



logger = logging.getLogger(__name__)

# グローバルverboseフラグ
_VERBOSE = False

def setup_logging(verbose: bool):
    """loggingの設定"""
    level = logging.DEBUG if verbose else logging.WARNING
    logging.basicConfig(
        level=level,
        format='[%(levelname)s] %(message)s',
    )

def set_verbose(verbose: bool):
    """verboseモードを設定"""
    global _VERBOSE
    _VERBOSE = verbose
    setup_logging(verbose)
    x2md.set_verbose(verbose)
    d2md.set_verbose(verbose)
    p2md.set_verbose(verbose)
    pdf2md.set_verbose(verbose)
    jtd2md.set_verbose(verbose)
    img2md.set_verbose(verbose)

def is_verbose() -> bool:
    """verboseモードかどうかを返す"""
    return _VERBOSE

def debug_print(*args, **kwargs):
    """verboseモード時のみ出力するデバッグ用print"""
    if _VERBOSE:
        print(*args, **kwargs)

def _is_auto_generated_heading(line: str, heading_patterns: list) -> bool:
    """Markdown見出し行がプログラム生成かどうかを判定する

    Args:
        line: 元のMarkdown行（見出し記号付き）
        heading_patterns: コンバータのget_auto_generated_patterns()が返したパターンリスト

    Returns:
        プログラム生成の見出しならTrue
    """
    import re
    m = re.match(r'^#{1,6}\s+(.+)$', line.rstrip())
    if not m:
        return False
    heading_text = m.group(1).strip()
    return any(p.match(heading_text) for p in heading_patterns)


def strip_markdown(text: str, auto_patterns: dict = None) -> str:
    """Markdownの書式記号を除去してプレーンテキストに変換する（公開API）

    Args:
        text: Markdownテキスト
        auto_patterns: コンバータから取得したプログラム生成パターン情報
            {'heading_patterns': list[re.Pattern], 'html_tags': list[str]}

    除去対象:
    - プログラム生成見出し（コンバータが定義したパターン）
    - コンバータが定義したHTMLタグ
    - コンバータが定義した行パターン（メタデータ等）
    - 見出し記号 (##)
    - 太字/斜体 (**text**, *text*)
    - テーブル区切り行 (| --- | --- |)
    - 画像リンク (![alt](path))
    - 行末の改行用スペース (trailing two spaces)
    - 水平線 (---, ***)
    """
    import re
    if auto_patterns is None:
        auto_patterns = {'heading_patterns': [], 'html_tags': [], 'line_patterns': []}
    heading_patterns = auto_patterns.get('heading_patterns', [])
    html_tags = set(auto_patterns.get('html_tags', []))
    line_patterns = auto_patterns.get('line_patterns', [])
    lines = text.split('\n')
    result = []
    for line in lines:
        # プログラムが付与した見出し行を除去
        if heading_patterns and _is_auto_generated_heading(line, heading_patterns):
            continue
        # コンバータが定義したHTMLタグを除去
        stripped = line.strip()
        if html_tags and stripped in html_tags:
            continue
        # <summary>...</summary> タグを汎用的に除去（html_tagsに<summary>系が含まれる場合）
        if html_tags and re.match(r'^<summary>.*</summary>$', stripped):
            continue
        # コンバータが定義した行パターンを除去（メタデータ等）
        if line_patterns and any(p.match(stripped) for p in line_patterns):
            continue
        # 画像リンク行を除去
        if re.match(r'^\s*!\[.*?\]\(.*?\)\s*$', line):
            continue
        # テーブル区切り行を除去 (| --- | --- | 形式)
        if re.match(r'^\s*\|[\s\-:|]+\|\s*$', line):
            continue
        # 水平線を除去
        if re.match(r'^\s*([-*_])\s*\1\s*\1[\s\1]*$', line):
            continue
        # 見出し記号を除去
        line = re.sub(r'^#{1,6}\s+', '', line)
        # 箇条書き記号を除去（インデント保持）
        line = re.sub(r'^(\s*)[-*+]\s+', r'\1', line)
        # 番号付きリスト記号を除去（インデント保持）
        line = re.sub(r'^(\s*)\d+\.\s+', r'\1', line)
        # 太字記号を除去
        line = re.sub(r'\*\*(.+?)\*\*', r'\1', line)
        # 斜体記号を除去
        line = re.sub(r'\*(.+?)\*', r'\1', line)
        # Markdownリンクをテキスト部分のみに変換 [text](url) -> text
        line = re.sub(r'\[([^\]]*)\]\([^)]*\)', r'\1', line)
        # テーブルのパイプ記号を除去してタブ区切りに変換
        if '|' in line and line.strip().startswith('|'):
            cells = [c.strip() for c in line.strip().strip('|').split('|')]
            line = '\t'.join(cells)
        # 行末のMarkdown改行用スペースを除去
        line = line.rstrip()
        result.append(line)
    return '\n'.join(result)


def _remove_image_links(text: str) -> str:
    """Markdownテキストから画像リンク行のみを除去する

    画像リンク(![alt](path))を含む行を削除し、
    それ以外のMarkdown書式はそのまま保持する。
    """
    import re
    lines = text.split('\n')
    result = []
    for line in lines:
        if re.match(r'^\s*!\[.*?\]\(.*?\)\s*$', line):
            continue
        result.append(line)
    return '\n'.join(result)


def convert_md_to_text(md_file_path: str, auto_patterns: dict = None,
                       remove_md: bool = False) -> str:
    """Markdownファイルをプレーンテキストに変換して.txtとして保存する

    Args:
        md_file_path: 変換元の.mdファイルパス
        auto_patterns: コンバータから取得したプログラム生成パターン情報
        remove_md: Trueの場合、変換後に元の.mdファイルを削除する

    Returns:
        出力した.txtファイルのパス
    """
    with open(md_file_path, 'r', encoding='utf-8') as f:
        md_content = f.read()

    text_content = strip_markdown(md_content, auto_patterns=auto_patterns)

    txt_file_path = md_file_path.rsplit('.md', 1)[0] + '.txt'
    with open(txt_file_path, 'w', encoding='utf-8') as f:
        f.write(text_content)

    if remove_md:
        os.remove(md_file_path)

    return txt_file_path


def strip_images_from_md(md_file_path: str):
    """Markdownファイルから画像リンク行を除去して上書きする

    Args:
        md_file_path: 対象の.mdファイルパス
    """
    with open(md_file_path, 'r', encoding='utf-8') as f:
        content = f.read()

    cleaned = _remove_image_links(content)

    with open(md_file_path, 'w', encoding='utf-8') as f:
        f.write(cleaned)


def detect_file_type(file_path: str) -> str:
    """ファイル拡張子からファイルタイプを判定
    
    Args:
        file_path: ファイルパス
        
    Returns:
        'excel', 'word', 'powerpoint', 'pdf', 'ichitaro', 'image', 'unknown'のいずれか
    """
    file_path_lower = file_path.lower()
    
    if file_path_lower.endswith(('.xlsx', '.xls')):
        return 'excel'
    elif file_path_lower.endswith(('.docx', '.doc')):
        return 'word'
    elif file_path_lower.endswith(('.pptx', '.ppt')):
        return 'powerpoint'
    elif file_path_lower.endswith('.pdf'):
        return 'pdf'
    elif file_path_lower.endswith(('.jtd', '.jtt')):
        return 'ichitaro'
    elif file_path_lower.endswith(('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.tif', '.webp')):
        return 'image'
    else:
        return 'unknown'


def convert_office_to_markdown(file_path: str, output_dir: str = None, **kwargs) -> tuple:
    """Officeファイルを自動判定してMarkdownに変換
    
    Args:
        file_path: 変換するOfficeファイルのパス
        output_dir: 出力ディレクトリ（省略時はデフォルト）
        **kwargs: 各変換クラス固有のオプション
            - use_heading_text: Word変換時に見出しテキストをリンクに使用（デフォルト: False）
            - shape_metadata: 図形メタデータを出力（デフォルト: False）
            - output_format: 出力画像形式 ('png' または 'svg'、デフォルト: 'png')
            
    Returns:
        (出力ファイルのパス, プログラム生成パターン情報の辞書, 出力画像数) の3-tuple
        パターン情報: {'heading_patterns': list[re.Pattern], 'html_tags': list[str]}
        出力画像数: int（実際にimages/に保存した画像ファイル数）
        
    Raises:
        ValueError: サポートされていないファイル形式
        FileNotFoundError: ファイルが見つからない
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(_("ファイルが見つかりません: {file}").format(file=file_path))
    
    file_type = detect_file_type(file_path)
    
    if file_type == 'unknown':
        raise ValueError(
            _("サポートされていないファイル形式です: {file}").format(file=file_path) + "\n"
            + _("対応形式: {formats}").format(formats=".xlsx, .xls, .docx, .doc, .pptx, .ppt, .pdf, .jtd, .jtt, .jpg, .jpeg, .png, .gif, .bmp, .tiff, .tif, .webp")
        )
    
    print(_("ファイルタイプを検出: {file_type}").format(file_type=file_type))
    
    converter = None
    output_file = None
    converted_file = None
    converted_temp_dir = None
    auto_patterns = {'heading_patterns': [], 'html_tags': []}
    
    try:
        if file_type == 'excel':
            # Excel変換
            processing_file = file_path
            
            # XLSファイルの場合は事前にXLSXに変換
            if file_path.lower().endswith('.xls'):
                converted_file = convert_xls_to_xlsx(file_path)
                if converted_file is None:
                    raise RuntimeError("XLS→XLSX変換に失敗しました。")
                processing_file = converted_file
                converted_temp_dir = Path(converted_file).parent
                print(_("XLS→XLSX変換完了: {file}").format(file=file_path))
            
            shape_metadata = kwargs.get('shape_metadata', False)
            output_format = kwargs.get('output_format', 'png')
            converter = ExcelToMarkdownConverter(
                processing_file, 
                output_dir=output_dir, 
                shape_metadata=shape_metadata,
                output_format=output_format
            )
            output_file = converter.convert()
            
        elif file_type == 'word':
            # Word変換
            processing_file = file_path
            
            # DOCファイルの場合は事前にDOCXに変換
            if file_path.lower().endswith('.doc'):
                converted_file = convert_doc_to_docx(file_path)
                if converted_file is None:
                    raise RuntimeError("DOC→DOCX変換に失敗しました。")
                processing_file = converted_file
                print(_("DOC→DOCX変換完了: {file}").format(file=file_path))
            
            use_heading_text = kwargs.get('use_heading_text', False)
            shape_metadata = kwargs.get('shape_metadata', False)
            output_format = kwargs.get('output_format', 'png')
            # DOC→DOCX変換時は元のファイルパスを表示用に渡す
            display_name = file_path if processing_file != file_path else None
            converter = WordToMarkdownConverter(
                processing_file, 
                use_heading_text=use_heading_text,
                output_dir=output_dir,
                shape_metadata=shape_metadata,
                output_format=output_format,
                display_name=display_name
            )
            output_file = converter.convert()
            
        elif file_type == 'powerpoint':
            # PowerPoint変換
            output_format = kwargs.get('output_format', 'png')
            converter = PowerPointToMarkdownConverter(
                file_path,
                output_dir=output_dir,
                output_format=output_format
            )
            output_file = converter.convert()
        
        elif file_type == 'pdf':
            # PDF変換
            output_format = kwargs.get('output_format', 'png')
            ocr_engine = kwargs.get('ocr_engine', 'tesseract')
            tessdata_dir = kwargs.get('tessdata_dir')
            use_docling = kwargs.get('use_docling', False)
            converter = PDFToMarkdownConverter(
                file_path,
                output_dir=output_dir,
                output_format=output_format,
                ocr_engine=ocr_engine,
                tessdata_dir=tessdata_dir,
                use_docling=use_docling
            )
            output_file = converter.convert()

        elif file_type == 'ichitaro':
            # 一太郎変換
            converter = JtdToMarkdownConverter(
                file_path,
                output_dir=output_dir,
            )
            output_file = converter.convert()

        elif file_type == 'image':
            # 画像OCR変換
            ocr_engine = kwargs.get('ocr_engine', 'tesseract')
            tessdata_dir = kwargs.get('tessdata_dir', None)
            converter = ImageToMarkdownConverter(
                file_path,
                output_dir=output_dir,
                ocr_engine=ocr_engine,
                tessdata_dir=tessdata_dir,
            )
            output_file = converter.convert()
        
        # コンバータからプログラム生成パターンを取得
        if converter and hasattr(converter, 'get_auto_generated_patterns'):
            auto_patterns['heading_patterns'] = converter.get_auto_generated_patterns()
        if converter and hasattr(converter, 'get_auto_generated_html_tags'):
            auto_patterns['html_tags'] = converter.get_auto_generated_html_tags()
        if converter and hasattr(converter, 'get_auto_generated_line_patterns'):
            auto_patterns['line_patterns'] = converter.get_auto_generated_line_patterns()

        # 出力画像カウントを取得
        output_image_count = 0
        if converter and hasattr(converter, 'output_image_count'):
            output_image_count = converter.output_image_count

        return output_file, auto_patterns, output_image_count
        
    finally:
        # PowerPointの一時ファイルクリーンアップ
        if file_type == 'powerpoint' and converter:
            converter.cleanup()

        # Excelのworkbookを明示的にクローズ
        if file_type == 'excel' and converter:
            try:
                if hasattr(converter, 'workbook') and converter.workbook:
                    converter.workbook.close()
            except Exception:
                pass

        # コンバータオブジェクトの参照を解放
        del converter

        # Excel/Wordの一時変換ファイルをクリーンアップ
        if converted_temp_dir:
            try:
                if converted_temp_dir.exists() and (
                    converted_temp_dir.name.startswith('xls2md_conversion_') or
                    converted_temp_dir.name.startswith('d2md_conversion_')
                ):
                    import shutil
                    shutil.rmtree(converted_temp_dir)
                    logger.info(f"一時ディレクトリを削除しました: {converted_temp_dir}")
            except Exception as e:
                logger.warning(f"一時ディレクトリの削除に失敗しました: {e}")
        
        if converted_file and file_type == 'word':
            try:
                parent_dir = Path(converted_file).parent
                if parent_dir.exists() and parent_dir.name.startswith('d2md_conversion_'):
                    import shutil
                    shutil.rmtree(parent_dir)
                    logger.info(f"一時ディレクトリを削除しました: {parent_dir}")
            except Exception as e:
                logger.warning(f"一時ディレクトリの削除に失敗しました: {e}")


# 対応拡張子一覧
SUPPORTED_EXTENSIONS = (
    '.xlsx', '.xls', '.docx', '.doc', '.pptx', '.ppt', '.pdf',
    '.jtd', '.jtt',
    '.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.tif', '.webp',
)


def collect_target_files(folder_path: str, recursive: bool = False) -> list:
    """フォルダ内の変換対象ファイルを収集

    Args:
        folder_path: 検索対象フォルダパス
        recursive: Trueの場合サブフォルダも再帰的に探索

    Returns:
        対象ファイルの絶対パスリスト（ソート済み）
    """
    target_files = []
    folder = Path(folder_path).resolve()
    if recursive:
        for root, _dirs, files in os.walk(folder):
            for fname in files:
                if fname.lower().endswith(SUPPORTED_EXTENSIONS):
                    target_files.append(os.path.join(root, fname))
    else:
        for fname in os.listdir(folder):
            fpath = os.path.join(folder, fname)
            if os.path.isfile(fpath) and fname.lower().endswith(SUPPORTED_EXTENSIONS):
                target_files.append(fpath)
    target_files.sort()
    return target_files


def convert_folder(folder_path: str, output_dir: str = None, recursive: bool = False, **kwargs) -> dict:
    """フォルダ内の全対象ファイルを一括変換

    Args:
        folder_path: 変換対象フォルダのパス
        output_dir: 出力ベースディレクトリ（省略時: ./output）
        recursive: Trueの場合サブフォルダも再帰的に処理
        **kwargs: convert_office_to_markdownに渡すオプション

    Returns:
        {'success': [...], 'failed': [...]} 形式の結果辞書
    """
    folder = Path(folder_path).resolve()
    if not folder.is_dir():
        raise NotADirectoryError(_("フォルダが見つかりません: {folder}").format(folder=folder_path))

    base_output = Path(output_dir) if output_dir else Path("output")
    target_files = collect_target_files(str(folder), recursive=recursive)

    if not target_files:
        logger.warning(f"変換対象ファイルが見つかりません: {folder_path}")
        return {'success': [], 'failed': []}

    total = len(target_files)
    print(_("フォルダ一括変換開始: {folder}").format(folder=folder))
    print(_("対象ファイル数: {total}").format(total=total))
    print("=" * 50)

    results = {'success': [], 'failed': []}

    from o2md.utils import is_text_only

    for idx, fpath in enumerate(target_files, 1):
        rel = Path(fpath).relative_to(folder)
        # 出力ディレクトリ: base_output / 元の相対パスのディレクトリ
        file_output_dir = str(base_output / rel.parent)

        print(f"\n[{idx}/{total}] {rel}")
        try:
            output_file, auto_patterns, img_count = convert_office_to_markdown(
                fpath,
                output_dir=file_output_dir,
                **kwargs
            )
            # テキストモードでは各コンバータが直接.txtを出力するため追加処理不要
            results['success'].append({'file': str(rel), 'output': output_file})

            # 個別ファイルの結果を表示
            print("\n" + "=" * 50)
            print(_("出力ファイル: {output_file}").format(output_file=output_file))
            if not is_text_only():
                file_type = detect_file_type(fpath)
                if file_type != 'image' and img_count > 0:
                    print(_("出力画像: {count}枚").format(count=img_count))
            print("=" * 50)
        except Exception as e:
            logger.error(f"変換失敗: {rel} - {e}")
            results['failed'].append({'file': str(rel), 'error': str(e)})
        finally:
            # 前のファイルのコンバータオブジェクトや中間データを解放
            gc.collect()

    return results


def main():
    """メイン関数"""
    parser = argparse.ArgumentParser(
        description='Office文書（Excel、Word、PowerPoint、PDF、一太郎）をMarkdownに変換',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
対応ファイル形式:
  Excel:      .xlsx, .xls
  Word:       .docx, .doc
  PowerPoint: .pptx, .ppt
  PDF:        .pdf
  一太郎:    .jtd, .jtt

使用例:
  python o2md.py data.xlsx
  python o2md.py document.docx --use-heading-text
  python o2md.py presentation.pptx -o custom_output
  python o2md.py document.pdf -o pdf_output
  python o2md.py ./input_folder/             # フォルダ内のファイルを一括変換
  python o2md.py ./input_folder/ -r          # サブフォルダも再帰的に変換
  python o2md.py ./input_folder/ -r -o out   # 出力先指定
        """
    )

    parser.add_argument('file', help='変換するOfficeファイルまたはフォルダ')
    parser.add_argument('-o', '--output-dir', type=str,
                       help='出力ディレクトリを指定（デフォルト: ./output）')
    parser.add_argument('-r', '--recursive', action='store_true',
                       help='[フォルダ指定時] サブフォルダも再帰的に処理する')
    parser.add_argument('--use-heading-text', action='store_true',
                       help='[Word専用] 章番号の代わりに見出しテキストをリンクに使用')
    parser.add_argument('--shape-metadata', action='store_true',
                       help='図形メタデータを画像の後に出力（テキスト形式とJSON形式）')
    parser.add_argument('--format', choices=['png', 'svg'], default='svg',
                       help='出力画像形式を指定（デフォルト: svg）')
    parser.add_argument('--ocr-engine', choices=['manga-ocr', 'tesseract', 'sarashina'],
                       default='tesseract',
                       help='[PDF専用] OCRエンジンを指定（デフォルト: tesseract）')
    parser.add_argument('--tessdata-dir', type=str,
                       help='[PDF専用] tessdataディレクトリを指定（tessdata_best使用時）')
    parser.add_argument('--docling', action='store_true',
                       help='[PDF専用] doclingによる表検出を有効にする')
    parser.add_argument('--text', action='store_true',
                       help='テキスト抽出モード（.txtのみ出力、.mdは生成しない）')
    parser.add_argument('--lang', choices=['ja', 'en'], default=None,
                       help='表示言語を指定（未指定時はLANG環境変数から判定、デフォルト: ja）')
    parser.add_argument('-v', '--verbose', action='store_true',
                       help='デバッグ情報を出力し、debug_workbooks/pdfs/diagnosticsフォルダを保存')

    args = parser.parse_args()

    # 多言語設定の初期化（他の出力より先に実行）
    setup_i18n(args.lang)

    set_verbose(args.verbose)

    # テキストモードの設定
    from o2md.utils import set_text_only, is_libreoffice_available, warn_libreoffice_not_available
    if args.text:
        set_text_only(True)
        print(_("テキストモード: .txtのみを出力します"))

    # LibreOfficeの利用可否をチェックし、利用できない場合は警告を表示
    if not is_libreoffice_available():
        warn_libreoffice_not_available()

    common_kwargs = dict(
        use_heading_text=args.use_heading_text,
        shape_metadata=args.shape_metadata,
        output_format=args.format,
        ocr_engine=args.ocr_engine,
        tessdata_dir=args.tessdata_dir,
        use_docling=args.docling,
    )

    # フォルダが指定された場合: 一括変換
    if os.path.isdir(args.file):
        try:
            results = convert_folder(
                args.file,
                output_dir=args.output_dir,
                recursive=args.recursive,
                **common_kwargs
            )

            print("\n" + "=" * 50)
            print(_("成功: {count}ファイル").format(count=len(results['success'])))
            if results['failed']:
                print(_("失敗: {count}ファイル").format(count=len(results['failed'])))
                for item in results['failed']:
                    print("  - {file}: {error}".format(file=item['file'], error=item['error']))
            print("=" * 50)

            if results['failed']:
                sys.exit(1)

        except NotADirectoryError as e:
            print(_("エラー: {message}").format(message=e))
            sys.exit(1)
        except Exception as e:
            print(_("変換エラー: {message}").format(message=e))
            import traceback
            traceback.print_exc()
            sys.exit(1)
    else:
        # 単一ファイル変換
        try:
            output_file, auto_patterns, output_image_count = convert_office_to_markdown(
                args.file,
                output_dir=args.output_dir,
                **common_kwargs
            )

            print("\n" + "=" * 50)
            print(_("出力ファイル: {output_file}").format(output_file=output_file))

            # 出力画像数を表示（画像OCR変換時・テキストモード時は除外）
            if not args.text and detect_file_type(args.file) != 'image':
                if output_image_count > 0:
                    print(_("出力画像: {count}枚").format(count=output_image_count))

            if args.use_heading_text:
                print(_("見出しテキストリンクモード: 有効"))

            print("=" * 50)

        except ValueError as e:
            print(_("エラー: {message}").format(message=e))
            sys.exit(1)
        except FileNotFoundError as e:
            print(_("エラー: {message}").format(message=e))
            sys.exit(1)
        except Exception as e:
            print(_("変換エラー: {message}").format(message=e))
            import traceback
            traceback.print_exc()
            sys.exit(1)


if __name__ == "__main__":
    main()
