#!/usr/bin/env python3
"""
Office to Markdown Converter (o2md)
Excel、Word、PowerPointファイルを自動判定してMarkdownに変換する統合ツール

機能:
- ファイル拡張子に基づいて自動的に適切な変換クラスを選択
- Excel (.xlsx, .xls) → x2md.ExcelToMarkdownConverter
- Word (.docx, .doc) → d2md.WordToMarkdownConverter
- PowerPoint (.pptx, .ppt) → p2md.PowerPointToMarkdownConverter
- 古い形式（.xls, .doc, .ppt）は自動的に新形式に変換してから処理

対応ファイル形式:
- Excel: .xlsx, .xls
- Word: .docx, .doc
- PowerPoint: .pptx, .ppt

使用例:
    # 基本的な使用方法
    python o2md.py input_files/data.xlsx
    python o2md.py input_files/document.docx
    python o2md.py input_files/presentation.pptx
    
    # 出力ディレクトリを指定
    python o2md.py input_files/data.xlsx -o custom_output
    
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

必要な依存関係:
- x2md.py, d2md.py, p2md.py
- openpyxl, python-docx, python-pptx
- Pillow (PIL)
- LibreOffice (古い形式の変換に必要)
"""

import os
import sys
import argparse
from pathlib import Path

# 各変換クラスをインポート
try:
    from x2md import ExcelToMarkdownConverter, convert_xls_to_xlsx
    import x2md
except ImportError as e:
    raise ImportError(
        "x2md.pyのインポートに失敗しました。必要な依存関係をインストールしてください: uv sync"
    ) from e

try:
    from d2md import WordToMarkdownConverter, convert_doc_to_docx
    import d2md
except ImportError as e:
    raise ImportError(
        "d2md.pyのインポートに失敗しました。必要な依存関係をインストールしてください: uv sync"
    ) from e

try:
    from p2md import PowerPointToMarkdownConverter
    import p2md
except ImportError as e:
    raise ImportError(
        "p2md.pyのインポートに失敗しました。必要な依存関係をインストールしてください: uv sync"
    ) from e



# グローバルverboseフラグ
_VERBOSE = False

def set_verbose(verbose: bool):
    """verboseモードを設定"""
    global _VERBOSE
    _VERBOSE = verbose
    x2md.set_verbose(verbose)
    d2md.set_verbose(verbose)
    p2md.set_verbose(verbose)

def is_verbose() -> bool:
    """verboseモードかどうかを返す"""
    return _VERBOSE

def debug_print(*args, **kwargs):
    """verboseモード時のみ出力するデバッグ用print"""
    if _VERBOSE:
        print(*args, **kwargs)

def detect_file_type(file_path: str) -> str:
    """ファイル拡張子からファイルタイプを判定
    
    Args:
        file_path: ファイルパス
        
    Returns:
        'excel', 'word', 'powerpoint', 'unknown'のいずれか
    """
    file_path_lower = file_path.lower()
    
    if file_path_lower.endswith(('.xlsx', '.xls')):
        return 'excel'
    elif file_path_lower.endswith(('.docx', '.doc')):
        return 'word'
    elif file_path_lower.endswith(('.pptx', '.ppt')):
        return 'powerpoint'
    else:
        return 'unknown'


def convert_office_to_markdown(file_path: str, output_dir: str = None, **kwargs) -> str:
    """Officeファイルを自動判定してMarkdownに変換
    
    Args:
        file_path: 変換するOfficeファイルのパス
        output_dir: 出力ディレクトリ（省略時はデフォルト）
        **kwargs: 各変換クラス固有のオプション
            - use_heading_text: Word変換時に見出しテキストをリンクに使用（デフォルト: False）
            - shape_metadata: 図形メタデータを出力（デフォルト: False）
            - output_format: 出力画像形式 ('png' または 'svg'、デフォルト: 'png')
            
    Returns:
        出力ファイルのパス
        
    Raises:
        ValueError: サポートされていないファイル形式
        FileNotFoundError: ファイルが見つからない
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"ファイルが見つかりません: {file_path}")
    
    file_type = detect_file_type(file_path)
    
    if file_type == 'unknown':
        raise ValueError(
            f"サポートされていないファイル形式です: {file_path}\n"
            "対応形式: .xlsx, .xls, .docx, .doc, .pptx, .ppt"
        )
    
    print(f"[INFO] ファイルタイプを検出: {file_type}")
    print(f"[INFO] 変換開始: {file_path}")
    
    converter = None
    output_file = None
    converted_file = None
    converted_temp_dir = None
    
    try:
        if file_type == 'excel':
            # Excel変換
            processing_file = file_path
            
            # XLSファイルの場合は事前にXLSXに変換
            if file_path.lower().endswith('.xls'):
                print("[INFO] XLSファイルが指定されました。XLSXに変換します...")
                converted_file = convert_xls_to_xlsx(file_path)
                if converted_file is None:
                    raise RuntimeError("XLS→XLSX変換に失敗しました。")
                processing_file = converted_file
                converted_temp_dir = Path(converted_file).parent
                print(f"[INFO] ✅ XLS→XLSX変換完了: {converted_file}")
            
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
                print("[INFO] DOCファイルが指定されました。DOCXに変換します...")
                converted_file = convert_doc_to_docx(file_path)
                if converted_file is None:
                    raise RuntimeError("DOC→DOCX変換に失敗しました。")
                processing_file = converted_file
                print(f"[INFO] ✅ DOC→DOCX変換完了: {converted_file}")
            
            use_heading_text = kwargs.get('use_heading_text', False)
            shape_metadata = kwargs.get('shape_metadata', False)
            output_format = kwargs.get('output_format', 'png')
            converter = WordToMarkdownConverter(
                processing_file, 
                use_heading_text=use_heading_text,
                output_dir=output_dir,
                shape_metadata=shape_metadata,
                output_format=output_format
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
        
        return output_file
        
    finally:
        # PowerPointの一時ファイルクリーンアップ
        if file_type == 'powerpoint' and converter:
            converter.cleanup()
        
        # Excel/Wordの一時変換ファイルをクリーンアップ
        if converted_temp_dir:
            try:
                if converted_temp_dir.exists() and (
                    converted_temp_dir.name.startswith('xls2md_conversion_') or
                    converted_temp_dir.name.startswith('d2md_conversion_')
                ):
                    import shutil
                    shutil.rmtree(converted_temp_dir)
                    print(f"[INFO] 一時ディレクトリを削除しました: {converted_temp_dir}")
            except Exception as e:
                print(f"[WARNING] 一時ディレクトリの削除に失敗しました: {e}")
        
        if converted_file and file_type == 'word':
            try:
                parent_dir = Path(converted_file).parent
                if parent_dir.exists() and parent_dir.name.startswith('d2md_conversion_'):
                    import shutil
                    shutil.rmtree(parent_dir)
                    print(f"[INFO] 一時ディレクトリを削除しました: {parent_dir}")
            except Exception as e:
                print(f"[WARNING] 一時ディレクトリの削除に失敗しました: {e}")


def main():
    """メイン関数"""
    parser = argparse.ArgumentParser(
        description='Office文書（Excel、Word、PowerPoint）をMarkdownに変換',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
対応ファイル形式:
  Excel:      .xlsx, .xls
  Word:       .docx, .doc
  PowerPoint: .pptx, .ppt

使用例:
  python o2md.py data.xlsx
  python o2md.py document.docx --use-heading-text
  python o2md.py presentation.pptx -o custom_output
        """
    )
    
    parser.add_argument('file', help='変換するOfficeファイル')
    parser.add_argument('-o', '--output-dir', type=str,
                       help='出力ディレクトリを指定（デフォルト: ./output）')
    parser.add_argument('--use-heading-text', action='store_true',
                       help='[Word専用] 章番号の代わりに見出しテキストをリンクに使用')
    parser.add_argument('--shape-metadata', action='store_true',
                       help='図形メタデータを画像の後に出力（テキスト形式とJSON形式）')
    parser.add_argument('--format', choices=['png', 'svg'], default='png',
                       help='出力画像形式を指定（デフォルト: png）')
    parser.add_argument('-v', '--verbose', action='store_true',
                       help='デバッグ情報を出力')
    
    args = parser.parse_args()
    
    set_verbose(args.verbose)
    
    try:
        output_file = convert_office_to_markdown(
            args.file,
            output_dir=args.output_dir,
            use_heading_text=args.use_heading_text,
            shape_metadata=args.shape_metadata,
            output_format=args.format
        )
        
        print("\n" + "=" * 50)
        print("✅ 変換完了!")
        print(f"📄 出力ファイル: {output_file}")
        
        # 画像ディレクトリの情報を表示
        if args.output_dir:
            images_dir = os.path.join(args.output_dir, "images")
        else:
            images_dir = os.path.join(os.getcwd(), "output", "images")
        
        if os.path.exists(images_dir) and os.listdir(images_dir):
            print(f"🖼️  画像フォルダ: {images_dir}")
        
        if args.use_heading_text:
            print("📝 見出しテキストリンクモード: 有効")
        
        print("=" * 50)
        
    except ValueError as e:
        print(f"❌ エラー: {e}")
        sys.exit(1)
    except FileNotFoundError as e:
        print(f"❌ エラー: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"❌ 変換エラー: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
