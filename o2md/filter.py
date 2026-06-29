#!/usr/bin/env python3
"""
o2md-filter: 検索エンジンインデックス用テキスト抽出フィルタ

ファイルパス指定またはstdinからOffice文書を読み込み、
プレーンテキストをstdoutに出力する。

使用例:
    # ファイルパス指定
    o2md-filter document.xlsx
    o2md-filter report.pdf

    # stdinから入力（マジックバイトで自動判別）
    cat document.xlsx | o2md-filter
    cat report.pdf | o2md-filter

    # 検索エンジン連携
    find /docs -name "*.xlsx" | while read f; do
        o2md-filter "$f" | index-tool --source "$f"
    done
"""

import argparse
import os
import sys
import tempfile
import logging

from o2md.utils import set_text_only


logger = logging.getLogger(__name__)


# マジックバイトによるファイルタイプ判定テーブル
_MAGIC_SIGNATURES = [
    # (バイト列, ファイルタイプ)
    (b'%PDF', 'pdf'),
    (b'\xd0\xcf\x11\xe0', 'ole'),  # OLE2 (doc/xls/ppt/jtd)
]

# ZIP系 (xlsx/docx/pptx) はPKシグネチャ + 内部ファイルで判別
_ZIP_SIGNATURE = b'PK\x03\x04'


def detect_type_from_bytes(header: bytes) -> str:
    """先頭バイトからファイルタイプを判定する

    Args:
        header: ファイル先頭のバイト列（最低4バイト）

    Returns:
        'excel', 'word', 'powerpoint', 'pdf', 'ole', 'zip', 'image', 'unknown'
    """
    if len(header) < 4:
        return 'unknown'

    # PDF判定
    if header[:4] == b'%PDF':
        return 'pdf'

    # OLE2判定 (doc/xls/ppt/jtd)
    if header[:4] == b'\xd0\xcf\x11\xe0':
        return 'ole'

    # ZIP系 (xlsx/docx/pptx)
    if header[:4] == _ZIP_SIGNATURE:
        return 'zip'

    # 画像判定
    if header[:8] == b'\x89PNG\r\n\x1a\n':
        return 'image'
    if header[:2] == b'\xff\xd8':
        return 'image'
    if header[:4] == b'GIF8':
        return 'image'
    if header[:2] == b'BM':
        return 'image'
    if header[:4] == b'RIFF' and len(header) >= 12 and header[8:12] == b'WEBP':
        return 'image'
    # TIFF
    if header[:2] in (b'II', b'MM') and len(header) >= 4:
        if header[:2] == b'II' and header[2:4] == b'\x2a\x00':
            return 'image'
        if header[:2] == b'MM' and header[2:4] == b'\x00\x2a':
            return 'image'

    return 'unknown'


def detect_zip_subtype(file_path: str) -> str:
    """ZIPファイルの中身からOfficeサブタイプを判定する

    Args:
        file_path: ZIPファイルのパス

    Returns:
        'excel', 'word', 'powerpoint', 'unknown'
    """
    import zipfile
    try:
        with zipfile.ZipFile(file_path, 'r') as zf:
            names = zf.namelist()
            # Content_Typesから判定
            if '[Content_Types].xml' in names:
                content_types = zf.read('[Content_Types].xml').decode('utf-8', errors='ignore')
                if 'spreadsheetml' in content_types or 'xl/' in ' '.join(names):
                    return 'excel'
                if 'wordprocessingml' in content_types or 'word/' in ' '.join(names):
                    return 'word'
                if 'presentationml' in content_types or 'ppt/' in ' '.join(names):
                    return 'powerpoint'
            # フォルダ構造から判定
            for name in names:
                if name.startswith('xl/'):
                    return 'excel'
                if name.startswith('word/'):
                    return 'word'
                if name.startswith('ppt/'):
                    return 'powerpoint'
    except (zipfile.BadZipFile, Exception):
        pass
    return 'unknown'


def detect_ole_subtype(file_path: str) -> str:
    """OLE2ファイルのサブタイプを判定する

    Args:
        file_path: OLE2ファイルのパス

    Returns:
        'word', 'excel', 'powerpoint', 'ichitaro', 'unknown'
    """
    try:
        import olefile
        with olefile.OleFileIO(file_path) as ole:
            streams = ole.listdir()
            stream_names = ['/'.join(s) for s in streams]
            joined = ' '.join(stream_names).lower()

            # 一太郎判定
            if any('jtdocument' in s.lower() or 'jtdbody' in s.lower()
                   or 'jsrv_segmentinformation' in s.lower()
                   for s in stream_names):
                return 'ichitaro'
            # Word判定
            if 'worddocument' in joined or '1table' in joined:
                return 'word'
            # Excel判定
            if 'workbook' in joined or 'book' in joined:
                return 'excel'
            # PowerPoint判定
            if 'powerpoint document' in joined or 'current user' in joined:
                return 'powerpoint'
    except Exception:
        pass
    return 'unknown'


def resolve_file_type(file_path: str) -> str:
    """ファイルのタイプを確定する

    拡張子 → マジックバイト → ZIP/OLE内部構造の順に判定する。

    Args:
        file_path: ファイルパス

    Returns:
        'excel', 'word', 'powerpoint', 'pdf', 'ichitaro', 'image', 'unknown'
    """
    # 拡張子で判定可能ならそれを使用
    from o2md.o2md import detect_file_type
    ext_type = detect_file_type(file_path)
    if ext_type != 'unknown':
        return ext_type

    # マジックバイトで判定
    with open(file_path, 'rb') as f:
        header = f.read(12)

    base_type = detect_type_from_bytes(header)

    if base_type == 'zip':
        return detect_zip_subtype(file_path)
    elif base_type == 'ole':
        return detect_ole_subtype(file_path)
    elif base_type in ('pdf', 'image'):
        return base_type

    return 'unknown'


def filter_file(file_path: str, ocr_engine: str = 'tesseract') -> str:
    """ファイルをプレーンテキストに変換する

    convert_office_to_markdownがファイル拡張子からタイプを判定するため、
    呼び出し前に正しい拡張子を設定しておく必要がある。

    Args:
        file_path: 変換対象ファイルパス（正しい拡張子であること）
        ocr_engine: OCRエンジン ('tesseract', 'manga-ocr', 'sarashina')

    Returns:
        プレーンテキスト文字列
    """
    from o2md.o2md import convert_office_to_markdown, strip_markdown

    # テキストモードを有効化（画像処理スキップ）
    set_text_only(True)

    # 一時出力ディレクトリを使用
    with tempfile.TemporaryDirectory(prefix='o2md_filter_') as tmp_dir:
        output_file, auto_patterns, _ = convert_office_to_markdown(
            file_path,
            output_dir=tmp_dir,
            ocr_engine=ocr_engine,
        )

        # 出力ファイルを読み込み
        with open(output_file, 'r', encoding='utf-8') as f:
            content = f.read()

    # テキストモードでは既に.txt出力されるが、念のためstrip_markdownを適用
    # .txt出力の場合はそのまま返す
    if output_file.endswith('.txt'):
        return content

    # .md出力の場合はMarkdown記法を除去
    return strip_markdown(content, auto_patterns=auto_patterns)


def main():
    """o2md-filter メイン関数"""
    parser = argparse.ArgumentParser(
        description='Office文書をプレーンテキストに変換してstdoutに出力（検索エンジン前処理用）',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用例:
  o2md-filter document.xlsx              # ファイル指定→stdout
  o2md-filter report.pdf                 # PDF→stdout
  o2md-filter input.xlsx output.txt      # ファイル出力
  cat file.docx | o2md-filter            # stdinから入力（自動判別）

対応形式:
  Excel (.xlsx, .xls), Word (.docx, .doc), PowerPoint (.pptx, .ppt),
  PDF (.pdf), 一太郎 (.jtd, .jtt), 画像 (.jpg, .png, etc.)
        """
    )

    parser.add_argument('file', nargs='?', default=None,
                        help='変換対象ファイル（省略時はstdinから読み込み）')
    parser.add_argument('output', nargs='?', default=None,
                        help='出力先ファイルパス（省略時はstdoutに出力）')
    parser.add_argument('--ocr-engine', choices=['manga-ocr', 'tesseract', 'sarashina'],
                        default='tesseract',
                        help='OCRエンジンを指定（デフォルト: tesseract）')

    args = parser.parse_args()

    # loggingを抑制
    logging.basicConfig(
        level=logging.CRITICAL,
        stream=sys.stderr,
    )

    # 変換中の進捗メッセージを全て抑制（stdoutは結果専用）
    original_stdout = sys.stdout
    sys.stdout = open(os.devnull, 'w')

    try:
        if args.file:
            # ファイルパス指定
            if not os.path.exists(args.file):
                print(f"エラー: ファイルが見つかりません: {args.file}",
                      file=sys.stderr)
                sys.exit(1)

            file_type = resolve_file_type(args.file)
            if file_type == 'unknown':
                print(f"エラー: ファイルタイプを判別できません: {args.file}",
                      file=sys.stderr)
                sys.exit(1)

            text = filter_file(args.file, ocr_engine=args.ocr_engine)

        else:
            # stdinから読み込み
            if sys.stdin.isatty():
                print("エラー: ファイルを指定するかstdinにデータを入力してください",
                      file=sys.stderr)
                parser.print_help(sys.stderr)
                sys.exit(1)

            # stdinからバイナリ読み込み→一時ファイルに保存
            stdin_data = sys.stdin.buffer.read()
            if not stdin_data:
                print("エラー: stdinからデータを読み取れませんでした",
                      file=sys.stderr)
                sys.exit(1)

            # マジックバイトでタイプ判定
            base_type = detect_type_from_bytes(stdin_data[:12])

            # 一時ファイルに保存して処理
            suffix = _get_suffix_for_type(base_type)
            with tempfile.NamedTemporaryFile(
                suffix=suffix, delete=False, prefix='o2md_stdin_'
            ) as tmp:
                tmp.write(stdin_data)
                tmp_path = tmp.name

            try:
                file_type = resolve_file_type(tmp_path)
                if file_type == 'unknown':
                    print("エラー: ファイルタイプを判別できません",
                          file=sys.stderr)
                    sys.exit(1)

                # 正しい拡張子にリネーム（convert_office_to_markdownが
                # 拡張子でタイプ判定するため）
                correct_suffix = _type_to_extension(file_type, base_type)
                new_path = tmp_path + correct_suffix
                os.rename(tmp_path, new_path)
                tmp_path = new_path

                text = filter_file(tmp_path, ocr_engine=args.ocr_engine)
            finally:
                if os.path.exists(tmp_path):
                    os.unlink(tmp_path)

    except Exception as e:
        print(f"変換エラー: {e}", file=sys.stderr)
        sys.exit(1)
    finally:
        # stdoutを元に戻す
        sys.stdout.close()
        sys.stdout = original_stdout

    # 結果を出力
    if args.output:
        with open(args.output, 'w', encoding='utf-8') as out_f:
            out_f.write(text)
    else:
        original_stdout.write(text)


def _type_to_extension(file_type: str, base_type: str = 'zip') -> str:
    """確定したファイルタイプとコンテナ形式から正しい拡張子を返す

    OLE形式の場合はレガシー拡張子(.xls, .doc, .ppt)を返し、
    ZIP形式の場合はモダン拡張子(.xlsx, .docx, .pptx)を返す。
    これにより convert_office_to_markdown が古い形式を検出して
    LibreOfficeによる変換ステップを実行できる。

    Args:
        file_type: 'excel', 'word', 'powerpoint', 'pdf', 'ichitaro', 'image'
        base_type: コンテナ形式 ('ole', 'zip', 'pdf', 'image', 'unknown')

    Returns:
        拡張子文字列（ドット付き）
    """
    if base_type == 'ole':
        ole_ext_map = {
            'excel': '.xls',
            'word': '.doc',
            'powerpoint': '.ppt',
            'ichitaro': '.jtd',
        }
        if file_type in ole_ext_map:
            return ole_ext_map[file_type]

    zip_ext_map = {
        'excel': '.xlsx',
        'word': '.docx',
        'powerpoint': '.pptx',
        'pdf': '.pdf',
        'ichitaro': '.jtd',
        'image': '.png',
    }
    return zip_ext_map.get(file_type, '.bin')


def _get_suffix_for_type(base_type: str) -> str:
    """マジックバイトのベースタイプから一時ファイルの拡張子を決定する

    Args:
        base_type: マジックバイトから判定したベースタイプ

    Returns:
        拡張子文字列（ドット付き）
    """
    type_suffix_map = {
        'pdf': '.pdf',
        'ole': '.bin',
        'zip': '.zip',
        'image': '.png',
    }
    return type_suffix_map.get(base_type, '.bin')


if __name__ == "__main__":
    main()
