#!/usr/bin/env python3
"""
一太郎 (Ichitaro) to Markdown Converter
一太郎文書ファイル (.jtd, .jtt) をMarkdownに変換するツール

対応形式:
- .jtd (一太郎 ver8以降)
- .jtt (一太郎テンプレート ver8以降)

内部構造:
- OLE2 Compound Document形式
- DocumentTextストリームにUTF-16BEテキストを格納
- Footnoteストリームに脚注テキストを格納
- 制御コード(001C, 001F等)でセクション/書式を管理
"""

import os
import sys
import argparse
from pathlib import Path
from typing import Optional

try:
    import olefile
except ImportError:
    print("olefileライブラリが必要です: pip install olefile")
    sys.exit(1)


# 一太郎ファイルの対応拡張子
JTD_EXTENSIONS = ('.jtd', '.jtt')

# OLE2ストリーム名
STREAM_DOCUMENT_TEXT = "DocumentText"
STREAM_FOOTNOTE = "Footnote"

# ストリームヘッダのマジック文字列
HEADER_MAGIC_SSMG = b"SsmgV.01"
HEADER_MAGIC_CTEXT = b"CTextV.01"


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


def _read_u16be(data: bytes, pos: int) -> int:
    """バイト列からUTF-16BEの1コードユニットを読み取る"""
    if pos + 1 >= len(data):
        return -1
    return (data[pos] << 8) | data[pos + 1]


def _safe_chr(code: int) -> str:
    """Unicodeコードポイントを安全に文字に変換する（サロゲートペア対応）"""
    if 0xD800 <= code <= 0xDFFF:
        return ''
    try:
        return chr(code)
    except (ValueError, OverflowError):
        return ''


def _find_content_start(data: bytes) -> int:
    """DocumentTextストリームのテキストコンテンツ開始位置を検出する

    ヘッダ構造:
    - SsmgV.01 (8バイト)
    - フラグ (数バイト)
    - CTextV.01\\x00 (10バイト)
    - パディング
    - テキスト開始(偶数アライメント)
    """
    # SsmgV.01ヘッダの確認
    if len(data) < 0x20:
        return -1
    if data[:8] != HEADER_MAGIC_SSMG:
        debug_print(f"[警告] SsmgV.01ヘッダが見つかりません: {data[:8]}")
        return -1

    # CTextV.01の位置を探す
    ctext_pos = data.find(HEADER_MAGIC_CTEXT)
    if ctext_pos < 0:
        debug_print("[警告] CTextV.01ヘッダが見つかりません")
        return -1

    # CTextV.01 + null終端の直後、偶数アライメントで開始
    start = ctext_pos + len(HEADER_MAGIC_CTEXT) + 1  # +1 for null terminator
    if start % 2 != 0:
        start += 1

    debug_print(f"[JTD] テキストコンテンツ開始位置: 0x{start:04x}")
    return start


def _flush_chars(chars: list[str], lines: list[str]):
    """文字バッファを行リストにフラッシュする（空行はスキップ）"""
    text = ''.join(chars)
    if text.strip():
        lines.append(text)
    chars.clear()


def _skip_format_block(data: bytes, pos: int) -> int:
    """001Cで始まるフォーマットブロックをスキップし、
    テキスト開始マーカー(001F)の次の位置を返す

    フォーマットブロックの構造:
    - 001C 0010 ... FFFF ... 001F (段落書式)
    - 001C 0030 ... 001F (表セル書式)
    - 001C 0000/0001 ... 001F (表構造制御)
    """
    total = len(data)
    i = pos + 2  # 001C の次から

    while i < total - 1:
        code = _read_u16be(data, i)
        if code == 0x001F:
            return i + 2  # 001F の次がテキスト開始
        i += 2

    return total  # ストリーム末尾に到達


def _extract_text_from_stream(data: bytes) -> list[str]:
    """DocumentText/Footnoteストリームからテキスト行を抽出する

    一太郎のDocumentTextストリーム構造:
    1. ヘッダ(SsmgV.01 + CTextV.01)
    2. テキストセクション(UTF-16BEエンコード)
       - 001C: フォーマットブロック開始
       - 001F: テキストコンテンツ開始マーカー
       - 000A: 改行
       - 000E: セクション/行末マーカー
       - その他: テキスト文字(UTF-16BE)
    """
    start = _find_content_start(data)
    if start < 0:
        return []

    total = len(data)
    lines = []
    current_chars = []
    i = start
    in_text_zone = False  # 001Fの後〜000Eまでがテキストゾーン

    while i < total - 1:
        code = _read_u16be(data, i)

        if code == -1:
            break

        # フォーマットブロック開始
        if code == 0x001C:
            # 現在のテキストをフラッシュ
            _flush_chars(current_chars, lines)
            i = _skip_format_block(data, i)
            in_text_zone = True  # 001C→001Fの後はテキストゾーン
            continue

        # テキスト開始マーカー(フォーマットブロック外で出現した場合)
        if code == 0x001F:
            in_text_zone = True
            i += 2
            continue

        # テキストゾーン外はスキップ
        if not in_text_zone:
            i += 2
            continue

        # 改行(テキストゾーン内で有効)
        if code == 0x000A:
            _flush_chars(current_chars, lines)
            i += 2
            continue

        # セクション/行末マーカー → テキストゾーン終了
        if code == 0x000E:
            _flush_chars(current_chars, lines)
            in_text_zone = False
            i += 2
            continue

        # 制御文字のスキップ(タブ以外)
        if code < 0x0020 and code != 0x0009:
            i += 2
            continue

        # サロゲートペア処理(U+10000以上の文字)
        if 0xD800 <= code <= 0xDBFF and i + 3 < total:
            lo = _read_u16be(data, i + 2)
            if 0xDC00 <= lo <= 0xDFFF:
                full_cp = 0x10000 + ((code - 0xD800) << 10) + (lo - 0xDC00)
                ch = _safe_chr(full_cp)
                if ch:
                    current_chars.append(ch)
                i += 4
                continue
            i += 2
            continue

        # 孤立した下位サロゲートのスキップ
        if 0xDC00 <= code <= 0xDFFF:
            i += 2
            continue

        # 通常のUTF-16BE文字
        ch = _safe_chr(code)
        if ch:
            current_chars.append(ch)
        i += 2

    # 末尾の残りテキストを追加(テキストゾーン内の場合のみ)
    if in_text_zone:
        _flush_chars(current_chars, lines)

    # 末尾のバイナリゴミ行を除去
    lines = _trim_trailing_binary(lines)

    debug_print(f"[JTD] 抽出行数: {len(lines)}")
    return lines


def _is_readable_japanese_line(text: str) -> bool:
    """行が読み取り可能な日本語テキストかどうかを判定する

    判定基準:
    - ひらがな・カタカナを含む(日本語テキストの特徴)
    - または全体の80%以上が標準的な文字範囲
    """
    if not text.strip():
        return False

    has_kana = False
    standard_count = 0
    total = len(text)

    for ch in text:
        cp = ord(ch)
        # ひらがな(U+3040-U+309F)またはカタカナ(U+30A0-U+30FF)
        if 0x3040 <= cp <= 0x30FF:
            has_kana = True
        # 標準的な文字範囲
        if (0x0020 <= cp <= 0x007E or  # ASCII印字可能
            0x3000 <= cp <= 0x9FFF or  # CJK、ひらがな、カタカナ
            0xFF01 <= cp <= 0xFF9F or  # 全角英数、半角カナ
            0x2010 <= cp <= 0x2070):   # 一般句読点・記号
            standard_count += 1

    ratio = standard_count / total if total > 0 else 0

    # CJK漢字(U+4E00-U+9FFF)を含むかチェック
    has_cjk = any(0x4E00 <= ord(ch) <= 0x9FFF for ch in text)

    # ひらがな/カタカナを含む場合は緩い基準
    if has_kana:
        return ratio >= 0.5
    # 漢字を含む場合は中程度の基準
    if has_cjk:
        return ratio >= 0.7
    # 日本語文字を含まない行は非テキストとみなす
    return False


def _trim_trailing_binary(lines: list[str]) -> list[str]:
    """ストリーム末尾のバイナリデータが文字化けした行を除去する

    一太郎のDocumentTextストリーム末尾には
    表レイアウト等のバイナリデータが含まれることがある
    """
    result = list(lines)
    while result:
        if _is_readable_japanese_line(result[-1]):
            break
        result.pop()
    return result


def extract_jtd_text(file_path: str) -> str:
    """一太郎ファイルからテキストを抽出する

    Args:
        file_path: JTDファイルのパス

    Returns:
        抽出されたテキスト(Unicode文字列)

    Raises:
        FileNotFoundError: ファイルが見つからない場合
        ValueError: JTDファイルとして解析できない場合
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"ファイルが見つかりません: {file_path}")

    if not olefile.isOleFile(file_path):
        raise ValueError(
            f"OLE2形式ではありません（一太郎ver8以降のファイルが必要です）: {file_path}"
        )

    ole = olefile.OleFileIO(file_path)
    all_lines = []

    try:
        # DocumentTextストリームからメインテキストを抽出
        if ole.exists(STREAM_DOCUMENT_TEXT):
            raw = ole.openstream(STREAM_DOCUMENT_TEXT).read()
            debug_print(
                f"[JTD] {STREAM_DOCUMENT_TEXT}ストリーム: {len(raw)} バイト"
            )
            doc_lines = _extract_text_from_stream(raw)
            all_lines.extend(doc_lines)
        else:
            debug_print(
                f"[警告] {STREAM_DOCUMENT_TEXT}ストリームが見つかりません"
            )

        # Footnoteストリームから脚注テキストを抽出
        if ole.exists(STREAM_FOOTNOTE):
            fn_raw = ole.openstream(STREAM_FOOTNOTE).read()
            debug_print(
                f"[JTD] {STREAM_FOOTNOTE}ストリーム: {len(fn_raw)} バイト"
            )
            fn_lines = _extract_text_from_stream(fn_raw)
            if fn_lines:
                all_lines.append("")
                all_lines.append("---")
                all_lines.append("**脚注:**")
                all_lines.extend(fn_lines)

    finally:
        ole.close()

    if not all_lines:
        raise ValueError(
            f"テキストを抽出できませんでした: {file_path}"
        )

    return '\n'.join(all_lines)


def _get_ole_metadata(ole: olefile.OleFileIO) -> dict:
    """OLE2メタデータを取得する"""
    meta = ole.get_metadata()
    result = {}

    for attr in ('title', 'subject', 'author', 'keywords', 'comments'):
        val = getattr(meta, attr, None)
        if val:
            if isinstance(val, bytes):
                # メタデータはShift-JIS(CP932)でエンコードされていることが多い
                try:
                    val = val.decode('cp932')
                except (UnicodeDecodeError, AttributeError):
                    try:
                        val = val.decode('utf-8', errors='replace')
                    except (UnicodeDecodeError, AttributeError):
                        val = str(val)
            result[attr] = val

    for attr in ('create_time', 'last_saved_time'):
        val = getattr(meta, attr, None)
        if val:
            result[attr] = str(val)

    return result


class JtdToMarkdownConverter:
    """一太郎ファイルをMarkdownに変換するコンバータ

    OLE2 Compound Document形式の一太郎ファイル(.jtd, .jtt)から
    テキストを抽出してMarkdown形式で出力する
    """

    def __init__(
        self,
        file_path: str,
        output_dir: Optional[str] = None,
    ):
        """
        Args:
            file_path: 一太郎ファイル(.jtd/.jtt)のパス
            output_dir: 出力ディレクトリ(省略時はカレントディレクトリ)
        """
        self.file_path = file_path
        self.base_name = Path(file_path).stem

        if output_dir:
            self.output_dir = output_dir
        else:
            self.output_dir = os.path.join(os.getcwd(), "output")

        os.makedirs(self.output_dir, exist_ok=True)

    def convert(self) -> str:
        """変換メイン処理

        Returns:
            出力Markdownファイルのパス
        """
        print(f"[INFO] 一太郎文書変換開始: {self.file_path}")

        # テキスト抽出
        text = extract_jtd_text(self.file_path)
        lines = text.split('\n')

        # Markdown生成
        md_lines = self._build_markdown(lines)

        # ファイル出力
        output_path = os.path.join(
            self.output_dir, f"{self.base_name}.md"
        )
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(md_lines))

        print(f"[INFO] 変換完了: {output_path}")
        print(f"[INFO] 抽出行数: {len(lines)}")
        return output_path

    def _build_markdown(self, lines: list[str]) -> list[str]:
        """テキスト行からMarkdown形式に変換する

        見出し候補の判定:
        - ■/●/◆ で始まる行 → ## (h2)
        - （様式... や 【...】 を含む行はタイトル補助
        """
        md = []
        md.append(f"# {self.base_name}")
        md.append("")

        # メタデータの追加
        metadata = self._get_metadata()
        if metadata:
            for key, value in metadata.items():
                md.append(f"- **{key}**: {value}")
            md.append("")
            md.append("---")
            md.append("")

        for line in lines:
            stripped = line.strip()
            if not stripped:
                md.append("")
                continue

            # 脚注セパレータはそのまま
            if stripped == "---":
                md.append("---")
                continue

            # 見出し判定: ■/●/◆ で始まる行
            if stripped.startswith(('■', '●', '◆')):
                md.append(f"## {stripped}")
                md.append("")
                continue

            # 太字マーカー付き脚注ヘッダ
            if stripped.startswith("**脚注"):
                md.append(stripped)
                md.append("")
                continue

            # 通常テキスト
            md.append(stripped)

        # 末尾の空行を整理
        while md and md[-1] == "":
            md.pop()
        md.append("")

        return md

    def _get_metadata(self) -> dict:
        """ファイルのOLE2メタデータを取得"""
        try:
            ole = olefile.OleFileIO(self.file_path)
            try:
                return _get_ole_metadata(ole)
            finally:
                ole.close()
        except Exception:
            return {}


def main():
    """コマンドラインエントリポイント"""
    parser = argparse.ArgumentParser(
        description='一太郎文書 (.jtd/.jtt) をMarkdownに変換'
    )
    parser.add_argument(
        'file', help='変換する一太郎ファイル (.jtd/.jtt)'
    )
    parser.add_argument(
        '-o', '--output-dir', type=str,
        help='出力ディレクトリを指定（デフォルト: ./output）'
    )
    parser.add_argument(
        '-v', '--verbose', action='store_true',
        help='デバッグ情報を出力'
    )
    parser.add_argument(
        '--text-only', action='store_true',
        help='Markdown変換せずテキストのみ出力'
    )

    args = parser.parse_args()

    set_verbose(args.verbose)

    if args.text_only:
        text = extract_jtd_text(args.file)
        print(text)
        return

    converter = JtdToMarkdownConverter(
        args.file,
        output_dir=args.output_dir,
    )
    output_file = converter.convert()

    print(f"\n変換完了!")
    print(f"出力ファイル: {output_file}")


if __name__ == "__main__":
    main()
