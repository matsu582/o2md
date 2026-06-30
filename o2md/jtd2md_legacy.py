#!/usr/bin/env python3
"""
旧一太郎ファイル (ver4-7) テキスト抽出モジュール

対応形式:
- .jsw (一太郎 ver4)
- .jaw/.jtw (一太郎 ver5)
- .jbw/.juw (一太郎 ver6)
- .jfw/.jvw (一太郎 ver7)

ver4-6: 独自バイナリ形式 (DOC\\x00シグネチャ、Shift-JISテキスト)
ver7: OLE2 Compound Document形式 (既存パーサへ委譲)
"""

import logging
import os

logger = logging.getLogger(__name__)

# 旧一太郎ファイルの対応拡張子
LEGACY_JTD_EXTENSIONS = (
    '.jsw',   # ver4
    '.jaw', '.jtw',   # ver5
    '.jbw', '.juw',   # ver6
    '.jfw', '.jvw',   # ver7
)

# ver7はOLE2形式のため既存パーサへ委譲する
VER7_EXTENSIONS = ('.jfw', '.jvw')

# ver4-6の独自バイナリ形式の拡張子
VER456_EXTENSIONS = ('.jsw', '.jaw', '.jtw', '.jbw', '.juw')

# ファイルシグネチャ
MAGIC_DOC = b"DOC\x00"

# ヘッダ内の検証値 (オフセット0x3C-0x3Fに格納)
HEADER_VALIDATION_OFFSET = 0x3C
HEADER_VALIDATION_VALUE = 0x22028919

# テキストデータの位置 (フォールバック)
TEXT_SIZE_OFFSET = 0x800
TEXT_START_OFFSET = 0x804

# ブロックテーブル関連のヘッダオフセット
BLOCK_TABLE_INFO_OFFSET = 0xE2
BLOCK_ENTRY_SIZE = 0x40

# 制御コード定義
CTRL_NEWLINE_PREFIX = 0xFE
CTRL_SKIP_SINGLE = 0x1E
CTRL_SKIP_VARIABLE = 0x1F
CTRL_SKIP_FORMAT = 0x1C
CTRL_KEISEN_PREFIX = 0xFD

# 改行を示す0xFE後続バイトの範囲
_NEWLINE_RANGES = (
    (0x41, 0x43),  # A-C
    (0x45, 0x47),  # E-G
    (0x49, 0x4B),  # I-K
)

# 罫線文字変換テーブル (0xFD後のバイト → Shift-JISバイトペア)
# 旧一太郎バイナリ解析結果に基づく対応表
_KEISEN_TABLE: dict[int, bytes] = {
    0x21: b'\x84\x9f',  # ─ (横線)
    0x22: b'\x84\xa0',  # │ (縦線)
    0x23: b'\x84\xa1',  # ┌ (左上角)
    0x24: b'\x84\xa2',  # ┐ (右上角)
    0x25: b'\x84\xa3',  # ┘ (右下角)
    0x26: b'\x84\xa4',  # └ (左下角)
    0x27: b'\x84\xa5',  # ├ (左T字)
    0x28: b'\x84\xa6',  # ┬ (上T字)
    0x29: b'\x84\xa7',  # ┤ (右T字)
}
# デフォルト: 中黒 (・)
_KEISEN_DEFAULT = b'\x81\x45'


def is_legacy_jtd_file(file_path: str) -> bool:
    """旧一太郎形式のファイルかどうかを判定する

    Args:
        file_path: ファイルパス

    Returns:
        旧一太郎形式ならTrue
    """
    ext = os.path.splitext(file_path)[1].lower()
    return ext in LEGACY_JTD_EXTENSIONS


def is_ver7_file(file_path: str) -> bool:
    """ver7形式(OLE2)のファイルかどうかを判定する

    Args:
        file_path: ファイルパス

    Returns:
        ver7形式ならTrue
    """
    ext = os.path.splitext(file_path)[1].lower()
    return ext in VER7_EXTENSIONS


def _is_newline_code(byte_val: int) -> bool:
    """0xFEに続くバイトが改行を示すかどうか判定"""
    for low, high in _NEWLINE_RANGES:
        if low <= byte_val <= high:
            return True
    return False


def _is_sjis_lead_byte(b: int) -> bool:
    """Shift-JISのマルチバイト文字の第1バイトかどうかを判定"""
    return (0x81 <= b <= 0x9F) or (0xE0 <= b <= 0xFC)


def _is_printable_byte(b: int) -> bool:
    """表示可能なASCII/半角カナ文字かどうかを判定"""
    return (0x20 <= b <= 0x7E) or (0xA1 <= b <= 0xDF)


def _read_u32_be(data: bytes, offset: int) -> int:
    """ビッグエンディアンで4バイト整数を読み取る"""
    if offset + 4 > len(data):
        return 0
    return (data[offset + 3] << 24 | data[offset + 2] << 16 |
            data[offset + 1] << 8 | data[offset])


def _read_u16_le(data: bytes, offset: int) -> int:
    """リトルエンディアンで2バイト整数を読み取る"""
    if offset + 2 > len(data):
        return 0
    return data[offset] | (data[offset + 1] << 8)


def _validate_header(data: bytes) -> bool:
    """旧一太郎ファイルのヘッダを検証する

    検証内容:
    - 先頭4バイトがDOC\\x00であること
    - オフセット0x3C-0x3Fの検証値が一致すること
    """
    if len(data) < TEXT_START_OFFSET:
        return False

    # シグネチャ確認
    if data[:4] != MAGIC_DOC:
        return False

    # 検証値確認
    val = _read_u32_be(data, HEADER_VALIDATION_OFFSET)
    if val != HEADER_VALIDATION_VALUE:
        logger.debug(
            f"[旧JTD] ヘッダ検証値不一致: "
            f"0x{val:08X} != 0x{HEADER_VALIDATION_VALUE:08X}"
        )
        return False

    return True


def _find_text_region(data: bytes) -> tuple[int, int]:
    """テキスト領域の開始オフセットとサイズを特定する

    ブロックテーブルを検索し、テキストブロックの位置を計算する。
    ブロックが見つからない場合はフォールバック位置(0x804)を使用する。

    Returns:
        (text_start_offset, text_size) のタプル
    """
    file_size = len(data)

    # ブロックテーブル情報の読み取り
    block_start_val = _read_u16_le(data, BLOCK_TABLE_INFO_OFFSET)
    block_count = _read_u16_le(data, BLOCK_TABLE_INFO_OFFSET + 2)
    base_offset_val = _read_u16_le(data, BLOCK_TABLE_INFO_OFFSET + 4)
    block_stride = _read_u16_le(data, BLOCK_TABLE_INFO_OFFSET + 6)

    # ブロックテーブルのスキャン
    text_block_offset = -1
    if block_count > 0 and block_start_val > 0:
        for i in range(block_count):
            entry_pos = block_start_val + i * BLOCK_ENTRY_SIZE
            if entry_pos + BLOCK_ENTRY_SIZE > file_size:
                break
            # マーカー検索: 最初の2バイトが0x0001
            marker = _read_u16_le(data, entry_pos)
            if marker == 0x0001:
                # 2番目のマーカー確認
                marker2_offset = entry_pos + 2
                if marker2_offset + 2 <= file_size:
                    marker2 = _read_u16_le(data, marker2_offset)
                    if marker2 == 0x0002:
                        text_block_offset = entry_pos
                        break

    if text_block_offset >= 0:
        # ブロックからページ数を読み取り
        page_count = _read_u32_be(data, text_block_offset + 8)
        if page_count > 0 and block_stride > 0:
            # テキストサイズの計算
            calc_offset = block_stride * (page_count - 1) + base_offset_val
            if calc_offset + 4 <= file_size:
                text_size = _read_u32_be(data, calc_offset)
                if 0 < text_size < file_size:
                    logger.debug(
                        f"[旧JTD] ブロックテーブルから特定: "
                        f"offset=0x{TEXT_START_OFFSET:04X}, size={text_size}"
                    )
                    return TEXT_START_OFFSET, text_size

    # フォールバック: オフセット0x800から4バイトでサイズ取得
    text_size = _read_u32_be(data, TEXT_SIZE_OFFSET)
    if text_size <= 0 or text_size > file_size:
        # サイズが不正な場合、ファイル末尾までを対象とする
        text_size = file_size - TEXT_START_OFFSET

    logger.debug(
        f"[旧JTD] フォールバック位置使用: "
        f"offset=0x{TEXT_START_OFFSET:04X}, size={text_size}"
    )
    return TEXT_START_OFFSET, text_size


def _extract_sjis_text(data: bytes, start: int, size: int) -> bytes:
    """バイナリデータからShift-JISテキストを抽出する

    制御コードを処理し、テキスト部分のみをShift-JISバイト列として返す。

    Args:
        data: ファイル全体のバイナリデータ
        start: テキスト開始オフセット
        size: テキスト領域のサイズ

    Returns:
        抽出されたShift-JISバイト列
    """
    output = bytearray()
    pos = 0
    end = min(size, len(data) - start)

    while pos < end:
        byte_val = data[start + pos]

        if byte_val == CTRL_NEWLINE_PREFIX:
            # 0xFE: 改行プレフィクス
            if pos + 1 < end:
                next_byte = data[start + pos + 1]
                if _is_newline_code(next_byte):
                    output.extend(b'\r\n')
                # 改行でもそうでなくても2バイト消費
                pos += 2
            else:
                pos += 1

        elif byte_val == CTRL_SKIP_SINGLE:
            # 0x1E: 1バイトスキップ
            pos += 1

        elif byte_val == CTRL_SKIP_VARIABLE:
            # 0x1F: 可変長スキップ
            if pos + 2 < end:
                skip_len = data[start + pos + 2]
                if skip_len > 0:
                    pos += skip_len
                else:
                    pos += 1
            else:
                pos += 1

        elif byte_val == CTRL_SKIP_FORMAT:
            # 0x1C: 書式制御 (3バイトスキップ)
            pos += 3

        elif byte_val in (0x11, 0x12):
            # 0x11-0x12: 書式制御 (3バイトスキップ)
            pos += 3

        elif byte_val == CTRL_KEISEN_PREFIX:
            # 0xFD: 罫線文字
            if pos + 1 < end:
                keisen_code = data[start + pos + 1]
                sjis_pair = _KEISEN_TABLE.get(keisen_code, _KEISEN_DEFAULT)
                output.extend(sjis_pair)
                pos += 2
            else:
                pos += 1

        elif _is_sjis_lead_byte(byte_val):
            # Shift-JISマルチバイト文字の第1バイト
            if pos + 1 < end:
                output.append(byte_val)
                output.append(data[start + pos + 1])
                pos += 2
            else:
                pos += 1

        elif _is_printable_byte(byte_val):
            # 表示可能な1バイト文字
            output.append(byte_val)
            pos += 1

        elif byte_val in (0x0D, 0x0A):
            # CR/LF: そのまま出力
            output.append(byte_val)
            pos += 1

        else:
            # その他の制御コード: スキップ
            pos += 1

    return bytes(output)


def extract_legacy_jtd_text(file_path: str) -> str:
    """旧一太郎ファイル(ver4-6)からテキストを抽出する

    Args:
        file_path: 旧一太郎ファイルのパス

    Returns:
        抽出されたテキスト(Unicode文字列)

    Raises:
        FileNotFoundError: ファイルが見つからない場合
        ValueError: ファイル形式が不正な場合
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"ファイルが見つかりません: {file_path}")

    with open(file_path, 'rb') as f:
        data = f.read()

    if not _validate_header(data):
        raise ValueError(
            f"旧一太郎形式(ver4-6)として認識できません: {file_path}"
        )

    # テキスト領域の特定
    text_start, text_size = _find_text_region(data)

    # Shift-JISテキストの抽出
    sjis_bytes = _extract_sjis_text(data, text_start, text_size)

    if not sjis_bytes:
        raise ValueError(
            f"テキストを抽出できませんでした: {file_path}"
        )

    # Shift-JIS → Unicode変換
    try:
        text = sjis_bytes.decode('cp932', errors='replace')
    except Exception as e:
        raise ValueError(
            f"テキストのデコードに失敗しました: {file_path}: {e}"
        )

    # CR+LFの正規化
    text = text.replace('\r\n', '\n').replace('\r', '\n')

    return text


def extract_legacy_jtd_lines(file_path: str) -> list[str]:
    """旧一太郎ファイルからテキスト行リストを抽出する

    Args:
        file_path: 旧一太郎ファイルのパス

    Returns:
        テキスト行のリスト
    """
    text = extract_legacy_jtd_text(file_path)
    lines = text.split('\n')
    # 末尾の空行を除去
    while lines and not lines[-1].strip():
        lines.pop()
    return lines
