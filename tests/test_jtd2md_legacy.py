"""jtd2md_legacy.py（旧一太郎テキスト抽出）のテストコード

テスト対象:
- ヘッダ検証
- 制御コード処理
- Shift-JISテキスト抽出
- 罫線文字変換
- LegacyJtdToMarkdownConverter統合
"""

import os
import sys
import tempfile
from pathlib import Path
import pytest

sys.path.insert(0, str(Path(__file__).parent.parent))

from o2md.jtd2md_legacy import (
    _validate_header,
    _is_sjis_lead_byte,
    _is_printable_byte,
    _is_newline_code,
    _extract_sjis_text,
    _find_text_region,
    extract_legacy_jtd_text,
    extract_legacy_jtd_lines,
    is_legacy_jtd_file,
    is_ver7_file,
    LEGACY_JTD_EXTENSIONS,
    VER7_EXTENSIONS,
    TEXT_START_OFFSET,
)


def _make_legacy_jtd(text_data: bytes) -> bytes:
    """テスト用の旧一太郎バイナリを生成するヘルパー"""
    buf = bytearray(TEXT_START_OFFSET + len(text_data))
    # シグネチャ
    buf[0:4] = b'DOC\x00'
    # 検証値
    buf[0x3C] = 0x19
    buf[0x3D] = 0x89
    buf[0x3E] = 0x02
    buf[0x3F] = 0x22
    # テキストサイズ (リトルエンディアンで格納)
    size = len(text_data)
    buf[0x800] = size & 0xFF
    buf[0x801] = (size >> 8) & 0xFF
    buf[0x802] = (size >> 16) & 0xFF
    buf[0x803] = (size >> 24) & 0xFF
    # テキストデータ
    buf[TEXT_START_OFFSET:TEXT_START_OFFSET + len(text_data)] = text_data
    return bytes(buf)


class TestHeaderValidation:
    """ヘッダ検証のテスト"""

    def test_valid_header(self):
        """正常なヘッダの検証"""
        data = _make_legacy_jtd(b'test')
        assert _validate_header(data) is True

    def test_invalid_magic(self):
        """不正なシグネチャ"""
        data = bytearray(_make_legacy_jtd(b'test'))
        data[0:4] = b'XXX\x00'
        assert _validate_header(bytes(data)) is False

    def test_invalid_validation_value(self):
        """不正な検証値"""
        data = bytearray(_make_legacy_jtd(b'test'))
        data[0x3C] = 0x00
        assert _validate_header(bytes(data)) is False

    def test_too_short(self):
        """データが短すぎる場合"""
        assert _validate_header(b'DOC\x00' + b'\x00' * 10) is False


class TestByteClassification:
    """バイト分類関数のテスト"""

    def test_sjis_lead_byte(self):
        """Shift-JIS第1バイト判定"""
        assert _is_sjis_lead_byte(0x81) is True
        assert _is_sjis_lead_byte(0x9F) is True
        assert _is_sjis_lead_byte(0xE0) is True
        assert _is_sjis_lead_byte(0xFC) is True
        assert _is_sjis_lead_byte(0x80) is False
        assert _is_sjis_lead_byte(0xA0) is False
        assert _is_sjis_lead_byte(0x41) is False

    def test_printable_byte(self):
        """表示可能バイト判定"""
        assert _is_printable_byte(0x20) is True  # スペース
        assert _is_printable_byte(0x41) is True  # 'A'
        assert _is_printable_byte(0x7E) is True  # '~'
        assert _is_printable_byte(0xA1) is True  # 半角カナ開始
        assert _is_printable_byte(0xDF) is True  # 半角カナ終了
        assert _is_printable_byte(0x00) is False
        assert _is_printable_byte(0x1F) is False
        assert _is_printable_byte(0x7F) is False

    def test_newline_code(self):
        """改行コード判定"""
        # A-C (0x41-0x43)
        assert _is_newline_code(0x41) is True
        assert _is_newline_code(0x43) is True
        # E-G (0x45-0x47)
        assert _is_newline_code(0x45) is True
        assert _is_newline_code(0x47) is True
        # I-K (0x49-0x4B)
        assert _is_newline_code(0x49) is True
        assert _is_newline_code(0x4B) is True
        # 範囲外
        assert _is_newline_code(0x44) is False
        assert _is_newline_code(0x48) is False
        assert _is_newline_code(0x50) is False


class TestTextExtraction:
    """テキスト抽出のテスト"""

    def test_plain_sjis_text(self):
        """通常のShift-JISテキスト抽出"""
        text = 'テスト'.encode('cp932')
        result = _extract_sjis_text(text, 0, len(text))
        assert result.decode('cp932') == 'テスト'

    def test_newline_handling(self):
        """改行制御コード (0xFE + newline_code)"""
        text = 'A'.encode('cp932') + bytes([0xFE, 0x41]) + 'B'.encode('cp932')
        result = _extract_sjis_text(text, 0, len(text))
        decoded = result.decode('cp932')
        assert 'A' in decoded
        assert 'B' in decoded
        assert '\r\n' in decoded

    def test_skip_single(self):
        """1バイトスキップ (0x1E)"""
        text = 'A'.encode('cp932') + bytes([0x1E]) + 'B'.encode('cp932')
        result = _extract_sjis_text(text, 0, len(text))
        assert result.decode('cp932') == 'AB'

    def test_skip_format(self):
        """書式制御スキップ (0x1C)"""
        text = (
            'A'.encode('cp932')
            + bytes([0x1C, 0x99, 0x99])
            + 'B'.encode('cp932')
        )
        result = _extract_sjis_text(text, 0, len(text))
        assert result.decode('cp932') == 'AB'

    def test_skip_variable(self):
        """可変長スキップ (0x1F)"""
        # 0x1F, type_byte, length=4 → 4バイトスキップ
        text = (
            'A'.encode('cp932')
            + bytes([0x1F, 0x00, 0x04, 0xFF])
            + 'B'.encode('cp932')
        )
        result = _extract_sjis_text(text, 0, len(text))
        assert result.decode('cp932') == 'AB'

    def test_keisen_character(self):
        """罫線文字変換 (0xFD)"""
        text = bytes([0xFD, 0x21])  # ─ (横線)
        result = _extract_sjis_text(text, 0, len(text))
        decoded = result.decode('cp932')
        assert decoded == '─'

    def test_keisen_default(self):
        """罫線文字デフォルト (未知のコード)"""
        text = bytes([0xFD, 0xFF])  # 不明なコード → 中黒
        result = _extract_sjis_text(text, 0, len(text))
        decoded = result.decode('cp932')
        assert decoded == '・'

    def test_cr_lf_passthrough(self):
        """CR/LFのパススルー"""
        text = b'A\r\nB'
        result = _extract_sjis_text(text, 0, len(text))
        assert result == b'A\r\nB'

    def test_unprintable_skip(self):
        """非表示バイトのスキップ"""
        text = b'A' + bytes([0x03, 0x04, 0x05]) + b'B'
        result = _extract_sjis_text(text, 0, len(text))
        assert result == b'AB'


class TestFileExtraction:
    """ファイルレベルの抽出テスト"""

    def test_extract_from_file(self, tmp_path):
        """ファイルからのテキスト抽出"""
        text_data = 'テスト文書'.encode('cp932')
        file_data = _make_legacy_jtd(text_data)
        file_path = tmp_path / "test.jsw"
        file_path.write_bytes(file_data)

        result = extract_legacy_jtd_text(str(file_path))
        assert result == 'テスト文書'

    def test_extract_lines(self, tmp_path):
        """行リスト抽出"""
        text_data = (
            '一行目'.encode('cp932')
            + bytes([0xFE, 0x41])
            + '二行目'.encode('cp932')
        )
        file_data = _make_legacy_jtd(text_data)
        file_path = tmp_path / "test.jaw"
        file_path.write_bytes(file_data)

        lines = extract_legacy_jtd_lines(str(file_path))
        assert lines == ['一行目', '二行目']

    def test_file_not_found(self):
        """存在しないファイル"""
        with pytest.raises(FileNotFoundError):
            extract_legacy_jtd_text('/nonexistent/file.jsw')

    def test_invalid_format(self, tmp_path):
        """不正なファイル形式"""
        file_path = tmp_path / "invalid.jsw"
        file_path.write_bytes(b'INVALID' + b'\x00' * 2048)
        with pytest.raises(ValueError):
            extract_legacy_jtd_text(str(file_path))


class TestExtensionDetection:
    """拡張子判定のテスト"""

    def test_legacy_extensions(self):
        """旧形式拡張子の判定"""
        for ext in LEGACY_JTD_EXTENSIONS:
            assert is_legacy_jtd_file(f"test{ext}") is True
        assert is_legacy_jtd_file("test.jtd") is False
        assert is_legacy_jtd_file("test.doc") is False

    def test_ver7_detection(self):
        """ver7形式の判定"""
        assert is_ver7_file("test.jfw") is True
        assert is_ver7_file("test.jvw") is True
        assert is_ver7_file("test.jsw") is False
        assert is_ver7_file("test.jtd") is False

    def test_case_insensitive(self):
        """大文字小文字を区別しない"""
        assert is_legacy_jtd_file("TEST.JSW") is True
        assert is_legacy_jtd_file("Test.Jaw") is True
