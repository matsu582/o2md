"""jtd2md.py（一太郎 to Markdown Converter）のテストコード

テスト対象:
- テキスト抽出ヘルパー関数
- OLE2ストリーム解析
- Markdown変換
- o2md.pyとの統合
"""

import os
import sys
import tempfile
import shutil
import struct
from pathlib import Path
import pytest

sys.path.insert(0, str(Path(__file__).parent.parent))

from jtd2md import (
    _read_u16be,
    _safe_chr,
    _find_content_start,
    _is_readable_japanese_line,
    _trim_trailing_binary,
    _extract_text_from_stream,
    extract_jtd_text,
    JtdToMarkdownConverter,
    HEADER_MAGIC_SSMG,
    HEADER_MAGIC_CTEXT,
)
from o2md import detect_file_type


class TestReadU16BE:
    """UTF-16BE読み取り関数のテスト"""

    def test_basic_read(self):
        """基本的なUTF-16BE読み取り"""
        data = b'\x00\x41'  # 'A'
        assert _read_u16be(data, 0) == 0x0041

    def test_japanese_char(self):
        """日本語文字の読み取り"""
        data = b'\x30\x42'  # 'あ'
        assert _read_u16be(data, 0) == 0x3042

    def test_boundary(self):
        """境界値テスト: データ末尾"""
        data = b'\x00'
        assert _read_u16be(data, 0) == -1

    def test_offset(self):
        """オフセット指定の読み取り"""
        data = b'\x00\x41\x30\x42'
        assert _read_u16be(data, 2) == 0x3042


class TestSafeChr:
    """安全な文字変換のテスト"""

    def test_normal_char(self):
        """通常文字の変換"""
        assert _safe_chr(0x0041) == 'A'
        assert _safe_chr(0x3042) == 'あ'

    def test_surrogate(self):
        """サロゲート領域は空文字を返す"""
        assert _safe_chr(0xD800) == ''
        assert _safe_chr(0xDBFF) == ''
        assert _safe_chr(0xDC00) == ''
        assert _safe_chr(0xDFFF) == ''

    def test_overflow(self):
        """無効なコードポイント"""
        assert _safe_chr(0x110000) == ''


class TestFindContentStart:
    """コンテンツ開始位置検出のテスト"""

    def test_valid_header(self):
        """正常なヘッダからの開始位置検出"""
        # SsmgV.01 + padding + CTextV.01\x00 + padding → 偶数アライメント
        header = HEADER_MAGIC_SSMG + b'\x00' * 5 + HEADER_MAGIC_CTEXT + b'\x00'
        # CTextV.01は位置13から、+9(length)+1(null)=23、奇数→24
        pos = _find_content_start(header + b'\x00' * 20)
        assert pos >= 0
        assert pos % 2 == 0

    def test_missing_ssmg(self):
        """SsmgV.01ヘッダがない場合"""
        data = b'InvalidHeader' + b'\x00' * 50
        assert _find_content_start(data) == -1

    def test_missing_ctext(self):
        """CTextV.01がない場合"""
        data = HEADER_MAGIC_SSMG + b'\x00' * 50
        assert _find_content_start(data) == -1

    def test_short_data(self):
        """データが短すぎる場合"""
        assert _find_content_start(b'\x00' * 10) == -1


class TestIsReadableJapaneseLine:
    """日本語行判定のテスト"""

    def test_hiragana_line(self):
        """ひらがなを含む行"""
        assert _is_readable_japanese_line("これはテストです") is True

    def test_katakana_line(self):
        """カタカナを含む行"""
        assert _is_readable_japanese_line("テレビ、ラジオ") is True

    def test_kanji_only(self):
        """漢字のみの行"""
        assert _is_readable_japanese_line("災害情報") is True

    def test_ascii_garbage(self):
        """ASCII文字のみのゴミ行"""
        assert _is_readable_japanese_line('!"#$%&()*+,-./0123456789') is False

    def test_empty_line(self):
        """空行"""
        assert _is_readable_japanese_line("") is False
        assert _is_readable_japanese_line("   ") is False

    def test_binary_garbage(self):
        """バイナリゴミ文字列"""
        garbage = ''.join(chr(c) for c in [0x0202, 0x02C1, 0x1404, 0x19FE])
        assert _is_readable_japanese_line(garbage) is False


class TestTrimTrailingBinary:
    """末尾バイナリ除去のテスト"""

    def test_clean_lines(self):
        """クリーンな行はそのまま"""
        lines = ["これはテストです", "避難情報"]
        assert _trim_trailing_binary(lines) == lines

    def test_trailing_garbage(self):
        """末尾のゴミ行を除去"""
        lines = ["これはテストです", '!"#$%&()*+']
        result = _trim_trailing_binary(lines)
        assert len(result) == 1
        assert result[0] == "これはテストです"

    def test_all_garbage(self):
        """すべてゴミ行の場合"""
        lines = ['!"#$%&()*+', "ABCDEFG"]
        result = _trim_trailing_binary(lines)
        assert len(result) == 0

    def test_empty_list(self):
        """空リスト"""
        assert _trim_trailing_binary([]) == []


class TestExtractTextFromStream:
    """ストリームテキスト抽出のテスト"""

    def _build_stream(self, text_segments):
        """テスト用ストリームデータを構築する

        テキストセグメントを001C/001Fマーカーで囲んで返す
        """
        # ヘッダ
        header = HEADER_MAGIC_SSMG + b'\x00\x00\x00\x01\x00\x00\x01\x00\x00\x00\x00'
        header += HEADER_MAGIC_CTEXT + b'\x00'
        # 偶数アライメントにパディング
        if len(header) % 2 != 0:
            header += b'\x00'

        body = b''
        for seg in text_segments:
            # 001C + 0010(段落書式) + FFFF + 001F
            body += b'\x00\x1C\x00\x10\xFF\xFF\x00\x1F'
            # テキストをUTF-16BEでエンコード
            for ch in seg:
                body += ch.encode('utf-16-be')
            # 000E(セクション終了)
            body += b'\x00\x0E'

        return header + body

    def test_single_line(self):
        """1行のテキスト抽出"""
        stream = self._build_stream(["テスト"])
        lines = _extract_text_from_stream(stream)
        assert len(lines) == 1
        assert lines[0] == "テスト"

    def test_multiple_lines(self):
        """複数行のテキスト抽出"""
        stream = self._build_stream(["行1", "行2", "行3"])
        lines = _extract_text_from_stream(stream)
        assert len(lines) == 3
        assert lines[0] == "行1"
        assert lines[1] == "行2"
        assert lines[2] == "行3"

    def test_empty_stream(self):
        """空ストリーム"""
        lines = _extract_text_from_stream(b'')
        assert lines == []


class TestExtractJtdText:
    """JTDファイルテキスト抽出のテスト"""

    def test_file_not_found(self):
        """存在しないファイル"""
        with pytest.raises(FileNotFoundError):
            extract_jtd_text("/nonexistent/file.jtd")

    def test_not_ole2_file(self):
        """OLE2形式でないファイル"""
        with tempfile.NamedTemporaryFile(suffix='.jtd', delete=False) as f:
            f.write(b'This is not an OLE2 file')
            temp_path = f.name
        try:
            with pytest.raises(ValueError, match="OLE2形式ではありません"):
                extract_jtd_text(temp_path)
        finally:
            os.unlink(temp_path)


class TestDetectFileTypeIchitaro:
    """o2md.pyでの一太郎ファイルタイプ検出テスト"""

    def test_detect_jtd(self):
        """JTDファイルを一太郎として検出"""
        assert detect_file_type("test.jtd") == 'ichitaro'
        assert detect_file_type("TEST.JTD") == 'ichitaro'
        assert detect_file_type("/path/to/file.jtd") == 'ichitaro'

    def test_detect_jtt(self):
        """JTTファイルを一太郎として検出"""
        assert detect_file_type("test.jtt") == 'ichitaro'
        assert detect_file_type("TEST.JTT") == 'ichitaro'

    def test_existing_types_unchanged(self):
        """既存のファイルタイプ検出に影響がないことを確認"""
        assert detect_file_type("test.xlsx") == 'excel'
        assert detect_file_type("test.docx") == 'word'
        assert detect_file_type("test.pptx") == 'powerpoint'
        assert detect_file_type("test.pdf") == 'pdf'
        assert detect_file_type("test.txt") == 'unknown'


class TestJtdToMarkdownConverter:
    """JtdToMarkdownConverterクラスのテスト"""

    @pytest.fixture
    def temp_output_dir(self):
        """一時出力ディレクトリ"""
        temp_dir = tempfile.mkdtemp()
        yield temp_dir
        shutil.rmtree(temp_dir, ignore_errors=True)

    def test_output_dir_creation(self, temp_output_dir):
        """出力ディレクトリが自動作成されること"""
        output = os.path.join(temp_output_dir, "subdir")
        converter = JtdToMarkdownConverter(
            "dummy.jtd", output_dir=output
        )
        assert os.path.isdir(output)

    def test_base_name(self):
        """ファイル名からベース名を取得"""
        converter = JtdToMarkdownConverter("/path/to/document.jtd")
        assert converter.base_name == "document"
