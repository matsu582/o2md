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

from o2md.jtd2md import (
    _read_u16be,
    _safe_chr,
    _find_content_start,
    _is_readable_japanese_line,
    _trim_trailing_binary,
    _extract_text_from_stream,
    extract_jtd_text,
    extract_jtd_structured,
    JtdToMarkdownConverter,
    HEADER_MAGIC_SSMG,
    HEADER_MAGIC_CTEXT,
    HEADER_MAGIC_TEXT,
)
from o2md.jtd2md_table import (
    TableCell,
    parse_cell_block,
    scan_stream_events,
    extract_tables_from_events,
    table_to_markdown,
    _merge_continuation_rows,
    count_rulers_in_para,
    extract_font_size_from_para,
    _is_table_section,
    _StreamEvent,
)
from o2md.o2md import detect_file_type


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


class TestFindContentStartTextV01:
    """TextV.01ヘッダでのコンテンツ開始位置検出テスト"""

    def test_textv01_header(self):
        """TextV.01ヘッダからの開始位置検出"""
        header = HEADER_MAGIC_SSMG + b'\x00' * 8 + HEADER_MAGIC_TEXT + b'\x00'
        pos = _find_content_start(header + b'\x00' * 20)
        assert pos >= 0
        assert pos % 2 == 0

    def test_ctextv01_preferred_over_textv01(self):
        """CTextV.01がある場合はそちらが優先される"""
        header = HEADER_MAGIC_SSMG + b'\x00' * 3 + HEADER_MAGIC_CTEXT + b'\x00'
        pos = _find_content_start(header + b'\x00' * 20)
        assert pos >= 0


class TestTableCell:
    """TableCellクラスのテスト"""

    def test_create_cell(self):
        """セル作成"""
        cell = TableCell(0x0002, 0x0048, 0x0001)
        assert cell.col_start == 0x0002
        assert cell.col_end == 0x0048
        assert cell.row_flag == 0x0001
        assert cell.text == ""

    def test_append_text(self):
        """テキスト追記"""
        cell = TableCell(0, 0, 0)
        cell.text = "テスト"
        cell.append_text("追加")
        assert cell.text == "テスト 追加"


class TestMergeContinuationRows:
    """継続行統合のテスト"""

    def test_no_merge_needed(self):
        """統合不要な行"""
        c1 = TableCell(0, 10, 0)
        c1.text = "A"
        c2 = TableCell(10, 20, 0)
        c2.text = "B"
        c3 = TableCell(0, 10, 0)
        c3.text = "C"
        c4 = TableCell(10, 20, 0)
        c4.text = "D"
        rows = [[c1, c2], [c3, c4]]
        result = _merge_continuation_rows(rows, 2)
        assert len(result) == 2

    def test_merge_continuation(self):
        """継続行を前の行に統合"""
        c1 = TableCell(0, 10, 0)
        c1.text = "行1"
        c2 = TableCell(10, 20, 0)
        c2.text = ""
        c3 = TableCell(0, 10, 0)
        c3.text = ""
        c4 = TableCell(10, 20, 0)
        c4.text = "継続"
        rows = [[c1, c2], [c3, c4]]
        result = _merge_continuation_rows(rows, 2)
        assert len(result) == 1
        assert result[0][1].text.strip() == "継続"

    def test_empty_rows(self):
        """空の行リスト"""
        assert _merge_continuation_rows([], 2) == []


class TestTableToMarkdown:
    """テーブルMarkdown変換のテスト"""

    def test_simple_table(self):
        """シンプルなテーブル"""
        c1 = TableCell(0, 10, 0)
        c1.text = "ヘッダ1"
        c2 = TableCell(10, 20, 0)
        c2.text = "ヘッダ2"
        c3 = TableCell(0, 10, 0)
        c3.text = "データ1"
        c4 = TableCell(10, 20, 0)
        c4.text = "データ2"
        table = {
            'rows': [[c1, c2], [c3, c4]],
            'num_cols': 2,
            'col_map': [(0, 10), (10, 20)],
        }
        md = table_to_markdown(table)
        assert len(md) == 3  # ヘッダ + セパレータ + データ行
        assert "ヘッダ1" in md[0]
        assert "---" in md[1]
        assert "データ1" in md[2]

    def test_pipe_escaping(self):
        """パイプ文字のエスケープ"""
        c1 = TableCell(0, 10, 0)
        c1.text = "A|B"
        table = {
            'rows': [[c1]],
            'num_cols': 1,
            'col_map': [(0, 10)],
        }
        md = table_to_markdown(table)
        assert "A\\|B" in md[0]


class TestCountRulersInPara:
    """PARAブロック内の罫線数カウントのテスト"""

    def test_no_008f_tag(self):
        """008Fタグなし → 罫線0"""
        # PARAブロック: 001C 0010 0019 0000 ... FFFF ... 0010 001F
        block = bytes([
            0x00, 0x1c, 0x00, 0x10,
            0x00, 0x19, 0x00, 0x00,
            0x00, 0x01, 0x00, 0x20,
            0xff, 0xff,
            0x00, 0x00, 0x00, 0x10, 0x00, 0x1f,
        ])
        assert count_rulers_in_para(block, 0, len(block)) == 0

    def test_008f_with_rulers(self):
        """008Fタグ内に罫線識別子(001B)が含まれる → 罫線カウント"""
        # 008F 0007 [data with 001B entries]
        block = bytes([
            0x00, 0x8f, 0x00, 0x07,
            0x01, 0x24, 0x00, 0x00,
            0x00, 0x00,
            0x00, 0x1b, 0x00, 0x00,
            0x00, 0x08, 0x00, 0x46,
        ])
        assert count_rulers_in_para(block, 0, len(block)) == 1

    def test_008f_with_multiple_rulers(self):
        """罫線識別子0013も含む場合"""
        block = bytes([
            0x00, 0x8f, 0x00, 0x09,
            0x01, 0x24, 0x00, 0x00,
            0x00, 0x00,
            0x00, 0x13, 0x00, 0x00,
            0x00, 0x13, 0x00, 0x00,
            0x00, 0x13, 0x00, 0x00,
        ])
        assert count_rulers_in_para(block, 0, len(block)) == 3

    def test_008f_zero_rulers(self):
        """008Fタグあり、罫線なし → 罫線0（枠線のみ）"""
        block = bytes([
            0x00, 0x8f, 0x00, 0x03,
            0x01, 0x24, 0x00, 0x00,
            0x00, 0x00,
        ])
        assert count_rulers_in_para(block, 0, len(block)) == 0


class TestIsTableSection:
    """セクションのテーブル判定テスト"""

    def test_section_with_rulers(self):
        """罫線≥1のPARAがあるセクション → テーブル"""
        section = [
            _StreamEvent('PARA', 0, ruler_count=3),
            _StreamEvent('CELL', 10, cell=TableCell(0, 10, 0)),
            _StreamEvent('SECTION_END', 20),
        ]
        assert _is_table_section(section) is True

    def test_section_without_rulers(self):
        """罫線なしのセクション → 非テーブル"""
        section = [
            _StreamEvent('PARA', 0, text='通常テキスト'),
            _StreamEvent('SECTION_END', 10),
        ]
        assert _is_table_section(section) is False

    def test_section_with_cell_but_no_rulers(self):
        """セルはあるが罫線なし → 非テーブル（枠線のみ）"""
        section = [
            _StreamEvent('PARA', 0, ruler_count=0),
            _StreamEvent('CELL', 10, cell=TableCell(0, 100, 0)),
            _StreamEvent('SECTION_END', 20),
        ]
        assert _is_table_section(section) is False


class TestExtractFontSizeFromPara:
    """PARAブロックからのフォントサイズ抽出テスト"""

    def test_font_size_tag_present(self):
        """TAG 0008がある場合 → フォントサイズを返す"""
        # 001C 0010 ... 0008 [size=850=0x0352] ...
        block = bytes([
            0x00, 0x1c, 0x00, 0x10,
            0x00, 0x19, 0x00, 0x00,
            0x00, 0x08, 0x03, 0x52,
            0x00, 0x1f,
        ])
        assert extract_font_size_from_para(block, 0, len(block)) == 850

    def test_font_size_tag_absent(self):
        """TAG 0008がない場合 → 0を返す"""
        block = bytes([
            0x00, 0x1c, 0x00, 0x10,
            0x00, 0x19, 0x00, 0x00,
            0x00, 0x09, 0x00, 0x01,
            0x00, 0x1f,
        ])
        assert extract_font_size_from_para(block, 0, len(block)) == 0


class TestDetectHeading:
    """見出し推定テスト（フォントサイズベース）"""

    def test_larger_font_h2(self):
        """フォントサイズが本文より大きい → h2"""
        result = JtdToMarkdownConverter._detect_heading(
            "セクションタイトル", 1200, 700)
        assert result == (2, "セクションタイトル")

    def test_default_font_not_heading(self):
        """フォントサイズ=0（デフォルト）は見出しではなく太字として扱う"""
        result = JtdToMarkdownConverter._detect_heading(
            "1.計画の目的", 0, 700)
        assert result is None

    def test_same_font_as_body(self):
        """本文と同じフォントサイズ → None"""
        result = JtdToMarkdownConverter._detect_heading(
            "（1） 情報収集及び情報伝達を担う担当者", 850, 850)
        assert result is None

    def test_smaller_font_than_body(self):
        """本文より小さいフォントサイズ → None"""
        result = JtdToMarkdownConverter._detect_heading(
            "注記テキスト", 500, 700)
        assert result is None

    def test_normal_text_no_heading(self):
        """通常テキスト（同サイズ） → None"""
        result = JtdToMarkdownConverter._detect_heading(
            "この計画は、土砂災害に関する法律です。", 700, 700)
        assert result is None

    def test_no_body_font_no_heading(self):
        """本文フォント不明 → None"""
        result = JtdToMarkdownConverter._detect_heading(
            "通常のテキスト行です", 0, 0)
        assert result is None

    def test_default_font_no_body_info(self):
        """本文フォント=0の場合は見出し判定しない"""
        result = JtdToMarkdownConverter._detect_heading(
            "テキスト", 850, 0)
        assert result is None


class TestIsBold:
    """太字検出テスト"""

    def test_default_font_is_bold(self):
        """フォントサイズ=0（デフォルト）で本文が明示サイズ → 太字"""
        assert JtdToMarkdownConverter._is_bold(0, 700) is True

    def test_explicit_font_not_bold(self):
        """明示フォントサイズ（本文と同じ） → 太字ではない"""
        assert JtdToMarkdownConverter._is_bold(700, 700) is False

    def test_larger_font_not_bold(self):
        """本文より大きいフォント → 見出し扱いなので太字判定はただす"""
        assert JtdToMarkdownConverter._is_bold(1200, 700) is False

    def test_no_body_font_not_bold(self):
        """本文フォント不明 → 太字判定しない"""
        assert JtdToMarkdownConverter._is_bold(0, 0) is False

    def test_smaller_font_not_bold(self):
        """本文より小さいフォント → 太字ではない"""
        assert JtdToMarkdownConverter._is_bold(500, 700) is False


class TestDetectBodyFontSize:
    """本文フォントサイズ推定テスト"""

    def test_most_frequent_size(self):
        """最頻出フォントサイズを返す"""
        blocks = [
            {'type': 'text', 'lines': [
                ("見出し", 0),
                ("本文1", 700),
                ("本文2", 700),
                ("本文3", 700),
                ("大きい文字", 1200),
            ]},
        ]
        assert JtdToMarkdownConverter._detect_body_font_size(
            blocks) == 700

    def test_no_explicit_font(self):
        """明示的フォントなし → 0"""
        blocks = [
            {'type': 'text', 'lines': [
                ("テキスト1", 0),
                ("テキスト2", 0),
            ]},
        ]
        assert JtdToMarkdownConverter._detect_body_font_size(
            blocks) == 0


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
