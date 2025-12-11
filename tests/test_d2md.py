"""d2md.py (Word to Markdown Converter) のテストコード

このテストコードは以下の機能をテストします：
- WordToMarkdownConverterクラスの初期化
- 見出し構造の解析
- アンカーID生成
- 章番号マッピング
- Word文書からMarkdownへの変換
"""

import os
import sys
import tempfile
import shutil
from pathlib import Path
import pytest

sys.path.insert(0, str(Path(__file__).parent.parent))

from d2md import WordToMarkdownConverter, convert_doc_to_docx
from utils import get_libreoffice_path


class TestWordToMarkdownConverter:
    """WordToMarkdownConverterクラスのテスト"""

    @pytest.fixture
    def sample_word_file(self):
        """テスト用のサンプルWordファイルのパスを返す"""
        input_dir = Path(__file__).parent.parent / "input_files"
        word_file = input_dir / "sample_document.docx"
        if word_file.exists():
            return str(word_file)
        else:
            pytest.skip("テスト用Wordファイルが存在しません")

    @pytest.fixture
    def temp_output_dir(self):
        """一時的な出力ディレクトリを作成し、テスト後にクリーンアップする"""
        temp_dir = tempfile.mkdtemp()
        yield temp_dir
        shutil.rmtree(temp_dir, ignore_errors=True)

    def test_converter_initialization(self, sample_word_file, temp_output_dir):
        """コンバータの初期化をテスト"""
        converter = WordToMarkdownConverter(
            sample_word_file,
            output_dir=temp_output_dir
        )
        
        assert converter.word_file == sample_word_file
        assert converter.output_dir == temp_output_dir
        assert os.path.exists(converter.images_dir)
        assert converter.markdown_lines == []
        assert converter.headings == []
        assert isinstance(converter.headings_map, dict)

    def test_generate_anchor_id(self, sample_word_file, temp_output_dir):
        """アンカーID生成のテスト"""
        converter = WordToMarkdownConverter(
            sample_word_file,
            output_dir=temp_output_dir
        )
        
        anchor_id = converter._generate_anchor_id("第1章 はじめに")
        assert anchor_id == "第1章-はじめに"
        
        anchor_id = converter._generate_anchor_id("概要 (Overview)")
        assert "概要" in anchor_id
        assert "Overview" in anchor_id
        
        anchor_id = converter._generate_anchor_id("これは テスト です")
        assert "-" in anchor_id

    def test_extract_heading_title(self, sample_word_file, temp_output_dir):
        """見出しタイトル抽出のテスト"""
        converter = WordToMarkdownConverter(
            sample_word_file,
            output_dir=temp_output_dir
        )
        
        title = converter._extract_heading_title("第1章 システム概要")
        assert title == "システム概要"
        
        title = converter._extract_heading_title("1. はじめに")
        assert title == "はじめに"
        
        title = converter._extract_heading_title("1.2 システム構成")
        assert title == "システム構成"
        
        title = converter._extract_heading_title("1.2.3 詳細設計")
        assert "詳細設計" in title

    def test_normalize_text(self, sample_word_file, temp_output_dir):
        """テキスト正規化のテスト（ユーティリティ関数）"""
        converter = WordToMarkdownConverter(
            sample_word_file,
            output_dir=temp_output_dir
        )
        
        text = "これは  テスト  です"
        if hasattr(converter, '_normalize_text'):
            normalized = converter._normalize_text(text)
            assert "  " not in normalized
            assert normalized.strip() == normalized

    def test_analyze_headings(self, sample_word_file, temp_output_dir):
        """見出し構造解析のテスト"""
        converter = WordToMarkdownConverter(
            sample_word_file,
            output_dir=temp_output_dir
        )
        
        converter._analyze_headings()
        
        assert len(converter.headings) > 0
        
        for heading in converter.headings:
            assert 'level' in heading
            assert 'text' in heading
            assert 'anchor' in heading
            assert isinstance(heading['level'], int)
            assert heading['level'] > 0

    def test_conversion_creates_markdown_file(self, sample_word_file, temp_output_dir):
        """Markdown変換がファイルを作成することをテスト"""
        converter = WordToMarkdownConverter(
            sample_word_file,
            output_dir=temp_output_dir
        )
        
        output_file = converter.convert()
        
        assert os.path.exists(output_file)
        assert output_file.endswith('.md')
        
        with open(output_file, 'r', encoding='utf-8') as f:
            content = f.read()
            assert len(content) > 0

    def test_conversion_with_heading_text_option(self, sample_word_file, temp_output_dir):
        """見出しテキスト使用オプションのテスト"""
        converter = WordToMarkdownConverter(
            sample_word_file,
            output_dir=temp_output_dir,
            use_heading_text=True
        )
        
        assert converter.use_heading_text is True
        
        output_file = converter.convert()
        
        assert os.path.exists(output_file)

    def test_image_directory_creation(self, sample_word_file, temp_output_dir):
        """画像ディレクトリが作成されることをテスト"""
        converter = WordToMarkdownConverter(
            sample_word_file,
            output_dir=temp_output_dir
        )
        
        assert os.path.exists(converter.images_dir)
        assert os.path.isdir(converter.images_dir)
        assert converter.images_dir == os.path.join(temp_output_dir, "images")

    def test_multiple_conversions_do_not_conflict(self, sample_word_file, temp_output_dir):
        """複数回の変換が競合しないことをテスト"""
        converter1 = WordToMarkdownConverter(
            sample_word_file,
            output_dir=temp_output_dir
        )
        output1 = converter1.convert()
        
        converter2 = WordToMarkdownConverter(
            sample_word_file,
            output_dir=temp_output_dir
        )
        output2 = converter2.convert()
        
        assert os.path.exists(output1)
        assert os.path.exists(output2)
        assert output1 == output2  # 同じパスに出力される


class TestConvertDocToDocx:
    """convert_doc_to_docx関数のテスト"""

    def test_function_exists(self):
        """convert_doc_to_docx関数が存在することを確認"""
        assert callable(convert_doc_to_docx)

    @pytest.mark.skipif(
        get_libreoffice_path() == "soffice",
        reason="LibreOfficeがインストールされていません"
    )
    def test_doc_conversion_requires_libreoffice(self):
        """doc変換にはLibreOfficeが必要であることを確認"""
        libreoffice_path = get_libreoffice_path()
        assert libreoffice_path != "soffice"
        assert os.path.exists(libreoffice_path) or shutil.which(libreoffice_path)


class TestEdgeCases:
    """エッジケースのテスト"""

    @pytest.fixture
    def temp_output_dir(self):
        """一時的な出力ディレクトリ"""
        temp_dir = tempfile.mkdtemp()
        yield temp_dir
        shutil.rmtree(temp_dir, ignore_errors=True)

    def test_empty_headings_map(self, temp_output_dir):
        """見出しマップが空の場合の動作をテスト"""
        pass

    def test_anchor_id_with_special_characters(self, temp_output_dir):
        """特殊文字を含むアンカーID生成のテスト"""
        input_dir = Path(__file__).parent.parent / "input_files"
        word_file = input_dir / "sample_document.docx"
        
        if not word_file.exists():
            pytest.skip("テスト用ファイルが存在しません")
        
        converter = WordToMarkdownConverter(
            str(word_file),
            output_dir=temp_output_dir
        )
        
        test_cases = [
            ("これは!テスト@です#", "これはテストです"),
            ("Test & Example", "Test--Example"),
            ("100% 完了", "100-完了"),
        ]
        
        for input_text, expected_partial in test_cases:
            anchor_id = converter._generate_anchor_id(input_text)
            assert '!' not in anchor_id
            assert '@' not in anchor_id
            assert '#' not in anchor_id


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
