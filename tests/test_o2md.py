"""o2md.py (Office to Markdown Converter) のテストコード

このテストコードは以下の機能をテストします：
- ファイルタイプの自動検出
- Excel/Word/PowerPoint各変換器への自動振り分け
- convert_office_to_markdown関数の動作
- エラーハンドリング
"""

import os
import sys
import tempfile
import shutil
from pathlib import Path
import pytest

sys.path.insert(0, str(Path(__file__).parent.parent))

from o2md import detect_file_type, convert_office_to_markdown


class TestDetectFileType:
    """ファイルタイプ検出機能のテスト"""

    def test_detect_excel_xlsx(self):
        """XLSXファイルをExcelとして検出"""
        assert detect_file_type("test.xlsx") == 'excel'
        assert detect_file_type("TEST.XLSX") == 'excel'
        assert detect_file_type("/path/to/file.xlsx") == 'excel'

    def test_detect_excel_xls(self):
        """XLSファイルをExcelとして検出"""
        assert detect_file_type("test.xls") == 'excel'
        assert detect_file_type("TEST.XLS") == 'excel'
        assert detect_file_type("/path/to/file.xls") == 'excel'

    def test_detect_word_docx(self):
        """DOCXファイルをWordとして検出"""
        assert detect_file_type("test.docx") == 'word'
        assert detect_file_type("TEST.DOCX") == 'word'
        assert detect_file_type("/path/to/file.docx") == 'word'

    def test_detect_word_doc(self):
        """DOCファイルをWordとして検出"""
        assert detect_file_type("test.doc") == 'word'
        assert detect_file_type("TEST.DOC") == 'word'
        assert detect_file_type("/path/to/file.doc") == 'word'

    def test_detect_powerpoint_pptx(self):
        """PPTXファイルをPowerPointとして検出"""
        assert detect_file_type("test.pptx") == 'powerpoint'
        assert detect_file_type("TEST.PPTX") == 'powerpoint'
        assert detect_file_type("/path/to/file.pptx") == 'powerpoint'

    def test_detect_powerpoint_ppt(self):
        """PPTファイルをPowerPointとして検出"""
        assert detect_file_type("test.ppt") == 'powerpoint'
        assert detect_file_type("TEST.PPT") == 'powerpoint'
        assert detect_file_type("/path/to/file.ppt") == 'powerpoint'

    def test_detect_unknown(self):
        """未対応ファイルをunknownとして検出"""
        assert detect_file_type("test.txt") == 'unknown'
        assert detect_file_type("test.pdf") == 'unknown'
        assert detect_file_type("test.csv") == 'unknown'
        assert detect_file_type("test") == 'unknown'


class TestConvertOfficeToMarkdown:
    """convert_office_to_markdown関数のテスト"""

    @pytest.fixture
    def sample_files(self):
        """テスト用のサンプルファイルパスを返す"""
        input_dir = Path(__file__).parent.parent / "input_files"
        
        files = {
            'excel': list(input_dir.glob("*.xlsx")),
            'word': list(input_dir.glob("*.docx")),
            'powerpoint': list(input_dir.glob("*.pptx")),
        }
        
        existing_files = {}
        if files['excel']:
            existing_files['excel'] = str(files['excel'][0])
        if files['word']:
            existing_files['word'] = str(files['word'][0])
        if files['powerpoint']:
            existing_files['powerpoint'] = str(files['powerpoint'][0])
        
        if not existing_files:
            pytest.skip("テスト用ファイルが存在しません")
        
        return existing_files

    @pytest.fixture
    def temp_output_dir(self):
        """一時的な出力ディレクトリを作成し、テスト後にクリーンアップする"""
        temp_dir = tempfile.mkdtemp()
        yield temp_dir
        shutil.rmtree(temp_dir, ignore_errors=True)

    def test_file_not_found_error(self, temp_output_dir):
        """存在しないファイルを指定した場合のエラーテスト"""
        with pytest.raises(FileNotFoundError):
            convert_office_to_markdown(
                "/nonexistent/file.xlsx",
                output_dir=temp_output_dir
            )

    def test_unsupported_file_error(self, temp_output_dir):
        """未対応ファイル形式を指定した場合のエラーテスト"""
        temp_file = os.path.join(temp_output_dir, "test.txt")
        with open(temp_file, 'w') as f:
            f.write("test")
        
        with pytest.raises(ValueError) as exc_info:
            convert_office_to_markdown(temp_file, output_dir=temp_output_dir)
        
        assert "サポートされていないファイル形式" in str(exc_info.value)

    def test_convert_excel_file(self, sample_files, temp_output_dir):
        """Excelファイルの変換テスト"""
        if 'excel' not in sample_files:
            pytest.skip("Excelテストファイルが存在しません")
        
        excel_file = sample_files['excel']
        output_file = convert_office_to_markdown(
            excel_file,
            output_dir=temp_output_dir
        )
        
        assert output_file is not None
        assert os.path.exists(output_file)
        assert output_file.endswith('.md')
        
        with open(output_file, 'r', encoding='utf-8') as f:
            content = f.read()
            assert len(content) > 0

    def test_convert_word_file(self, sample_files, temp_output_dir):
        """Wordファイルの変換テスト"""
        if 'word' not in sample_files:
            pytest.skip("Wordテストファイルが存在しません")
        
        word_file = sample_files['word']
        output_file = convert_office_to_markdown(
            word_file,
            output_dir=temp_output_dir
        )
        
        assert output_file is not None
        assert os.path.exists(output_file)
        assert output_file.endswith('.md')
        
        with open(output_file, 'r', encoding='utf-8') as f:
            content = f.read()
            assert len(content) > 0

    def test_convert_powerpoint_file(self, sample_files, temp_output_dir):
        """PowerPointファイルの変換テスト"""
        if 'powerpoint' not in sample_files:
            pytest.skip("PowerPointテストファイルが存在しません")
        
        pptx_file = sample_files['powerpoint']
        output_file = convert_office_to_markdown(
            pptx_file,
            output_dir=temp_output_dir
        )
        
        assert output_file is not None
        assert os.path.exists(output_file)
        assert output_file.endswith('.md')
        
        with open(output_file, 'r', encoding='utf-8') as f:
            content = f.read()
            assert len(content) > 0

    def test_convert_word_with_heading_text_option(self, sample_files, temp_output_dir):
        """Wordファイルの変換テスト（見出しテキストオプション付き）"""
        if 'word' not in sample_files:
            pytest.skip("Wordテストファイルが存在しません")
        
        word_file = sample_files['word']
        output_file = convert_office_to_markdown(
            word_file,
            output_dir=temp_output_dir,
            use_heading_text=True
        )
        
        assert output_file is not None
        assert os.path.exists(output_file)
        assert output_file.endswith('.md')

    def test_output_directory_creation(self, sample_files, temp_output_dir):
        """出力ディレクトリが正しく作成されることをテスト"""
        if 'excel' not in sample_files:
            pytest.skip("Excelテストファイルが存在しません")
        
        custom_output = os.path.join(temp_output_dir, "custom_output")
        excel_file = sample_files['excel']
        
        output_file = convert_office_to_markdown(
            excel_file,
            output_dir=custom_output
        )
        
        assert os.path.exists(custom_output)
        assert os.path.isdir(custom_output)
        assert output_file.startswith(custom_output)

    def test_images_directory_creation(self, sample_files, temp_output_dir):
        """画像ディレクトリが作成されることをテスト"""
        if 'excel' not in sample_files:
            pytest.skip("Excelテストファイルが存在しません")
        
        excel_file = sample_files['excel']
        output_file = convert_office_to_markdown(
            excel_file,
            output_dir=temp_output_dir
        )
        
        images_dir = os.path.join(temp_output_dir, "images")
        assert os.path.exists(images_dir)
        assert os.path.isdir(images_dir)

    def test_multiple_conversions_same_directory(self, sample_files, temp_output_dir):
        """同じディレクトリへの複数変換が競合しないことをテスト"""
        available_files = []
        
        if 'excel' in sample_files:
            available_files.append(sample_files['excel'])
        if 'word' in sample_files:
            available_files.append(sample_files['word'])
        if 'powerpoint' in sample_files:
            available_files.append(sample_files['powerpoint'])
        
        if len(available_files) < 2:
            pytest.skip("複数種類のテストファイルが存在しません")
        
        output_files = []
        for file_path in available_files[:2]:
            output_file = convert_office_to_markdown(
                file_path,
                output_dir=temp_output_dir
            )
            output_files.append(output_file)
        
        for output_file in output_files:
            assert os.path.exists(output_file)
        
        assert len(set(output_files)) == len(output_files)


class TestIntegrationScenarios:
    """統合シナリオテスト"""

    @pytest.fixture
    def temp_output_dir(self):
        """一時出力ディレクトリ"""
        temp_dir = tempfile.mkdtemp()
        yield temp_dir
        shutil.rmtree(temp_dir, ignore_errors=True)

    def test_all_file_types_detection(self):
        """すべてのサポートファイルタイプが正しく検出されることをテスト"""
        test_cases = [
            ("file.xlsx", "excel"),
            ("file.xls", "excel"),
            ("file.docx", "word"),
            ("file.doc", "word"),
            ("file.pptx", "powerpoint"),
            ("file.ppt", "powerpoint"),
            ("file.txt", "unknown"),
        ]
        
        for filename, expected_type in test_cases:
            assert detect_file_type(filename) == expected_type

    def test_case_insensitive_detection(self):
        """拡張子の大文字小文字を区別しないことをテスト"""
        test_cases = [
            "file.XLSX",
            "file.XLSx",
            "file.Docx",
            "file.PPTX",
        ]
        
        for filename in test_cases:
            file_type = detect_file_type(filename)
            assert file_type != 'unknown'


class TestEdgeCases:
    """エッジケースのテスト"""

    @pytest.fixture
    def temp_output_dir(self):
        """一時出力ディレクトリ"""
        temp_dir = tempfile.mkdtemp()
        yield temp_dir
        shutil.rmtree(temp_dir, ignore_errors=True)

    def test_filename_with_multiple_dots(self):
        """複数のドットを含むファイル名の検出テスト"""
        assert detect_file_type("my.file.name.xlsx") == 'excel'
        assert detect_file_type("document.v1.0.docx") == 'word'
        assert detect_file_type("presentation.final.pptx") == 'powerpoint'

    def test_filename_with_spaces(self):
        """スペースを含むファイル名の検出テスト"""
        assert detect_file_type("my file.xlsx") == 'excel'
        assert detect_file_type("word document.docx") == 'word'

    def test_path_with_special_characters(self):
        """特殊文字を含むパスの検出テスト"""
        assert detect_file_type("/path/to/file-name_123.xlsx") == 'excel'
        assert detect_file_type("C:\\Users\\Test\\document (1).docx") == 'word'


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
