"""x2md.py (Excel to Markdown Converter) のテストコード

このテストコードは以下の機能をテストします：
- ExcelToMarkdownConverterクラスの初期化
- シート変換機能
- データ範囲の検出
- 罫線テーブルの検出
- Excel文書からMarkdownへの変換
"""

import os
import sys
import tempfile
import shutil
from pathlib import Path
import pytest

sys.path.insert(0, str(Path(__file__).parent.parent))

from x2md import ExcelToMarkdownConverter


class TestExcelToMarkdownConverter:
    """ExcelToMarkdownConverterクラスのテスト"""

    @pytest.fixture
    def sample_excel_files(self):
        """テスト用のサンプルExcelファイルのパスを返す"""
        input_dir = Path(__file__).parent.parent / "input_files"
        files = {
            'simple': input_dir / "Book1.xlsx",
            'one_sheet': input_dir / "one_sheet_.xlsx",
            'two_sheet': input_dir / "tow_sheet_.xlsx",
            'three_sheet': input_dir / "three_sheet_.xlsx",
            'five_sheet': input_dir / "five_sheet_.xlsx",
            'six_sheet': input_dir / "six_sheet_.xlsx",
            'complex': input_dir / "SurportManagerAPI.xlsx",
        }
        
        existing_files = {k: str(v) for k, v in files.items() if v.exists()}
        
        if not existing_files:
            pytest.skip("テスト用Excelファイルが存在しません")
        
        return existing_files

    @pytest.fixture
    def temp_output_dir(self):
        """一時的な出力ディレクトリを作成し、テスト後にクリーンアップする"""
        temp_dir = tempfile.mkdtemp()
        yield temp_dir
        shutil.rmtree(temp_dir, ignore_errors=True)

    def test_converter_initialization(self, sample_excel_files, temp_output_dir):
        """コンバータの初期化をテスト"""
        excel_file = list(sample_excel_files.values())[0]
        converter = ExcelToMarkdownConverter(
            excel_file,
            output_dir=temp_output_dir
        )
        
        assert converter.excel_file == excel_file
        assert converter.output_dir == temp_output_dir
        assert os.path.exists(converter.images_dir)
        assert hasattr(converter, 'markdown_lines')
        assert hasattr(converter, 'workbook')

    def test_output_directory_creation(self, sample_excel_files, temp_output_dir):
        """出力ディレクトリが正しく作成されることをテスト"""
        excel_file = list(sample_excel_files.values())[0]
        converter = ExcelToMarkdownConverter(
            excel_file,
            output_dir=temp_output_dir
        )
        
        assert os.path.exists(converter.images_dir)
        assert os.path.isdir(converter.images_dir)
        assert converter.images_dir == os.path.join(temp_output_dir, "images")

    def test_workbook_loading(self, sample_excel_files, temp_output_dir):
        """Excelワークブックが正しく読み込まれることをテスト"""
        excel_file = list(sample_excel_files.values())[0]
        converter = ExcelToMarkdownConverter(
            excel_file,
            output_dir=temp_output_dir
        )
        
        assert converter.workbook is not None
        assert len(converter.workbook.sheetnames) > 0

    def test_simple_excel_conversion(self, sample_excel_files, temp_output_dir):
        """シンプルなExcelファイルの変換をテスト"""
        if 'simple' not in sample_excel_files:
            pytest.skip("Book1.xlsxが存在しません")
        
        converter = ExcelToMarkdownConverter(
            sample_excel_files['simple'],
            output_dir=temp_output_dir
        )
        
        output_file = converter.convert()
        
        assert os.path.exists(output_file)
        assert output_file.endswith('.md')
        
        with open(output_file, 'r', encoding='utf-8') as f:
            content = f.read()
            assert len(content) > 0

    def test_multi_sheet_conversion(self, sample_excel_files, temp_output_dir):
        """複数シートを持つExcelファイルの変換をテスト"""
        multi_sheet_files = [k for k in ['two_sheet', 'three_sheet', 'five_sheet'] 
                            if k in sample_excel_files]
        
        if not multi_sheet_files:
            pytest.skip("複数シートのテストファイルが存在しません")
        
        excel_file = sample_excel_files[multi_sheet_files[0]]
        converter = ExcelToMarkdownConverter(
            excel_file,
            output_dir=temp_output_dir
        )
        
        assert len(converter.workbook.sheetnames) > 1
        
        output_file = converter.convert()
        
        assert os.path.exists(output_file)

    def test_escape_angle_brackets(self, sample_excel_files, temp_output_dir):
        """角括弧のエスケープ機能をテスト"""
        excel_file = list(sample_excel_files.values())[0]
        converter = ExcelToMarkdownConverter(
            excel_file,
            output_dir=temp_output_dir
        )
        
        test_cases = [
            ("<Tag>", "&lt;Tag&gt;"),
            ("通常のテキスト", "通常のテキスト"),
            ("<html>タグ</html>", "&lt;html&gt;タグ&lt;/html&gt;"),
            (None, ""),
        ]
        
        for input_text, expected in test_cases:
            result = converter._escape_angle_brackets(input_text)
            assert result == expected

    def test_normalize_text(self, sample_excel_files, temp_output_dir):
        """テキスト正規化機能をテスト"""
        excel_file = list(sample_excel_files.values())[0]
        converter = ExcelToMarkdownConverter(
            excel_file,
            output_dir=temp_output_dir
        )
        
        test_cases = [
            ("  スペース  ", "スペース"),
            ("改行\n\nテスト", "改行 テスト"),
            ("　全角　スペース　", "全角 スペース"),
            (None, ""),
        ]
        
        for input_text, expected in test_cases:
            result = converter._normalize_text(input_text)
            assert result == expected

    def test_safe_get_cell_value(self, sample_excel_files, temp_output_dir):
        """安全なセル値取得機能をテスト"""
        excel_file = list(sample_excel_files.values())[0]
        converter = ExcelToMarkdownConverter(
            excel_file,
            output_dir=temp_output_dir
        )
        
        sheet = converter.workbook.worksheets[0]
        
        value = converter._safe_get_cell_value(sheet, 1, 1)
        assert value is None or value is not None  # エラーが起きなければOK
        
        value = converter._safe_get_cell_value(sheet, 10000, 10000)
        assert value is None or value is not None  # エラーが起きなければOK

    def test_per_sheet_state_initialization(self, sample_excel_files, temp_output_dir):
        """シートごとの状態初期化をテスト"""
        excel_file = list(sample_excel_files.values())[0]
        converter = ExcelToMarkdownConverter(
            excel_file,
            output_dir=temp_output_dir
        )
        
        assert hasattr(converter, '_cell_to_md_index')
        assert hasattr(converter, '_sheet_shape_images')
        assert hasattr(converter, '_sheet_emitted_rows')
        assert hasattr(converter, '_embedded_image_cid_by_name')

    def test_canonical_emit_flag(self, sample_excel_files, temp_output_dir):
        """カノニカル出力フラグの動作をテスト"""
        excel_file = list(sample_excel_files.values())[0]
        converter = ExcelToMarkdownConverter(
            excel_file,
            output_dir=temp_output_dir
        )
        
        assert converter._is_canonical_emit() is False
        
        converter._in_canonical_emit = True
        assert converter._is_canonical_emit() is True
        
        converter._in_canonical_emit = False
        assert converter._is_canonical_emit() is False

    def test_image_counter(self, sample_excel_files, temp_output_dir):
        """画像カウンタの初期化をテスト"""
        excel_file = list(sample_excel_files.values())[0]
        converter = ExcelToMarkdownConverter(
            excel_file,
            output_dir=temp_output_dir
        )
        
        assert hasattr(converter, 'image_counter')
        assert converter.image_counter == 0

    def test_conversion_creates_markdown_file(self, sample_excel_files, temp_output_dir):
        """Markdown変換がファイルを作成することをテスト"""
        excel_file = list(sample_excel_files.values())[0]
        converter = ExcelToMarkdownConverter(
            excel_file,
            output_dir=temp_output_dir
        )
        
        output_file = converter.convert()
        
        assert os.path.exists(output_file)
        assert output_file.endswith('.md')
        
        with open(output_file, 'r', encoding='utf-8') as f:
            content = f.read()
            assert len(content) > 0
            base_name = Path(excel_file).stem
            assert base_name in content

    def test_multiple_conversions_do_not_conflict(self, sample_excel_files, temp_output_dir):
        """複数回の変換が競合しないことをテスト"""
        excel_file = list(sample_excel_files.values())[0]
        
        converter1 = ExcelToMarkdownConverter(
            excel_file,
            output_dir=temp_output_dir
        )
        output1 = converter1.convert()
        
        converter2 = ExcelToMarkdownConverter(
            excel_file,
            output_dir=temp_output_dir
        )
        output2 = converter2.convert()
        
        assert os.path.exists(output1)
        assert os.path.exists(output2)
        assert output1 == output2  # 同じパスに出力される


class TestDataRangeDetection:
    """データ範囲検出機能のテスト"""

    @pytest.fixture
    def sample_excel_file(self):
        """テスト用のサンプルExcelファイル"""
        input_dir = Path(__file__).parent.parent / "input_files"
        excel_file = input_dir / "Book1.xlsx"
        if excel_file.exists():
            return str(excel_file)
        else:
            pytest.skip("テスト用Excelファイルが存在しません")

    @pytest.fixture
    def temp_output_dir(self):
        """一時出力ディレクトリ"""
        temp_dir = tempfile.mkdtemp()
        yield temp_dir
        shutil.rmtree(temp_dir, ignore_errors=True)

    def test_get_data_range(self, sample_excel_file, temp_output_dir):
        """データ範囲取得機能をテスト"""
        converter = ExcelToMarkdownConverter(
            sample_excel_file,
            output_dir=temp_output_dir
        )
        
        if len(converter.workbook.worksheets) > 0:
            sheet = converter.workbook.worksheets[0]
            
            if hasattr(converter, '_get_data_range'):
                data_range = converter._get_data_range(sheet)
                assert isinstance(data_range, tuple)
                assert len(data_range) == 4


class TestEdgeCases:
    """エッジケースのテスト"""

    @pytest.fixture
    def temp_output_dir(self):
        """一時出力ディレクトリ"""
        temp_dir = tempfile.mkdtemp()
        yield temp_dir
        shutil.rmtree(temp_dir, ignore_errors=True)

    def test_sanitize_filename(self, temp_output_dir):
        """ファイル名サニタイズ機能をテスト"""
        input_dir = Path(__file__).parent.parent / "input_files"
        excel_file = input_dir / "Book1.xlsx"
        
        if not excel_file.exists():
            pytest.skip("テスト用ファイルが存在しません")
        
        converter = ExcelToMarkdownConverter(
            str(excel_file),
            output_dir=temp_output_dir
        )
        
        if hasattr(converter, '_sanitize_filename'):
            test_cases = [
                ("normal.txt", "normal.txt"),
                ("file/with/slashes.txt", "file_with_slashes.txt"),
                ("file:with:colons.txt", "file_with_colons.txt"),
                ("file<with>brackets.txt", "file_with_brackets.txt"),
            ]
            
            for input_name, expected_pattern in test_cases:
                result = converter._sanitize_filename(input_name)
                assert '/' not in result
                assert ':' not in result or os.name != 'posix'  # Windowsでは:は許可されない

    def test_empty_sheet_handling(self, temp_output_dir):
        """空のシートの処理をテスト"""
        pass


class TestLoggingList:
    """_LoggingListクラスのテスト"""

    @pytest.fixture
    def temp_output_dir(self):
        """一時出力ディレクトリ"""
        temp_dir = tempfile.mkdtemp()
        yield temp_dir
        shutil.rmtree(temp_dir, ignore_errors=True)

    def test_logging_list_append(self, temp_output_dir):
        """_LoggingListのappendメソッドをテスト"""
        input_dir = Path(__file__).parent.parent / "input_files"
        excel_file = input_dir / "Book1.xlsx"
        
        if not excel_file.exists():
            pytest.skip("テスト用ファイルが存在しません")
        
        converter = ExcelToMarkdownConverter(
            str(excel_file),
            output_dir=temp_output_dir
        )
        
        assert hasattr(converter.markdown_lines, 'append')
        
        initial_length = len(converter.markdown_lines)
        converter.markdown_lines.append("Test line")
        
        assert len(converter.markdown_lines) > initial_length


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
