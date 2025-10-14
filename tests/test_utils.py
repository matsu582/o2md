"""utils.py (共通ユーティリティ) のテストコード

このテストコードは以下の機能をテストします：
- get_libreoffice_path(): プラットフォームに応じたLibreOfficeパスの取得
- col_letter(): Excel列番号から列文字への変換
"""

import os
import sys
import platform
from pathlib import Path
import pytest
from unittest.mock import patch, MagicMock

sys.path.insert(0, str(Path(__file__).parent.parent))

from utils import get_libreoffice_path, col_letter


class TestGetLibreOfficePath:
    """get_libreoffice_path関数のテスト"""
    
    def test_function_exists(self):
        """関数が存在することを確認"""
        assert callable(get_libreoffice_path)
    
    def test_returns_string(self):
        """戻り値が文字列であることを確認"""
        result = get_libreoffice_path()
        assert isinstance(result, str)
        assert len(result) > 0
    
    @patch('platform.system')
    @patch('os.path.exists')
    def test_macos_path(self, mock_exists, mock_system):
        """macOSのLibreOfficeパスを返すことを確認"""
        mock_system.return_value = 'Darwin'
        mock_exists.return_value = True
        
        result = get_libreoffice_path()
        assert result == '/Applications/LibreOffice.app/Contents/MacOS/soffice'
    
    @patch('platform.system')
    @patch('os.path.exists')
    def test_macos_path_not_found(self, mock_exists, mock_system):
        """macOSでLibreOfficeが見つからない場合"""
        mock_system.return_value = 'Darwin'
        mock_exists.return_value = False
        
        result = get_libreoffice_path()
        assert result == 'soffice'
    
    @patch('platform.system')
    @patch('os.path.exists')
    def test_linux_usr_bin_soffice(self, mock_exists, mock_system):
        """Linux環境で/usr/bin/sofficeが見つかる場合"""
        mock_system.return_value = 'Linux'
        
        def exists_side_effect(path):
            return path == '/usr/bin/soffice'
        
        mock_exists.side_effect = exists_side_effect
        
        result = get_libreoffice_path()
        assert result == '/usr/bin/soffice'
    
    @patch('platform.system')
    @patch('os.path.exists')
    def test_linux_usr_bin_libreoffice(self, mock_exists, mock_system):
        """Linux環境で/usr/bin/libreofficeが見つかる場合"""
        mock_system.return_value = 'Linux'
        
        def exists_side_effect(path):
            return path == '/usr/bin/libreoffice'
        
        mock_exists.side_effect = exists_side_effect
        
        result = get_libreoffice_path()
        assert result == '/usr/bin/libreoffice'
    
    @patch('platform.system')
    @patch('os.path.exists')
    @patch('subprocess.run')
    def test_linux_which_command(self, mock_run, mock_exists, mock_system):
        """Linux環境でwhichコマンドで検出する場合"""
        mock_system.return_value = 'Linux'
        mock_exists.return_value = False
        
        mock_result = MagicMock()
        mock_result.returncode = 0
        mock_result.stdout = '/custom/path/soffice\n'
        mock_run.return_value = mock_result
        
        result = get_libreoffice_path()
        assert result == '/custom/path/soffice'
    
    @patch('platform.system')
    @patch('os.path.exists')
    def test_windows_program_files(self, mock_exists, mock_system):
        """Windows環境でProgram FilesのLibreOfficeを検出"""
        mock_system.return_value = 'Windows'
        
        def exists_side_effect(path):
            return path == r'C:\Program Files\LibreOffice\program\soffice.exe'
        
        mock_exists.side_effect = exists_side_effect
        
        result = get_libreoffice_path()
        assert result == r'C:\Program Files\LibreOffice\program\soffice.exe'
    
    @patch('platform.system')
    @patch('os.path.exists')
    def test_unknown_platform_fallback(self, mock_exists, mock_system):
        """未知のプラットフォームではデフォルト値を返す"""
        mock_system.return_value = 'Unknown'
        mock_exists.return_value = False
        
        result = get_libreoffice_path()
        assert result == 'soffice'


class TestColLetter:
    """col_letter関数のテスト"""
    
    def test_function_exists(self):
        """関数が存在することを確認"""
        assert callable(col_letter)
    
    def test_single_letter_columns(self):
        """1文字の列（A-Z）のテスト"""
        assert col_letter(1) == 'A'
        assert col_letter(2) == 'B'
        assert col_letter(3) == 'C'
        assert col_letter(26) == 'Z'
    
    def test_double_letter_columns(self):
        """2文字の列（AA-AZ, BA-ZZ）のテスト"""
        assert col_letter(27) == 'AA'
        assert col_letter(28) == 'AB'
        assert col_letter(52) == 'AZ'
        assert col_letter(53) == 'BA'
        assert col_letter(702) == 'ZZ'
    
    def test_triple_letter_columns(self):
        """3文字の列（AAA-ZZZ）のテスト"""
        assert col_letter(703) == 'AAA'
        assert col_letter(704) == 'AAB'
        assert col_letter(728) == 'AAZ'
        assert col_letter(729) == 'ABA'
    
    def test_common_excel_columns(self):
        """よく使われるExcel列のテスト"""
        assert col_letter(1) == 'A'
        assert col_letter(5) == 'E'
        assert col_letter(10) == 'J'
        assert col_letter(26) == 'Z'
        assert col_letter(27) == 'AA'
        assert col_letter(100) == 'CV'
        assert col_letter(256) == 'IV'  # Excel 2003の最大列
    
    def test_returns_string(self):
        """戻り値が文字列であることを確認"""
        result = col_letter(1)
        assert isinstance(result, str)
    
    def test_returns_uppercase(self):
        """戻り値が大文字であることを確認"""
        result = col_letter(1)
        assert result.isupper()
    
    def test_large_column_numbers(self):
        """大きな列番号のテスト"""
        assert col_letter(16384) == 'XFD'
        
        assert col_letter(1000) == 'ALL'
        assert col_letter(10000) == 'NTP'


class TestEdgeCases:
    """エッジケースのテスト"""
    
    def test_col_letter_boundary_values(self):
        """col_letterの境界値テスト"""
        assert col_letter(26) == 'Z'
        assert col_letter(27) == 'AA'
        assert col_letter(52) == 'AZ'
        assert col_letter(53) == 'BA'
    
    def test_col_letter_sequential(self):
        """col_letterの連続値が正しいことを確認"""
        expected_sequence = ['A', 'B', 'C', 'D', 'E']
        for i, expected in enumerate(expected_sequence, start=1):
            assert col_letter(i) == expected
    
    @patch('platform.system')
    def test_get_libreoffice_path_caching(self, mock_system):
        """LibreOfficeパス取得が複数回呼ばれても安定していることを確認"""
        mock_system.return_value = 'Linux'
        
        result1 = get_libreoffice_path()
        result2 = get_libreoffice_path()
        
        assert result1 == result2


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
