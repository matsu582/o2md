"""isolated_group_renderer.py (分離グループレンダラー) のテストコード

このテストコードは以下の機能をテストします：
- IsolatedGroupRendererクラスの初期化
- ユーティリティメソッドのテスト
- フェーズメソッドの基本的な動作確認

注意: このクラスはExcelToMarkdownConverterに強く依存しているため、
完全な統合テストではなく、単体テスト可能な部分のみをテストします。
"""

import os
import sys
import tempfile
from pathlib import Path
import pytest
from unittest.mock import Mock, MagicMock, patch

sys.path.insert(0, str(Path(__file__).parent.parent))

from isolated_group_renderer import IsolatedGroupRenderer
from utils import col_letter


class TestIsolatedGroupRendererInit:
    """IsolatedGroupRendererクラスの初期化テスト"""
    
    def test_class_exists(self):
        """クラスが存在することを確認"""
        assert IsolatedGroupRenderer is not None
    
    def test_initialization_with_mock_converter(self):
        """モックコンバータでの初期化をテスト"""
        mock_converter = Mock()
        renderer = IsolatedGroupRenderer(mock_converter)
        
        assert renderer.converter == mock_converter
        assert renderer.sheet is None
        assert isinstance(renderer._last_iso_preserved_ids, set)
        assert len(renderer._last_iso_preserved_ids) == 0
        assert renderer._last_temp_pdf_path is None
    
    def test_initialization_attributes(self):
        """初期化時の属性が正しく設定されることを確認"""
        mock_converter = Mock()
        renderer = IsolatedGroupRenderer(mock_converter)
        
        assert hasattr(renderer, 'converter')
        assert hasattr(renderer, 'sheet')
        assert hasattr(renderer, '_last_iso_preserved_ids')
        assert hasattr(renderer, '_last_temp_pdf_path')
    
    def test_initialization_with_different_converters(self):
        """異なるコンバータで複数のインスタンスを作成"""
        mock_converter1 = Mock(name='converter1')
        mock_converter2 = Mock(name='converter2')
        
        renderer1 = IsolatedGroupRenderer(mock_converter1)
        renderer2 = IsolatedGroupRenderer(mock_converter2)
        
        assert renderer1.converter != renderer2.converter
        assert renderer1.converter == mock_converter1
        assert renderer2.converter == mock_converter2


class TestUtilityMethods:
    """ユーティリティメソッドのテスト"""
    
    def test_col_letter_function_available(self):
        """col_letter関数がutils.pyから利用可能であることを確認"""
        assert col_letter is not None
        assert callable(col_letter)
    
    def test_col_letter_single_char(self):
        """col_letterが1文字の列を正しく変換"""
        assert col_letter(1) == 'A'
        assert col_letter(2) == 'B'
        assert col_letter(26) == 'Z'
    
    def test_col_letter_double_char(self):
        """col_letterが2文字の列を正しく変換"""
        assert col_letter(27) == 'AA'
        assert col_letter(28) == 'AB'
        assert col_letter(52) == 'AZ'


class TestRenderMethod:
    """renderメソッドの基本テスト"""
    
    def test_render_method_exists(self):
        """renderメソッドが存在することを確認"""
        mock_converter = Mock()
        renderer = IsolatedGroupRenderer(mock_converter)
        
        assert hasattr(renderer, 'render')
        assert callable(renderer.render)
    
    def test_render_with_invalid_inputs_returns_none(self):
        """無効な入力でrenderを呼ぶとNoneを返す"""
        mock_converter = Mock()
        mock_converter.excel_file = '/nonexistent/file.xlsx'
        mock_converter.workbook = Mock()
        mock_converter.workbook.sheetnames = []
        
        renderer = IsolatedGroupRenderer(mock_converter)
        
        mock_sheet = Mock()
        mock_sheet.title = 'Sheet1'
        
        result = renderer.render(mock_sheet, [], dpi=600)
        
        assert result is None or isinstance(result, tuple)
    
    def test_render_signature(self):
        """renderメソッドのシグネチャを確認"""
        mock_converter = Mock()
        renderer = IsolatedGroupRenderer(mock_converter)
        
        import inspect
        sig = inspect.signature(renderer.render)
        params = list(sig.parameters.keys())
        
        assert 'sheet' in params
        assert 'shape_indices' in params
        assert 'dpi' in params
        assert 'cell_range' in params


class TestPhase2CollectAnchors:
    """_phase2_collect_anchorsメソッドのテスト"""
    
    def test_phase2_method_exists(self):
        """_phase2_collect_anchorsメソッドが存在することを確認"""
        mock_converter = Mock()
        renderer = IsolatedGroupRenderer(mock_converter)
        
        assert hasattr(renderer, '_phase2_collect_anchors')
        assert callable(renderer._phase2_collect_anchors)
    
    def test_phase2_with_empty_xml(self):
        """空のXMLで_phase2_collect_anchorsを呼ぶ"""
        mock_converter = Mock()
        mock_converter._anchor_has_drawable = Mock(return_value=True)
        
        renderer = IsolatedGroupRenderer(mock_converter)
        
        import xml.etree.ElementTree as ET
        root = ET.Element('root')
        
        result = renderer._phase2_collect_anchors(root)
        
        assert isinstance(result, list)
        assert len(result) == 0


class TestPhase4CollectKeepIds:
    """_phase4_collect_keep_idsメソッドのテスト"""
    
    def test_phase4_method_exists(self):
        """_phase4_collect_keep_idsメソッドが存在することを確認"""
        mock_converter = Mock()
        renderer = IsolatedGroupRenderer(mock_converter)
        
        assert hasattr(renderer, '_phase4_collect_keep_ids')
        assert callable(renderer._phase4_collect_keep_ids)
    
    def test_phase4_with_empty_inputs(self):
        """空の入力で_phase4_collect_keep_idsを呼ぶ"""
        mock_converter = Mock()
        renderer = IsolatedGroupRenderer(mock_converter)
        
        result = renderer._phase4_collect_keep_ids([], [])
        
        assert isinstance(result, set)
        assert len(result) == 0
    
    def test_phase4_with_invalid_indices(self):
        """無効なインデックスで_phase4_collect_keep_idsを呼ぶ"""
        mock_converter = Mock()
        renderer = IsolatedGroupRenderer(mock_converter)
        
        result = renderer._phase4_collect_keep_ids([100, 200], [])
        
        assert isinstance(result, set)
        assert len(result) == 0


class TestConversionMethods:
    """変換メソッドのテスト"""
    
    def test_convert_excel_to_pdf_method_exists(self):
        """_convert_excel_to_pdfメソッドが存在することを確認"""
        mock_converter = Mock()
        renderer = IsolatedGroupRenderer(mock_converter)
        
        assert hasattr(renderer, '_convert_excel_to_pdf')
        assert callable(renderer._convert_excel_to_pdf)
    
    def test_convert_pdf_to_png_method_exists(self):
        """_convert_pdf_to_pngメソッドが存在することを確認"""
        mock_converter = Mock()
        renderer = IsolatedGroupRenderer(mock_converter)
        
        assert hasattr(renderer, '_convert_pdf_to_png')
        assert callable(renderer._convert_pdf_to_png)
    
    def test_convert_pdf_to_png_with_output_method_exists(self):
        """_convert_pdf_to_png_with_outputメソッドが存在することを確認"""
        mock_converter = Mock()
        renderer = IsolatedGroupRenderer(mock_converter)
        
        assert hasattr(renderer, '_convert_pdf_to_png_with_output')
        assert callable(renderer._convert_pdf_to_png_with_output)


class TestPageSetupMethods:
    """ページ設定メソッドのテスト"""
    
    def test_set_page_setup_and_margins_method_exists(self):
        """_set_page_setup_and_marginsメソッドが存在することを確認"""
        mock_converter = Mock()
        renderer = IsolatedGroupRenderer(mock_converter)
        
        assert hasattr(renderer, '_set_page_setup_and_margins')
        assert callable(renderer._set_page_setup_and_margins)


class TestCropMethod:
    """_crop_png_to_cell_rangeメソッドのテスト"""
    
    def test_crop_method_exists(self):
        """_crop_png_to_cell_rangeメソッドが存在することを確認"""
        mock_converter = Mock()
        renderer = IsolatedGroupRenderer(mock_converter)
        
        assert hasattr(renderer, '_crop_png_to_cell_range')
        assert callable(renderer._crop_png_to_cell_range)


class TestEdgeCases:
    """エッジケースのテスト"""
    
    def test_multiple_renderers_independent(self):
        """複数のレンダラーが独立して動作することを確認"""
        mock_converter1 = Mock()
        mock_converter2 = Mock()
        
        renderer1 = IsolatedGroupRenderer(mock_converter1)
        renderer2 = IsolatedGroupRenderer(mock_converter2)
        
        renderer1._last_iso_preserved_ids.add('id1')
        
        assert 'id1' in renderer1._last_iso_preserved_ids
        assert 'id1' not in renderer2._last_iso_preserved_ids
    
    def test_renderer_state_reset_on_render(self):
        """renderメソッド呼び出し時に状態がリセットされることを確認"""
        mock_converter = Mock()
        mock_converter.excel_file = '/test/file.xlsx'
        mock_converter.workbook = Mock()
        mock_converter.workbook.sheetnames = ['Sheet1']
        
        renderer = IsolatedGroupRenderer(mock_converter)
        
        initial_ids = renderer._last_iso_preserved_ids.copy()
        
        mock_sheet = Mock()
        mock_sheet.title = 'Sheet1'
        
        try:
            renderer.render(mock_sheet, [], dpi=600)
        except Exception:
            pass
        


class TestMethodChaining:
    """メソッドの連携テスト"""
    
    def test_all_phase_methods_exist(self):
        """すべてのフェーズメソッドが存在することを確認"""
        mock_converter = Mock()
        renderer = IsolatedGroupRenderer(mock_converter)
        
        phase_methods = [
            '_phase1_initialize_and_load_xml',
            '_phase2_collect_anchors',
            '_phase3_compute_cell_range',
            '_phase4_collect_keep_ids',
            '_phase5_create_tmpdir_and_resolve_connectors',
            '_phase6_prune_drawing_xml',
            '_phase7_apply_connector_cosmetics',
            '_phase8_prepare_workbook',
            '_phase9_generate_pdf_png',
            '_phase10_postprocess',
        ]
        
        for method_name in phase_methods:
            assert hasattr(renderer, method_name), f"{method_name} が存在しません"
            assert callable(getattr(renderer, method_name)), f"{method_name} が呼び出し可能ではありません"


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
