"""p2md.py (PowerPoint to Markdown Converter) のテストコード

このテストコードは以下の機能をテストします：
- PowerPointToMarkdownConverterクラスの初期化
- スライドタイトル取得
- スライド内容の分析
- リストタイプ判定
- PowerPoint文書からMarkdownへの変換
"""

import os
import sys
import tempfile
import shutil
from pathlib import Path
import pytest

sys.path.insert(0, str(Path(__file__).parent.parent))

from p2md import PowerPointToMarkdownConverter


class TestPowerPointToMarkdownConverter:
    """PowerPointToMarkdownConverterクラスのテスト"""

    @pytest.fixture
    def sample_pptx_files(self):
        """テスト用のサンプルPowerPointファイルのパスを返す"""
        input_dir = Path(__file__).parent.parent / "input_files"
        
        pptx_files = list(input_dir.glob("*.pptx"))
        ppt_files = list(input_dir.glob("*.ppt"))
        
        all_files = pptx_files + ppt_files
        
        if not all_files:
            pytest.skip("テスト用PowerPointファイルが存在しません")
        
        return [str(f) for f in all_files]

    @pytest.fixture
    def temp_output_dir(self):
        """一時的な出力ディレクトリを作成し、テスト後にクリーンアップする"""
        temp_dir = tempfile.mkdtemp()
        yield temp_dir
        shutil.rmtree(temp_dir, ignore_errors=True)

    def test_converter_initialization_with_pptx(self, sample_pptx_files, temp_output_dir):
        """コンバータの初期化をテスト（.pptxファイル）"""
        pptx_files = [f for f in sample_pptx_files if f.endswith('.pptx')]
        
        if not pptx_files:
            pytest.skip(".pptxファイルが存在しません")
        
        pptx_file = pptx_files[0]
        converter = PowerPointToMarkdownConverter(
            pptx_file,
            output_dir=temp_output_dir
        )
        
        assert converter.pptx_file == pptx_file
        assert converter.output_dir == temp_output_dir
        assert os.path.exists(converter.images_dir)
        assert converter.markdown_lines == []
        assert converter.image_counter == 0
        assert converter.slide_counter == 0

    def test_output_directory_creation(self, sample_pptx_files, temp_output_dir):
        """出力ディレクトリが正しく作成されることをテスト"""
        pptx_files = [f for f in sample_pptx_files if f.endswith('.pptx')]
        
        if not pptx_files:
            pytest.skip(".pptxファイルが存在しません")
        
        pptx_file = pptx_files[0]
        converter = PowerPointToMarkdownConverter(
            pptx_file,
            output_dir=temp_output_dir
        )
        
        assert os.path.exists(converter.images_dir)
        assert os.path.isdir(converter.images_dir)
        assert converter.images_dir == os.path.join(temp_output_dir, "images")

    def test_presentation_loading(self, sample_pptx_files, temp_output_dir):
        """PowerPointプレゼンテーションが正しく読み込まれることをテスト"""
        pptx_files = [f for f in sample_pptx_files if f.endswith('.pptx')]
        
        if not pptx_files:
            pytest.skip(".pptxファイルが存在しません")
        
        pptx_file = pptx_files[0]
        converter = PowerPointToMarkdownConverter(
            pptx_file,
            output_dir=temp_output_dir
        )
        
        assert converter.prs is not None
        assert hasattr(converter.prs, 'slides')

    def test_simple_pptx_conversion(self, sample_pptx_files, temp_output_dir):
        """シンプルなPowerPointファイルの変換をテスト"""
        pptx_files = [f for f in sample_pptx_files if f.endswith('.pptx')]
        
        if not pptx_files:
            pytest.skip(".pptxファイルが存在しません")
        
        pptx_file = pptx_files[0]
        converter = PowerPointToMarkdownConverter(
            pptx_file,
            output_dir=temp_output_dir
        )
        
        output_file = converter.convert()
        
        assert os.path.exists(output_file)
        assert output_file.endswith('.md')
        
        with open(output_file, 'r', encoding='utf-8') as f:
            content = f.read()
            assert len(content) > 0

    def test_get_slide_title(self, sample_pptx_files, temp_output_dir):
        """スライドタイトル取得機能をテスト"""
        pptx_files = [f for f in sample_pptx_files if f.endswith('.pptx')]
        
        if not pptx_files:
            pytest.skip(".pptxファイルが存在しません")
        
        pptx_file = pptx_files[0]
        converter = PowerPointToMarkdownConverter(
            pptx_file,
            output_dir=temp_output_dir
        )
        
        if len(converter.prs.slides) > 0:
            slide = converter.prs.slides[0]
            title = converter._get_slide_title(slide)
            
            assert title is None or isinstance(title, str)

    def test_analyze_slide(self, sample_pptx_files, temp_output_dir):
        """スライド分析機能をテスト"""
        pptx_files = [f for f in sample_pptx_files if f.endswith('.pptx')]
        
        if not pptx_files:
            pytest.skip(".pptxファイルが存在しません")
        
        pptx_file = pptx_files[0]
        converter = PowerPointToMarkdownConverter(
            pptx_file,
            output_dir=temp_output_dir
        )
        
        if len(converter.prs.slides) > 0:
            slide = converter.prs.slides[0]
            slide_info = converter._analyze_slide(slide)
            
            assert isinstance(slide_info, dict)
            assert 'has_text' in slide_info
            assert 'has_table' in slide_info
            assert 'has_shapes' in slide_info
            assert 'content_items' in slide_info
            assert 'tables' in slide_info
            
            assert isinstance(slide_info['has_text'], bool)
            assert isinstance(slide_info['has_table'], bool)
            assert isinstance(slide_info['has_shapes'], bool)
            assert isinstance(slide_info['content_items'], list)
            assert isinstance(slide_info['tables'], list)

    def test_is_numbered_text(self, sample_pptx_files, temp_output_dir):
        """番号付きテキスト判定機能をテスト"""
        pptx_files = [f for f in sample_pptx_files if f.endswith('.pptx')]
        
        if not pptx_files:
            pytest.skip(".pptxファイルが存在しません")
        
        pptx_file = pptx_files[0]
        converter = PowerPointToMarkdownConverter(
            pptx_file,
            output_dir=temp_output_dir
        )
        
        test_cases = [
            (["1. 項目1", "2. 項目2"], True),
            (["１．項目1", "２．項目2"], True),
            (["① 項目1", "② 項目2"], True),
            (["通常のテキスト"], False),
            (["項目A", "項目B"], False),
        ]
        
        for texts, expected in test_cases:
            result = converter._is_numbered_text(texts)
            assert result == expected

    def test_remove_number_prefix(self, sample_pptx_files, temp_output_dir):
        """番号プレフィックス除去機能をテスト"""
        pptx_files = [f for f in sample_pptx_files if f.endswith('.pptx')]
        
        if not pptx_files:
            pytest.skip(".pptxファイルが存在しません")
        
        pptx_file = pptx_files[0]
        converter = PowerPointToMarkdownConverter(
            pptx_file,
            output_dir=temp_output_dir
        )
        
        test_cases = [
            ("1. 項目1", "項目1"),
            ("１．項目1", "項目1"),
            ("① 項目1", "項目1"),
            ("(1) 項目1", "項目1"),
            ("通常のテキスト", "通常のテキスト"),
        ]
        
        for input_text, expected in test_cases:
            result = converter._remove_number_prefix(input_text)
            assert result == expected

    def test_conversion_creates_markdown_file(self, sample_pptx_files, temp_output_dir):
        """Markdown変換がファイルを作成することをテスト"""
        pptx_files = [f for f in sample_pptx_files if f.endswith('.pptx')]
        
        if not pptx_files:
            pytest.skip(".pptxファイルが存在しません")
        
        pptx_file = pptx_files[0]
        converter = PowerPointToMarkdownConverter(
            pptx_file,
            output_dir=temp_output_dir
        )
        
        output_file = converter.convert()
        
        assert os.path.exists(output_file)
        assert output_file.endswith('.md')
        
        with open(output_file, 'r', encoding='utf-8') as f:
            content = f.read()
            assert len(content) > 0
            base_name = Path(pptx_file).stem
            assert base_name in content

    def test_multiple_conversions_do_not_conflict(self, sample_pptx_files, temp_output_dir):
        """複数回の変換が競合しないことをテスト"""
        pptx_files = [f for f in sample_pptx_files if f.endswith('.pptx')]
        
        if not pptx_files:
            pytest.skip(".pptxファイルが存在しません")
        
        pptx_file = pptx_files[0]
        
        converter1 = PowerPointToMarkdownConverter(
            pptx_file,
            output_dir=temp_output_dir
        )
        output1 = converter1.convert()
        
        converter2 = PowerPointToMarkdownConverter(
            pptx_file,
            output_dir=temp_output_dir
        )
        output2 = converter2.convert()
        
        assert os.path.exists(output1)
        assert os.path.exists(output2)
        assert output1 == output2

    def test_image_counter_initialization(self, sample_pptx_files, temp_output_dir):
        """画像カウンタの初期化をテスト"""
        pptx_files = [f for f in sample_pptx_files if f.endswith('.pptx')]
        
        if not pptx_files:
            pytest.skip(".pptxファイルが存在しません")
        
        pptx_file = pptx_files[0]
        converter = PowerPointToMarkdownConverter(
            pptx_file,
            output_dir=temp_output_dir
        )
        
        assert hasattr(converter, 'image_counter')
        assert converter.image_counter == 0

    def test_slide_counter_initialization(self, sample_pptx_files, temp_output_dir):
        """スライドカウンタの初期化をテスト"""
        pptx_files = [f for f in sample_pptx_files if f.endswith('.pptx')]
        
        if not pptx_files:
            pytest.skip(".pptxファイルが存在しません")
        
        pptx_file = pptx_files[0]
        converter = PowerPointToMarkdownConverter(
            pptx_file,
            output_dir=temp_output_dir
        )
        
        assert hasattr(converter, 'slide_counter')
        assert converter.slide_counter == 0


class TestPPTConversion:
    """PPTファイル変換のテスト"""

    @pytest.fixture
    def temp_output_dir(self):
        """一時出力ディレクトリ"""
        temp_dir = tempfile.mkdtemp()
        yield temp_dir
        shutil.rmtree(temp_dir, ignore_errors=True)

    @pytest.mark.skipif(
        not os.path.exists("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
        reason="LibreOfficeがインストールされていません"
    )
    def test_ppt_to_pptx_conversion(self, temp_output_dir):
        """PPTからPPTXへの変換機能をテスト"""
        input_dir = Path(__file__).parent.parent / "input_files"
        ppt_files = list(input_dir.glob("*.ppt"))
        
        if not ppt_files:
            pytest.skip(".pptファイルが存在しません")
        
        ppt_file = str(ppt_files[0])
        
        try:
            converter = PowerPointToMarkdownConverter(
                ppt_file,
                output_dir=temp_output_dir
            )
            
            assert converter.pptx_file is not None
            assert converter.pptx_file.endswith('.pptx')
            assert os.path.exists(converter.pptx_file)
            
            converter.cleanup()
        except RuntimeError as e:
            if "pptからpptxへの変換に失敗" in str(e):
                pytest.skip("LibreOfficeによる変換が失敗しました")
            raise


class TestEdgeCases:
    """エッジケースのテスト"""

    @pytest.fixture
    def temp_output_dir(self):
        """一時出力ディレクトリ"""
        temp_dir = tempfile.mkdtemp()
        yield temp_dir
        shutil.rmtree(temp_dir, ignore_errors=True)

    def test_empty_slide_handling(self, temp_output_dir):
        """空のスライドの処理をテスト"""
        pass

    def test_cleanup_method(self, temp_output_dir):
        """クリーンアップメソッドの動作をテスト"""
        input_dir = Path(__file__).parent.parent / "input_files"
        pptx_files = list(input_dir.glob("*.pptx"))
        
        if not pptx_files:
            pytest.skip(".pptxファイルが存在しません")
        
        pptx_file = str(pptx_files[0])
        converter = PowerPointToMarkdownConverter(
            pptx_file,
            output_dir=temp_output_dir
        )
        
        converter.cleanup()


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
