#!/usr/bin/env python3
"""
docling統合モジュール
doclingのTableFormerを使用してPDFから表を検出・抽出する機能を提供

doclingがインストールされていない場合は警告を出してスキップ
"""

from typing import List, Dict, Optional, Tuple
import os


# doclingの遅延インポート用フラグ
_DOCLING_AVAILABLE = None


def is_docling_available() -> bool:
    """doclingが利用可能かどうかを確認"""
    global _DOCLING_AVAILABLE
    if _DOCLING_AVAILABLE is None:
        try:
            from docling.document_converter import DocumentConverter
            _DOCLING_AVAILABLE = True
        except ImportError:
            _DOCLING_AVAILABLE = False
    return _DOCLING_AVAILABLE


class DoclingTableExtractor:
    """doclingを使用してPDFから表を検出・抽出するクラス
    
    doclingのTableFormerモデルを使用して、罫線のない表も含めて
    高精度な表検出・テキスト抽出を行う。
    """
    
    def __init__(self, verbose: bool = False):
        """初期化
        
        Args:
            verbose: デバッグ出力を有効にするかどうか
        """
        self.verbose = verbose
        self._converter = None
        
        if not is_docling_available():
            if self.verbose:
                print("[WARNING] doclingがインストールされていません")
                print("[WARNING] インストール: pip install docling")
    
    def _get_converter(self):
        """DocumentConverterインスタンスを取得（遅延初期化）
        
        OCRエンジンをrapidocr（torch backend）に固定（ocrmacでの問題を回避）
        MPSが利用可能な場合はMPSを使用
        """
        if self._converter is None and is_docling_available():
            from docling.document_converter import DocumentConverter, PdfFormatOption
            from docling.datamodel.pipeline_options import PdfPipelineOptions, RapidOcrOptions
            from docling.datamodel.base_models import InputFormat
            
            # OCRエンジンをrapidocr（torch backend）に固定（ocrmacでの問題を回避）
            # アクセラレータはデフォルト（MPS/CUDA/CPU自動選択）
            pipeline_options = PdfPipelineOptions()
            pipeline_options.ocr_options = RapidOcrOptions(
                backend="torch"
            )
            
            self._converter = DocumentConverter(
                format_options={
                    InputFormat.PDF: PdfFormatOption(pipeline_options=pipeline_options)
                }
            )
            if self.verbose:
                print("[INFO] docling DocumentConverterを初期化しました（rapidocr使用）")
        return self._converter
    
    def _table_data_to_markdown(self, table_data) -> str:
        """TableDataオブジェクトからMarkdown形式の表を生成
        
        doclingのexport_to_markdown()が空を返す場合のフォールバック
        
        Args:
            table_data: doclingのTableDataオブジェクト
            
        Returns:
            Markdown形式の表文字列
        """
        if not table_data:
            return ""
        
        # 方法A: gridから生成（docling 2.67.0以前）
        if hasattr(table_data, 'grid') and table_data.grid:
            return self._grid_to_markdown(table_data.grid)
        
        # 方法B: table_cellsから生成（docling 2.68.0）
        if hasattr(table_data, 'table_cells') and table_data.table_cells:
            return self._table_cells_to_markdown(
                table_data.table_cells,
                getattr(table_data, 'num_rows', 0),
                getattr(table_data, 'num_cols', 0)
            )
        
        return ""
    
    def _grid_to_markdown(self, grid) -> str:
        """gridデータからMarkdownを生成"""
        if not grid:
            return ""
        
        rows = []
        for row in grid:
            cells = []
            for cell in row:
                text = cell.text if hasattr(cell, 'text') else ""
                text = text.replace("|", "\\|").replace("\n", " ")
                cells.append(text.strip())
            rows.append(cells)
        
        if not rows:
            return ""
        
        max_cols = max(len(row) for row in rows)
        for row in rows:
            while len(row) < max_cols:
                row.append("")
        
        lines = []
        header = "| " + " | ".join(rows[0]) + " |"
        lines.append(header)
        separator = "| " + " | ".join(["---"] * max_cols) + " |"
        lines.append(separator)
        for row in rows[1:]:
            line = "| " + " | ".join(row) + " |"
            lines.append(line)
        
        return "\n".join(lines)
    
    def _table_cells_to_markdown(self, table_cells, num_rows: int, num_cols: int) -> str:
        """table_cellsデータからMarkdownを生成（docling 2.68.0対応）"""
        if not table_cells:
            return ""
        
        # num_rowsとnum_colsが0の場合、table_cellsから推定
        if num_rows == 0 or num_cols == 0:
            for cell in table_cells:
                if hasattr(cell, 'end_row_offset_idx'):
                    num_rows = max(num_rows, cell.end_row_offset_idx)
                if hasattr(cell, 'end_col_offset_idx'):
                    num_cols = max(num_cols, cell.end_col_offset_idx)
        
        if num_rows == 0 or num_cols == 0:
            return ""
        
        # グリッドを初期化
        grid = [["" for _ in range(num_cols)] for _ in range(num_rows)]
        
        # セルデータをグリッドに配置
        for cell in table_cells:
            row_idx = getattr(cell, 'start_row_offset_idx', 0)
            col_idx = getattr(cell, 'start_col_offset_idx', 0)
            text = getattr(cell, 'text', "") or ""
            text = text.replace("|", "\\|").replace("\n", " ").strip()
            
            if 0 <= row_idx < num_rows and 0 <= col_idx < num_cols:
                grid[row_idx][col_idx] = text
        
        # Markdown生成
        lines = []
        header = "| " + " | ".join(grid[0]) + " |"
        lines.append(header)
        separator = "| " + " | ".join(["---"] * num_cols) + " |"
        lines.append(separator)
        for row in grid[1:]:
            line = "| " + " | ".join(row) + " |"
            lines.append(line)
        
        return "\n".join(lines)
    
    def extract_tables_from_page(
        self, 
        pdf_path: str, 
        page_num: int
    ) -> List[str]:
        """指定ページから表を抽出してMarkdown形式で返す
        
        Args:
            pdf_path: PDFファイルのパス
            page_num: ページ番号（1-indexed）
            
        Returns:
            Markdown形式の表文字列のリスト
        """
        if not is_docling_available():
            return []
        
        converter = self._get_converter()
        if converter is None:
            return []
        
        tables_md = []
        
        try:
            # 指定ページのみを変換
            result = converter.convert(
                pdf_path, 
                page_range=(page_num, page_num)
            )
            
            # 表を抽出
            doc = result.document
            if self.verbose:
                print(f"[DEBUG] docling検出: {len(doc.tables)}個の表要素を検出")
            for i, table in enumerate(doc.tables):
                try:
                    md = None
                    
                    # 方法1: export_to_markdown(doc=doc)を試行
                    try:
                        md = table.export_to_markdown(doc=doc)
                        if self.verbose:
                            print(f"[DEBUG] 表{i+1} 方法1結果: {repr(md[:50]) if md else 'None/空'}")
                    except TypeError as e:
                        if self.verbose:
                            print(f"[DEBUG] 表{i+1} 方法1エラー: {e}")
                    
                    # 方法2: export_to_markdown()を試行（doc引数なし）
                    if not md or not md.strip():
                        try:
                            md = table.export_to_markdown()
                            if self.verbose:
                                print(f"[DEBUG] 表{i+1} 方法2結果: {repr(md[:50]) if md else 'None/空'}")
                        except Exception as e:
                            if self.verbose:
                                print(f"[DEBUG] 表{i+1} 方法2エラー: {e}")
                    
                    # 方法3: table.dataから直接Markdownを生成
                    if not md or not md.strip():
                        if self.verbose:
                            has_data = hasattr(table, 'data') and table.data is not None
                            has_grid = has_data and hasattr(table.data, 'grid') and table.data.grid
                            has_cells = has_data and hasattr(table.data, 'table_cells') and table.data.table_cells
                            num_rows = getattr(table.data, 'num_rows', 0) if has_data else 0
                            num_cols = getattr(table.data, 'num_cols', 0) if has_data else 0
                            print(f"[DEBUG] 表{i+1} 方法3: has_grid={has_grid}, has_cells={has_cells}, rows={num_rows}, cols={num_cols}")
                        md = self._table_data_to_markdown(table.data)
                        if self.verbose:
                            print(f"[DEBUG] 表{i+1} 方法3結果: {repr(md[:50]) if md else 'None/空'}")
                    
                    if self.verbose:
                        print(f"[DEBUG] 表{i+1} 最終結果: {repr(md[:100]) if md else 'None/空'}")
                    
                    if md and md.strip():
                        tables_md.append(md.strip())
                        if self.verbose:
                            lines = md.strip().split('\n')
                            print(f"[DEBUG] docling表検出: {len(lines)}行")
                except Exception as e:
                    if self.verbose:
                        print(f"[DEBUG] 表エクスポートエラー: {e}")
                    continue
                    
        except Exception as e:
            if self.verbose:
                print(f"[DEBUG] docling変換エラー (ページ{page_num}): {e}")
        
        return tables_md
    
    def extract_tables_from_pdf(
        self, 
        pdf_path: str,
        page_nums: Optional[List[int]] = None
    ) -> Dict[int, List[str]]:
        """PDFから表を抽出してページごとにMarkdown形式で返す
        
        Args:
            pdf_path: PDFファイルのパス
            page_nums: 抽出するページ番号のリスト（1-indexed、省略時は全ページ）
            
        Returns:
            ページ番号をキー、Markdown形式の表文字列のリストを値とする辞書
        """
        if not is_docling_available():
            return {}
        
        converter = self._get_converter()
        if converter is None:
            return {}
        
        result_dict: Dict[int, List[str]] = {}
        
        try:
            # PDF全体を変換
            result = converter.convert(pdf_path)
            
            # 表を抽出してページごとに分類
            doc = result.document
            for table in doc.tables:
                try:
                    # ページ番号を取得
                    table_page = 1
                    if hasattr(table, 'prov') and table.prov:
                        table_page = table.prov[0].page_no
                    
                    # 指定ページのみを抽出
                    if page_nums is not None and table_page not in page_nums:
                        continue
                    
                    md = None
                    
                    # 方法1: export_to_markdown(doc=doc)を試行
                    try:
                        md = table.export_to_markdown(doc=doc)
                    except TypeError:
                        pass
                    
                    # 方法2: export_to_markdown()を試行（doc引数なし）
                    if not md or not md.strip():
                        try:
                            md = table.export_to_markdown()
                        except Exception:
                            pass
                    
                    # 方法3: table.dataから直接Markdownを生成
                    if not md or not md.strip():
                        md = self._table_data_to_markdown(table.data)
                    
                    if md and md.strip():
                        if table_page not in result_dict:
                            result_dict[table_page] = []
                        result_dict[table_page].append(md.strip())
                        if self.verbose:
                            lines = md.strip().split('\n')
                            print(f"[DEBUG] docling表検出 (ページ{table_page}): {len(lines)}行")
                except Exception as e:
                    if self.verbose:
                        print(f"[DEBUG] 表エクスポートエラー: {e}")
                    continue
                    
        except Exception as e:
            if self.verbose:
                print(f"[DEBUG] docling変換エラー: {e}")
        
        return result_dict


def extract_slide_tables_with_docling(
    pdf_path: str,
    page_num: int,
    verbose: bool = False
) -> List[str]:
    """スライドPDFから表を抽出するヘルパー関数
    
    Args:
        pdf_path: PDFファイルのパス
        page_num: ページ番号（1-indexed）
        verbose: デバッグ出力を有効にするかどうか
        
    Returns:
        Markdown形式の表文字列のリスト
    """
    extractor = DoclingTableExtractor(verbose=verbose)
    return extractor.extract_tables_from_page(pdf_path, page_num)
