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
        """DocumentConverterインスタンスを取得（遅延初期化）"""
        if self._converter is None and is_docling_available():
            from docling.document_converter import DocumentConverter
            self._converter = DocumentConverter()
            if self.verbose:
                print("[INFO] docling DocumentConverterを初期化しました")
        return self._converter
    
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
            for table in doc.tables:
                try:
                    # Markdown形式でエクスポート（doc引数を渡す）
                    try:
                        md = table.export_to_markdown(doc=doc)
                    except TypeError:
                        # 古いバージョンのdoclingではdoc引数がない
                        md = table.export_to_markdown()
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
                    
                    # Markdown形式でエクスポート（doc引数を渡す）
                    try:
                        md = table.export_to_markdown(doc=doc)
                    except TypeError:
                        # 古いバージョンのdoclingではdoc引数がない
                        md = table.export_to_markdown()
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
