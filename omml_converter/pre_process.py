# -*- coding: utf-8 -*-
"""
DOCX ファイルの前処理モジュール

DOCX ファイル内の数式 (OMML) を LaTeX 形式に変換する前処理を行います。
markitdown (MIT License) の実装を参考にしています。
"""

import re
import zipfile
from io import BytesIO
from typing import BinaryIO
from xml.etree import ElementTree as ET

from .omml import OMML_NS, oMath2Latex

# XML 名前空間の定義
NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
}

# 名前空間プレフィックスを登録
for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)


def _convert_omath_element_to_latex(element: ET.Element) -> str:
    """
    OMML 要素を LaTeX 形式に変換します。

    Args:
        element: ElementTree の Element オブジェクト (oMath 要素)

    Returns:
        str: LaTeX 形式の文字列
    """
    try:
        latex = oMath2Latex(element).latex
        return latex if latex else ""
    except Exception as e:
        print(f"[WARNING] 数式変換エラー: {e}")
        return ""


def _process_math_in_xml(content: bytes) -> bytes:
    """
    DOCX の XML ファイル内の数式を処理します。

    OMML 要素を LaTeX 形式に変換し、テキスト要素として置換します。

    Args:
        content: DOCX ファイルの XML コンテンツ (バイト列)

    Returns:
        bytes: 処理済みのコンテンツ (バイト列)
    """
    # XML をパース
    try:
        root = ET.fromstring(content)
    except ET.ParseError as e:
        print(f"[WARNING] XML パースエラー: {e}")
        return content
    
    # oMathPara (ブロック数式) を処理
    omath_para_ns = OMML_NS + "oMathPara"
    omath_ns = OMML_NS + "oMath"
    
    # すべての oMathPara 要素を検索して処理
    for parent in root.iter():
        children_to_remove = []
        children_to_add = []
        
        for i, child in enumerate(parent):
            if child.tag == omath_para_ns:
                # oMathPara 内の oMath 要素を処理
                latex_parts = []
                for omath in child.iter(omath_ns):
                    latex = _convert_omath_element_to_latex(omath)
                    if latex:
                        latex_parts.append(latex)
                
                if latex_parts:
                    # 新しい段落要素を作成
                    new_p = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
                    new_r = ET.SubElement(new_p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
                    new_t = ET.SubElement(new_r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
                    new_t.text = "$$" + " ".join(latex_parts) + "$$"
                    children_to_remove.append(child)
                    children_to_add.append((i, new_p))
            
            elif child.tag == omath_ns:
                # インライン oMath を処理
                latex = _convert_omath_element_to_latex(child)
                if latex:
                    new_r = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
                    new_t = ET.SubElement(new_r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
                    new_t.text = f"${latex}$"
                    children_to_remove.append(child)
                    children_to_add.append((i, new_r))
        
        # 要素を置換
        for child in children_to_remove:
            parent.remove(child)
        
        for i, new_elem in children_to_add:
            parent.insert(i, new_elem)
    
    # XML を文字列に変換
    return ET.tostring(root, encoding='unicode').encode('utf-8')


def pre_process_docx(input_docx: BinaryIO) -> BinaryIO:
    """
    DOCX ファイルを前処理します。

    DOCX ファイルをメモリ上で解凍し、特定の XML ファイル
    (数式を含む可能性のあるファイル) を変換してから、
    再度 DOCX ファイルとして圧縮します。

    Args:
        input_docx: DOCX ファイルのバイナリ入力ストリーム

    Returns:
        BinaryIO: 処理済み DOCX ファイルのバイナリ出力ストリーム
    """
    output_docx = BytesIO()
    
    # 前処理対象の XML ファイル
    pre_process_enable_files = [
        "word/document.xml",
        "word/footnotes.xml",
        "word/endnotes.xml",
    ]
    
    with zipfile.ZipFile(input_docx, mode="r") as zip_input:
        files = {name: zip_input.read(name) for name in zip_input.namelist()}
        
        with zipfile.ZipFile(output_docx, mode="w") as zip_output:
            zip_output.comment = zip_input.comment
            
            for name, content in files.items():
                if name in pre_process_enable_files:
                    try:
                        # コンテンツを前処理
                        updated_content = _pre_process_math(content)
                        zip_output.writestr(name, updated_content)
                    except Exception as e:
                        # 処理エラーの場合は元のコンテンツを書き込む
                        print(f"[WARNING] {name} の前処理エラー: {e}")
                        zip_output.writestr(name, content)
                else:
                    zip_output.writestr(name, content)
    
    output_docx.seek(0)
    return output_docx


def has_math_content(docx_path: str) -> bool:
    """
    DOCX ファイルに数式が含まれているかチェックします。

    Args:
        docx_path: DOCX ファイルのパス

    Returns:
        bool: 数式が含まれている場合 True
    """
    try:
        with zipfile.ZipFile(docx_path, mode="r") as z:
            if "word/document.xml" in z.namelist():
                content = z.read("word/document.xml").decode()
                return "oMath" in content or "oMathPara" in content
    except Exception:
        pass
    return False
