#!/usr/bin/env python3
"""
共通ユーティリティモジュール
Office to Markdown変換ツールで使用される共通機能
"""

import os
import platform
import subprocess
import shutil
import zipfile
import xml.etree.ElementTree as ET
from typing import Optional, Tuple, List, Set, Dict, Any
import logging


def get_libreoffice_path():
    """プラットフォームに応じたLibreOfficeのパスを取得"""
    system = platform.system()
    
    if system == "Darwin":  # macOS
        path = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
        if os.path.exists(path):
            return path
    elif system == "Linux":  # Ubuntu/Linux
        common_paths = [
            "/usr/bin/soffice",
            "/usr/bin/libreoffice",
            "/snap/bin/libreoffice",
        ]
        for path in common_paths:
            if os.path.exists(path):
                return path
        try:
            result = subprocess.run(["which", "soffice"], capture_output=True, text=True)
            if result.returncode == 0 and result.stdout.strip():
                return result.stdout.strip()
            result = subprocess.run(["which", "libreoffice"], capture_output=True, text=True)
            if result.returncode == 0 and result.stdout.strip():
                return result.stdout.strip()
        except Exception:
            pass  # コマンド検索失敗は無視
    elif system == "Windows":
        common_paths = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        for path in common_paths:
            if os.path.exists(path):
                return path
    
    return "soffice"


def col_letter(n: int) -> str:
    """列番号をExcelの列文字に変換（1 -> A, 27 -> AA, etc.）"""
    letters = ''
    while n > 0:
        n, rem = divmod(n-1, 26)
        letters = chr(65 + rem) + letters
    return letters


# ============================================================================
# ============================================================================

def normalize_excel_path(path: str) -> str:
    """Excelファイル内のパスを正規化する
    
    Args:
        path: 正規化するパス（例: '../drawings/drawing1.xml'）
    
    Returns:
        正規化されたパス（例: 'xl/drawings/drawing1.xml'）
    """
    if path.startswith('..'):
        path = path.replace('../', 'xl/')
    if path.startswith('/'):
        path = path.lstrip('/')
    return path


def get_xml_from_zip(z: zipfile.ZipFile, path: str) -> Optional[ET.Element]:
    """ZIPファイルからXMLを取得してパースする
    
    Args:
        z: ZipFileオブジェクト
        path: XMLファイルのパス
    
    Returns:
        パースされたXMLのルート要素、失敗時はNone
    """
    try:
        if path not in z.namelist():
            return None
        xml_bytes = z.read(path)
        return ET.fromstring(xml_bytes)
    except Exception as e:
        logging.debug(f"Failed to get XML from zip: {path}, error: {e}")
        return None


# ============================================================================
# ============================================================================

def extract_anchor_id(anchor: ET.Element, allow_idx: bool = False) -> Optional[str]:
    """アンカー要素からcNvPr IDを抽出する
    
    Args:
        anchor: アンカー要素（twoCellAnchor, oneCellAnchor等）
        allow_idx: Trueの場合、idが見つからない時にidx属性も確認する
    
    Returns:
        cNvPr ID文字列、見つからない場合はNone
    """
    try:
        for sub in anchor.iter():
            if sub.tag.split('}')[-1].lower() == 'cnvpr':
                cid = sub.attrib.get('id')
                if cid is not None:
                    return str(cid)
                if allow_idx:
                    cid = sub.attrib.get('idx')
                    if cid is not None:
                        return str(cid)
        return None
    except Exception:
        return None


def anchor_is_hidden(anchor: ET.Element) -> bool:
    """アンカーが非表示かどうかを判定する
    
    Args:
        anchor: アンカー要素
    
    Returns:
        非表示の場合True
    """
    try:
        for sub in anchor.iter():
            if sub.tag.split('}')[-1].lower() == 'cnvpr':
                hidden = sub.attrib.get('hidden')
                if hidden in ('1', 'true'):
                    return True
                break
        return False
    except Exception:
        return False


def anchor_has_drawable(anchor: ET.Element) -> bool:
    """アンカーが描画可能な要素を持つか判定
    
    Args:
        anchor: アンカー要素
    
    Returns:
        描画可能な要素を持つ場合True
    """
    try:
        for child in anchor:
            tag_local = child.tag.split('}')[-1].lower()
            if tag_local in ('pic', 'sp', 'grpsp', 'graphicframe', 'cxnsp'):
                return True
        return False
    except Exception:
        return False


def collect_anchors(drawing_xml: ET.Element, anchor_filter_func=None) -> List[ET.Element]:
    """drawing XMLから描画可能なアンカーリストを取得
    
    Args:
        drawing_xml: drawing XMLのルート要素
        anchor_filter_func: アンカーをフィルタする関数（オプション）
    
    Returns:
        アンカー要素のリスト
    """
    anchors = []
    try:
        for node in drawing_xml:
            lname = node.tag.split('}')[-1].lower()
            if lname in ('twocellanchor', 'onecellanchor'):
                if anchor_filter_func is None or anchor_filter_func(node):
                    anchors.append(node)
    except Exception:
        pass
    return anchors


# ============================================================================
# ============================================================================

def compute_sheet_cell_pixel_map(sheet, dpi: int = 300) -> Tuple[List[float], List[float]]:
    """Excelシートのセル→ピクセル座標変換マップを計算
    
    Args:
        sheet: openpyxlのWorksheetオブジェクト
        dpi: DPI設定
    
    Returns:
        (列座標リスト, 行座標リスト)のタプル
    """
    try:
        col_pixels = [0.0]
        for col_idx in range(1, sheet.max_column + 1):
            col_letter_str = col_letter(col_idx)
            col_dim = sheet.column_dimensions.get(col_letter_str)
            if col_dim and col_dim.width:
                width_chars = col_dim.width
            else:
                width_chars = 8.43  # デフォルト幅
            width_px = width_chars * 7.0 * (dpi / 96.0)
            col_pixels.append(col_pixels[-1] + width_px)
        
        row_pixels = [0.0]
        for row_idx in range(1, sheet.max_row + 1):
            row_dim = sheet.row_dimensions.get(row_idx)
            if row_dim and row_dim.height:
                height_pt = row_dim.height
            else:
                height_pt = 15.0  # デフォルト高さ
            height_px = height_pt * (dpi / 72.0)
            row_pixels.append(row_pixels[-1] + height_px)
        
        return col_pixels, row_pixels
    except Exception as e:
        logging.debug(f"compute_sheet_cell_pixel_map failed: {e}")
        return [0.0], [0.0]
