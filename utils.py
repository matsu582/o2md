#!/usr/bin/env python3
"""
共通ユーティリティモジュール
Office文書変換で使用される共通関数を提供
"""

import os
import platform
import subprocess
import shutil
from typing import Optional


def get_libreoffice_path() -> str:
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


def get_imagemagick_command() -> str:
    """ImageMagickのコマンド名を取得（バージョンに応じて'magick'または'convert'）"""
    try:
        if shutil.which('magick'):
            return 'magick'
        elif shutil.which('convert'):
            return 'convert'
        else:
            return 'convert'
    except Exception:
        return 'convert'


def col_letter(n: int) -> str:
    """
    列番号（1始まり）をExcel形式の列文字（A, B, ..., Z, AA, AB, ...）に変換
    
    Args:
        n: 列番号（1始まり）
    
    Returns:
        Excel形式の列文字列
    
    Examples:
        >>> col_letter(1)
        'A'
        >>> col_letter(26)
        'Z'
        >>> col_letter(27)
        'AA'
    """
    letters = ''
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def sanitize_filename(s: str) -> str:
    """
    ファイル名として安全な文字列に変換
    
    Args:
        s: 元の文字列
    
    Returns:
        サニタイズされた文字列
    """
    safe = []
    for c in s:
        if c.isalnum() or c in (' ', '_', '-'):
            safe.append(c)
        else:
            safe.append('_')
    result = ''.join(safe).strip()
    while '__' in result:
        result = result.replace('__', '_')
    return result


def detect_image_format(image_data: bytes) -> str:
    """
    画像データから画像フォーマットを検出
    
    Args:
        image_data: 画像のバイトデータ
    
    Returns:
        画像フォーマット（'png', 'jpeg', 'emf', 'wmf', 'unknown'）
    """
    if not image_data or len(image_data) < 8:
        return 'unknown'
    
    if image_data[:8] == b'\x89PNG\r\n\x1a\n':
        return 'png'
    
    if image_data[:2] == b'\xff\xd8':
        return 'jpeg'
    
    if len(image_data) >= 4:
        if image_data[:4] == b'\x01\x00\x00\x00':
            if len(image_data) >= 44:
                signature = image_data[40:44]
                if signature == b' EMF':
                    return 'emf'
    
    if len(image_data) >= 4:
        if image_data[:4] == b'\xd7\xcd\xc6\x9a':
            return 'wmf'
        if image_data[:2] == b'\x01\x00' or image_data[:2] == b'\x02\x00':
            return 'wmf'
    
    return 'unknown'


def to_positive(value: int, orig_ext: int, orig_ch_ext: int, target_px: int) -> int:
    """
    負のオフセット値を正の値に変換（Excel図形の座標変換用）
    
    Args:
        value: 変換対象の値
        orig_ext: 元の範囲
        orig_ch_ext: 元の子要素の範囲
        target_px: ターゲットピクセル値
    
    Returns:
        正の値に変換された結果
    """
    if value >= 0:
        return value
    
    try:
        if orig_ext > 0 and orig_ch_ext > 0:
            scale = target_px / float(orig_ext)
            neg_px = scale * abs(value)
            new_ext = target_px - neg_px
            if new_ext > 0:
                ratio = float(orig_ch_ext) / orig_ext
                return int(new_ext * ratio)
    except (ZeroDivisionError, ValueError):
        pass
    
    return max(0, value)
