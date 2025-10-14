#!/usr/bin/env python3
"""
共通ユーティリティモジュール
Office to Markdown変換ツールで使用される共通機能
"""

import os
import platform
import subprocess
import shutil


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
