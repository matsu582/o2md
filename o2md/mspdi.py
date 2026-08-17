"""MS Project Data Interchange XML の判定処理。"""

import xml.etree.ElementTree as ET
import re


MSPDI_NAMESPACE = "schemas.microsoft.com/project"


def is_mspdi_xml(data: bytes) -> bool:
    """XMLデータがMSPDIのProject要素か判定する。"""
    if not data or b"<" not in data:
        return False
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        match = re.search(
            rb"<(?:[A-Za-z_][\w.-]*:)?Project\b[^>]*\bxmlns(?::[A-Za-z_][\w.-]*)?"
            rb"=[\"'][^\"']*schemas\.microsoft\.com/project[^\"']*[\"']",
            data,
        )
        return bool(match)
    local_name = root.tag.rsplit("}", 1)[-1]
    namespace = root.tag.split("}", 1)[0].lstrip("{") if "}" in root.tag else ""
    return local_name == "Project" and MSPDI_NAMESPACE in namespace


def is_mspdi_file(file_path: str, limit: int = 16384) -> bool:
    """ファイル先頭の限定された範囲からMSPDIか判定する。"""
    try:
        with open(file_path, "rb") as file:
            return is_mspdi_xml(file.read(limit))
    except OSError:
        return False
