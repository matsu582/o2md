"""DTDを拒否するXMLパース共通処理。"""

import xml.etree.ElementTree as ET


def fromstring(data: bytes | str):
    """DOCTYPEを含まないXMLだけをElementTreeで解析する。"""
    source = data.encode("utf-8") if isinstance(data, str) else data
    if b"<!DOCTYPE" in source.upper():
        raise ET.ParseError("DOCTYPEを含むXMLは解析しません")
    return ET.fromstring(data)
