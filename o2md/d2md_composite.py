"""Word複合図形の一時文書に必要な座標情報を扱う補助関数。"""

import xml.etree.ElementTree as ET

from docx.oxml.ns import qn


def paragraph_properties_xml(paragraph):
    """元段落の段落設定を一時文書用XMLとして返す。"""
    properties = paragraph._p.find(qn("w:pPr"))
    if properties is None:
        return ""
    return ET.tostring(properties, encoding="unicode")


def section_properties_xml(document):
    """元文書のセクション設定を一時文書用XMLとして返す。"""
    section_properties = document._element.body.sectPr
    if section_properties is None:
        return (
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1440" '
            'w:bottom="1440" w:left="1440"/></w:sectPr>'
        )
    allowed = {"pgSz", "pgMar", "cols", "titlePg", "docGrid"}
    children = [
        ET.tostring(child, encoding="unicode")
        for child in section_properties
        if child.tag.rsplit("}", 1)[-1] in allowed
    ]
    return f"<w:sectPr>{''.join(children)}</w:sectPr>"


def composite_paragraphs_xml(paragraphs, drawing_xml_groups):
    """元段落ごとの設定を保った複合図形本文XMLを作る。"""
    body_parts = []
    for paragraph, drawing_xmls in zip(paragraphs, drawing_xml_groups):
        properties = paragraph_properties_xml(paragraph)
        drawings = "".join(drawing_xmls)
        body_parts.append(f"<w:p>{properties}<w:r>{drawings}</w:r></w:p>")
    return "".join(body_parts)
