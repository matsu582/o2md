"""Word複合図形の一時文書に必要な座標情報を扱う補助関数。"""

import xml.etree.ElementTree as ET

from docx.oxml.ns import qn


WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"


def _tag(namespace, name):
    return f"{{{namespace}}}{name}"


def _local_name(tag):
    return tag.rsplit("}", 1)[-1]


def _position_offset(position):
    offset = position.find(_tag(WP_NS, "posOffset"))
    if offset is None or offset.text is None:
        return None
    try:
        return int(offset.text)
    except ValueError:
        return None


def _set_margin_position(position, offset):
    position.set("relativeFrom", "margin")
    align = position.find(_tag(WP_NS, "align"))
    if align is not None:
        position.remove(align)
    pos_offset = position.find(_tag(WP_NS, "posOffset"))
    if pos_offset is None:
        pos_offset = ET.Element(_tag(WP_NS, "posOffset"))
        position.append(pos_offset)
    pos_offset.text = str(offset)


def _rewrite_shape_anchor(anchor):
    extent = anchor.find(_tag(WP_NS, "extent"))
    position_h = anchor.find(_tag(WP_NS, "positionH"))
    position_v = anchor.find(_tag(WP_NS, "positionV"))
    if extent is None or position_h is None or position_v is None:
        return False

    if extent.get("cx") is None or extent.get("cy") is None:
        return False

    horizontal = position_h.get("relativeFrom")
    if horizontal == "column":
        offset = _position_offset(position_h)
        if offset is None:
            return False
        _set_margin_position(position_h, offset)
    elif horizontal not in {"page", "margin"}:
        return False

    vertical = position_v.get("relativeFrom")
    if vertical in {"paragraph", "line"}:
        offset = _position_offset(position_v)
        if offset is None:
            return False
        _set_margin_position(position_v, offset)
    elif vertical not in {"page", "margin"}:
        return False

    return True


def _inline_to_margin_anchor(inline, drawing_index):
    extent = inline.find(_tag(WP_NS, "extent"))
    if extent is None or extent.get("cx") is None or extent.get("cy") is None:
        return None

    anchor_attributes = {
        "distT": "0",
        "distB": "0",
        "distL": "0",
        "distR": "0",
        "simplePos": "0",
        "relativeHeight": "0",
        "behindDoc": "1",
        "locked": inline.get("locked", "0"),
        "layoutInCell": "1",
        "allowOverlap": "1",
    }
    for name in ("anchorId", "editId"):
        if inline.get(name) is not None:
            anchor_attributes[name] = inline.get(name)
    anchor = ET.Element(_tag(WP_NS, "anchor"), anchor_attributes)
    ET.SubElement(anchor, _tag(WP_NS, "simplePos"), {"x": "0", "y": "0"})

    position_h = ET.SubElement(
        anchor, _tag(WP_NS, "positionH"), {"relativeFrom": "margin"}
    )
    ET.SubElement(position_h, _tag(WP_NS, "posOffset")).text = "0"
    position_v = ET.SubElement(
        anchor, _tag(WP_NS, "positionV"), {"relativeFrom": "margin"}
    )
    ET.SubElement(position_v, _tag(WP_NS, "posOffset")).text = "0"

    ET.SubElement(anchor, _tag(WP_NS, "wrapNone"))
    for child in inline:
        if _local_name(child.tag) in {
            "extent",
            "effectExtent",
            "docPr",
            "cNvGraphicFramePr",
            "graphic",
        }:
            anchor.append(child)

    doc_pr = anchor.find(_tag(WP_NS, "docPr"))
    if doc_pr is not None:
        doc_pr.set("id", str(drawing_index))
    return anchor


def absolute_drawing_xml(drawing_xmls, logger=None):
    """合成図形を余白基準の絶対配置へ変換する。"""
    parsed_drawings = []
    used_doc_pr_ids = set()
    for drawing_xml in drawing_xmls:
        try:
            drawing = ET.fromstring(drawing_xml)
        except ET.ParseError:
            if logger:
                logger.warning("合成図形のXMLを解析できないため、従来配置へフォールバックします")
            return None
        parsed_drawings.append(drawing)
        for doc_pr in drawing.findall(f".//{_tag(WP_NS, 'docPr')}"):
            if doc_pr.get("id") is not None:
                used_doc_pr_ids.add(doc_pr.get("id"))

    converted = []
    next_doc_pr_id = 1
    for drawing in parsed_drawings:
        while str(next_doc_pr_id) in used_doc_pr_ids:
            next_doc_pr_id += 1

        inline = drawing.find(f".//{_tag(WP_NS, 'inline')}")
        anchor = drawing.find(f".//{_tag(WP_NS, 'anchor')}")
        if inline is not None and anchor is not None:
            if logger:
                logger.warning("inlineとanchorが同一drawingにあるため、従来配置へフォールバックします")
            return None

        if inline is not None:
            replacement = _inline_to_margin_anchor(inline, next_doc_pr_id)
            if replacement is None:
                if logger:
                    logger.warning("画像のextentを取得できないため、従来配置へフォールバックします")
                return None
            parent = drawing
            parent.remove(inline)
            parent.append(replacement)
            used_doc_pr_ids.add(str(next_doc_pr_id))
            next_doc_pr_id += 1
        elif anchor is not None:
            if not _rewrite_shape_anchor(anchor):
                if logger:
                    logger.warning("図形の座標情報を取得できないため、従来配置へフォールバックします")
                return None
        else:
            if logger:
                logger.warning("anchorまたはinlineを含まないdrawingのため、従来配置へフォールバックします")
            return None

        converted.append(ET.tostring(drawing, encoding="unicode"))
    return converted


def absolute_composite_paragraph_xml(drawing_xmls):
    """絶対配置用に高さを持たない空段落を作る。"""
    drawings = "".join(drawing_xmls)
    return (
        "<w:p>"
        '<w:pPr><w:spacing w:line="1" w:lineRule="exact"/>'
        '<w:rPr><w:sz w:val="1"/><w:szCs w:val="1"/></w:rPr></w:pPr>'
        f"<w:r>{drawings}</w:r>"
        "</w:p>"
    )


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
