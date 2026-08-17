"""Word複合図形の一時文書に必要な座標情報を扱う補助関数。"""

import xml.etree.ElementTree as ET

from docx.oxml.ns import qn


WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
EMU_PER_TWIP = 635


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


def _parse_fragment(xml_text):
    if not xml_text:
        return None
    try:
        return ET.fromstring(xml_text)
    except ET.ParseError:
        if xml_text.lstrip().startswith("<w:"):
            try:
                start = xml_text.find(">")
                if start < 0:
                    return None
                return ET.fromstring(
                    xml_text[:start]
                    + f' xmlns:w="{W_NS}"'
                    + xml_text[start:]
                )
            except ET.ParseError:
                return None
        return None


def _twips_attribute(element, name):
    value = element.get(_tag(W_NS, name))
    if value is None:
        return None
    try:
        return int(value)
    except ValueError:
        return None


def _paragraph_layout(paragraph_xml, section_xml):
    paragraph = _parse_fragment(paragraph_xml)
    section = _parse_fragment(section_xml)
    if paragraph is None or section is None:
        return None

    paragraph_properties = (
        paragraph
        if _local_name(paragraph.tag) == "pPr"
        else paragraph.find(_tag(W_NS, "pPr"))
    )
    section_properties = section
    page_size = section_properties.find(_tag(W_NS, "pgSz"))
    page_margins = section_properties.find(_tag(W_NS, "pgMar"))
    if page_size is None or page_margins is None:
        return None

    page_width = _twips_attribute(page_size, "w")
    margin_left = _twips_attribute(page_margins, "left")
    margin_right = _twips_attribute(page_margins, "right")
    if None in {page_width, margin_left, margin_right}:
        return None

    left_indent = 0
    right_indent = 0
    alignment = None
    if paragraph_properties is not None:
        alignment_element = paragraph_properties.find(_tag(W_NS, "jc"))
        if alignment_element is not None:
            alignment = alignment_element.get(_tag(W_NS, "val"))
        indent = paragraph_properties.find(_tag(W_NS, "ind"))
        if indent is not None:
            left_indent = _twips_attribute(indent, "left")
            if left_indent is None:
                left_indent = _twips_attribute(indent, "start")
            right_indent = _twips_attribute(indent, "right")
            if right_indent is None:
                right_indent = _twips_attribute(indent, "end")
            left_indent = left_indent or 0
            right_indent = right_indent or 0

    text_width = page_width - margin_left - margin_right - left_indent - right_indent
    if text_width < 0:
        return None
    snap_to_grid = True
    if paragraph_properties is not None:
        snap_element = paragraph_properties.find(_tag(W_NS, "snapToGrid"))
        if snap_element is not None:
            snap_to_grid = snap_element.get(_tag(W_NS, "val")) not in {
                "0",
                "false",
                "off",
            }
    doc_grid = section_properties.find(_tag(W_NS, "docGrid"))
    line_pitch = None
    if (
        doc_grid is not None
        and doc_grid.get(_tag(W_NS, "type")) in {"lines", "lineAndChar"}
    ):
        line_pitch_twips = _twips_attribute(doc_grid, "linePitch")
        if line_pitch_twips is not None and line_pitch_twips > 0:
            line_pitch = line_pitch_twips * EMU_PER_TWIP
    return {
        "alignment": alignment,
        "left_indent": left_indent * EMU_PER_TWIP,
        "text_width": text_width * EMU_PER_TWIP,
        "line_pitch": line_pitch if snap_to_grid else None,
    }


def _inline_horizontal_offset(image_width, layout):
    if layout is None:
        return None
    available_width = layout["text_width"]
    left_indent = layout["left_indent"]
    alignment = layout["alignment"]
    if image_width > available_width:
        return 0
    if alignment in {"center", "distribute"}:
        offset = (available_width - image_width) // 2 + left_indent
    elif alignment in {"right", "end"}:
        offset = available_width - image_width + left_indent
    else:
        offset = left_indent
    return max(0, offset)


def _vertical_grid_gap(image_heights, layout):
    line_pitch = layout.get("line_pitch") if layout else None
    if not line_pitch or not image_heights:
        return 0
    image_height = max(image_heights)
    remainder = image_height % line_pitch
    if remainder == 0:
        return 0
    gap = (line_pitch - remainder) // 2
    if gap > image_height:
        return None
    return gap


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


def _rewrite_shape_anchor(anchor, left_indent=0, vertical_gap=0):
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
        _set_margin_position(position_h, offset + left_indent)
    elif horizontal not in {"page", "margin"}:
        return False

    vertical = position_v.get("relativeFrom")
    if vertical in {"paragraph", "line"}:
        offset = _position_offset(position_v)
        if offset is None:
            return False
        _set_margin_position(position_v, offset - vertical_gap)
    elif vertical not in {"page", "margin"}:
        return False

    return True


def _inline_to_margin_anchor(inline, drawing_index, horizontal_offset=0):
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
    ET.SubElement(position_h, _tag(WP_NS, "posOffset")).text = str(
        horizontal_offset
    )
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


def absolute_drawing_xml(
    drawing_xmls, logger=None, paragraph_xml="", section_xml=""
):
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

    layout = _paragraph_layout(paragraph_xml, section_xml)
    if paragraph_xml or section_xml:
        if layout is None:
            if logger:
                logger.warning("段落またはセクションの座標情報を取得できないため、従来配置へフォールバックします")
            return None

    inline_heights = []
    for drawing in parsed_drawings:
        inline = drawing.find(f".//{_tag(WP_NS, 'inline')}")
        if inline is None:
            continue
        extent = inline.find(_tag(WP_NS, "extent"))
        image_height = (
            int(extent.get("cy"))
            if extent is not None and extent.get("cy", "").isdigit()
            else None
        )
        if image_height is None:
            if logger:
                logger.warning("画像の高さを取得できないため、従来配置へフォールバックします")
            return None
        inline_heights.append(image_height)
    vertical_gap = _vertical_grid_gap(inline_heights, layout)
    if vertical_gap is None:
        if logger:
            logger.warning("行グリッド補正量が画像高を超えるため、従来配置へフォールバックします")
        return None
    if vertical_gap and logger:
        line_pitch = layout["line_pitch"]
        logger.debug(
            "[DEBUG] 行グリッド補正: 画像高=%d EMU, linePitch=%d EMU, gap=%d EMU (%.3fpt)",
            max(inline_heights),
            line_pitch,
            vertical_gap,
            vertical_gap / 12700,
        )

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
            extent = inline.find(_tag(WP_NS, "extent"))
            image_width = (
                int(extent.get("cx"))
                if extent is not None and extent.get("cx", "").isdigit()
                else None
            )
            horizontal_offset = (
                _inline_horizontal_offset(image_width, layout)
                if image_width is not None and layout is not None
                else 0
            )
            if horizontal_offset is None:
                if logger:
                    logger.warning("画像の水平配置を計算できないため、従来配置へフォールバックします")
                return None
            replacement = _inline_to_margin_anchor(
                inline, next_doc_pr_id, horizontal_offset
            )
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
            left_indent = layout["left_indent"] if layout is not None else 0
            if not _rewrite_shape_anchor(anchor, left_indent, vertical_gap):
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
