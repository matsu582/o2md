"""Word field codeを失わずにMarkdownテキストへ変換する。"""

import logging
import shlex
import urllib.parse


logger = logging.getLogger(__name__)
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def qn(name: str) -> str:
    """WordprocessingML要素名を生成する。"""
    return f"{{{W}}}{name}"


QN = qn


def parse_instruction(instruction: str) -> tuple[str, list[str], dict[str, str | bool]]:
    """field instructionを種別、引数、スイッチへ分解する。"""
    try:
        tokens = shlex.split(instruction.strip(), posix=False)
    except ValueError:
        tokens = instruction.split()
    tokens = [token[1:-1] if len(token) >= 2 and token[0] == token[-1] == '"' else token for token in tokens]
    if not tokens:
        return "", [], {}
    switches = {}
    args = []
    index = 1
    while index < len(tokens):
        token = tokens[index]
        if token.startswith("\\"):
            key = token[1:].lower()
            value = True
            if index + 1 < len(tokens) and not tokens[index + 1].startswith("\\"):
                value = tokens[index + 1]
                index += 1
            switches[key] = value
        else:
            args.append(token)
        index += 1
    return tokens[0].upper(), args, switches


def _field_text(element, formatter):
    parts = []
    for child in element.iter():
        if child.tag == QN("t") and child.text:
            parts.append(formatter(child.text, element))
        elif child.tag == QN("noBreakHyphen"):
            parts.append("-")
        elif child.tag == QN("tab"):
            parts.append("\t")
        elif child.tag == QN("br"):
            parts.append("\n")
    return "".join(parts)


def convert_paragraph(
    paragraph,
    formatter=None,
    reference_handler=None,
    hyperlink_resolver=None,
) -> str:
    """段落内の通常テキストとfield resultを順序どおりに抽出する。"""
    formatter = formatter or (lambda text, _element: text)
    stack = []
    output = []

    def add_text(text, element, sink):
        if not text:
            return
        if stack:
            stack[-1]["result"].append(formatter(text, element))
        else:
            sink.append(formatter(text, element))

    def finish(field, sink):
        result = "".join(field["result"])
        kind, args, switches = parse_instruction(field["instruction"])
        if kind == "HYPERLINK" and (args or switches.get("l")):
            target = str(switches.get("l") or args[0])
            if switches.get("l"):
                target = "#" + target
            else:
                target = urllib.parse.unquote(target)
            value = f"[{result}]({target})" if result else ""
        else:
            value = result
        if stack:
            stack[-1]["result"].append(value)
        else:
            sink.append(value)

    def convert_nodes(nodes):
        nested_output = []
        for child in nodes:
            if child.tag == QN("pPr"):
                continue
            if child.tag == QN("hyperlink"):
                saved_stack = stack[:]
                stack.clear()
                text = convert_nodes(child)
                stack.extend(saved_stack)
                if not text:
                    continue
                anchor = child.get(QN("anchor"))
                relation_id = child.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                target = hyperlink_resolver(relation_id) if hyperlink_resolver and relation_id else None
                if anchor:
                    target = "#" + anchor
                value = f"[{text}]({target})" if target else text
                if stack:
                    stack[-1]["result"].append(value)
                else:
                    nested_output.append(value)
                continue
            if child.tag in {
                QN("smartTag"),
                QN("sdt"),
                QN("sdtContent"),
                QN("ins"),
                QN("del"),
                QN("customXml"),
                QN("moveFrom"),
                QN("moveTo"),
            }:
                nested_output.append(convert_nodes(child))
                continue
            if child.tag == QN("fldSimple"):
                field = {"instruction": child.get(QN("instr"), ""), "result": []}
                for nested in child.iter():
                    if nested.tag == QN("t") and nested.text:
                        field["result"].append(formatter(nested.text, child))
                finish(field, nested_output)
                continue
            if child.tag != QN("r"):
                continue
            fld_char = child.find(QN("fldChar"))
            if fld_char is not None:
                kind = fld_char.get(QN("fldCharType"))
                if kind == "begin":
                    stack.append({"instruction": "", "result": [], "separate": False})
                elif kind == "separate" and stack:
                    stack[-1]["separate"] = True
                elif kind == "end" and stack:
                    finish(stack.pop(), nested_output)
                continue
            instruction = child.find(QN("instrText"))
            if instruction is not None and instruction.text:
                if stack:
                    stack[-1]["instruction"] += instruction.text
                continue
            reference = child.find(f".//{QN('footnoteReference')}")
            endnote = child.find(f".//{QN('endnoteReference')}")
            if reference_handler and (reference is not None or endnote is not None):
                note_element = reference if reference is not None else endnote
                note_type = "fn" if reference is not None else "en"
                note_id = note_element.get(QN("id"))
                add_text(reference_handler(note_type, note_id), child, nested_output)
                continue
            add_text(_field_text(child, lambda text, _element: text), child, nested_output)
        return "".join(nested_output)

    for child in paragraph._element:
        if child.tag == QN("pPr"):
            continue
        output.append(convert_nodes([child]))
    return "".join(output)
