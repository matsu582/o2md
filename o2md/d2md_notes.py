"""DOCX脚注・文末脚注の読み込みと参照管理。"""

import logging
import posixpath
import zipfile
import xml.etree.ElementTree as ET


logger = logging.getLogger(__name__)
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def qn(name: str) -> str:
    """WordprocessingML要素名を生成する。"""
    return f"{{{W}}}{name}"


QN = qn


class NoteManager:
    """脚注パートを読み込み、本文参照と末尾定義を管理する。"""

    def __init__(self, docx_path: str):
        self.notes = {"fn": {}, "en": {}}
        self.note_nodes = {"fn": {}, "en": {}}
        self.references = {}
        try:
            with zipfile.ZipFile(docx_path) as package:
                names = set(package.namelist())
                targets = self._relationship_targets(package)
                for kind, conventional in (("fn", "word/footnotes.xml"), ("en", "word/endnotes.xml")):
                    name = targets.get(kind, conventional)
                    if name not in names:
                        continue
                    data = package.read(name)
                    if b"<!DOCTYPE" in data.upper():
                        logger.warning(
                            "DTDを含む%sパートをスキップします: %s",
                            kind,
                            name,
                        )
                        continue
                    root = ET.fromstring(data)
                    for node in root:
                        if node.tag != QN("footnote") and node.tag != QN("endnote"):
                            continue
                        note_id = node.get(QN("id"))
                        if note_id in ("-1", "0") or node.get(QN("type")) in (
                            "separator", "continuationSeparator", "continuationNotice"
                        ):
                            continue
                        self.notes[kind][note_id] = self._note_text(node)
                        self.note_nodes[kind][note_id] = node
        except (OSError, zipfile.BadZipFile, ET.ParseError) as exc:
            logger.warning("脚注パートを読み込めません。脚注をスキップします: %s", exc)

    def _relationship_targets(self, package):
        """document.xmlのrelationshipから脚注パートの場所を解決する。"""
        rels_name = "word/_rels/document.xml.rels"
        if rels_name not in package.namelist():
            return {}
        try:
            root = ET.fromstring(package.read(rels_name))
            targets = {}
            for relation in root:
                rel_type = relation.get("Type", "")
                target = relation.get("Target", "")
                if "footnotes" in rel_type:
                    kind = "fn"
                elif "endnotes" in rel_type:
                    kind = "en"
                else:
                    continue
                if target.startswith("/"):
                    target = target.lstrip("/")
                else:
                    target = posixpath.join("word", target)
                targets[kind] = posixpath.normpath(target)
            return targets
        except ET.ParseError as exc:
            logger.warning("脚注relationshipを読み込めません: %s", exc)
            return {}

    def _note_text(self, node):
        """脚注内のブロックを出現順にテキスト化する。"""
        blocks = []
        for block in node:
            if block.tag == QN("p"):
                text = "".join(t.text for t in block.iter(QN("t")) if t.text)
                if text:
                    blocks.append(text)
                continue
            if block.tag != QN("tbl"):
                continue
            table_rows = []
            for row in block.findall(QN("tr")):
                cells = []
                for cell in row.findall(QN("tc")):
                    cells.append("".join(t.text for t in cell.iter(QN("t")) if t.text))
                if cells:
                    table_rows.append(cells)
            if table_rows:
                blocks.append("| " + " | ".join(table_rows[0]) + " |")
                blocks.append("| " + " | ".join(["---"] * len(table_rows[0])) + " |")
                blocks.extend("| " + " | ".join(row) + " |" for row in table_rows[1:])
        return "\n".join(blocks)

    def reference(self, kind: str, note_id: str) -> str:
        """本文内の参照を出現順のMarkdown脚注マーカーに変換する。"""
        if not self.notes.get(kind, {}).get(note_id, "").strip():
            logger.warning(
                "%s note id=%sの本文が見つからないため参照マーカーを省略します",
                kind,
                note_id,
            )
            return ""
        key = (kind, note_id)
        if key not in self.references:
            self.references[key] = len([item for item in self.references if item[0] == kind]) + 1
        number = self.references[key]
        return f"[^{kind}{number}]"

    def definitions(self, text_only=False, renderer=None) -> list[str]:
        """本文参照された脚注の定義をMarkdown行として返す。"""
        lines = []
        for (kind, note_id), number in self.references.items():
            text = self.notes.get(kind, {}).get(note_id, "")
            if not text:
                logger.warning("%s note id=%sの本文が見つからないため定義を省略します", kind, note_id)
                continue
            if renderer is not None and note_id in self.note_nodes.get(kind, {}):
                text = renderer(self.note_nodes[kind][note_id])
            if not text:
                logger.warning("%s note id=%sの変換結果が空のため定義を省略します", kind, note_id)
                continue
            marker = f"{kind}{number}"
            if text_only:
                label = "脚注" if kind == "fn" else "文末注"
                lines.append(f"{label}{number}: {text}")
            else:
                lines.append(f"[^{marker}]: {text}")
        return lines
