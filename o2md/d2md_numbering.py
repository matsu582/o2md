"""Word番号付け定義を解析し、Markdown用の番号を生成する。"""

from dataclasses import dataclass
import logging
import re
import xml.etree.ElementTree as ET


logger = logging.getLogger(__name__)
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def qn(name: str) -> str:
    """WordprocessingML要素名を生成する。"""
    return f"{{{W}}}{name}"


QN = qn


@dataclass
class LevelDefinition:
    start: int = 1
    num_fmt: str = "decimal"
    level_text: str = "%1."
    restart: int | None = None
    bullet: bool = False
    suffix: str = "tab"


def _value(element, name: str, default=None):
    if element is None:
        return default
    return element.get(qn(name), default)


def _roman(value: int, upper: bool) -> str:
    result = ""
    for number, text in ((1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
                         (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
                         (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")):
        result += text * (value // number)
        value %= number
    return result if upper else result.lower()


def _letters(value: int, upper: bool) -> str:
    result = ""
    while value:
        value, remainder = divmod(value - 1, 26)
        result = chr((65 if upper else 97) + remainder) + result
    return result or ("A" if upper else "a")


def format_number(value: int, num_fmt: str) -> str:
    """Wordの番号形式を文字列表現へ変換する。"""
    if num_fmt == "decimal":
        return str(value)
    if num_fmt == "upperRoman":
        return _roman(value, True)
    if num_fmt == "lowerRoman":
        return _roman(value, False)
    if num_fmt == "upperLetter":
        return _letters(value, True)
    if num_fmt == "lowerLetter":
        return _letters(value, False)
    if num_fmt == "decimalEnclosedCircle":
        return chr(0x2460 + value - 1) if 1 <= value <= 20 else str(value)
    if num_fmt == "decimalFullWidth":
        return str(value).translate(str.maketrans("0123456789", "０１２３４５６７８９"))
    if num_fmt in ("aiueo", "aiueoFullWidth"):
        return "あいうえおかきくけこさしすせそたちつてと"[value - 1:value] or str(value)
    if num_fmt in ("iroha", "irohaFullWidth"):
        return "いろはにほへとちりぬるをわかよたれそつねならむうゐのおくやま"[value - 1:value] or str(value)
    if num_fmt == "bullet":
        return "•"
    if num_fmt == "none":
        return ""
    logger.warning("未知のWord番号形式 '%s'。decimalにフォールバックします", num_fmt)
    return str(value)


class NumberingResolver:
    """numbering.xmlの定義と段落ごとの番号状態を保持する。"""

    def __init__(self, blob: bytes | None, style_root=None):
        self.valid = False
        self.levels = {}
        self.instances = {}
        self.counters = {}
        self.restart_seen = {}
        self.style_root = style_root
        if blob:
            try:
                self._parse(blob)
                self.valid = True
            except (ET.ParseError, ValueError, TypeError) as exc:
                logger.warning("numbering.xmlを読み込めません。従来の判定へ退避します: %s", exc)

    def _parse(self, blob: bytes):
        root = ET.fromstring(blob)
        abstracts = {}
        for abstract in root.findall(f".//{QN('abstractNum')}"):
            abstract_id = _value(abstract, "abstractNumId")
            levels = {}
            for level in abstract.findall(QN("lvl")):
                ilvl = int(_value(level, "ilvl", "0"))
                fmt = level.find(QN("numFmt"))
                text = level.find(QN("lvlText"))
                start = level.find(QN("start"))
                restart = level.find(QN("lvlRestart"))
                levels[ilvl] = LevelDefinition(
                    start=int(_value(start, "val", "1")),
                    num_fmt=_value(fmt, "val", "decimal"),
                    level_text=_value(text, "val", f"%{ilvl + 1}."),
                    restart=int(_value(restart, "val")) if restart is not None else None,
                    bullet=_value(fmt, "val") in ("bullet", "none"),
                    suffix=_value(level.find(QN("suff")), "val", "tab"),
                )
            abstracts[abstract_id] = levels
        for num in root.findall(f".//{QN('num')}"):
            num_id = _value(num, "numId")
            abstract = num.find(QN("abstractNumId"))
            if abstract is None:
                continue
            abstract_id = _value(abstract, "val")
            levels = {key: LevelDefinition(**vars(value)) for key, value in abstracts.get(abstract_id, {}).items()}
            for override in num.findall(QN("lvlOverride")):
                ilvl = int(_value(override, "ilvl", "0"))
                level = override.find(QN("lvl"))
                if level is not None:
                    levels[ilvl] = self._level_from_xml(level, levels.get(ilvl))
                start_override = override.find(QN("startOverride"))
                if start_override is not None:
                    levels.setdefault(ilvl, LevelDefinition()).start = int(_value(start_override, "val", "1"))
            self.instances[num_id] = levels
        self.levels = abstracts
        if not self.instances:
            raise ValueError("num定義がありません")

    def _level_from_xml(self, element, fallback=None):
        base = fallback or LevelDefinition()
        return LevelDefinition(
            start=int(_value(element.find(QN("start")), "val", str(base.start))),
            num_fmt=_value(element.find(QN("numFmt")), "val", base.num_fmt),
            level_text=_value(element.find(QN("lvlText")), "val", base.level_text),
            restart=int(_value(element.find(QN("lvlRestart")), "val")) if element.find(QN("lvlRestart")) is not None else base.restart,
            bullet=_value(element.find(QN("numFmt")), "val", base.num_fmt) in ("bullet", "none"),
            suffix=_value(element.find(QN("suff")), "val", base.suffix),
        )

    def _paragraph_num(self, paragraph):
        num_pr = paragraph._element.find(f".//{QN('numPr')}")
        if num_pr is None:
            style = getattr(paragraph, "style", None)
            visited = set()
            while style is not None and id(style) not in visited:
                visited.add(id(style))
                style_element = getattr(style, "_element", None)
                num_pr = style_element.find(f".//{QN('numPr')}") if style_element is not None else None
                if num_pr is not None:
                    break
                style = getattr(style, "base_style", None)
        if num_pr is None:
            return None
        num_id = num_pr.find(QN("numId"))
        ilvl = num_pr.find(QN("ilvl"))
        if num_id is None:
            return None
        return _value(num_id, "val"), int(_value(ilvl, "val", "0"))

    def marker(self, paragraph):
        """段落に対応するMarkdownリストマーカーを返す。"""
        selected = self._paragraph_num(paragraph)
        if selected is None:
            return None
        num_id, ilvl = selected
        levels = self.instances.get(num_id)
        if levels is None:
            return None
        definition = levels.get(ilvl, LevelDefinition(level_text=f"%{ilvl + 1}."))
        state = self.counters.setdefault(num_id, [0] * 9)
        seen = self.restart_seen.setdefault(num_id, [None] * 9)
        if definition.restart is not None and definition.restart < ilvl:
            higher_count = state[definition.restart]
            if seen[ilvl] is not None and seen[ilvl] != higher_count:
                state[ilvl] = 0
            seen[ilvl] = higher_count
        for index in range(ilvl + 1, 9):
            state[index] = 0
        state[ilvl] = max(definition.start, state[ilvl] + 1)
        values = {
            index + 1: format_number(
                state[index] or levels.get(index, definition).start,
                levels.get(index, definition).num_fmt,
            )
            for index in range(9)
        }
        label = re.sub(
            r"%([1-9])", lambda match: values[int(match.group(1))], definition.level_text
        )
        if definition.bullet or definition.num_fmt in ("bullet", "none"):
            return ilvl, "-", True
        placeholders = re.findall(r"%([1-9])", definition.level_text)
        if not placeholders:
            return ilvl, "-", True
        if len(placeholders) == 1 and re.fullmatch(
            r"%[1-9][\W_]*", definition.level_text
        ):
            number = state[int(placeholders[0]) - 1] or levels.get(
                int(placeholders[0]) - 1, definition
            ).start
            return ilvl, f"{number}.", False
        return ilvl, label, False
