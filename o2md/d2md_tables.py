"""Word表をOOXMLのgrid定義からMarkdownへ変換する。"""

from dataclasses import dataclass

from docx.table import _Cell


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def qn(name: str) -> str:
    """WordprocessingML要素名を生成する。"""
    return f"{{{W}}}{name}"


@dataclass
class Slot:
    """表のgrid上のセル位置を表す。"""

    cell: object | None = None
    covered: bool = False
    covered_by: object | None = None


def _int_value(element, name: str, default: int = 0) -> int:
    """Word属性を整数として取得する。"""
    if element is None:
        return default
    try:
        return int(element.get(qn(name), str(default)))
    except (TypeError, ValueError):
        return default


def _cell_span(tc_pr) -> int:
    """横結合幅を取得する。"""
    if tc_pr is None:
        return 1
    return max(1, _int_value(tc_pr.find(qn("gridSpan")), "val", 1))


def render_table(table, process_cell):
    """結合セルのorigin値をcovered slotにも繰り返したMarkdown表を返す。"""
    tbl = table._element
    grid = tbl.find(qn("tblGrid"))
    if grid is not None:
        columns = len(grid.findall(qn("gridCol")))
    else:
        columns = max((len(tr.tc_lst) for tr in tbl.tr_lst), default=0)
    if columns == 0:
        return []

    rows: list[list[Slot]] = []
    header_flags: list[bool] = []
    active_vertical: list[object | None] = [None] * columns
    for tr in tbl.tr_lst:
        tr_pr = tr.trPr
        before = _int_value(tr_pr.find(qn("gridBefore")) if tr_pr is not None else None, "val")
        after = _int_value(tr_pr.find(qn("gridAfter")) if tr_pr is not None else None, "val")
        slots = [Slot(covered=True) for _ in range(columns)]
        next_vertical: list[object | None] = [None] * columns
        cursor = before
        hmerge_spans = {}
        hmerge_continuations = set()
        for index, tc in enumerate(tr.tc_lst):
            tc_pr = tc.tcPr
            merge = tc_pr.find(qn("hMerge")) if tc_pr is not None else None
            if merge is None or merge.get(qn("val"), "continue") != "restart":
                continue
            span = _cell_span(tc_pr)
            following = index + 1
            while following < len(tr.tc_lst):
                next_pr = tr.tc_lst[following].tcPr
                next_merge = next_pr.find(qn("hMerge")) if next_pr is not None else None
                if next_merge is None or next_merge.get(qn("val"), "continue") == "restart":
                    break
                span += _cell_span(next_pr)
                hmerge_continuations.add(following)
                following += 1
            hmerge_spans[index] = span
        for index, tc in enumerate(tr.tc_lst):
            tc_pr = tc.tcPr
            merge = tc_pr.find(qn("vMerge")) if tc_pr is not None else None
            horizontal = tc_pr.find(qn("hMerge")) if tc_pr is not None else None
            if index in hmerge_continuations:
                continue
            span = hmerge_spans.get(index, _cell_span(tc_pr))
            if horizontal is not None and horizontal.get(qn("val"), "continue") != "restart":
                span = _cell_span(tc_pr)
                cell = None
            else:
                cell = _Cell(tc, table)
            is_vertical_continuation = (
                merge is not None and merge.get(qn("val"), "continue") != "restart"
            )
            origin = (
                active_vertical[cursor]
                if is_vertical_continuation and cursor < columns
                else None
            )
            required = cursor + span
            for slot_index in range(cursor, min(required, len(slots))):
                covered_by = origin if is_vertical_continuation else cell
                slots[slot_index] = Slot(
                    cell=cell if slot_index == cursor and not is_vertical_continuation else None,
                    covered=slot_index != cursor or is_vertical_continuation,
                    covered_by=covered_by if slot_index != cursor or is_vertical_continuation else None,
                )
                if merge is not None:
                    next_vertical[slot_index] = cell if not is_vertical_continuation else origin
            cursor += span
        rows.append(slots)
        active_vertical = next_vertical
        header_flags.append(
            tr_pr is not None and tr_pr.find(qn("tblHeader")) is not None
        )

    def row_text(row: list[Slot]) -> list[str]:
        return [
            process_cell(slot.cell or slot.covered_by)
            if slot.cell is not None or slot.covered_by is not None
            else ""
            for slot in row
        ]

    header_count = 0
    for is_header in header_flags:
        if not is_header:
            break
        header_count += 1
    rendered_rows = [row_text(row) for row in rows]
    if header_count > 1:
        def header_cell_text(row: int, column: int) -> str:
            return rendered_rows[row][column]

        header = [
            " ".join(
                text
                for row in range(header_count)
                for text in [header_cell_text(row, column)]
                if text
            ).strip()
            for column in range(columns)
        ]
        rendered_rows = [header] + rendered_rows[header_count:]
    if not rendered_rows:
        return []

    lines = [
        "| " + " | ".join(rendered_rows[0]) + " |",
        "| " + " | ".join(["---"] * columns) + " |",
    ]
    for row in rendered_rows[1:]:
        values = row[:columns] + [""] * max(0, columns - len(row))
        lines.append("| " + " | ".join(values) + " |")
    return lines
