"""
一太郎テーブル構造解析モジュール

DocumentTextストリーム内のフォーマットブロックから
罫線情報（008Fタグ）を読み取り、テーブル行を判定する。

テーブル判定基準:
- PARAブロック(001C 0010)内の008Fタグに罫線データ(001B/0013)が≥1個
  → テーブル行（カラム分割の罫線あり）
- 008Fタグなし or 罫線データ=0
  → 通常テキスト（枠線のみ、またはテーブル外）

テーブル構造:
- 001C 0030: セル開始（カラム位置・行情報を含む）
- 001C 0000: 行末マーカー（ROW_END）
- 001C 0001: 行開始マーカー（ROW_START, 複数行セルの継続行）
- 001C 0010: 段落書式（罫線情報を含む）
- 000E: セクション終了
"""

from typing import Optional


def _read_u16be(data: bytes, pos: int) -> int:
    """バイト列からUTF-16BEの1コードユニットを読み取る"""
    if pos + 1 >= len(data):
        return -1
    return (data[pos] << 8) | data[pos + 1]


def _safe_chr(code: int) -> str:
    """Unicodeコードポイントを安全に文字に変換する"""
    if 0xD800 <= code <= 0xDFFF:
        return ''
    try:
        return chr(code)
    except (ValueError, OverflowError):
        return ''


# 罫線タイプ識別子
_RULER_TAGS = frozenset((0x001b, 0x0013))

# フォントサイズタグ
_TAG_FONT_SIZE = 0x0008


def extract_font_size_from_para(data: bytes, pos: int,
                               block_len: int) -> int:
    """PARAブロックからフォントサイズを抽出する

    TAG 0008の最初の値をフォントサイズ(1/100pt単位)として返す。
    フォントサイズタグが見つからない場合は0を返す（デフォルトサイズ使用）。
    """
    block = data[pos:pos + block_len]
    for k in range(8, len(block) - 3, 2):
        tag = _read_u16be(block, k)
        if tag == _TAG_FONT_SIZE:
            return _read_u16be(block, k + 2)
        if tag == 0xFFFF or tag == 0x001F:
            break
    return 0


def count_rulers_in_para(data: bytes, pos: int,
                         block_len: int) -> int:
    """PARAブロック(001C 0010)から罫線の本数を数える

    008Fタグ内のデータに含まれる罫線識別子(001B/0013)の数を返す。
    罫線≥1 → テーブル行（カラム分割線あり）
    罫線=0 → テーブル外（枠線のみ or 非テーブル）
    """
    block = data[pos:pos + block_len]
    # 008Fタグを探す
    for k in range(0, len(block) - 3, 2):
        if _read_u16be(block, k) == 0x008f:
            data_len = _read_u16be(block, k + 2)
            if data_len <= 0:
                return 0
            # 008Fデータ内の罫線識別子をカウント
            ruler_data_start = k + 4
            ruler_data_end = min(
                ruler_data_start + data_len * 2,
                len(block)
            )
            count = 0
            for m in range(ruler_data_start, ruler_data_end - 1, 2):
                val = _read_u16be(block, m)
                if val in _RULER_TAGS:
                    count += 1
            return count
    return 0


class TableCell:
    """テーブルセルの情報"""
    __slots__ = ('col_start', 'col_end', 'row_flag', 'text')

    def __init__(self, col_start: int, col_end: int, row_flag: int):
        self.col_start = col_start
        self.col_end = col_end
        self.row_flag = row_flag
        self.text = ""

    def append_text(self, text: str):
        """セルにテキストを追加する"""
        if self.text:
            self.text += " "
        self.text += text

    def __repr__(self):
        return (
            f"Cell(col={self.col_start:04x}-{self.col_end:04x},"
            f" row={self.row_flag:04x}, text='{self.text[:20]}')"
        )


class _StreamEvent:
    """ストリーム解析で検出されるイベント"""
    __slots__ = (
        'kind', 'offset', 'cell', 'text',
        'ruler_count', 'font_size',
    )

    def __init__(self, kind: str, offset: int,
                 cell: Optional[TableCell] = None,
                 text: str = "",
                 ruler_count: int = 0,
                 font_size: int = 0):
        self.kind = kind
        self.offset = offset
        self.cell = cell
        self.text = text
        self.ruler_count = ruler_count
        self.font_size = font_size


def _read_text_at(data: bytes, pos: int,
                  limit: int = 500) -> tuple[str, int]:
    """指定位置からUTF-16BEテキストを読み取る

    制御コード(< 0x0020, ただしタブは除く)で停止する。
    Returns: (抽出テキスト, 読み取り終了位置)
    """
    chars = []
    i = pos
    total = len(data)
    count = 0
    while i < total - 1 and count < limit:
        code = _read_u16be(data, i)
        if code < 0:
            break
        if code < 0x0020 and code != 0x0009:
            break
        # サロゲートペア処理
        if 0xD800 <= code <= 0xDBFF and i + 3 < total:
            lo = _read_u16be(data, i + 2)
            if 0xDC00 <= lo <= 0xDFFF:
                full_cp = (
                    0x10000
                    + ((code - 0xD800) << 10)
                    + (lo - 0xDC00)
                )
                ch = _safe_chr(full_cp)
                if ch:
                    chars.append(ch)
                i += 4
                count += 1
                continue
        if 0xDC00 <= code <= 0xDFFF:
            i += 2
            count += 1
            continue
        ch = _safe_chr(code)
        if ch:
            chars.append(ch)
        i += 2
        count += 1
    return ''.join(chars), i


def _skip_to_marker(data: bytes, pos: int,
                    marker: int = 0x001F) -> int:
    """指定マーカーコードまでスキップし、その次の位置を返す"""
    total = len(data)
    i = pos
    while i < total - 1:
        if _read_u16be(data, i) == marker:
            return i + 2
        i += 2
    return total


def parse_cell_block(data: bytes, pos: int) -> tuple[TableCell, int]:
    """001C 0030セルブロックを解析する

    ブロック構造(24バイト固定):
    001C 0030 000C 0000 [col_start 2B] [col_end 2B]
    [row_flag 2B] 0000 000C 0000 0030 001F

    Returns: (TableCell, 001F直後の位置)
    """
    col_start = _read_u16be(data, pos + 8)
    col_end = _read_u16be(data, pos + 10)
    row_flag = _read_u16be(data, pos + 12)
    cell = TableCell(col_start, col_end, row_flag)
    after = _skip_to_marker(data, pos + 2, 0x001F)
    return cell, after


def scan_stream_events(data: bytes,
                       content_start: int) -> list[_StreamEvent]:
    """ストリームを走査してイベント列を構築する

    PARAブロックの罫線情報を含めてイベントを収集する
    """
    events = []
    total = len(data)
    i = content_start

    while i < total - 1:
        code = _read_u16be(data, i)
        if code < 0:
            break

        if code == 0x001C:
            block_type = _read_u16be(data, i + 2)

            if block_type == 0x0030:
                cell, after = parse_cell_block(data, i)
                text, text_end = _read_text_at(data, after)
                cell.text = text.strip()
                events.append(_StreamEvent(
                    'CELL', i, cell=cell, text=text.strip()))
                i = after
                continue

            elif block_type == 0x0010:
                # 001Fまでの位置を取得
                j = i + 2
                while j < total - 1:
                    if _read_u16be(data, j) == 0x001F:
                        break
                    j += 2
                block_len = j - i + 2
                # 罫線数とフォントサイズを取得
                rulers = count_rulers_in_para(data, i, block_len)
                fsize = extract_font_size_from_para(
                    data, i, block_len)
                after = j + 2
                text, text_end = _read_text_at(data, after)
                events.append(_StreamEvent(
                    'PARA', i, text=text.strip(),
                    ruler_count=rulers,
                    font_size=fsize))
                i = after
                continue

            elif block_type == 0x0000:
                after = _skip_to_marker(data, i + 2, 0x001F)
                events.append(_StreamEvent('ROW_END', i))
                i = after
                continue

            elif block_type == 0x0001:
                after = _skip_to_marker(data, i + 2, 0x001F)
                events.append(_StreamEvent('ROW_START', i))
                i = after
                continue

            else:
                after = _skip_to_marker(data, i + 2, 0x001F)
                i = after
                continue

        elif code == 0x000A:
            events.append(_StreamEvent('NEWLINE', i))
            i += 2
            continue

        elif code == 0x000E:
            events.append(_StreamEvent('SECTION_END', i))
            i += 2
            continue

        else:
            i += 2

    return events


def _split_sections(events: list[_StreamEvent]) -> list[list[_StreamEvent]]:
    """イベント列をセクション(SECTION_END区切り)に分割する"""
    sections = []
    current = []
    for ev in events:
        current.append(ev)
        if ev.kind == 'SECTION_END':
            sections.append(current)
            current = []
    if current:
        sections.append(current)
    return sections


def _is_table_section(section: list[_StreamEvent]) -> bool:
    """セクションがテーブル行かどうかを罫線情報で判定する

    PARAブロックに罫線≥1がある場合にテーブル行とみなす
    """
    for ev in section:
        if ev.kind == 'PARA' and ev.ruler_count >= 1:
            return True
    return False


def _build_column_map(
    table_sections: list[list[_StreamEvent]],
) -> list[tuple[int, int]]:
    """テーブルセクション群からユニークなカラム位置をソートして返す"""
    col_set = set()
    for section in table_sections:
        for ev in section:
            if ev.kind == 'CELL' and ev.cell:
                col_set.add((ev.cell.col_start, ev.cell.col_end))
    return sorted(col_set, key=lambda c: c[0])


def _find_col_index(col_map: list[tuple[int, int]],
                    col_start: int, col_end: int) -> int:
    """カラムマップからカラムインデックスを返す"""
    for idx, (cs, ce) in enumerate(col_map):
        if cs == col_start:
            return idx
    # フォールバック: 開始位置で最近傍
    best = 0
    best_dist = abs(col_map[0][0] - col_start) if col_map else 999999
    for idx, (cs, ce) in enumerate(col_map):
        dist = abs(cs - col_start)
        if dist < best_dist:
            best = idx
            best_dist = dist
    return best


def extract_tables_from_events(
    events: list[_StreamEvent],
) -> list[dict]:
    """イベント列からテーブルデータを抽出する（罫線ベース判定）

    Returns:
        テーブル情報のリスト。各テーブルは:
        - 'col_map': カラム位置リスト
        - 'rows': 行リスト。各行はセルリスト
        - 'section_range': (開始セクションidx, 終了セクションidx)
        - 'event_range': (開始イベントidx, 終了イベントidx)
    """
    sections = _split_sections(events)
    tables = []

    # 連続するテーブルセクションをグループ化
    table_groups = []
    current_group = []
    current_start_idx = -1

    for sec_idx, section in enumerate(sections):
        if _is_table_section(section):
            if not current_group:
                current_start_idx = sec_idx
            current_group.append(section)
        else:
            if current_group:
                table_groups.append(
                    (current_start_idx, sec_idx, current_group))
                current_group = []

    if current_group:
        table_groups.append(
            (current_start_idx, len(sections), current_group))

    # 各グループをテーブルとして処理
    for grp_start, grp_end, grp_sections in table_groups:
        col_map = _build_column_map(grp_sections)
        if not col_map:
            continue

        num_cols = len(col_map)
        rows = []

        for section in grp_sections:
            current_row = [None] * num_cols
            has_cells = False

            for ev in section:
                if ev.kind == 'CELL' and ev.cell:
                    has_cells = True
                    cell = ev.cell
                    col_idx = _find_col_index(
                        col_map, cell.col_start, cell.col_end)
                    if current_row[col_idx] is not None:
                        current_row[col_idx].append_text(cell.text)
                    else:
                        current_row[col_idx] = cell

            if has_cells:
                rows.append(current_row)

        if rows:
            # イベント範囲を計算
            first_ev = sections[grp_start][0]
            last_section = sections[min(grp_end - 1, len(sections) - 1)]
            last_ev = last_section[-1]
            ev_start = events.index(first_ev)
            ev_end = events.index(last_ev) + 1

            tables.append({
                'col_map': col_map,
                'rows': rows,
                'num_cols': num_cols,
                'section_range': (grp_start, grp_end),
                'event_range': (ev_start, ev_end),
            })

    return tables


def _merge_continuation_rows(
    rows: list[list], num_cols: int,
) -> list[list]:
    """継続行を前の行に統合する

    一太郎では1つの論理行が複数の物理行に分割されることがある。
    前の行と同じカラムが空の場合、テキストを前の行に追記する。
    """
    if len(rows) <= 1:
        return rows

    merged = [rows[0]]
    for row in rows[1:]:
        prev = merged[-1]
        is_continuation = True

        for ci in range(num_cols):
            if prev[ci] is not None and row[ci] is not None:
                prev_text = (
                    prev[ci].text.strip() if prev[ci] else "")
                curr_text = (
                    row[ci].text.strip() if row[ci] else "")
                if prev_text and curr_text:
                    is_continuation = False
                    break

        if is_continuation:
            for ci in range(num_cols):
                if row[ci] is not None and row[ci].text.strip():
                    if prev[ci] is None:
                        prev[ci] = row[ci]
                    else:
                        prev[ci].append_text(row[ci].text)
        else:
            merged.append(row)

    return merged


def table_to_markdown(table: dict) -> list[str]:
    """テーブルデータをMarkdownテーブル行に変換する"""
    rows = table['rows']
    num_cols = table['num_cols']

    merged = _merge_continuation_rows(rows, num_cols)

    if not merged:
        return []

    md_lines = []

    for row_idx, row in enumerate(merged):
        cells_text = []
        for ci in range(num_cols):
            cell = row[ci]
            if cell is not None:
                text = cell.text.strip()
                text = text.replace('|', '\\|')
                cells_text.append(text)
            else:
                cells_text.append("")

        md_lines.append("| " + " | ".join(cells_text) + " |")

        if row_idx == 0:
            sep = "| " + " | ".join(
                "---" for _ in range(num_cols)) + " |"
            md_lines.append(sep)

    return md_lines
