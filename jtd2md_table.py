"""
一太郎テーブル構造解析モジュール

DocumentTextストリーム内のフォーマットブロック(001C 0030)から
テーブルセルの位置情報を読み取り、Markdownテーブルに変換する

テーブル構造:
- 001C 0030: セル開始（カラム位置・行情報を含む）
- 001C 0000: 行末マーカー（ROW_END）
- 001C 0001: 行開始マーカー（ROW_START, 複数行セルの継続行）
- 001C 0010: 段落書式（非テーブル段落）
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
    __slots__ = ('kind', 'offset', 'cell', 'text')

    def __init__(self, kind: str, offset: int,
                 cell: Optional[TableCell] = None,
                 text: str = ""):
        self.kind = kind      # 'CELL', 'PARA', 'ROW_END', 'ROW_START',
                               # 'NEWLINE', 'SECTION_END'
        self.offset = offset
        self.cell = cell
        self.text = text


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
        # 制御コードで停止
        if code < 0x0020 and code != 0x0009:
            break
        # サロゲートペア処理
        if 0xD800 <= code <= 0xDBFF and i + 3 < total:
            lo = _read_u16be(data, i + 2)
            if 0xDC00 <= lo <= 0xDFFF:
                full_cp = 0x10000 + ((code - 0xD800) << 10) + (lo - 0xDC00)
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
    # 001Fまでスキップ
    after = _skip_to_marker(data, pos + 2, 0x001F)
    return cell, after


def scan_stream_events(data: bytes,
                       content_start: int) -> list[_StreamEvent]:
    """ストリームを走査してイベント列を構築する

    テーブル解析のために、フォーマットブロックの種類と
    テキスト内容をイベントとして収集する
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
                # テーブルセル
                cell, after = parse_cell_block(data, i)
                text, text_end = _read_text_at(data, after)
                cell.text = text.strip()
                events.append(_StreamEvent(
                    'CELL', i, cell=cell, text=text.strip()))
                i = after
                continue

            elif block_type == 0x0010:
                # 段落書式
                after = _skip_to_marker(data, i + 2, 0x001F)
                text, text_end = _read_text_at(data, after)
                events.append(_StreamEvent(
                    'PARA', i, text=text.strip()))
                i = after
                continue

            elif block_type == 0x0000:
                # 行末(ROW_END)
                after = _skip_to_marker(data, i + 2, 0x001F)
                events.append(_StreamEvent('ROW_END', i))
                i = after
                continue

            elif block_type == 0x0001:
                # 行開始(ROW_START, 継続行)
                after = _skip_to_marker(data, i + 2, 0x001F)
                events.append(_StreamEvent('ROW_START', i))
                i = after
                continue

            else:
                # その他のフォーマットブロック
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


def _build_column_map(events: list[_StreamEvent]) -> list[tuple[int, int]]:
    """イベント列からユニークなカラム位置をソートして返す

    全テーブル範囲で一貫したカラムインデックスを付与するために
    ユニークなカラム範囲を収集してソートする
    """
    col_set = set()
    for ev in events:
        if ev.kind == 'CELL' and ev.cell:
            col_set.add((ev.cell.col_start, ev.cell.col_end))
    return sorted(col_set, key=lambda c: c[0])


def _find_col_index(col_map: list[tuple[int, int]],
                    col_start: int, col_end: int) -> int:
    """カラムマップからカラムインデックスを返す

    完全一致しない場合（結合セル等）は、最も近い開始位置を使用
    """
    for idx, (cs, ce) in enumerate(col_map):
        if cs == col_start:
            return idx
    # フォールバック: 開始位置で最近傍を探す
    best = 0
    best_dist = abs(col_map[0][0] - col_start) if col_map else 999999
    for idx, (cs, ce) in enumerate(col_map):
        dist = abs(cs - col_start)
        if dist < best_dist:
            best = idx
            best_dist = dist
    return best


def _find_col_span(col_map: list[tuple[int, int]],
                   col_start: int, col_end: int) -> int:
    """セルのカラムスパン数を計算する（結合セル対応）"""
    start_idx = _find_col_index(col_map, col_start, col_end)
    span = 1
    for idx in range(start_idx + 1, len(col_map)):
        if col_map[idx][0] < col_end:
            span += 1
        else:
            break
    return span


def extract_tables_from_events(
    events: list[_StreamEvent],
) -> list[dict]:
    """イベント列からテーブルデータを抽出する

    Returns:
        テーブル情報のリスト。各テーブルは:
        - 'col_map': カラム位置リスト
        - 'rows': 行リスト。各行はセルリスト
        - 'event_range': (開始イベントidx, 終了イベントidx)
    """
    tables = []

    # テーブル領域を検出（連続するCELLイベントの塊）
    table_ranges = []
    i = 0
    n = len(events)
    while i < n:
        # CELLイベントを探す
        if events[i].kind == 'CELL':
            start_idx = i
            # CELLが途切れるまで進む（PARA, ROW_END等は許容）
            j = i
            last_cell_idx = i
            while j < n:
                if events[j].kind == 'CELL':
                    last_cell_idx = j
                elif events[j].kind == 'PARA':
                    # PARAの後にCELLが来なければテーブル終了
                    k = j + 1
                    has_more_cells = False
                    # 次の3イベント内にCELLがあるかチェック
                    while k < min(j + 5, n):
                        if events[k].kind == 'CELL':
                            has_more_cells = True
                            break
                        k += 1
                    if not has_more_cells:
                        break
                j += 1
            table_ranges.append((start_idx, last_cell_idx + 1))
            i = last_cell_idx + 1
        else:
            i += 1

    for tbl_start, tbl_end in table_ranges:
        tbl_events = events[tbl_start:tbl_end]
        col_map = _build_column_map(tbl_events)
        if not col_map:
            continue

        num_cols = len(col_map)
        rows = []
        current_row = [None] * num_cols

        for ev in tbl_events:
            if ev.kind == 'CELL' and ev.cell:
                cell = ev.cell
                col_idx = _find_col_index(
                    col_map, cell.col_start, cell.col_end)
                if current_row[col_idx] is not None:
                    # 同じカラムにすでにデータ → テキスト追記
                    current_row[col_idx].append_text(cell.text)
                else:
                    current_row[col_idx] = cell

            elif ev.kind == 'SECTION_END':
                # 行を確定
                if any(c is not None for c in current_row):
                    rows.append(current_row)
                    current_row = [None] * num_cols

            elif ev.kind == 'ROW_START':
                # 複数行セルの継続
                # ROW_STARTの後のCELLは同じ論理行に追加
                pass

            elif ev.kind == 'NEWLINE':
                # セル内改行（テキスト追記として処理済み）
                pass

        # 最後の行が残っていれば追加
        if any(c is not None for c in current_row):
            rows.append(current_row)

        if rows:
            tables.append({
                'col_map': col_map,
                'rows': rows,
                'event_range': (tbl_start, tbl_end),
                'num_cols': num_cols,
            })

    return tables


def _merge_continuation_rows(rows: list[list], num_cols: int) -> list[list]:
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

        # 前の行にデータがあるカラムに、現在行にも新たなデータがある場合は新規行
        for ci in range(num_cols):
            if prev[ci] is not None and row[ci] is not None:
                prev_text = prev[ci].text.strip() if prev[ci] else ""
                curr_text = row[ci].text.strip() if row[ci] else ""
                if prev_text and curr_text:
                    is_continuation = False
                    break

        if is_continuation:
            # 前の行にテキストを追記
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

    # 継続行を統合
    merged = _merge_continuation_rows(rows, num_cols)

    if not merged:
        return []

    md_lines = []

    for row_idx, row in enumerate(merged):
        cells_text = []
        for ci in range(num_cols):
            cell = row[ci]
            if cell is not None:
                # セル内テキストを整形（改行をスペースに）
                text = cell.text.strip()
                # Markdownテーブルのパイプをエスケープ
                text = text.replace('|', '\\|')
                cells_text.append(text)
            else:
                cells_text.append("")

        md_lines.append("| " + " | ".join(cells_text) + " |")

        # ヘッダ行の後にセパレータを挿入
        if row_idx == 0:
            sep = "| " + " | ".join(
                "---" for _ in range(num_cols)) + " |"
            md_lines.append(sep)

    return md_lines
