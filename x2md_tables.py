#!/usr/bin/env python3
"""
テーブル処理Mixinモジュール

ExcelToMarkdownConverterクラスのテーブル検出・構築・出力機能を提供します。
このモジュールはMixinクラスとして設計されており、メインクラスから継承されます。

機能:
- テーブル領域の検出と境界判定
- テーブルデータの構築と結合セル処理
- Markdown形式でのテーブル出力
"""

from typing import List, Dict, Tuple, Optional, Any, Set

# デバッグ_printはx2mdモジュールからインポート
# 注意: 循環インポートを避けるため、関数レベルでインポートするか、
# または遅延インポートを使用する必要がある場合があります
def _get_debug_print():
    """debug_print関数を取得（循環インポート回避）"""
    from x2md import debug_print
    return debug_print

# モジュールレベルでdebug_printを定義（遅延評価）
def debug_print(*args, **kwargs):
    """デバッグ出力（x2mdモジュールに委譲）"""
    try:
        from x2md import debug_print as _dp
        _dp(*args, **kwargs)
    except ImportError:
        pass  # インポートエラー時は無視


class _TablesMixin:
    """テーブル処理機能を提供するMixinクラス
    
    このクラスはExcelToMarkdownConverterに継承され、
    テーブルの検出、構築、出力機能を提供します。
    
    注意: このクラスは単独では使用できません。
    ExcelToMarkdownConverterクラスと組み合わせて使用してください。
    """

    def _detect_bordered_tables(self, sheet, min_row, max_row, min_col, max_col):
        """外枠罫線のみで囲まれた最大矩形をテーブルと判定（内部罫線は無視）"""
        tables = []
        
        visited = set()
        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                if (row, col) in visited:
                    continue
                
                cell = sheet.cell(row=row, column=col)
                if (cell.border and cell.border.left and cell.border.left.style and
                    cell.border and cell.border.top and cell.border.top.style):
                    region = self._find_bordered_region(sheet, row, col, min_row, max_row, min_col, max_col, visited)
                    if region:
                        r1, r2, c1, c2 = region
                        if r1 == 1 and r2 <= 4:
                            debug_print(f"[DEBUG][{sheet.title}] Top region detected: rows {r1}-{r2}, cols {c1}-{c2}")
                        if r2 - r1 >= 1 or c2 - c1 >= 2:
                            is_valid = self._is_valid_bordered_table(sheet, region)
                            if r1 == 1 and r2 <= 4:
                                debug_print(f"[DEBUG][{sheet.title}] Top region validation: {is_valid}")
                            if is_valid:
                                tables.append(region)
        
        debug_print(f"[DEBUG] _detect_bordered_tables found {len(tables)} tables")
        return tables

    def _find_discrete_data_regions(self, sheet, min_row: int, max_row: int, min_col: int, max_col: int, occupied_cells: Optional[Set[Tuple[int, int]]] = None) -> List[Tuple[int, int, int, int]]:
        """空白行/列で区切られた離散データ領域を検出する
        
        doclingの実装を参考に、シート内の非空セルをスキャンし、
        空白行/列で区切られた独立したデータ領域を個別のテーブルとして検出します。
        
        Args:
            sheet: ワークシートオブジェクト
            min_row: スキャン開始行
            max_row: スキャン終了行
            min_col: スキャン開始列
            max_col: スキャン終了列
            occupied_cells: 既に罫線テーブルで占有されているセルのセット（除外対象）
            
        Returns:
            検出された離散データ領域のリスト [(start_row, end_row, start_col, end_col), ...]
        """
        debug_print(f"[DEBUG][_find_discrete_data_regions] sheet={sheet.title} range=({min_row}-{max_row}, {min_col}-{max_col}) occupied_cells_count={len(occupied_cells) if occupied_cells else 0}")
        
        tables: List[Tuple[int, int, int, int]] = []
        visited: Set[Tuple[int, int]] = set()
        
        # 占有セルがある場合は訪問済みとしてマーク
        if occupied_cells:
            visited.update(occupied_cells)
        
        # シート内の非空セルをスキャン
        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                if (row, col) in visited:
                    continue
                
                cell = sheet.cell(row=row, column=col)
                if cell.value is None or str(cell.value).strip() == '':
                    continue
                
                # 新しいテーブル領域の開始点を発見
                table_bounds, visited_cells = self._find_discrete_table_bounds(
                    sheet, row, col, max_row, max_col
                )
                
                if table_bounds:
                    tables.append(table_bounds)
                    visited.update(visited_cells)
        
        debug_print(f"[DEBUG][_find_discrete_data_regions] found {len(tables)} discrete regions: {tables[:5]}")
        return tables

    def _find_discrete_table_bounds(
        self, sheet, start_row: int, start_col: int, max_row: int, max_col: int
    ) -> Tuple[Optional[Tuple[int, int, int, int]], Set[Tuple[int, int]]]:
        """離散テーブルの境界を検出する
        
        開始セルから空白行/列で区切られるまで領域を拡張し、
        テーブルの境界と訪問済みセルのセットを返します。
        
        Args:
            sheet: ワークシートオブジェクト
            start_row: 開始行
            start_col: 開始列
            max_row: 最大行
            max_col: 最大列
            
        Returns:
            (テーブル境界タプル, 訪問済みセルのセット)
        """
        # 下方向に拡張（空白行で停止）
        table_max_row = self._find_discrete_table_bottom(sheet, start_row, start_col, max_row)
        
        # 右方向に拡張（空白列で停止）
        table_max_col = self._find_discrete_table_right(sheet, start_row, start_col, max_col)
        
        # 訪問済みセルを収集
        visited_cells: Set[Tuple[int, int]] = set()
        for row in range(start_row, table_max_row + 1):
            for col in range(start_col, table_max_col + 1):
                visited_cells.add((row, col))
        
        # 結合セルを考慮して境界を拡張
        for merged_range in sheet.merged_cells.ranges:
            if (merged_range.min_row <= table_max_row and merged_range.max_row >= start_row and
                merged_range.min_col <= table_max_col and merged_range.max_col >= start_col):
                table_max_row = max(table_max_row, merged_range.max_row)
                table_max_col = max(table_max_col, merged_range.max_col)
                # 結合セル内のセルも訪問済みとしてマーク
                for r in range(merged_range.min_row, merged_range.max_row + 1):
                    for c in range(merged_range.min_col, merged_range.max_col + 1):
                        visited_cells.add((r, c))
        
        return (start_row, table_max_row, start_col, table_max_col), visited_cells

    def _find_discrete_table_bottom(self, sheet, start_row: int, start_col: int, max_row: int) -> int:
        """離散テーブルの下端を検出する（空白行で停止）
        
        Args:
            sheet: ワークシートオブジェクト
            start_row: 開始行
            start_col: 開始列
            max_row: 最大行
            
        Returns:
            テーブルの下端行番号
        """
        table_max_row = start_row
        
        for row in range(start_row + 1, max_row + 1):
            cell = sheet.cell(row=row, column=start_col)
            
            # 結合セルの一部かどうかをチェック
            merged_range = None
            for mr in sheet.merged_cells.ranges:
                if mr.min_row <= row <= mr.max_row and mr.min_col <= start_col <= mr.max_col:
                    merged_range = mr
                    break
            
            if cell.value is None and not merged_range:
                # 空白セルで結合セルでもない場合、停止
                break
            
            # 結合セルの場合、その範囲の最大行まで拡張
            if merged_range:
                table_max_row = max(table_max_row, merged_range.max_row)
            else:
                table_max_row = row
        
        return table_max_row

    def _find_discrete_table_right(self, sheet, start_row: int, start_col: int, max_col: int) -> int:
        """離散テーブルの右端を検出する（空白列で停止）
        
        Args:
            sheet: ワークシートオブジェクト
            start_row: 開始行
            start_col: 開始列
            max_col: 最大列
            
        Returns:
            テーブルの右端列番号
        """
        table_max_col = start_col
        
        for col in range(start_col + 1, max_col + 1):
            cell = sheet.cell(row=start_row, column=col)
            
            # 結合セルの一部かどうかをチェック
            merged_range = None
            for mr in sheet.merged_cells.ranges:
                if mr.min_row <= start_row <= mr.max_row and mr.min_col <= col <= mr.max_col:
                    merged_range = mr
                    break
            
            if cell.value is None and not merged_range:
                # 空白セルで結合セルでもない場合、停止
                break
            
            # 結合セルの場合、その範囲の最大列まで拡張
            if merged_range:
                table_max_col = max(table_max_col, merged_range.max_col)
            else:
                table_max_col = col
        
        return table_max_col
    
    def _is_valid_bordered_table(self, sheet, region):
        """罫線テーブルが有効かどうかをチェック（空行・空列が多すぎる場合は無効）"""
        r1, r2, c1, c2 = region
        total_rows = r2 - r1 + 1
        total_cols = c2 - c1 + 1
        empty_rows = 0
        empty_cols = 0
        
        cols_with_data = 0
        for col in range(c1, c2 + 1):
            value_count = 0
            for row in range(r1, r2 + 1):
                cell = sheet.cell(row=row, column=col)
                if cell.value and str(cell.value).strip():
                    value_count += 1
                    if value_count >= 2:
                        cols_with_data += 1
                        break
        
        # 結合セルを検出
        merged_cols = set()
        merged_rows = set()
        for merged_range in sheet.merged_cells.ranges:
            if (merged_range.min_row <= r2 and merged_range.max_row >= r1 and
                merged_range.min_col <= c2 and merged_range.max_col >= c1):
                for col in range(max(merged_range.min_col, c1), min(merged_range.max_col, c2) + 1):
                    merged_cols.add(col)
                for row in range(max(merged_range.min_row, r1), min(merged_range.max_row, r2) + 1):
                    merged_rows.add(row)
        
        for row in range(r1, r2 + 1):
            is_empty = True
            for col in range(c1, c2 + 1):
                cell = sheet.cell(row=row, column=col)
                if cell.value and str(cell.value).strip():
                    is_empty = False
                    break
            if is_empty:
                empty_rows += 1
        
        for col in range(c1, c2 + 1):
            is_empty = True
            for row in range(r1, r2 + 1):
                cell = sheet.cell(row=row, column=col)
                if cell.value and str(cell.value).strip():
                    is_empty = False
                    break
            if is_empty:
                empty_cols += 1
        
        empty_row_ratio = empty_rows / total_rows if total_rows > 0 else 0
        empty_col_ratio = empty_cols / total_cols if total_cols > 0 else 0
        debug_print(f"[DEBUG] Table region {region}: empty_rows={empty_rows}/{total_rows} (ratio={empty_row_ratio:.2f}), empty_cols={empty_cols}/{total_cols} (ratio={empty_col_ratio:.2f}), cols_with_data={cols_with_data}")
        
        has_many_merged_cells = len(merged_cols) > total_cols * 0.3 or len(merged_rows) > total_rows * 0.3
        is_small_table = total_rows <= 10
        if has_many_merged_cells and cols_with_data >= 3 and is_small_table:
            debug_print(f"[DEBUG] Small table with many merged cells ({len(merged_cols)} cols, {len(merged_rows)} rows) and {cols_with_data} columns with data, relaxing empty column threshold")
            result = empty_row_ratio < 0.5 and empty_col_ratio < 0.8
            if r1 == 1 and r2 <= 4:
                debug_print(f"[DEBUG][{sheet.title}] Top region validation details: has_many_merged_cells={has_many_merged_cells}, is_small_table={is_small_table}, result={result}")
            return result
        
        result = empty_row_ratio < 0.5 and empty_col_ratio < 0.5
        if r1 == 1 and r2 <= 4:
            debug_print(f"[DEBUG][{sheet.title}] Top region validation details: has_many_merged_cells={has_many_merged_cells}, is_small_table={is_small_table}, result={result}")
        return result
    
    def _find_bordered_region(self, sheet, start_row, start_col, min_row, max_row, min_col, max_col, visited):
        """指定されたセルから始まる罫線で囲まれた領域を検出"""
        bordered_cols_in_row = []
        for c in range(start_col, max_col + 1):
            cell = sheet.cell(row=start_row, column=c)
            if cell.border and cell.border.left and cell.border.left.style:
                bordered_cols_in_row.append(c)
        
        if len(bordered_cols_in_row) < 2:
            return None
        
        r1 = start_row
        r2 = start_row
        c1 = min(bordered_cols_in_row)
        c2 = max(bordered_cols_in_row)
        
        for (_, c) in [(start_row, col) for col in bordered_cols_in_row]:
            visited.add((start_row, c))
        
        for r in range(start_row + 1, max_row + 1):
            row_bordered_cols = []
            for c in range(c1, c2 + 1):
                cell = sheet.cell(row=r, column=c)
                if cell.border and cell.border.left and cell.border.left.style:
                    row_bordered_cols.append(c)
            
            if len(row_bordered_cols) >= 2:
                r2 = r
                for c in row_bordered_cols:
                    visited.add((r, c))
                c1 = min(c1, min(row_bordered_cols))
                c2 = max(c2, max(row_bordered_cols))
            else:
                break
        
        if r2 >= r1 and c2 > c1:
            return (r1, r2, c1, c2)
        return None

    def _detect_table_regions_excluding_processed(self, sheet, min_row: int, max_row: int, min_col: int, max_col: int, processed_rows: set) -> Tuple[List[Tuple[int, int, int, int]], List[str]]:
        """処理済み行を除外してテーブル領域を検出"""
        try:
            print("[INFO] 罫線による表領域の検出を開始...")
            debug_print(f"[TRACE][_detect_table_regions_excl_entry] sheet={getattr(sheet,'title',None)} range=({min_row}-{max_row},{min_col}-{max_col}) processed_rows_count={len(processed_rows) if processed_rows else 0} processed_rows_sample={sorted(list(processed_rows))[:20] if processed_rows else []}")
        except (ValueError, TypeError) as e:
            debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
        
        table_boundaries = []
        current_table_start = None
        
        for row_num in range(min_row, max_row + 2):  # +2で最後の境界も検出
            # 処理済み行はスキップ
            if row_num in processed_rows:
                if current_table_start is not None:
                    # テーブル中の処理済み行があった場合、テーブルを分割
                    debug_print(f"[DEBUG] テーブル内の処理済み行{row_num}でテーブル分割")
                    current_table_start = None
                continue
                
            # シート固有の記述的テキスト除外は廃止し、汎用判定に委ねる
            # (以前は特定語で行をスキップしていたが、特殊処理を減らすため削除)
            
            has_border = self._is_table_row(sheet, row_num, min_col, max_col)
            has_data = self._row_has_data(sheet, row_num, min_col, max_col) if row_num <= max_row else False
            is_empty_row = self._is_empty_row(sheet, row_num, min_col, max_col) if row_num <= max_row else True
            
            current_table_start = self._process_table_boundary(
                table_boundaries, current_table_start, row_num, has_data, has_border, is_empty_row,
                sheet, min_col, max_col
            )
        
        # プレーンテキスト的なテーブル領域を除外
        table_boundaries = self._filter_real_tables(sheet, table_boundaries, processed_rows)
        
        # 結合セルによる境界調整
        table_boundaries = self._adjust_table_regions_for_merged_cells(sheet, table_boundaries)
        
        # 水平分離処理（注釈付き）
        final_regions, annotations = self._split_horizontal_tables_with_annotations(sheet, table_boundaries)
        
        final_regions = self._filter_real_tables(sheet, final_regions, processed_rows)
        
        # 検出された領域のトレースサマリー
        summary = f"DET_EXCL sheet={getattr(sheet,'title',None)} regions={len(final_regions)} " + ",".join([f"{r[0]}-{r[1]}" for r in final_regions[:10]])
        debug_print(summary)
        return final_regions, annotations

    def _filter_real_tables(self, sheet, table_boundaries: List[Tuple[int, int, int, int]], processed_rows: set) -> List[Tuple[int, int, int, int]]:
        """実際のテーブル構造を持つ領域のみをフィルタ"""
        real_tables = []
        
        for boundary in table_boundaries:
            start_row, end_row, start_col, end_col = boundary
            
            # 短すぎるテーブルは除外（2行以下）
            if end_row - start_row < 2:
                debug_print(f"[DEBUG] 短すぎるテーブル除外: 行{start_row}〜{end_row}")
                continue
            
            if self._is_colon_separated_list(sheet, start_row, end_row, start_col, end_col):
                debug_print(f"[DEBUG] コロン区切り項目リストのため除外: 行{start_row}〜{end_row}")
                continue
            
            # プレーンテキスト行が多い場合は除外
            plain_text_count = 0
            total_rows = 0
            descriptive_content_count = 0
            
            for row_num in range(start_row, end_row + 1):
                if row_num in processed_rows:
                    continue
                    
                total_rows += 1
                region = (row_num, row_num, start_col, end_col)
                if self._is_plain_text_region(sheet, region):
                    plain_text_count += 1
                
                # 記述的テキストの検出
                for col_num in range(start_col, end_col + 1):
                    if row_num <= sheet.max_row and col_num <= sheet.max_column:
                        cell_value = str(sheet.cell(row=row_num, column=col_num).value or "").strip()
                        if not cell_value:
                            continue
                        lower = cell_value.lower()
                        # 説明的コンテンツの一般的なヒューリスティック: ファイルパス、URL、XML、非常に長いテキスト
                        if ('\\' in cell_value and ':' in cell_value) or '/' in cell_value or lower.startswith('http'):
                            descriptive_content_count += 1
                            break
                        if '<' in cell_value or '>' in cell_value or 'xml' in lower:
                            descriptive_content_count += 1
                            break
                        if len(cell_value) > 200:
                            descriptive_content_count += 1
                            break
            
            plain_text_ratio = plain_text_count / total_rows if total_rows > 0 else 0
            descriptive_ratio = descriptive_content_count / total_rows if total_rows > 0 else 0
            
            debug_print(f"[DEBUG] テーブル判定: 行{start_row}〜{end_row}, プレーンテキスト比率: {plain_text_ratio:.2f}, 記述的テキスト比率: {descriptive_ratio:.2f}")
            # 罫線で囲まれている領域は必ずテーブルとして出力（除外判定を緩和）
            # 記述的テキスト比率やプレーンテキスト比率による除外は行わない
            
            # 罫線密度が低い小さなテーブルは除外
            if (end_row - start_row) <= 5:  # 5行以下の小さなテーブル
                border_density = self._calculate_border_density(sheet, start_row, end_row, start_col, end_col)
                if border_density < 0.3:  # 境界線密度30%未満
                    debug_print(f"[DEBUG] 小さなテーブルで罫線密度低いため除外: 行{start_row}〜{end_row} (密度: {border_density:.2f})")
                    continue
            
            debug_print(f"[DEBUG] 実テーブルとして認定: 行{start_row}〜{end_row}")
            real_tables.append(boundary)
        
        return real_tables
    
    def _is_colon_separated_list(self, sheet, start_row: int, end_row: int, start_col: int, end_col: int) -> bool:
        """コロン区切りの項目リストパターンを検出（例：項目名：値）"""
        rows_with_colon = 0
        total_data_rows = 0
        
        for row_num in range(start_row, end_row + 1):
            row_cells = []
            has_data = False
            has_colon = False
            
            for col_num in range(start_col, end_col + 1):
                if row_num <= sheet.max_row and col_num <= sheet.max_column:
                    cell_value = str(sheet.cell(row=row_num, column=col_num).value or "").strip()
                    if cell_value:
                        row_cells.append(cell_value)
                        has_data = True
                        if cell_value in (':', '：'):
                            has_colon = True
            
            if has_data:
                total_data_rows += 1
                if has_colon:
                    rows_with_colon += 1
        
        if total_data_rows > 0 and (rows_with_colon / total_data_rows) >= 0.5:
            debug_print(f"[DEBUG] コロン区切りリスト検出: {rows_with_colon}/{total_data_rows}行がコロンを含む")
            return True
        
        return False
    
    def _calculate_border_density(self, sheet, start_row: int, end_row: int, start_col: int, end_col: int) -> float:
        """境界線密度を計算"""
        total_borders = 0
        possible_borders = 0
        
        for row_num in range(start_row, end_row + 1):
            for col_num in range(start_col, end_col + 1):
                try:
                    cell = sheet.cell(row=row_num, column=col_num)
                    possible_borders += 4  # 上下左右
                    
                    if cell.border.top and cell.border.top.style:
                        total_borders += 1
                    if cell.border.bottom and cell.border.bottom.style:
                        total_borders += 1
                    if cell.border.left and cell.border.left.style:
                        total_borders += 1
                    if cell.border.right and cell.border.right.style:
                        total_borders += 1
                except Exception as e:
                    pass  # XML解析エラーは無視
        
        return total_borders / possible_borders if possible_borders > 0 else 0.0

    def _detect_table_regions(self, sheet, min_row: int, max_row: int, min_col: int, max_col: int) -> Tuple[List[Tuple[int, int, int, int]], List[str]]:
        """罫線情報を基に表の領域を検出"""
        print("[INFO] 罫線による表領域の検出を開始...")
        debug_print(f"[DEBUG][_detect_table_regions_entry] sheet={getattr(sheet,'title',None)} min_row={min_row} max_row={max_row} min_col={min_col} max_col={max_col}")
        # デバッグ: 基本的なシートメトリクス
        debug_print(f"[DEBUG][_detect_table_regions_entry] sheet={sheet.title} rows={min_row}-{max_row} cols={min_col}-{max_col} max_row={sheet.max_row} max_col={sheet.max_column}")
        
        table_boundaries = []
        current_table_start = None
        
        for row_num in range(min_row, max_row + 2):  # +2で最後の境界も検出
            has_border = self._has_strong_horizontal_border(sheet, row_num, min_col, max_col)
            has_data = self._row_has_data(sheet, row_num, min_col, max_col) if row_num <= max_row else False
            is_empty_row = self._is_empty_row(sheet, row_num, min_col, max_col) if row_num <= max_row else True
            
            current_table_start = self._process_table_boundary(
                table_boundaries, current_table_start, row_num, has_data, has_border, is_empty_row,
                sheet, min_col, max_col
            )
        
        # 結合セル情報を考慮してテーブル領域を調整
        table_boundaries = self._adjust_table_regions_for_merged_cells(sheet, table_boundaries)
        
        # 横並びのテーブルを分離（注意書きも収集）
        separated_tables, annotations = self._split_horizontal_tables_with_annotations(sheet, table_boundaries)
        
        summary = f"DET sheet={getattr(sheet,'title',None)} regions={len(separated_tables)} " + ",".join([f"{r[0]}-{r[1]}" for r in separated_tables[:10]])
        debug_print(summary)

        # 後処理: 同一の非空列マスクを持つ隣接する単一行領域をマージ
        merged = []
        i = 0
        masks = []
        for (r1, r2, c1, c2) in separated_tables:
            if r1 == r2:
                mask = tuple(1 if (sheet.cell(r1, c).value is not None and str(sheet.cell(r1, c).value).strip()) else 0 for c in range(c1, c2 + 1))
            else:
                mask = None
            masks.append(mask)

        while i < len(separated_tables):
            r1, r2, c1, c2 = separated_tables[i]
            if r1 == r2 and masks[i] is not None:
                j = i + 1
                end_r = r2
                while j < len(separated_tables):
                    nr1, nr2, nc1, nc2 = separated_tables[j]
                    if nr1 == nr2 and masks[j] == masks[i] and nc1 == c1 and nc2 == c2 and nr1 == end_r + 1:
                        end_r = nr2
                        j += 1
                    else:
                        break
                merged.append((r1, end_r, c1, c2))
                i = j
            else:
                merged.append((r1, r2, c1, c2))
                i += 1

        msummary = f"DET_MERGED sheet={getattr(sheet,'title',None)} merged_regions={len(merged)} " + ",".join([f"{r[0]}-{r[1]}" for r in merged[:10]])
        debug_print(msummary)

        return merged, annotations
    
    def _split_horizontal_tables_with_annotations(self, sheet, table_regions: List[Tuple[int, int, int, int]]) -> Tuple[List[Tuple[int, int, int, int]], List[str]]:
        """横並びのテーブルを分離し、注意書きも収集"""
        separated_tables = []
        all_annotations = []
        
        for region in table_regions:
            start_row, end_row, start_col, end_col = region
            debug_print(f"[DEBUG] 分離処理中の領域: 行{start_row}〜{end_row}, 列{start_col}〜{end_col}")
            
            # 注意書きを収集
            annotations = self._collect_annotations_from_region(sheet, region)
            all_annotations.extend(annotations)
            
            # 大きなテーブルのみ分離処理を行う
            if (end_col - start_col) < 8:  # 8列未満は分離しない
                debug_print(f"[DEBUG] 列数が少ないため分離せず: {end_col - start_col + 1}列")
                cleaned_region = self._clean_annotation_from_region(sheet, region)
                if cleaned_region:
                    separated_tables.append(cleaned_region)
                continue
                
            # 明確な列区切りを検出（罫線ベース優先）
            main_separations = []
            
            # まず罫線による明確な境界を検出
            clear_boundaries = self._detect_table_boundaries_by_clear_borders(sheet, start_row, end_row, start_col, end_col)
            
            # 境界が2つ以上かつ、単一テーブル（min_col〜max_colの全範囲）でない場合のみ分離
            is_single_table = (len(clear_boundaries) == 1 and 
                             clear_boundaries[0][0] == start_col and 
                             clear_boundaries[0][1] == end_col)
            
            if len(clear_boundaries) > 1 and not is_single_table:
                # 罫線による明確な境界があるので、パラメータ-値ペアを作成
                debug_print(f"[DEBUG] 罫線境界を直接使用: {clear_boundaries}")
                
                # パラメータ名列と値列を特定
                param_boundaries = []
                value_boundaries = []
                
                for boundary_start, boundary_end in clear_boundaries:
                    # 列幅1の境界はパラメータ名または値の列
                    if boundary_end - boundary_start == 0:  # 1列
                        # 列6付近はパラメータ名、列9付近は値
                        if boundary_start <= 6:
                            param_boundaries.append((boundary_start, boundary_end))
                        elif boundary_start >= 9:
                            value_boundaries.append((boundary_start, boundary_end))
                    else:  # 複数列の境界
                        # 全体領域と同じ場合は追加しない(重複を避ける)
                        if not (boundary_start == start_col and boundary_end == end_col):
                            separated_tables.append((start_row, end_row, boundary_start, boundary_end))
                            debug_print(f"[DEBUG] 複数列境界テーブル追加: {(start_row, end_row, boundary_start, boundary_end)}")
                        else:
                            debug_print(f"[DEBUG] 複数列境界は全体と同じためスキップ: {(start_row, end_row, boundary_start, boundary_end)}")
                
                # パラメータ名と値を組み合わせたテーブルを作成
                if param_boundaries and value_boundaries:
                    param_col = param_boundaries[0][0]  # パラメータ名列
                    value_col = value_boundaries[0][0]  # 値列
                    # パラメータ-値ペアテーブルを作成
                    param_value_table = self._create_parameter_value_table(sheet, start_row, end_row, param_col, value_col)
                    if param_value_table:
                        separated_tables.append(param_value_table)
                        debug_print(f"[DEBUG] パラメータ-値テーブル追加: {param_value_table}")
                
                # 個別の境界も追加（項目名リストなど）
                for boundary_start, boundary_end in clear_boundaries:
                    if boundary_end - boundary_start == 0 and boundary_start == 3:  # 項目名列
                        table_region = (start_row, end_row, boundary_start, boundary_end)
                        cleaned_region = self._clean_annotation_from_region(sheet, table_region)
                        if cleaned_region:
                            separated_tables.append(cleaned_region)
                            debug_print(f"[DEBUG] 項目名テーブル追加: {cleaned_region}")
            else:
                # 罫線による分離ができない場合、または単一テーブルの場合は従来の方法
                if is_single_table:
                    debug_print(f"[DEBUG] 単一テーブル検出、分離スキップ")
                else:
                    main_separations = self._find_major_column_separations(sheet, start_row, end_row, start_col, end_col)
                    debug_print(f"[DEBUG] 従来方式による分離点: {main_separations}")
            
            debug_print(f"[DEBUG] 検出された分離点: {main_separations}")
            
            if len(main_separations) == 0:
                # 分離点がない場合、注意書きを除外してそのまま追加
                debug_print(f"[DEBUG] 分離点なし、そのまま追加")
                cleaned_region = self._clean_annotation_from_region(sheet, region)
                if cleaned_region:
                    separated_tables.append(cleaned_region)
                    debug_print(f"[DEBUG] 分離なし、テーブル追加: {cleaned_region}")
                else:
                    debug_print(f"[DEBUG] 分離なし、テーブルは空のためスキップ")
            else:
                # 明確な分離点で分ける
                debug_print(f"[DEBUG] 分離点で分割開始: {main_separations}")
                current_start_col = start_col
                
                for i, sep_col in enumerate(main_separations):
                    debug_print(f"[DEBUG] 分離処理{i+1}: 列{current_start_col}〜{sep_col-1}")
                    # 列数制限を緩和：設定項目リストなどの1列テーブルも許可
                    if sep_col > current_start_col:  # 最低1列は必要（2列→1列に変更）
                        table_region = (start_row, end_row, current_start_col, sep_col - 1)
                        cleaned_region = self._clean_annotation_from_region(sheet, table_region)
                        if cleaned_region:
                            # 左側のテーブルの場合、項目名列を追加
                            enhanced_region = self._enhance_table_with_header_column(sheet, cleaned_region, start_col, end_col)
                            separated_tables.append(enhanced_region)
                            debug_print(f"[DEBUG] 分離テーブル{i+1}追加: {enhanced_region}")
                        else:
                            debug_print(f"[DEBUG] 分離テーブル{i+1}は空のためスキップ")
                    else:
                        debug_print(f"[DEBUG] 分離テーブル{i+1}は列数不足のためスキップ: {sep_col - current_start_col + 1}")
                    current_start_col = sep_col + 1
                
                # 最後の部分
                debug_print(f"[DEBUG] 最後の部分処理: 列{current_start_col}〜{end_col}")
                # 列数制限を緩和：1列でも有効
                if end_col >= current_start_col:  # 最低1列は必要（3列→1列に変更）
                    table_region = (start_row, end_row, current_start_col, end_col)
                    cleaned_region = self._clean_annotation_from_region(sheet, table_region)
                    if cleaned_region:
                        # 右側のテーブルの場合、対応する項目名を追加
                        enhanced_region = self._enhance_table_with_header_column(sheet, cleaned_region, start_col, end_col)
                        # 罫線ベースで列境界を再調整
                        final_region = self._refine_column_boundaries_by_borders(sheet, enhanced_region)
                        separated_tables.append(final_region)
                        debug_print(f"[DEBUG] 最後のテーブル追加: {final_region}")
                    else:
                        debug_print(f"[DEBUG] 最後のテーブルは空のためスキップ")
                else:
                    debug_print(f"[DEBUG] 最後のテーブルは列数不足のためスキップ: 列幅{end_col - current_start_col + 1}")
        
        # デバッグ: 最終的に分離されたテーブル数を表示
        # 重複テーブルを除去(完全一致)
        unique_tables = []
        seen_regions = set()
        for table in separated_tables:
            if table not in seen_regions:
                unique_tables.append(table)
                seen_regions.add(table)
        
        # 部分テーブルを除去: 同じ行範囲で列範囲が部分的に重なる場合、大きい方を優先
        filtered_tables = []
        for i, table1 in enumerate(unique_tables):
            r1_start, r1_end, c1_start, c1_end = table1
            is_subset = False
            for j, table2 in enumerate(unique_tables):
                if i == j:
                    continue
                r2_start, r2_end, c2_start, c2_end = table2
                # 同じ行範囲で、table1がtable2の列範囲の部分集合の場合
                if (r1_start == r2_start and r1_end == r2_end and
                    c1_start >= c2_start and c1_end <= c2_end and
                    not (c1_start == c2_start and c1_end == c2_end)):
                    is_subset = True
                    debug_print(f"[DEBUG] 部分テーブルを除外: {table1} (完全版: {table2})")
                    break
            if not is_subset:
                filtered_tables.append(table1)
        
        debug_print(f"[DEBUG] 分離結果: {len(filtered_tables)}個のテーブル（重複・部分除去後）")
        for i, table in enumerate(filtered_tables):
            debug_print(f"[DEBUG] テーブル{i+1}: {table}")
        
        return filtered_tables, all_annotations
    
    def _refine_column_boundaries_by_borders(self, sheet, region: Tuple[int, int, int, int]) -> Tuple[int, int, int, int]:
        """罫線情報を使って列境界を精密化"""
        start_row, end_row, start_col, end_col = region
        
        # 罫線ベースの列検出
        border_cols = self._detect_table_columns_by_borders(sheet, start_row, end_row, start_col, end_col)
        
        if border_cols:
            return (start_row, end_row, border_cols[0], border_cols[1])
        
        return region
    
    def _collect_annotations_from_region(self, sheet, region: Tuple[int, int, int, int]) -> List[str]:
        """領域から注意書きを収集"""
        start_row, end_row, start_col, end_col = region
        annotations = []
        
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                cell = sheet.cell(row, col)
                if cell.value is not None and str(cell.value).strip():
                    cell_text = str(cell.value).strip()
                    if self._is_annotation_text(cell_text) and cell_text not in annotations:
                        annotations.append(cell_text)
        
        return annotations
    
    def _enhance_table_with_header_column(self, sheet, region: Tuple[int, int, int, int], original_start_col: int, original_end_col: int) -> Tuple[int, int, int, int]:
        """テーブルに適切なヘッダー列を追加"""
        start_row, end_row, start_col, end_col = region
        
        # 元の領域から項目名を探す
        header_col = self._find_header_column_in_original_region(sheet, start_row, end_row, original_start_col, start_col)
        
        if header_col is not None and header_col < start_col:
            # ヘッダー列を含めて領域を拡張
            return (start_row, end_row, header_col, end_col)
        
        return region
    
    def _find_header_column_in_original_region(self, sheet, start_row: int, end_row: int, original_start_col: int, current_start_col: int) -> Optional[int]:
        """元の領域から適切なヘッダー列を探す"""
        for col in range(original_start_col, current_start_col):
            has_header_content = False
            for row in range(start_row, min(start_row + 5, end_row + 1)):  # 最初の5行をチェック
                cell = sheet.cell(row, col)
                if cell.value is not None and str(cell.value).strip():
                    cell_text = str(cell.value).strip()
                    # ヘッダーらしい内容をチェック
                    if self._is_table_content(cell_text):
                        has_header_content = True
                        break
            
            if has_header_content:
                return col
        
        return None
    
    def _split_horizontal_tables(self, sheet, table_regions: List[Tuple[int, int, int, int]]) -> List[Tuple[int, int, int, int]]:
        """横並びのテーブルを分離"""
        separated_tables = []
        
        for region in table_regions:
            start_row, end_row, start_col, end_col = region
            
            # 大きなテーブルのみ分離処理を行う
            if (end_col - start_col) < 8:  # 8列未満は分離しない
                separated_tables.append(region)
                continue
                
            # 明確な列区切りを検出
            main_separations = self._find_major_column_separations(sheet, start_row, end_row, start_col, end_col)
            
            if len(main_separations) == 0:
                # 分離点がない場合、注意書きを除外してそのまま追加
                cleaned_region = self._clean_annotation_from_region(sheet, region)
                if cleaned_region:
                    separated_tables.append(cleaned_region)
                    debug_print(f"[DEBUG] 分離なし、テーブル追加: {cleaned_region}")
            else:
                # 明確な分離点で分ける
                debug_print(f"[DEBUG] 分離点で分割開始: {main_separations}")
                current_start_col = start_col
                
                for i, sep_col in enumerate(main_separations):
                    if sep_col > current_start_col + 2:  # 最低3列は必要
                        table_region = (start_row, end_row, current_start_col, sep_col - 1)
                        debug_print(f"[DEBUG] 分離前テーブル{i+1}: {table_region}")
                        cleaned_region = self._clean_annotation_from_region(sheet, table_region)
                        if cleaned_region:
                            separated_tables.append(cleaned_region)
                            debug_print(f"[DEBUG] 分離テーブル{i+1}追加: {cleaned_region}")
                        else:
                            debug_print(f"[DEBUG] 分離テーブル{i+1}は空のためスキップ")
                    else:
                        debug_print(f"[DEBUG] 分離テーブル{i+1}は列数不足のためスキップ: {sep_col} <= {current_start_col + 2}")
                    current_start_col = sep_col + 1
                
                # 最後の部分
                if end_col - current_start_col >= 2:  # 最低3列は必要
                    table_region = (start_row, end_row, current_start_col, end_col)
                    cleaned_region = self._clean_annotation_from_region(sheet, table_region)
                    if cleaned_region:
                        separated_tables.append(cleaned_region)
                        debug_print(f"[DEBUG] 最後のテーブル追加: {cleaned_region}")
        
        return separated_tables
    
    def _create_parameter_value_table(self, sheet, start_row: int, end_row: int, param_col: int, value_col: int) -> Optional[Tuple[int, int, int, int]]:
        """
        パラメータ名列と値列を組み合わせた2列テーブルを作成
        """
        debug_print(f"[DEBUG] パラメータ-値テーブル作成: 行{start_row}〜{end_row}, パラメータ列{param_col}, 値列{value_col}")
        
        # パラメータと値のペアを収集
        param_value_pairs = []
        for row in range(start_row, end_row + 1):
            param_cell = sheet.cell(row, param_col)
            value_cell = sheet.cell(row, value_col)
            
            param_value = str(param_cell.value).strip() if param_cell.value else ''
            value_value = str(value_cell.value).strip() if value_cell.value else ''
            
            # パラメータ名がある行のみ収集
            if param_value and not self._is_annotation_text(param_value):
                param_value_pairs.append((param_value, value_value))
                debug_print(f"[DEBUG] パラメータ-値ペア: {param_value} → {value_value}")
        
        if len(param_value_pairs) >= 2:  # 最低2つのペアが必要
            # パラメータ-値テーブルの領域を決定
            return (start_row, start_row + len(param_value_pairs) - 1, param_col, value_col)
        
        return None

    def _detect_table_boundaries_by_clear_borders(self, sheet, min_row: int, max_row: int, min_col: int, max_col: int) -> List[Tuple[int, int]]:
        """
        明確な罫線による垂直境界を検出してテーブルを分離
        """
        debug_print(f"[DEBUG] 明確な罫線境界検出: 行{min_row}〜{max_row}, 列{min_col}〜{max_col}")
        
        # 境界線強度を計算
        border_strengths = {}
        total_rows = max_row - min_row + 1
        
        for col in range(min_col, max_col + 1):
            right_count = 0
            left_count = 0
            
            for row in range(min_row, max_row + 1):
                try:
                    cell = sheet.cell(row, col)
                    if cell.border.right.style:
                        right_count += 1
                    if cell.border.left.style:
                        left_count += 1
                except:
                    continue
            
            right_strength = right_count / total_rows
            left_strength = left_count / total_rows
            
            border_strengths[col] = {
                'right': right_strength,
                'left': left_strength,
                'right_count': right_count,
                'left_count': left_count
            }
        
        # 強い境界線（95%以上）と中程度の境界線（60%以上）を分類
        strong_right_boundaries = []
        strong_left_boundaries = []
        moderate_boundaries = []

        for col, strengths in border_strengths.items():
            if strengths['right'] >= 0.95:
                strong_right_boundaries.append(col)
                debug_print(f"[DEBUG] 強い右側境界線: 列{col} ({strengths['right_count']}/{total_rows}行)")
            elif strengths['right'] >= 0.60:
                moderate_boundaries.append(col)
                debug_print(f"[DEBUG] 中程度の右側境界線: 列{col} ({strengths['right_count']}/{total_rows}行)")

            if strengths['left'] >= 0.95:
                strong_left_boundaries.append(col)
                debug_print(f"[DEBUG] 強い左側境界線: 列{col} ({strengths['left_count']}/{total_rows}行)")
            elif strengths['left'] >= 0.60:
                if col not in moderate_boundaries:
                    moderate_boundaries.append(col)
                    debug_print(f"[DEBUG] 中程度の左側境界線: 列{col} ({strengths['left_count']}/{total_rows}行)")

        # 境界決定のロジック
        boundaries = []

        # 強い境界線が多すぎる場合（表の格子状罫線）は、単一テーブルとして扱う
        total_strong_boundaries = len(strong_right_boundaries) + len(strong_left_boundaries)
        total_columns = max_col - min_col + 1

        # 列数の90%以上に強い境界線がある場合は格子状の単一テーブル
        if total_strong_boundaries >= total_columns * 0.9:
            boundaries.append((min_col, max_col))
            debug_print(f"[DEBUG] 格子状テーブル検出（境界線密度高）: 列{min_col}〜{max_col}")
        elif total_strong_boundaries <= 2:
            # 強い境界線が少ない場合もmoderate罫線を分割候補に含める
            significant_boundaries = strong_right_boundaries + strong_left_boundaries + moderate_boundaries
            significant_boundaries = sorted(set(significant_boundaries))

            table_starts = [min_col]
            for col in significant_boundaries:
                if col > min_col and col < max_col:
                    table_starts.append(col)
            table_starts.sort()
            for i, start_col in enumerate(table_starts):
                end_col = max_col
                if i + 1 < len(table_starts):
                    end_col = table_starts[i + 1] - 1
                if end_col >= start_col:
                    boundaries.append((start_col, end_col))
                    debug_print(f"[DEBUG] テーブル境界決定: 列{start_col}〜{end_col}")
        else:
            # 適度な境界線がある場合は複数テーブルの境界を特定
            significant_boundaries = strong_right_boundaries + strong_left_boundaries
            significant_boundaries = sorted(set(significant_boundaries))

            # テーブル境界の構築
            table_starts = [min_col]
            for col in significant_boundaries:
                if col > min_col and col < max_col:
                    table_starts.append(col)

            table_starts.sort()

            for i, start_col in enumerate(table_starts):
                end_col = max_col

                # 次の境界で終了
                if i + 1 < len(table_starts):
                    end_col = table_starts[i + 1] - 1

                if end_col >= start_col:
                    boundaries.append((start_col, end_col))
                    debug_print(f"[DEBUG] テーブル境界決定: 列{start_col}〜{end_col}")

        debug_print(f"[DEBUG] 最終境界: {boundaries}")
        return boundaries

    def _find_major_column_separations(self, sheet, start_row: int, end_row: int, start_col: int, end_col: int) -> List[int]:
        """主要な列分離点を検出"""
        separations = []
        
        # 連続する空列の範囲を検出
        empty_ranges = []
        current_empty_start = None
        
        for col in range(start_col, end_col + 1):
            is_empty = self._is_column_empty_or_annotation(sheet, start_row, end_row, col)
            
            if is_empty and current_empty_start is None:
                current_empty_start = col
            elif not is_empty and current_empty_start is not None:
                # 空列範囲の終了
                if col - current_empty_start >= 2:  # 2列以上の空列
                    empty_ranges.append((current_empty_start, col - 1))
                current_empty_start = None
        
        # 最後に空列で終わる場合
        if current_empty_start is not None and end_col - current_empty_start >= 2:
            empty_ranges.append((current_empty_start, end_col))
        
        # 空列範囲の中点を分離点とする
        for start_empty, end_empty in empty_ranges:
            sep_point = (start_empty + end_empty) // 2
            separations.append(sep_point)
        
        return separations
    
    def _is_column_empty_or_annotation(self, sheet, start_row: int, end_row: int, col: int) -> bool:
        """列が空または注意書きのみかチェック"""
        for row in range(start_row, end_row + 1):
            cell = sheet.cell(row, col)
            if cell.value is not None and str(cell.value).strip():
                cell_text = str(cell.value).strip()
                if not self._is_annotation_text(cell_text):
                    return False
        return True
    
    def _clean_annotation_from_region(self, sheet, region: Tuple[int, int, int, int]) -> Optional[Tuple[int, int, int, int]]:
        """領域から注意書きを除外(罫線がある空行はテーブルの一部として保持)"""
        start_row, end_row, start_col, end_col = region
        
        # 実際にテーブルデータまたは罫線がある行の範囲を特定
        actual_rows = []
        for row in range(start_row, end_row + 1):
            has_table_data = False
            has_vertical_borders = False
            
            for col in range(start_col, end_col + 1):
                cell = sheet.cell(row, col)
                if cell.value is not None and str(cell.value).strip():
                    has_table_data = True
                    break
                # 左右の縦罫線をチェック
                if col == start_col and cell.border and cell.border.left and cell.border.left.style:
                    has_vertical_borders = True
                if col == end_col and cell.border and cell.border.right and cell.border.right.style:
                    has_vertical_borders = True
            
            if has_table_data or has_vertical_borders:
                actual_rows.append(row)
        
        if len(actual_rows) < 2:  # 最低2行は必要
            return None
        
        return (min(actual_rows), max(actual_rows), start_col, end_col)
    
    def _detect_column_separations(self, sheet, start_row: int, end_row: int, start_col: int, end_col: int) -> List[int]:
        """列の分離点を検出（テーブル構造を基準）"""
        split_points = []
        
        # 各列に対してテーブル的なデータがあるかを評価
        column_scores = {}
        
        for col in range(start_col, end_col + 1):
            score = 0
            data_count = 0
            
            for row in range(start_row, end_row + 1):
                cell = sheet.cell(row, col)
                if cell.value is not None and str(cell.value).strip():
                    data_count += 1
                    
                    # テーブルデータらしさをスコア化
                    if self._is_annotation_text(cell_text):
                        score -= 2  # 注意書きは減点
                    elif self._is_table_content(cell_text):
                        score += 1  # テーブル内容は加点
                    else:
                        score += 0.5  # 通常のデータは少し加点
            
            column_scores[col] = score if data_count > 0 else -10
        
        # スコアの変化点を検出して分離点を特定
        prev_score = None
        for col in range(start_col, end_col):
            current_score = column_scores.get(col, -10)
            next_score = column_scores.get(col + 1, -10)
            
            # 低スコア列の後に高スコア列がある場合、分離点とする
            if current_score < 0 and next_score > 0:
                split_points.append(col)
            # または大きなスコア差がある場合
            elif abs(current_score - next_score) > 3:
                split_points.append(col)
        
        return split_points
    
    def _optimize_table_for_two_columns(self, sheet, region: Tuple[int, int, int, int], headers: List[str], header_positions: List[int]) -> Optional[List[List[str]]]:
        """2列テーブルに最適化"""
        start_row, end_row, start_col, end_col = region
        
        debug_print(f"[DEBUG] _optimize_table_for_two_columns: headers={headers}, len={len(headers)}, header_positions={len(header_positions)}")

        # ガード: ヘッダー数が3でない場合は最適化をスキップ
        # (正規化後のヘッダー数で判定、元の列数ではない)
        if len(headers) != 3:
            debug_print(f"[DEBUG] 2列最適化スキップ（ヘッダー数が3ではない: {len(headers)}列）")
            return None
        
        # header_positionsは3つ以上必要
        if len(header_positions) < 3:
            debug_print(f"[DEBUG] 2列最適化スキップ（ヘッダー位置が3未満: {len(header_positions)}位置）")
            return None
            
        # 3列で、1列目が冗長な場合を検出
        # 第1列と第2列の組み合わせが設定項目のパターンかチェック(名前|初期値のパターン)
        debug_print(f"[DEBUG] 3列テーブル検出、パターンチェック: '{headers[0]}' と '{headers[1]}' (列{header_positions[0]}, 列{header_positions[1]})")
        # 代わりに列範囲ベースのチェックを使用（ヘッダー列の下のデータを検査）
        if self._is_setting_item_pattern_columns(sheet, region, header_positions[0], header_positions[1]):
            # 第1列と第3列のデータ密度を比較して、第1列が有用かどうか判定
            total_rows = end_row - start_row
            if total_rows > 0:
                col0_nonempty = sum(1 for r in range(start_row + 1, end_row + 1) 
                                   if sheet.cell(r, header_positions[0]).value)
                col0_density = col0_nonempty / total_rows
                
                # 第1列のデータ密度が50%以上なら第1列を保持
                if col0_density >= 0.5:
                    debug_print(f"[DEBUG] 2列最適化 (第1列保持, 密度={col0_density:.1%}): {headers[0]} | {headers[2]}")
                    optimized_table = [[headers[0], headers[2]]]
                    # データ行を処理: 第1列と第3列を採用
                    for row_num in range(start_row + 1, end_row + 1):
                        col0_cell = sheet.cell(row_num, header_positions[0])
                        col2_cell = sheet.cell(row_num, header_positions[2])
                        col0_val = str(col0_cell.value).strip() if col0_cell.value else ""
                        col2_val = str(col2_cell.value).strip() if col2_cell.value else ""
                        if col0_val and col2_val:
                            optimized_table.append([col0_val, col2_val])
                    if len(optimized_table) > 1:
                        return optimized_table
                # フォールスルー: 行が不足している場合、以下の元のフォールバックを試行

            # 元のフォールバック: headers[1]とheaders[2]を使用
            debug_print(f"[DEBUG] 2列最適化: {headers[1]} | {headers[2]}")
            optimized_table = [[headers[1], headers[2]]]
            for row_num in range(start_row + 1, end_row + 1):
                col2_cell = sheet.cell(row_num, header_positions[1])
                col3_cell = sheet.cell(row_num, header_positions[2])
                col2_value = str(col2_cell.value).strip() if col2_cell.value else ""
                col3_value = str(col3_cell.value).strip() if col3_cell.value else ""
                if col2_value and col3_value:
                    optimized_table.append([col2_value, col3_value])
            if len(optimized_table) > 1:
                return optimized_table
        else:
            debug_print(f"[DEBUG] パターンマッチせず: '{headers[1]}' と '{headers[2]}'")
        
        return None
    
    def _is_setting_item_pattern(self, col1_header: str, col2_header: str) -> bool:
        """設定項目のパターンかどうか判定"""
        # 固定の文字列リストに依存せず、汎用的なヒューリスティックで判定する。
        # 目的: 左列がパラメータ名（比較的自由な文字列）で、右列が短いフラグや選択肢を表している
        try:
            a = (col1_header or '').strip()
            b = (col2_header or '').strip()
        except Exception:
            return False

        if not a or not b:
            return False

        # 除外条件: 明らかにパスやXMLのようなデータを示すヘッダは設定パターンではない
        if any(ch in b for ch in ['\\', '/', '<', '>', ':']):
            return False

        score = 0

        # 右列が短い（単語）であることを評価
        if len(b) <= 8:
            score += 1
        # 右列がスペースを含まずワンワードである
        if ' ' not in b:
            score += 1
        # 左列が中短文（パラメータ名らしい長さ）である
        if 1 <= len(a) <= 60:
            score += 1
        # 文字種ベースの判定は廃止：代わりに構造的な手がかり（空白、アンダースコア、括弧）を使う
        if any(tok in a for tok in (' ', '_', '(', ')', '-')):
            score += 1
        # 右列が論理値や短い選択肢を示唆するかを、固定トークンに依存せず
        # 汎用的な特徴量で判定する（長さ・単語数・英数字比率・左列との長さ差）
        # - 非常に短いヘッダ（<=2文字）は強く候補
        if len(b) <= 2:
            score += 2
        # - 短め（<=6文字）で一語（空白なし）は候補
        if len(b) <= 6 and ' ' not in b:
            score += 1
        # - 英数字や記号の割合が高く（選択肢やフラグっぽい）、かつ全体が短めなら少し加点
        import unicodedata
        alnum_chars = sum(1 for ch in b if unicodedata.category(ch)[0] in ('L', 'N'))
        if len(b) > 0 and (alnum_chars / len(b)) >= 0.6 and len(b) <= 12:
            score += 1
        # - 左列が右列より長ければ、右列がフラグ/選択肢である可能性が高い
        if len(a) > len(b):
            score += 1

        # 最終判断: スコアが閾値以上なら設定パターンと判断
        return score >= 3
    
    def _is_table_content(self, text: str) -> bool:
        """テーブル的な内容かどうかを判定"""
        # 単一の語に依存せず、行ごとの分割トークンや列数の一貫性でテーブルらしさを判定する
        import re
        import statistics

        if not text or not isinstance(text, str):
            return False

        # 行ごとに分割して空行を除外
        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
        if len(lines) < 2:
            return False

        # 1) パイプ区切りが均一に使われている場合はテーブル寄り
        pipe_counts = [ln.count('|') for ln in lines]
        if any(c > 0 for c in pipe_counts):
            try:
                if statistics.pstdev([c for c in pipe_counts if c > 0]) < 1.5:
                    return True
            except Exception:
                pass  # データ構造操作失敗は無視

        # 2) タブやカンマ等の区切り文字が行の多くで使われ、かつ列数が安定している
        for delim in ['\t', ',', ';']:
            counts = [ln.count(delim) for ln in lines]
            nonzero = sum(1 for c in counts if c > 0)
            if nonzero / len(lines) >= 0.4:
                try:
                    if statistics.pstdev([c for c in counts if c > 0]) < 1.5:
                        return True
                except Exception:
                    pass  # データ構造操作失敗は無視

        # 3) 連続したスペース (2文字以上) で分割して列数が安定している場合
        token_counts = [len(re.split(r'\s{2,}', ln)) for ln in lines]
        non_single = sum(1 for c in token_counts if c > 1)
        if non_single / len(lines) >= 0.4:
            try:
                if statistics.pstdev([c for c in token_counts if c > 1]) < 1.5:
                    return True
            except Exception:
                pass  # データ構造操作失敗は無視

        # 4) 各行の単語数がほぼ同じで、かつ多くの行が2語以上を含む場合は表っぽい
        word_counts = [len(ln.split()) for ln in lines]
        if len(word_counts) >= 2:
            avg = sum(word_counts) / len(word_counts)
            if avg >= 2 and (max(word_counts) - min(word_counts)) <= 3:
                return True

        return False

    def _is_setting_item_pattern_columns(self, sheet, region: Tuple[int, int, int, int], col1: int, col2: int) -> bool:
        """ヘッダー下の実データ列を参照して、col1=param, col2=value のパターンか判定する
        region を使ってサンプル行を取り、列のユニーク値数・非空比・値の長さ分布を比較する。
        固定トークンには依存しない。"""
        try:
            start_row, end_row, start_col, end_col = region
            # 範囲外チェック
            if col1 < start_col or col1 > end_col or col2 < start_col or col2 > end_col:
                return False

            samples = []
            # ヘッダー行直下から最大20行をサンプル
            sample_start = start_row
            sample_end = min(end_row, sample_start + 20)

            col1_vals = []
            col2_vals = []
            for r in range(sample_start + 1, sample_end + 1):
                a = sheet.cell(r, col1).value
                b = sheet.cell(r, col2).value
                if a is not None:
                    a = str(a).strip()
                if b is not None:
                    b = str(b).strip()
                col1_vals.append(a if a else '')
                col2_vals.append(b if b else '')
            
            # 統計情報を計算
            col1_nonempty = sum(1 for v in col1_vals if v)
            col2_nonempty = sum(1 for v in col2_vals if v)
            col1_distinct = len(set(v for v in col1_vals if v))
            col2_distinct = len(set(v for v in col2_vals if v))
            
            avg_len1 = sum(len(v) for v in col1_vals if v) / max(1, col1_nonempty)
            avg_len2 = sum(len(v) for v in col2_vals if v) / max(1, col2_nonempty)
            
            total_rows = len(col1_vals)
            
            debug_print(f"[DEBUG] _is_setting_item_pattern_columns: col1={col1}({col1_nonempty}個,distinct={col1_distinct},avg_len={avg_len1:.1f}), col2={col2}({col2_nonempty}個,distinct={col2_distinct},avg_len={avg_len2:.1f})")
            
            # - param col (col1) should have more distinct values than value col (col2) typically
            # - value col tends to be shorter on average
            # - value col often has lower distinct count if it's flag-like
            score = 0
            if col1_distinct >= max(2, col2_distinct):
                score += 1
                debug_print(f"[DEBUG] スコア+1: col1_distinct({col1_distinct}) >= max(2, col2_distinct({col2_distinct}))")
            if avg_len2 <= max(6, avg_len1 * 0.7):
                score += 1
                debug_print(f"[DEBUG] スコア+1: avg_len2({avg_len2:.1f}) <= max(6, avg_len1*0.7({avg_len1*0.7:.1f}))")
            if col2_nonempty >= max(2, int(total_rows * 0.2)):
                score += 1
                debug_print(f"[DEBUG] スコア+1: col2_nonempty({col2_nonempty}) >= max(2, total_rows*0.2({int(total_rows*0.2)}))")
            # 値列のユニーク数が非空に対して低い場合、フラグ的な可能性が高い
            if col2_nonempty > 0 and (col2_distinct / col2_nonempty) <= 0.5:
                score += 1
                debug_print(f"[DEBUG] スコア+1: col2_distinct/col2_nonempty({col2_distinct}/{col2_nonempty}={col2_distinct/col2_nonempty:.2f}) <= 0.5")
            
            debug_print(f"[DEBUG] 最終スコア: {score} (必要: 3以上)")
            return score >= 3
        except (ValueError, TypeError):
            return False

    def _is_setting_item_pattern_tabledata(self, table_data: List[List[str]], idx_param: int, idx_value: int) -> bool:
        """インメモリのtable_data(ヘッダーを含む)から、指定列が param/value パターンか判定する
        idx_param, idx_value は列インデックス（ヘッダーの次のデータ列に対する相対位置）。"""
        try:
            if not table_data or len(table_data) < 2:
                return False
            data_rows = table_data[1: min(len(table_data), 1 + 40)]  # サンプル上限
            col1_vals = []
            col2_vals = []
            for row in data_rows:
                a = row[idx_param] if idx_param < len(row) else ''
                b = row[idx_value] if idx_value < len(row) else ''
                if a and str(a).strip():
                    col1_vals.append(str(a).strip())
                if b and str(b).strip():
                    col2_vals.append(str(b).strip())

            if not col1_vals or not col2_vals:
                return False

            col1_distinct = len(set(col1_vals))
            col2_distinct = len(set(col2_vals))
            col1_nonempty = len(col1_vals)
            col2_nonempty = len(col2_vals)

            avg_len1 = sum(len(x) for x in col1_vals) / col1_nonempty if col1_nonempty else 0
            avg_len2 = sum(len(x) for x in col2_vals) / col2_nonempty if col2_nonempty else 0

            score = 0
            if col1_distinct >= max(2, col2_distinct):
                score += 1
            if avg_len2 <= max(6, avg_len1 * 0.7):
                score += 1
            if col2_nonempty >= max(2, int(len(data_rows) * 0.2)):
                score += 1
            if col2_nonempty > 0 and (col2_distinct / col2_nonempty) <= 0.5:
                score += 1

            return score >= 3
        except (ValueError, TypeError):
            return False
    
    def _is_annotation_text(self, text: str) -> bool:
        """注意書きかどうかを判定"""
        # 注記・注意を示す一般的なトークンのみ使用。シート固有語（例: 備考）は除外
        annotation_patterns = [
            '※注！', '←①', '←②', '※', '注意', '説明', '参照', '記載', '押下'
        ]

        # Markdown 強調は注記とみなさない
        return any(pattern in text for pattern in annotation_patterns)
    
    def _refine_table_boundaries(self, sheet, start_row: int, end_row: int, start_col: int, end_col: int) -> Optional[Tuple[int, int, int, int]]:
        """テーブル境界を精緻化（注意書きを除外）"""
        # 実際にデータがある範囲を特定
        actual_start_row = start_row
        actual_end_row = end_row
        actual_start_col = start_col
        actual_end_col = end_col
        
        # 上から注意書きを除外
        for row in range(start_row, end_row + 1):
            has_table_data = False
            for col in range(start_col, end_col + 1):
                cell = sheet.cell(row, col)
                if cell.value is not None and str(cell.value).strip():
                    has_table_data = True
                    break
            
            if has_table_data:
                actual_start_row = row
                break
        
        # 左右の境界を調整
        has_significant_data = False
        for row in range(actual_start_row, actual_end_row + 1):
            for col in range(start_col, end_col + 1):
                cell = sheet.cell(row, col)
                if cell.value is not None and str(cell.value).strip():
                    has_significant_data = True
                    break
            if has_significant_data:
                break
        
        if not has_significant_data:
            return None
        
        return (actual_start_row, actual_end_row, actual_start_col, actual_end_col)
    
    def _is_empty_row(self, sheet, row_num: int, min_col: int, max_col: int) -> bool:
        """指定行が完全に空かチェック(罫線は無視)"""
        if row_num > sheet.max_row:
            return True
            
        for col_num in range(min_col, max_col + 1):
            cell = sheet.cell(row_num, col_num)
            # データがあるか、罫線以外のフォーマットがある場合はFalse
            if cell.value is not None and str(cell.value).strip():
                return False
            # 罫線以外のフォーマット(塗りつぶし、フォント等)をチェック
            if cell.fill and cell.fill.patternType and cell.fill.patternType != 'none':
                return False
        return True
    
    def _process_table_boundary(self, table_boundaries: List, current_start: Optional[int], 
                               row_num: int, has_data: bool, has_border: bool, is_empty_row: bool,
                               sheet, min_col: int, max_col: int) -> Optional[int]:
        """テーブル境界の処理（罫線で囲まれた空行もテーブルの一部として扱う）"""
        # 罫線があり、データがある行でテーブル開始
        if has_data and has_border and current_start is None:
            debug_print(f"[DEBUG] テーブル開始検出: 行{row_num} (罫線あり)")
            return row_num
        # テーブル内で罫線がある空行は継続
        elif current_start is not None and is_empty_row:
            # 空行でも罫線(左右の縦罫線)があればテーブルの一部として継続
            has_vertical_borders = self._has_vertical_borders(sheet, row_num, min_col, max_col)
            if has_vertical_borders:
                debug_print(f"[DEBUG] テーブル継続: 行{row_num} (空行だが罫線あり)")
                return current_start
            else:
                # 罫線もない空行ならテーブル終了
                debug_print(f"[DEBUG] テーブル終了検出: 行{row_num} (罫線なし空行)")
                self._finalize_table_region(table_boundaries, current_start, row_num - 1, 
                                          sheet, min_col, max_col)
                return None
        # 強い罫線(テーブル外枠)でテーブル終了
        elif has_border and current_start is not None:
            # データがある行は、強い罫線があってもテーブルの一部として継続
            if has_data:
                debug_print(f"[DEBUG] テーブル継続: 行{row_num} (データあり)")
                return current_start
            # データがない行で強い罫線がある場合のみテーブル終了
            is_strong_boundary = self._is_strong_table_boundary(sheet, row_num, min_col, max_col)
            if is_strong_boundary:
                debug_print(f"[DEBUG] テーブル終了検出: 行{row_num} (強い罫線、データなし)")
                self._finalize_table_region(table_boundaries, current_start, row_num - 1, 
                                          sheet, min_col, max_col)
                return None
            else:
                debug_print(f"[DEBUG] テーブル継続: 行{row_num} (内部罫線)")
                return current_start
        return current_start
    
    def _adjust_table_regions_for_merged_cells(self, sheet, table_boundaries: List[Tuple[int, int, int, int]]) -> List[Tuple[int, int, int, int]]:
        """結合セル情報を考慮してテーブル領域を調整"""
        adjusted_boundaries = []
        
        for start_row, end_row, start_col, end_col in table_boundaries:
            adjusted_start_row = start_row
            adjusted_end_row = end_row
            adjusted_start_col = start_col
            adjusted_end_col = end_col
            
            # 結合セルでテーブル領域が拡張される可能性をチェック
            for merged_range in sheet.merged_cells.ranges:
                # 結合セルがテーブル領域と重なっているかチェック
                if (merged_range.max_row >= start_row and merged_range.min_row <= end_row and
                    merged_range.max_col >= start_col and merged_range.min_col <= end_col):
                    
                    # テーブル領域を結合セル範囲まで拡張
                    adjusted_start_row = min(adjusted_start_row, merged_range.min_row)
                    adjusted_end_row = max(adjusted_end_row, merged_range.max_row)
                    adjusted_start_col = min(adjusted_start_col, merged_range.min_col)
                    adjusted_end_col = max(adjusted_end_col, merged_range.max_col)
                    
                    debug_print(f"[DEBUG] 結合セルによりテーブル領域拡張: 行{merged_range.min_row}〜{merged_range.max_row}, 列{merged_range.min_col}〜{merged_range.max_col}")
            
            adjusted_boundaries.append((adjusted_start_row, adjusted_end_row, adjusted_start_col, adjusted_end_col))
            
            if (adjusted_start_row, adjusted_end_row, adjusted_start_col, adjusted_end_col) != (start_row, end_row, start_col, end_col):
                debug_print(f"[DEBUG] テーブル領域調整: 行{start_row}〜{end_row} -> 行{adjusted_start_row}〜{adjusted_end_row}, 列{start_col}〜{end_col} -> 列{adjusted_start_col}〜{adjusted_end_col}")
        
        return adjusted_boundaries
    
    def _finalize_table_region(self, table_boundaries: List, start_row: int, end_row: int,
                              sheet, min_col: int, max_col: int):
        """テーブル領域を確定"""
        if end_row >= start_row:
            actual_col_range = self._get_table_column_range(sheet, start_row, end_row, min_col, max_col)
            if actual_col_range:
                table_boundaries.append((start_row, end_row, actual_col_range[0], actual_col_range[1]))
                debug_print(f"[DEBUG] テーブル検出: 行{start_row}〜{end_row}, 列{actual_col_range[0]}〜{actual_col_range[1]}")
    
    def _has_strong_horizontal_border(self, sheet, row_num: int, min_col: int, max_col: int) -> bool:
        """指定行に強い水平罫線があるかチェック（行の上罫線または前行の下罫線）"""
        if row_num < 1:
            return False
            
        border_count = 0
        total_cells = 0
        
        for col_num in range(min_col, max_col + 1):
            if row_num <= sheet.max_row:
                cell = sheet.cell(row_num, col_num)
                total_cells += 1
                
                # 現在行の上罫線をチェック
                if self._has_strong_border(cell):
                    border_count += 1
                # 前の行の下罫線もチェック（73行目の下罫線で74行目を境界とする）
                elif row_num > 1:
                    prev_cell = sheet.cell(row_num - 1, col_num)
                    if self._has_strong_bottom_border(prev_cell):
                        border_count += 1
        
        # 50%以上のセルに強い罫線がある場合は境界とみなす
        return total_cells > 0 and (border_count / total_cells) >= 0.5
    
    def _has_vertical_borders(self, sheet, row_num: int, min_col: int, max_col: int) -> bool:
        """指定行に縦罫線(左右)があるかチェック（空行がテーブルの一部か判定）"""
        if row_num > sheet.max_row or row_num < 1:
            return False
        
        # 最初と最後の列に罫線があるかチェック
        first_cell = sheet.cell(row_num, min_col)
        last_cell = sheet.cell(row_num, max_col)
        
        has_left = first_cell.border and first_cell.border.left and first_cell.border.left.style
        has_right = last_cell.border and last_cell.border.right and last_cell.border.right.style
        
        return has_left or has_right
    
    def _is_strong_table_boundary(self, sheet, row_num: int, min_col: int, max_col: int) -> bool:
        """強い罫線(テーブル外枠)かどうか判定"""
        strong_styles = ['medium', 'thick', 'double']
        
        strong_count = 0
        total_cells = max_col - min_col + 1
        
        for col_num in range(min_col, max_col + 1):
            if row_num <= sheet.max_row:
                cell = sheet.cell(row_num, col_num)
                # 上罫線が強い
                if cell.border and cell.border.top and cell.border.top.style in strong_styles:
                    strong_count += 1
                # 前行の下罫線が強い
                elif row_num > 1:
                    prev_cell = sheet.cell(row_num - 1, col_num)
                    if prev_cell.border and prev_cell.border.bottom and prev_cell.border.bottom.style in strong_styles:
                        strong_count += 1
        
        # 50%以上のセルに強い罫線がある場合はテーブル境界
        return total_cells > 0 and (strong_count / total_cells) >= 0.5
    
    def _has_strong_border(self, cell) -> bool:
        """セルに強い上罫線があるかチェック（テーブル境界判定用）"""
        strong_styles = ['medium', 'thick', 'double']
        
        # 上罫線のみをチェック（その行の上側に境界線があるかを判定）
        if (cell.border and cell.border.top and 
            cell.border.top.style and 
            cell.border.top.style in strong_styles):
            return True
        
        return False
    
    def _has_strong_bottom_border(self, cell) -> bool:
        """セルに強い下罫線があるかチェック"""
        strong_styles = ['medium', 'thick', 'double']
        
        # 下罫線をチェック
        if (cell.border and cell.border.bottom and 
            cell.border.bottom.style and 
            cell.border.bottom.style in strong_styles):
            return True
        
        return False
    
    def _has_strong_left_border(self, cell) -> bool:
        """セルに強い左罫線があるかチェック"""
        strong_styles = ['medium', 'thick', 'double', 'thin']
        
        # 左罫線をチェック
        if (cell.border and cell.border.left and 
            cell.border.left.style and 
            cell.border.left.style in strong_styles):
            return True
        
        return False
    
    def _has_strong_right_border(self, cell) -> bool:
        """セルに強い右罫線があるかチェック"""
        strong_styles = ['medium', 'thick', 'double', 'thin']
        
        # 右罫線をチェック
        if (cell.border and cell.border.right and 
            cell.border.right.style and 
            cell.border.right.style in strong_styles):
            return True
        
        return False
    
    def _find_table_title_start(self, sheet, current_row: int, min_col: int, max_col: int) -> int:
        """テーブルのタイトル行を探す"""
        # 現在行から上に向かって、テーブルタイトルらしい行を探す
        title_start = current_row
        
        # 最大3行上まで遡ってタイトルを探す
        for check_row in range(max(1, current_row - 3), current_row):
            if self._is_potential_table_title(sheet, check_row, min_col, max_col):
                title_start = check_row
                break
        
        return title_start
    
    def _is_potential_table_title(self, sheet, row: int, min_col: int, max_col: int) -> bool:
        """テーブルタイトルらしい行かどうか判定"""
        try:
            # セルの内容をチェック
            for col in range(min_col, min_col + 5):  # 最初の5列をチェック
                if col > max_col:
                    break
                cell = sheet.cell(row, col)
                if cell.value and isinstance(cell.value, str):
                    text = str(cell.value).strip()
                    # マークダウン強調や太字はタイトル候補として扱う（特定キーワードには依存しない）
                    if text.startswith('**') and text.endswith('**') and len(text) > 4:
                        return True
                    if cell.font and cell.font.bold:
                        return True
            return False
        except (ValueError, TypeError):
            return False
    
    def _row_has_content(self, sheet, row: int, min_col: int, max_col: int) -> bool:
        """行にコンテンツがあるかチェック"""
        try:
            for col in range(min_col, max_col + 1):
                cell = sheet.cell(row, col)
                if cell.value is not None:
                    return True
            return False
        except Exception:
            return False
    
    def _is_table_separator_row(self, sheet, row: int, min_col: int, max_col: int) -> bool:
        """テーブル区切り行かどうか判定"""
        # 空行が連続している場合はテーブル区切りとみなす
        try:
            # 前後の行もチェック
            for check_row in [row - 1, row, row + 1]:
                if check_row < 1:
                    continue
                has_content = False
                for col in range(min_col, min(min_col + 10, max_col + 1)):  # 最初の10列をチェック
                    cell = sheet.cell(check_row, col)
                    if cell.value is not None:
                        has_content = True
                        break
                if has_content:
                    return False
            return True
        except Exception:
            return True
    
    def _row_has_data(self, sheet, row_num: int, min_col: int, max_col: int) -> bool:
        """指定行にデータがあるかチェック"""
        if row_num > sheet.max_row:
            return False
            
        for col_num in range(min_col, max_col + 1):
            cell = sheet.cell(row_num, col_num)
            if cell.value is not None and str(cell.value).strip():
                return True
        return False
    
    def _get_table_column_range(self, sheet, start_row: int, end_row: int, min_col: int, max_col: int) -> Optional[Tuple[int, int]]:
        """テーブルの実際の列範囲を取得（罫線情報を考慮）"""
        # 罫線ベースの列検出を試行
        border_based_range = self._detect_table_columns_by_borders(sheet, start_row, end_row, min_col, max_col)
        if border_based_range:
            return border_based_range
        
        # フォールバック：データベースの列検出
        actual_min_col = None
        actual_max_col = None
        
        for row_num in range(start_row, end_row + 1):
            row_range = self._get_row_column_range(sheet, row_num, min_col, max_col)
            if row_range:
                actual_min_col, actual_max_col = self._update_column_bounds(
                    actual_min_col, actual_max_col, row_range[0], row_range[1]
                )
        
        return (actual_min_col, actual_max_col) if actual_min_col is not None else None
    
    def _detect_table_columns_by_borders(self, sheet, start_row: int, end_row: int, min_col: int, max_col: int) -> Optional[Tuple[int, int]]:
        """罫線情報を使ってテーブルの列範囲を検出（左罫線・右罫線を正確に判定）"""
        debug_print(f"[DEBUG] 列範囲検出: 行{start_row}〜{end_row}, 列{min_col}〜{max_col}")
        
        # 左境界の検出：列の左罫線または前列の右罫線をチェック
        table_start_col = None
        for col in range(min_col, max_col + 1):
            border_count = 0
            total_cells = 0
            
            for row in range(start_row, min(start_row + 5, end_row + 1)):
                cell = sheet.cell(row, col)
                total_cells += 1
                
                # 現在列の左罫線をチェック
                if self._has_strong_left_border(cell):
                    border_count += 1
                # 前の列の右罫線もチェック
                elif col > min_col:
                    prev_cell = sheet.cell(row, col - 1)
                    if self._has_strong_right_border(prev_cell):
                        border_count += 1
            
            # 50%以上のセルに境界線がある場合はテーブル開始
            if total_cells > 0 and (border_count / total_cells) >= 0.5:
                table_start_col = col
                debug_print(f"[DEBUG] テーブル開始列検出: 列{col} (境界線密度: {border_count}/{total_cells})")
                break
        
        # 右境界の検出：列の右罫線または次列の左罫線をチェック
        table_end_col = None
        for col in range(max_col, min_col - 1, -1):  # 逆順でチェック
            border_count = 0
            total_cells = 0
            
            for row in range(start_row, min(start_row + 5, end_row + 1)):
                cell = sheet.cell(row, col)
                total_cells += 1
                
                # 現在列の右罫線をチェック
                if self._has_strong_right_border(cell):
                    border_count += 1
                # 次の列の左罫線もチェック
                elif col < max_col:
                    next_cell = sheet.cell(row, col + 1)
                    if self._has_strong_left_border(next_cell):
                        border_count += 1
            
            # 50%以上のセルに境界線がある場合はテーブル終了
            if total_cells > 0 and (border_count / total_cells) >= 0.5:
                table_end_col = col
                debug_print(f"[DEBUG] テーブル終了列検出: 列{col} (境界線密度: {border_count}/{total_cells})")
                break
        
        if table_start_col is not None and table_end_col is not None and table_start_col <= table_end_col:
            debug_print(f"[DEBUG] 罫線ベース列範囲: 列{table_start_col}〜{table_end_col}")
            return (table_start_col, table_end_col)
        
        debug_print("[DEBUG] 罫線ベース列検出失敗")
        return None
    
    def _has_table_borders(self, cell) -> bool:
        """セルに表らしい罫線があるかチェック"""
        try:
            if not cell.border:
                return False
            
            # 上下左右のいずれかに罫線があるかチェック
            borders = [
                cell.border.left,
                cell.border.right,
                cell.border.top,
                cell.border.bottom
            ]
            
            border_count = 0
            for border in borders:
                if border and border.style:
                    border_count += 1
            
            # 2つ以上の辺に罫線がある場合は表の一部とみなす
            return border_count >= 2
            
        except Exception:
            return False
    
    def _get_row_column_range(self, sheet, row_num: int, min_col: int, max_col: int) -> Optional[Tuple[int, int]]:
        """1行の列範囲を取得"""
        row_min_col = None
        row_max_col = None
        
        for col_num in range(min_col, max_col + 1):
            cell = sheet.cell(row_num, col_num)
            if cell.value is not None or self._has_cell_formatting(cell):
                if row_min_col is None:
                    row_min_col = col_num
                row_max_col = col_num
        
        return (row_min_col, row_max_col) if row_min_col is not None else None
    
    def _update_column_bounds(self, current_min: Optional[int], current_max: Optional[int],
 new_min: int, new_max: int) -> Tuple[int, int]:
        """列の境界を更新"""
        updated_min = new_min if current_min is None or new_min < current_min else current_min
        updated_max = new_max if current_max is None or new_max > current_max else current_max
        return updated_min, updated_max
    
    def _convert_single_table(self, sheet, min_row: int, max_row: int, min_col: int, max_col: int):
        """単一テーブルとして変換（従来の処理）"""
        table_data = []
        
        for row_num in range(min_row, max_row + 1):
            row_data = []
            for col_num in range(min_col, max_col + 1):
                cell = sheet.cell(row_num, col_num)
                cell_content = self._format_cell_content(cell)
                row_data.append(cell_content)
            table_data.append(row_data)
        
        if table_data:
            # 出力前にデバッグ用にtable_dataをダンプ
            try:
                cols = max(len(r) for r in table_data) if table_data else 0
            except Exception:
                cols = 0
            debug_print(f"[DEBUG] _output_markdown_table called (single_table path): rows={len(table_data)}, max_cols={cols}")
            for i, r in enumerate(table_data[:10]):
                debug_print(f"[DEBUG] table_data row {i} cols={len(r)}: {r}")
            # min_row..max_rowの仮定からsource_rowsを順次構築
                try:
                    source_rows = list(range(min_row, max_row + 1))[:len(table_data)]
                except (ValueError, TypeError):
                    source_rows = None
                # シートで既に出力済みの行を削除（データ前の行）
                try:
                    debug_print(f"[DEBUG][_prune_call_single] sheet={sheet.title} before_prune rows={len(table_data) if table_data else 0} source_rows_sample={source_rows[:10] if source_rows else None}")
                    table_data, source_rows = self._prune_emitted_rows(sheet.title, table_data, source_rows)
                    debug_print(f"[DEBUG][_prune_result_single] sheet={sheet.title} after_prune rows={len(table_data) if table_data else 0} source_rows_sample={source_rows[:10] if source_rows else None}")
                except (ValueError, TypeError) as e:
                    debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
                # デバッグ用の出力前決定論的ダンプ: table_dataとsource_rowsの小さなプレビューをキャプチャ
                try:
                    src_sample = source_rows[:10] if source_rows else None
                    rows_len = len(table_data) if table_data else 0
                    debug_print(f"[DEBUG][_pre_output_call] path=single_table sheet={sheet.title} rows={rows_len} source_rows_sample={src_sample}")
                except (ValueError, TypeError) as e:
                    debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
                # 正規パスまでテーブル出力を遅延させ、権威的マッピングが
                # そのパス中にのみ記録されるようにする。最初のソース行を
                # アンカーとして使用。後方互換性のある形状のために
                # オプションのメタデータ（このパスではタイトルなし）を含める。
                try:
                    anchor = (source_rows[0] if source_rows else min_row)
                except (ValueError, TypeError):
                    anchor = min_row
                try:
                    meta = None
                    self._sheet_deferred_tables.setdefault(sheet.title, []).append((anchor, table_data, source_rows, meta))
                    debug_print(f"DEFER_TABLE single_table sheet={sheet.title} anchor={anchor} rows={len(table_data)}")
                except (ValueError, TypeError):
                    # 失敗時はデータ損失を避けるため即時出力にフォールバック
                    self._output_markdown_table(table_data, source_rows=source_rows, sheet_title=sheet.title)
    
    def _convert_table_region(self, sheet, region: Tuple[int, int, int, int], table_number: int,
                              strict_column_bounds: bool = False):
        """指定された領域をテーブルとして変換（結合セル対応、ヘッダー行検出）
        
        Args:
            strict_column_bounds: Trueの場合、列範囲の拡張を制限（離散データ領域検出用）
        """
        start_row, end_row, start_col, end_col = region
        # 診断エントリログ: 領域と生セル値の小さなサンプルを出力
        try:
            debug_print(f"[DEBUG][_convert_table_region_entry] sheet={getattr(sheet, 'title', None)} region={start_row}-{end_row},{start_col}-{end_col}")
            # テーブルが検出されたかどうかを識別するために最大5行の生値をダンプ
            max_dump = min(5, end_row - start_row + 1)
            for rr in range(start_row, start_row + max_dump):
                rowvals = []
                for cc in range(start_col, end_col + 1):
                    try:
                        v = sheet.cell(rr, cc).value
                    except (ValueError, TypeError):
                        v = None
                    rowvals.append((cc, v))
                debug_print(f"[DEBUG][_convert_table_region_entry] raw row {rr}: {rowvals}")
        except (ValueError, TypeError) as e:
            debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
        
        # 小さすぎるテーブル（1-2行のみ）で、タイトルのみを含む場合はスキップ
        if end_row - start_row <= 1:
            # この領域がタイトルのみかチェック
            title_text = self._find_table_title_in_region(sheet, region, strict_column_bounds)
            if title_text:
                # タイトルのみの小さなテーブルはスキップ
                debug_print(f"[DEBUG] タイトルのみの小さなテーブルをスキップ: '{title_text}' at 行{start_row}-{end_row}")
                return
        
        # 非表形式のテキスト（対象分析装置など）をチェック
        if self._is_plain_text_region(sheet, region):
            debug_print(f"[DEBUG] 非表形式テキストとして処理: 行{start_row}-{end_row}")
            self._convert_plain_text_region(sheet, region)
            return
        
        # ヘッダー行を検出
        header_info = self._find_table_header_row(sheet, region)
        header_row = None
        header_height = 1
        if header_info:
            header_row, header_height = header_info
        
        # テーブルタイトルを常に検出（OnlineQC、StartupReportなど）
        # strict_column_boundsがTrueの場合は列範囲を制限
        title_text = self._find_table_title_in_region(sheet, region, strict_column_bounds)
        
        # この領域のタイトルテキストを検出した場合、ローカルに保持し、
        # table_data構築後に遅延テーブルメタデータに添付
        # これにより、タイトルを別の遅延テキストエントリ（_sheet_deferred_texts内）
        # として出力することを避け、重複抑制と順序付けを複雑にしない。
        safe_title = None
        if title_text:
            safe_title = self._escape_angle_brackets(str(title_text))

        # 検出されたタイトルが実際に領域の一部（先頭行に表示）の場合、
        # テーブル構築時にその行をスキップし、タイトルがヘッダーセルとして誤解されないようにする。
        try:
            found_title_in_region = False
            for c in range(start_col, end_col + 1):
                cell_val = sheet.cell(start_row, c).value
                if cell_val and str(cell_val).strip() == str(title_text).strip():
                    found_title_in_region = True
                    break
            if found_title_in_region:
                debug_print(f"[DEBUG] タイトル行が領域先頭に含まれているためスキップ: '{title_text}' at 行{start_row}")
                start_row = start_row + 1
                region = (start_row, end_row, start_col, end_col)
        except (ValueError, TypeError) as e:
            debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
        
        # 結合セル情報を取得
        merged_cells = self._get_merged_cell_info(sheet, region)
        # table_dataは以下の一部の条件分岐でのみ割り当てられる可能性がある; 定義を確保
        # 後続コードで`if table_data:`をチェックする際のUnboundLocalErrorを回避
        table_data = None

        # フォールバック: ヘッダー行が無い場合、領域内の非空セルが存在する列の集合を使って
        # テーブルを組み立てる（最大8列）。これにより、見た目上複数列に分かれている表を復元する。
        if not header_row:
            unique_cols = [c for c in range(start_col, end_col + 1)
                           if any((sheet.cell(r, c).value is not None and str(sheet.cell(r, c).value).strip())
                                  for r in range(start_row, end_row + 1))]
            # 限度: 2〜8列のときのみ適用
            if 1 < len(unique_cols) <= 8:
                # ヒューリスティック: 左端のユニーク列が繰り返しセクションラベル（多くの行で同じ値）で、
                # 次の列が異なるプロパティ名を含む場合、左端の列を削除して
                # プロパティ列が最初のデータ列として使用されるようにする。これにより、
                # 最初の出力列が'TransferFileList'のようなセクション名で埋められることを防ぐ。
                try:
                    col_stats = []
                    for c in unique_cols:
                        values = [str(sheet.cell(r, c).value).strip() for r in range(start_row, end_row + 1) if sheet.cell(r, c).value]
                        distinct = len(set(values))
                        nonempty = len(values)
                        col_stats.append({'col': c, 'distinct': distinct, 'nonempty': nonempty})
                    if len(col_stats) >= 2:
                        left = col_stats[0]
                        right = col_stats[1]
                        # 左が単一の繰り返し非空値を持ち、右が複数の異なる非空値を持つ場合、
                        # かつ左が行の95%未満に存在する場合（真のデータ列の削除を避けるため）、左を削除
                        total_rows = end_row - start_row + 1
                        if left['distinct'] == 1 and right['distinct'] > 1 and left['nonempty'] / max(1, total_rows) < 0.95:
                            debug_print(f"[DEBUG] unique_cols heuristic: dropping left repeated column {left['col']} in favor of {right['col']}")
                            unique_cols = unique_cols[1:]
                except (ValueError, TypeError) as e:
                    debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
                # 行ごとの平均非空セル数（対象列内）
                total_rows = end_row - header_row + 1 if header_row else end_row - start_row + 1
                row_counts = []
                for r in range(start_row, end_row + 1):
                    cnt = sum(1 for c in unique_cols if (sheet.cell(r, c).value is not None and str(sheet.cell(r, c).value).strip()))
                    row_counts.append(cnt)
                avg_nonempty = sum(row_counts) / len(row_counts) if row_counts else 0
                # 平均が0.5以上なら列ベース表と見なす
                if avg_nonempty >= 0.5:
                    debug_print(f"CONV_UNIQUECOLS sheet={getattr(sheet,'title',None)} region={region} unique_cols={unique_cols} avg_nonempty={avg_nonempty:.2f}")
                    table_data = []
                    source_rows = []
                    # ヘッダー行があればヘッダーとして使う（短いテキスト行）、なければ空ヘッダー
                    first_row_vals = [str(sheet.cell(start_row, c).value).strip() if sheet.cell(start_row, c).value else '' for c in unique_cols]
                    # 判定: 最初の行がヘッダーっぽい（全て短いテキストかつ複数非空）ならヘッダー行として使う
                    nonempty_in_first = sum(1 for v in first_row_vals if v)
                    if nonempty_in_first >= max(1, len(unique_cols)//3) and all(len(v) < 120 for v in first_row_vals if v):
                        # first_row_valsはヘッダー行として扱われるが、時々
                        # 空のエントリ（例: ['', '']）が含まれ、左の
                        # 非空ヘッダー（'名前'など）にマージすべき。そのような空ヘッダー列を
                        # 左隣にマージして空ヘッダー列の生成を回避。
                        headers_candidate = list(first_row_vals)

                        # マージする列を決定: ヘッダーが空で左ヘッダーが存在する場合
                        merge_into_left = set()
                        for idx in range(1, len(headers_candidate)):
                            if not headers_candidate[idx] and headers_candidate[idx-1]:
                                merge_into_left.add(idx)

                        if merge_into_left:
                            # 元のunique_colsインデックスから新しい列へのマッピングを構築
                            new_unique_cols = []
                            merge_map = {}  # col_idx -> target_new_indexのマッピング
                            new_idx = 0
                            for idx, col in enumerate(unique_cols):
                                if idx in merge_into_left:
                                    # この列を前のnew_idx-1にマージ
                                    merge_map[idx] = new_idx - 1
                                else:
                                    new_unique_cols.append(col)
                                    merge_map[idx] = new_idx
                                    new_idx += 1

                            # マージされた列からテキストをマージして新しい列のヘッダー行を構築
                            new_headers = []
                            for old_idx, col in enumerate(unique_cols):
                                target = merge_map[old_idx]
                                while len(new_headers) <= target:
                                    new_headers.append('')
                                val = headers_candidate[old_idx] or ''
                                if new_headers[target]:
                                    if val:
                                        new_headers[target] = (new_headers[target] + ' ' + val).strip()
                                else:
                                    new_headers[target] = val

                            # unique_colsとヘッダー行をマージ版に置換
                            unique_cols = new_unique_cols
                            first_row_vals = new_headers

                        table_data.append(first_row_vals)
                        source_rows.append(start_row)
                        data_start_row = start_row + 1
                    else:
                        data_start_row = start_row

                    for r in range(data_start_row, end_row + 1):
                        row_vals = [str(sheet.cell(r, c).value).strip() if sheet.cell(r, c).value else '' for c in unique_cols]
                        # 出力する行は少なくとも1つの非空を含むこと
                        if any(v for v in row_vals):
                            table_data.append(row_vals)
                            source_rows.append(r)
            if table_data:
                debug_print(f"[DEBUG] unique_cols-based table used: cols={unique_cols}, rows={len(table_data)}")
                # 追加ダンプ: unique_cols フォールバック時の内部状態確認
                # 安全なダンプ: このスコープに存在しない変数がある可能性（header_positionsなど）
                try:
                    ctx = {}
                    ctx['unique_cols'] = unique_cols
                    ctx['table_data_rows'] = len(table_data)
                    # header_positions/final_groups/compressed_headersはここで定義されていない可能性
                    if 'header_positions' in locals():
                        ctx['header_positions'] = header_positions
                    if 'final_groups' in locals():
                        ctx['final_groups'] = final_groups
                    if 'compressed_headers' in locals():
                        ctx['compressed_headers'] = compressed_headers
                    debug_print(f"[DEBUG-DUMP] unique_cols context: {ctx}")
                    for i, r in enumerate(table_data[:5]):
                        debug_print(f"[DEBUG-DUMP] unique_cols table_data row {i}: {r}")
                except Exception as _e:
                    debug_print(f"[DEBUG-DUMP] failed unique_cols dump: {_e}")
                # コンパクトな機械可読トレースをファイルに書き込み（デバッグログが利用可能な場合）
                try:
                    sheet_name = getattr(sheet, 'title', None)
                    first_row_sample = first_row_vals[:8] if 'first_row_vals' in locals() else None
                    merge_info_sample = None
                    if 'merge_into_left' in locals():
                        merge_info_sample = sorted(list(merge_into_left))
                except (ValueError, TypeError) as e:
                    debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
                # シート固有の追加ダンプ: 'XMLファイル自動生成' の場合はより詳細に出力
                try:
                    sheet_name = getattr(sheet, 'title', None)
                    title_in_region = None
                    try:
                        title_in_region = self._find_table_title_in_region(sheet, region)
                    except Exception:
                        title_in_region = None
                    if sheet_name == 'XMLファイル自動生成' or title_in_region == 'XMLファイル自動生成':
                        debug_print('[DEBUG-TRACE] Detected target sheet/region for deep dump: XMLファイル自動生成')
                        debug_print(f"[DEBUG-TRACE] region={region}")
                        debug_print(f"[DEBUG-TRACE] unique_cols={unique_cols}")
                        # first_row_valsが存在する場合はダンプ
                        if 'first_row_vals' in locals():
                            debug_print(f"[DEBUG-TRACE] first_row_vals={first_row_vals}")
                        if 'merge_into_left' in locals():
                            debug_print(f"[DEBUG-TRACE] merge_into_left={merge_into_left}")
                        if 'merge_map' in locals():
                            debug_print(f"[DEBUG-TRACE] merge_map={merge_map}")
                        # ヘッダー関連の構造体が存在する場合
                        for name in ('header_positions', 'final_groups', 'compressed_headers', 'group_positions'):
                            if name in locals():
                                debug_print(f"[DEBUG-TRACE] {name}={locals()[name]}")
                        # クロスチェック用に領域の生のセル値をいくつかダンプ
                        try:
                            for rr in range(region[0], min(region[0]+6, region[1]+1)):
                                rowvals = []
                                for c in range(region[2], region[3]+1):
                                    try:
                                        v = sheet.cell(rr, c).value
                                    except (ValueError, TypeError):
                                        v = None
                                    rowvals.append((c, v))
                                debug_print(f"[DEBUG-TRACE] raw row {rr}: {rowvals}")
                        except Exception as _e:
                            debug_print(f"[DEBUG-TRACE] failed to dump raw rows: {_e}")
                except Exception as _e:
                    debug_print(f"[DEBUG-TRACE] deep dump failed: {_e}")
                # 列ヘッダーが無ければプレーンな表として出力
                debug_print(f"[DEBUG] 出力前テーブルプレビュー(unique_cols): rows={len(table_data)}, first_row={table_data[0] if table_data else None}")
                # デバッグ用にtable_dataの形状と最初の行をダンプ
                try:
                    cols = max(len(r) for r in table_data) if table_data else 0
                except (ValueError, TypeError):
                    cols = 0
                debug_print(f"[DEBUG] _output_markdown_table called (unique_cols path): rows={len(table_data)}, max_cols={cols}")
                for i, r in enumerate(table_data[:10]):
                    debug_print(f"[DEBUG] table_data row {i} cols={len(r)}: {r}")
                try:
                    # 以前の行と重複する可能性のある事前出力行を削除
                    debug_print(f"[DEBUG][_prune_call_unique] sheet={sheet.title} before_prune rows={len(table_data) if table_data else 0} source_rows_sample={source_rows[:10] if source_rows else None}")
                    table_data, source_rows = self._prune_emitted_rows(sheet.title, table_data, source_rows)
                    debug_print(f"[DEBUG][_prune_result_unique] sheet={sheet.title} after_prune rows={len(table_data) if table_data else 0} source_rows_sample={source_rows[:10] if source_rows else None}")
                except (ValueError, TypeError) as e:
                    debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
                # デバッグ用の出力前決定論的ダンプ（unique_colsパス）
                try:
                    src_sample = source_rows[:10] if source_rows else None
                    rows_len = len(table_data) if table_data else 0
                    debug_print(f"[DEBUG][_pre_output_call] path=unique_cols sheet={getattr(sheet, 'title', None)} rows={rows_len} source_rows_sample={src_sample}")
                except (ValueError, TypeError) as e:
                    debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
                try:
                    # 正規パスまでテーブル出力を遅延。利用可能な場合は最初のソース行を
                    # アンカーとして使用し、ここではタイトルメタを含めない。
                    try:
                        anchor = (source_rows[0] if source_rows else start_row)
                    except (ValueError, TypeError):
                        anchor = start_row
                    try:
                        meta = None
                        self._sheet_deferred_tables.setdefault(sheet.title, []).append((anchor, table_data, source_rows, meta))
                        debug_print(f"DEFER_TABLE unique_cols sheet={sheet.title} anchor={anchor} rows={len(table_data)}")
                    except (ValueError, TypeError):
                        # データ損失を避けるため失敗時は即時出力にフォールバック
                        try:
                            self._output_markdown_table(table_data, source_rows=source_rows)
                        except (ValueError, TypeError):
                            self._output_markdown_table(table_data)
                except (ValueError, TypeError):
                    # 外側のtry - 他に失敗した場合、直接出力を試行
                    try:
                        self._output_markdown_table(table_data)
                    except Exception as e:
                        pass  # XML解析エラーは無視
                return

        # テーブルデータを結合セル考慮で構築
        if header_row:
            # ヘッダー行を考慮した構築
            table_data = self._build_table_with_header_row(sheet, region, header_row, merged_cells, header_height=header_height, strict_column_bounds=strict_column_bounds)
            # ヘッダー行から開始するため、approx_rowsもheader_rowから計算
            actual_start_row = header_row
        else:
            # 従来の方法
            table_data = self._build_table_data_with_merges(sheet, region, merged_cells, strict_column_bounds=strict_column_bounds)
            actual_start_row = start_row
        
        if table_data:
            debug_print(f"[DEBUG] 出力前テーブルプレビュー: rows={len(table_data)}, first_row={table_data[0] if table_data else None}")
            # 出力前にtable_dataの形状をダンプ
            try:
                cols = max(len(r) for r in table_data) if table_data else 0
            except (ValueError, TypeError):
                cols = 0
            debug_print(f"[DEBUG] _output_markdown_table called (header/data path): rows={len(table_data)}, max_cols={cols}")
            for i, r in enumerate(table_data[:10]):
                debug_print(f"[DEBUG] table_data row {i} cols={len(r)}: {r}")
            try:
                # actual_start_rowから開始（header_rowまたはstart_row）
                # regionのend_rowを使用し実際のテーブル範囲全体をカバー
                # これにより、table_dataから除外された行(空行など)も含めて、
                # テーブル領域全体がprocessed_rowsとして記録される
                approx_rows = list(range(actual_start_row, region[1] + 1))  # region[1]はend_row
            except (ValueError, TypeError):
                approx_rows = None
            try:
                debug_print(f"[DEBUG][_prune_call_headerdata] sheet={sheet.title} before_prune rows={len(table_data) if table_data else 0} approx_rows_sample={approx_rows[:10] if approx_rows else None}")
                table_data, approx_rows = self._prune_emitted_rows(sheet.title, table_data, approx_rows)
                debug_print(f"[DEBUG][_prune_result_headerdata] sheet={sheet.title} after_prune rows={len(table_data) if table_data else 0} approx_rows_sample={approx_rows[:10] if approx_rows else None}")
            except (ValueError, TypeError) as e:
                debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
            # デバッグ用の出力前決定論的ダンプ（ヘッダー/データパス）
            try:
                src_sample = approx_rows[:10] if approx_rows else None
                rows_len = len(table_data) if table_data else 0
                debug_print(f"[DEBUG][_pre_output_call] path=header_data sheet={sheet.title} rows={rows_len} source_rows_sample={src_sample}")
            except (ValueError, TypeError) as e:
                debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
            # 正規パスまでテーブル出力を遅延させ、権威的マップが
            # そのパス中にのみ記録されるようにする。アンカー行 = 最初のソース行を保存
            try:
                # 利用可能な場合は検出されたタイトル行をテーブルアンカーとして優先。
                title_anchor = getattr(self, '_last_table_title_row', None) if safe_title else None
                if title_anchor and isinstance(title_anchor, int):
                    anchor = title_anchor
                else:
                    anchor = (approx_rows[0] if approx_rows else start_row)
            except Exception:
                anchor = start_row
            try:
                # 遅延テーブルにオプションのメタデータ（タイトル）を含め、
                # 正規エミッターがタイトルとテーブルを単一のアトミックイベントで
                # 出力できるようにする。遅延テーブルの後方互換性のある形状:
                # (anchor, table_data, approx_rows) -> (anchor, table_data, approx_rows, meta_dict)
                meta = {'title': safe_title} if safe_title else None
                self._sheet_deferred_tables.setdefault(sheet.title, []).append((anchor, table_data, approx_rows, meta))
                # 延期後に一時的なタイトル行をクリア
                try:
                    self._last_table_title_row = None
                except Exception as e:
                    pass  # XML解析エラーは無視
                debug_print(f"DEFER_TABLE sheet={sheet.title} anchor={anchor} rows={len(table_data)} title_present={bool(safe_title)}")
            except (ValueError, TypeError):
                # 延期が失敗した場合は即時出力にフォールバック
                try:
                    self._output_markdown_table(table_data, source_rows=approx_rows, sheet_title=sheet.title)
                except (ValueError, TypeError) as e:
                    debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
            # テーブル右隣の記述的テキストを検出・出力 (this will be deferred by _emit_free_text)
            # _last_group_positionsが存在する場合は、実際に使用された最大列を使用
            try:
                if hasattr(self, '_last_group_positions') and self._last_group_positions:
                    actual_end_col = max(self._last_group_positions)
                    debug_print(f"[DEBUG] _output_right_side_plain_text: actual_end_col={actual_end_col} (from group_positions={self._last_group_positions})")
                else:
                    actual_end_col = end_col
                    debug_print(f"[DEBUG] _output_right_side_plain_text: actual_end_col={actual_end_col} (from region end_col)")
            except Exception as e:
                actual_end_col = end_col
                debug_print(f"[DEBUG] _output_right_side_plain_text: actual_end_col={actual_end_col} (exception: {e})")
            self._output_right_side_plain_text(sheet, region, actual_end_col)
        else:
            self.markdown_lines.append("*空のテーブル*")
            self.markdown_lines.append("")

    def _output_right_side_plain_text(self, sheet, region: Tuple[int, int, int, int], actual_end_col: int = None):
        """テーブル領域の右隣にある記述的テキストを検出・出力"""
        start_row, end_row, start_col, end_col = region
        # 実際に使用された最終列が指定されている場合はそれを使用
        if actual_end_col is not None:
            end_col = actual_end_col
        max_col = sheet.max_column
        debug_print(f"[DEBUG] _output_right_side_plain_text: rows={start_row}-{end_row}, cols={end_col+1}-{max_col}")
        for row_num in range(start_row, end_row + 1):
            right_texts = []
            for col_num in range(end_col + 1, max_col + 1):
                cell = sheet.cell(row=row_num, column=col_num)
                if cell.value is not None:
                    text = str(cell.value).strip()
                    if text:
                        right_texts.append(text)
                        debug_print(f"[DEBUG] _output_right_side_plain_text: 行{row_num}列{col_num} text='{text}'")
            # 右側にテキストがあれば出力
            if right_texts:
                # 重複と出力済み行を追跡するため集中エミッタ経由で出力
                for text in right_texts:
                    try:
                        self._emit_free_text(sheet, row_num, text)
                    except (ValueError, TypeError):
                        # エミッタが何らかの理由で失敗した場合は直接追加にフォールバック
                        self.markdown_lines.append(f"{text}  ")
        # テーブル右隣のテキストがあれば空行で区切る
        if any(sheet.cell(row=row_num, column=col_num).value for row_num in range(start_row, end_row + 1) for col_num in range(end_col + 1, max_col + 1)):
            self.markdown_lines.append("")
    
    def _is_plain_text_region(self, sheet, region: Tuple[int, int, int, int]) -> bool:
        """領域が通常のテキスト（非表形式）かどうかを判定"""
        start_row, end_row, start_col, end_col = region
        # 早期デバッグ: エントリと単純なメトリクスを報告
        try:
            rows = end_row - start_row + 1
            cols = end_col - start_col + 1
            debug_print(f"[DEBUG][_is_plain_text_region_entry] sheet={getattr(sheet,'title',None)} region={start_row}-{end_row},{start_col}-{end_col} rows={rows} cols={cols}")
        except (ValueError, TypeError) as e:
            debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
        
        # 領域のサイズが小さい場合（行数が少ない）
        row_count = end_row - start_row + 1
        col_count = end_col - start_col + 1
        
        # この領域のすべてのテキスト内容を収集
        texts = []
        non_empty_cells = 0
        total_cells = 0
        
        for row_num in range(start_row, end_row + 1):
            for col_num in range(start_col, end_col + 1):
                cell = sheet.cell(row_num, col_num)
                total_cells += 1
                if cell.value:
                    text = str(cell.value).strip()
                    if text:
                        texts.append(text)
                        non_empty_cells += 1
        
        # データが1セルでもあれば判定対象
        if non_empty_cells < 1:
            return False
        # デバッグ: この領域の計算されたヒューリスティックを報告
        debug_print(f"PLAIN_ENTRY sheet={getattr(sheet,'title',None)} region={start_row}-{end_row},{start_col}-{end_col} non_empty={non_empty_cells} total={total_cells}")
        text_content = ' '.join(texts)
        
        avg_len = sum(len(t) for t in texts) / non_empty_cells if non_empty_cells > 0 else 0
        
        # トークンベースのヒューリスティック: 複数の短いトークンを含む単一行
        # はコンパクトなテーブルヘッダーまたはデータ行の可能性が高い（例: "名前 初期値 設定値"）
        # 保守的に: 平均セル長が大きすぎないことを要求し、
        # 説明文をテーブルとして誤分類しないようにする。
        try:
            tokens = [tok for tok in text_content.split() if tok]
            if row_count == 1 and len(tokens) >= 2 and avg_len <= 60:
                debug_print(f"[DEBUG] 単一行トークン複数 -> 表扱い: 行{start_row}〜{end_row}, tokens={len(tokens)}, avg_len={avg_len:.1f}")
                return False
        except (ValueError, TypeError) as e:
            debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
        # プレーンテキスト判定: キーワードベースを廃止し、汎用的な構造的ヒューリスティックを使用する
        # - ファイルパス・URL・XMLやタグなどの記述的コンテンツが多い -> プレーンテキスト
        # - セルの平均長が大きい（長文が多い） -> プレーンテキスト
        # - 列ごとの非空セル分布が均一で、各行に同程度の列数のデータがある -> 表形式
        long_count = sum(1 for t in texts if len(t) > 120)
        path_like_count = sum(1 for t in texts if ('\\' in t and ':' in t) or '/' in t or t.lower().startswith('http') or 'xml' in t.lower() or ('<' in t and '>' in t))

        # 列ごとの非空セル数を数える（構造性の指標）
        col_nonempty = {c: 0 for c in range(start_col, end_col + 1)}
        row_nonempty = {r: 0 for r in range(start_row, end_row + 1)}
        for r in range(start_row, end_row + 1):
            for c in range(start_col, end_col + 1):
                v = sheet.cell(r, c).value
                if v is not None and str(v).strip():
                    col_nonempty[c] += 1
                    row_nonempty[r] += 1

        # 列のうち非空セルがある列数と、行ごとの非空カウントの分散を計算
        cols_with_content = sum(1 for v in col_nonempty.values() if v > 0)
        import statistics
        row_counts = [row_nonempty[r] for r in row_nonempty]
        row_std = statistics.pstdev(row_counts) if len(row_counts) > 0 else 0
        avg_row_nonempty = sum(row_counts) / len(row_counts) if len(row_counts) > 0 else 0

        # 例外: 2列の番号付きリストパターン -> プレーンテキストとして扱う
        # 左列が主に番号/マーカー（①、1、A、iなど）を含み、
        # 右列がより長い説明テキストを含む場合、領域をテーブルではなく
        # 番号付きリスト/説明行として扱うことを優先。
        try:
            content_cols = sorted([c for c, v in col_nonempty.items() if v > 0])
            if len(content_cols) == 2:
                left_col, right_col = content_cols[0], content_cols[1]
                left_texts = []
                right_texts = []
                for r in range(start_row, end_row + 1):
                    try:
                        lv = sheet.cell(r, left_col).value
                        rv = sheet.cell(r, right_col).value
                    except Exception:
                        lv = None
                        rv = None
                    if lv is not None and str(lv).strip():
                        left_texts.append(str(lv).strip())
                    if rv is not None and str(rv).strip():
                        right_texts.append(str(rv).strip())

                if left_texts and right_texts and len(left_texts) >= 2:
                    import re
                    import unicodedata
                    num_matches = 0
                    # 一般的な丸数字を明示的に含める（①〜⑳）
                    circled = '①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳'
                    for t in left_texts:
                        tt = t.strip()
                        # 全角数字/句読点をASCII相当に正規化
                        try:
                            nn = unicodedata.normalize('NFKC', tt)
                        except Exception:
                            nn = tt

                        # 丸数字を最初にチェック（ASCIIに正規化されない）
                        if any(ch in circled for ch in tt):
                            num_matches += 1
                            continue

                        # 以下のパターンを受け入れる:
                        #  - (1) / （1） / 1) / 1）
                        #  - 1. / 1．
                        #  - 1 / １ (NFKCで正規化された全角)
                        #  - (a) / a)
                        #  - ローマ数字 I, II, III（オプションで句読点付き）
                        # 正規化された文字列を正規表現に使用し、全角句読点を処理
                        # オプションの括弧（ASCIIと全角の両方）と
                        # オプションの末尾句読点（'.'や'．'など）を許可
                        try:
                            if re.match(r'^[\(\（]?\s*(?:\d+|[IVXivx]+|[A-Za-z])\s*[\)\）]?[\.．]?$', nn):
                                num_matches += 1
                                continue
                        except Exception:
                            pass  # データ構造操作失敗は無視

                        # フォールバック: 単一文字マーカー（例: '-', 'a', '1'）
                        try:
                            if len(nn.strip()) == 1 and re.match(r'^[A-Za-z0-9\-]$', nn.strip()):
                                num_matches += 1
                                continue
                        except Exception:
                            pass  # データ構造操作失敗は無視

                    ratio = (num_matches / len(left_texts)) if left_texts else 0.0
                    right_avg = sum(len(s) for s in right_texts) / len(right_texts) if right_texts else 0
                    # ヒューリスティック閾値: 左の80%以上が番号のようで、右の平均長が10以上
                    if ratio >= 0.8 and right_avg >= 10:
                        debug_print(f"[DEBUG] 番号付きリスト検出: 行{start_row}〜{end_row} 左番号率={num_matches}/{len(left_texts)} 右平均長={right_avg:.1f}")
                        return True
        except (ValueError, TypeError) as e:
            debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")

        # ルール1: ファイルパス/URL/XMLが多い場合はプレーンテキスト（説明的な列）
        if non_empty_cells > 0 and (path_like_count / non_empty_cells) > 0.25:
            # 複数列を示す強い縦罫線がある場合、テーブル解釈を優先
            try:
                border_cols = self._detect_table_columns_by_borders(sheet, start_row, end_row, start_col, end_col)
            except (ValueError, TypeError):
                border_cols = None

            if border_cols:
                debug_print(f"[DEBUG] パス/XML多だが縦罫線で列境界が検出されたため表として扱います: {border_cols}")
                return False

            debug_print(f"[DEBUG] プレーンテキスト判定(パス/XML多): 行{start_row}〜{end_row}, path_like={path_like_count}/{non_empty_cells}")
            return True

        # ルール2: 平均セル長が大きく、行数が少なめなら説明文ブロック
        if row_count <= 8 and avg_len > 60:
            debug_print(f"[DEBUG] プレーンテキスト判定(長文多): 行{start_row}〜{end_row}, avg_len={avg_len:.1f}")
            return True

        # 例外: 複数の短い列を持つ単一行はコンパクトなテーブルを表す可能性が高い
        # 例: 'A  B  C'のような短いラベルの単一行はテーブルとして扱うべき
        try:
            if row_count == 1 and cols_with_content >= 2 and avg_len < 40:
                debug_print(f"[DEBUG] 単一行短文複数列は表扱い: 行{start_row}〜{end_row}, cols_with_content={cols_with_content}, avg_len={avg_len:.1f}")
                return False
        except (ValueError, TypeError) as e:
            debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")

        # ルール3: 非常に少ない行・セルで長文が混在している場合はプレーンテキスト
        if row_count <= 2 and non_empty_cells <= 6 and long_count > 0:
            debug_print(f"[DEBUG] 単純テキスト判定(少行で長文): 行{start_row}〜{end_row}, long_count={long_count}")
            return True

        # ルール4: 列ごとの分布が均一で、各行に複数列のデータがある -> 表形式とみなす
        # 平均非空セル数が2以上かつ行ごとの分散が小さい場合は表
        if avg_row_nonempty >= 2 and row_std <= max(1.5, avg_row_nonempty * 0.6) and cols_with_content >= 2:
            # 表の可能性が高いのでプレーンテキストにはしない
            debug_print(f"[DEBUG] 表構造検出: 行{start_row}〜{end_row}, avg_row_nonempty={avg_row_nonempty:.1f}, row_std={row_std:.2f}, cols_with_content={cols_with_content}")
            return False

        # それ以外は保守的にプレーンテキストと判定しない（表として扱う）
        return False
    
    def _convert_plain_text_region(self, sheet, region: Tuple[int, int, int, int]):
        """非表形式の領域を通常のテキストとして変換（Excelの印刷イメージを保持）"""
        start_row, end_row, start_col, end_col = region
        
        text_lines = []  # 改行を含むテキスト行を収集
        
        # 集中エミッターを使用して各ソース行を単一の結合行として出力
        for row_num in range(start_row, end_row + 1):
            row_texts = []
            for col_num in range(start_col, min(start_col + 10, end_col + 1)):
                cell = sheet.cell(row_num, col_num)
                if cell.value:
                    text = str(cell.value).strip()
                    if text and text not in row_texts:
                        if cell.font and cell.font.bold:
                            text = f"**{text}**"
                        row_texts.append(text)

            if row_texts:
                combined = " ".join(row_texts)
                try:
                    self._emit_free_text(sheet, row_num, combined)
                except (ValueError, TypeError):
                    # エミッタが失敗した場合は直接追加にフォールバック
                    # エミッタが失敗した場合は直接追加にフォールバック
                    # 正規出力パス中でない限り権威的マッピングを変更しない。
                    # 行/テキストを早期に出力済みとしてマークすると、
                    # 正当なテーブル行が削除される原因となった。
                    try:
                        self.markdown_lines.append(self._escape_angle_brackets(combined) + "  ")
                        if getattr(self, '_in_canonical_emit', False):
                            try:
                                # 正規パス中のみ権威的マッピングを記録
                                md_idx = len(self.markdown_lines) - 1
                                self._mark_sheet_map(sheet.title, row_num, md_idx)
                            except Exception as e:
                                pass  # XML解析エラーは無視
                            try:
                                self._mark_emitted_row(sheet.title, row_num)
                            except Exception as e:
                                pass  # XML解析エラーは無視
                            try:
                                self._mark_emitted_text(sheet.title, self._normalize_text(combined))
                            except Exception as e:
                                pass  # XML解析エラーは無視
                        else:
                            # 非正規コンテキスト: 正規パスがインデックスを割り当てる
                            debug_print(f"[TRACE] Skipping authoritative mapping for plain-text fallback row={row_num} (non-canonical)")
                    except (ValueError, TypeError) as e:
                        debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")

        # 行が出力された場合は区切りの空行を追加
        try:
            emitted = self._sheet_emitted_rows.get(sheet.title, set())
            any_emitted = any(r in emitted for r in range(start_row, end_row + 1))
        except (ValueError, TypeError):
            any_emitted = True
        if any_emitted:
            self.markdown_lines.append("")  # セクション区切りの空行を追加
    
    def _build_table_with_header_row(self, sheet, region: Tuple[int, int, int, int], 
                                   header_row: int, merged_info: Dict[str, Any], header_height: int = 1,
                                   strict_column_bounds: bool = False) -> List[List[str]]:
        """ヘッダー行を基にテーブルを正しく構築
        
        Args:
            header_height: ヘッダーの高さ（行数）。_find_table_header_rowから渡される
            strict_column_bounds: Trueの場合、列範囲の拡張を制限（離散データ領域検出用）
        """
        start_row, end_row, start_col, end_col = region
        
        debug_print(f"[DEBUG] ヘッダー行{header_row}でテーブルを構築中...")
        
        # ヘッダー行の実際の行・列範囲を確認し、regionを拡張
        # (header_rowがregion外の場合や、「名前」など範囲外のヘッダーを含めるため)
        # strict_column_boundsがTrueの場合は列範囲の拡張を制限
        actual_start_row = min(start_row, header_row)
        actual_end_row = max(end_row, header_row + header_height - 1)
        
        header_min_col = start_col
        header_max_col = end_col
        if not strict_column_bounds:
            # 列範囲の拡張を許可（従来の動作）
            for col_num in range(1, sheet.max_column + 1):
                cell = sheet.cell(header_row, col_num)
                if cell.value is not None and str(cell.value).strip():
                    header_min_col = min(header_min_col, col_num)
                    header_max_col = max(header_max_col, col_num)
        
        if header_min_col < start_col or header_max_col > end_col or actual_start_row < start_row:
            debug_print(f"[DEBUG] ヘッダー行により範囲を拡張: 行{start_row}-{end_row} → {actual_start_row}-{actual_end_row}, 列{start_col}-{end_col} → {header_min_col}-{header_max_col}")
            start_row = actual_start_row
            end_row = actual_end_row
            start_col = header_min_col
            end_col = header_max_col
            # 拡張された範囲で結合セル情報を再取得
            merged_info = self._get_merged_cell_info(sheet, (start_row, end_row, start_col, end_col))
        
        # ヘッダー行からカラム情報を取得
        headers = []
        header_positions = []

        # ヘッダー高さを使用して複数行を結合したヘッダー文字列を生成する
        # （引数で渡されない場合は、以前のロジックをフォールバックとして使用）
        if header_height is None or header_height <= 0:
            header_height = int(getattr(self, '_detected_header_height', 1) or 1)
        # 上限を3行に制限（保守的）
        header_height = max(1, min(header_height, 3))
        # _output_markdown_tableで使用するために保存
        self._detected_header_height = header_height

        # 結合セルも考慮してヘッダーを検出（複数行を結合）
        for col in range(start_col, end_col + 1):
            parts = []
            for r in range(header_row, min(header_row + header_height, end_row + 1)):
                key = f"{r}_{col}"
                if key in merged_info:
                    m = merged_info[key]
                    master_cell = sheet.cell(m['master_row'], m['master_col'])
                    raw_text = (str(master_cell.value) if master_cell.value is not None else '')
                else:
                    cell = sheet.cell(r, col)
                    raw_text = (str(cell.value) if cell.value is not None else '')

                # 改行を<br>に正規化し、冗長な<br>トークンを折りたたみ/トリム
                try:
                    import re as _re
                    text = raw_text.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '<br>')
                    # 複数の連続した<br>を1つに折りたたむ
                    text = _re.sub(r'(<br>\s*){2,}', '<br>', text)
                    # 先頭/末尾の<br>を削除
                    text = _re.sub(r'^(?:<br>\s*)+', '', text)
                    text = _re.sub(r'(?:\s*<br>)+$', '', text)
                    text = text.strip()
                except Exception:
                    text = raw_text.replace('\n', '<br>').strip() if raw_text else ''

                if text and len(text) <= 120:
                    if not parts or parts[-1] != text:
                        parts.append(text)

            # 繰り返し連結を避けるため連続した重複部分を削除
            dedup_parts = []
            for p in parts:
                if not dedup_parts or dedup_parts[-1] != p:
                    dedup_parts.append(p)
            # データ行に属する可能性が高い部分をフィルタリング（ヘッダー下で頻繁に出現）
            try:
                # この列がヘッダー行内に結合/マスターヘッダーセルを含むかどうかを判定。
                is_master_col = False
                for r in range(header_row, min(header_row + header_height, end_row + 1)):
                    keym = f"{r}_{col}"
                    if keym in merged_info:
                        mi = merged_info[keym]
                        # マスターセルがヘッダー候補行内にあるか、複数の行/列にまたがる場合、
                        # この列をマスター/ヘッダー列として扱い、積極的なデータトークン削除を避ける。
                        mr = int(mi.get('master_row', header_row))
                        mc = int(mi.get('master_col', col))
                        span_r = int(mi.get('span_rows', 1) or 1)
                        span_c = int(mi.get('span_cols', 1) or 1)
                        if (header_row <= mr < header_row + header_height) or span_r > 1 or span_c > 1 or (mc != col):
                            is_master_col = True
                            break

                filtered_parts = []
                # この列がマスター/ヘッダー列の場合、または複数行ヘッダー（height>1）の場合、
                # サンプリングベースの削除をスキップし、dedup_partsをそのまま保持
                if is_master_col or header_height > 1:
                    filtered_parts = list(dedup_parts)
                else:
                    for p in dedup_parts:
                        # 非常に短いトークンを保持（例: '-', 単一文字マーカー）
                        if not p or len(p.strip()) <= 2:
                            filtered_parts.append(p)
                            continue

                        # ヘッダー領域下の行でこのトークンの出現回数をカウント
                        cnt = 0
                        total = 0
                        for rr in range(header_row + header_height, end_row + 1):
                            if rr > sheet.max_row:
                                continue
                            v = sheet.cell(rr, col).value
                            if v is None:
                                continue
                            vv = str(v).replace('\n', '<br>').strip()
                            if not vv:
                                continue
                            total += 1
                            # 完全一致または部分一致を証拠として考慮
                            if vv == p or vv == p.replace('<br>', '\n') or vv in p or p in vv:
                                cnt += 1

                        frac = (cnt / total) if total > 0 else 0.0
                        # 保守的ルール: 比較的長く、データ行に頻繁に出現するトークンのみを削除。
                        # これにより、データサンプルにも再出現する可能性のある
                        # '装置名'や'PSコード'のような短いヘッダーラベルの削除を避ける。
                        drop = False
                        try:
                            plen = len(p.strip()) if p else 0
                            # 強い証拠: 非常に頻繁（>=90%）-> 適切な長さのトークンのみ削除
                            # （'装置名'のような短いラベル的トークンの削除を避ける）
                            if frac >= 0.9 and plen >= 4:
                                drop = True
                            # 中程度の証拠: 頻繁（>=60%）でトークンが非常に短くない -> 削除
                            elif frac >= 0.6 and plen >= 8:
                                drop = True
                        except Exception:
                            drop = False

                        if drop:
                            debug_print(f"[DEBUG] ヘッダからデータトークン除外: '{p}' at 列{col} (occurrence_fraction={frac:.2f}, len={plen})")
                            continue
                        filtered_parts.append(p)

                combined = '<br>'.join(filtered_parts) if filtered_parts else ''
                # さらに順序を保持しながら繰り返しサブパーツを削除し、
                # 複数行結合による'A<br>B<br>A<br>B<br>A<br>B'のようなパターンを回避
                try:
                    if combined:
                        subs = [s.strip() for s in combined.split('<br>') if s.strip()]
                        # まず順序を保持しながら連続した重複を削除
                        seen = set()
                        uniq = []
                        for s in subs:
                            if not uniq or uniq[-1] != s:
                                uniq.append(s)
                        # 次に[A,B,A,B,A,B] -> [A,B]のような完全な繰り返しシーケンスを折りたたむ
                        collapsed = self._collapse_repeated_sequence(uniq)
                        combined = '<br>'.join(collapsed)
                except Exception:
                    pass  # 一時ファイルの削除失敗は無視
            except Exception:
                combined = '<br>'.join(dedup_parts) if dedup_parts else ''

            # マスターセルの存在する列（header rows のいずれかで master_col==col）を header_positions に優先して登録
            is_master_col = False
            for r in range(header_row, min(header_row + header_height, end_row + 1)):
                key = f"{r}_{col}"
                if key in merged_info:
                    mi = merged_info[key]
                    if mi.get('master_col') == col and (mi.get('master_row') >= header_row and mi.get('master_row') < header_row + header_height):
                        is_master_col = True
                        break

            # 除外判定（注記っぽい列の排除）は結合後の文字列で行う
            if combined:
                col_ratio = self._column_nonempty_fraction(sheet, start_row, end_row, col)
                keep_despite_low_ratio = False
                try:
                    # 太字や左右罫線、塗りつぶしがあればヘッダーとみなす
                    head_cell = sheet.cell(header_row, col)
                    if head_cell.font and getattr(head_cell.font, 'bold', False):
                        keep_despite_low_ratio = True
                    else:
                        try:
                            if head_cell.border and (getattr(head_cell.border.left, 'style', None) or getattr(head_cell.border.right, 'style', None)):
                                keep_despite_low_ratio = True
                        except Exception as e:
                            pass  # XML解析エラーは無視
                        
                        # 塗りつぶしがある列も保持
                        if not keep_despite_low_ratio:
                            try:
                                if head_cell.fill and head_cell.fill.patternType and head_cell.fill.patternType != 'none':
                                    keep_despite_low_ratio = True
                            except Exception:
                                pass  # エラーは無視

                        if not keep_despite_low_ratio:
                            right_count = 0
                            total_check = 0
                            for rr in range(header_row, end_row + 1):
                                try:
                                    c = sheet.cell(rr, col)
                                    total_check += 1
                                    if c.border and c.border.right and getattr(c.border.right, 'style', None):
                                        right_count += 1
                                except Exception as e:
                                    pass  # XML解析エラーは無視
                            if total_check > 0 and (right_count / total_check) >= 0.5:
                                keep_despite_low_ratio = True
                        
                        # データ行(header_row+1以降)で塗りつぶしや強い罫線を持つセルがあれば保持
                        if not keep_despite_low_ratio:
                            for rr in range(header_row + 1, min(header_row + 5, end_row + 1)):
                                try:
                                    data_cell = sheet.cell(rr, col)
                                    # 塗りつぶしチェック
                                    if data_cell.fill and data_cell.fill.patternType and data_cell.fill.patternType != 'none':
                                        keep_despite_low_ratio = True
                                        break
                                    # 強い左罫線チェック(medium, thick, double)
                                    if data_cell.border and data_cell.border.left:
                                        border_style = getattr(data_cell.border.left, 'style', None)
                                        if border_style in ('medium', 'thick', 'double'):
                                            keep_despite_low_ratio = True
                                            break
                                except Exception as e:
                                    pass  # XML解析エラーは無視
                except Exception:
                    keep_despite_low_ratio = False

                if col_ratio < 0.2 and not keep_despite_low_ratio:
                    # 注記っぽい列としてスキップ
                    debug_print(f"[DEBUG] ヘッダー候補除外(注記っぽい列): '{combined}' at 列{col} (col_nonempty={col_ratio:.2f})")
                    continue

                headers.append(combined)
                header_positions.append(col)
                if is_master_col:
                    debug_print(f"[DEBUG] 結合ヘッダー検出・展開(マスター含む): '{combined}' at 列{col}")
                else:
                    debug_print(f"[DEBUG] ヘッダー検出(結合): '{combined}' at 列{col}")
        
        debug_print(f"[DEBUG] 最終ヘッダー: {headers}")
        debug_print(f"[DEBUG] ヘッダー位置: {header_positions}")

        # フォールバック: 検出されたヘッダーがほとんど空の場合、実際のヘッダー
        # コンテンツが1行下にシフトしている可能性がある（タイトル行がスキップされた場合に一般的）。
        # 1行下にシフトして単純なヘッダーテキストを保守的に再抽出する。
        # これは、ヘッダートークンが実際に次の行に表示される場合に
        # 列を失わないための小さな低リスクのヒューリスティック。
        try:
            nonempty_headers = sum(1 for h in headers if h and str(h).strip())
            total_headers = len(headers) if headers else 0
            if total_headers > 0 and (nonempty_headers / total_headers) < 0.20 and header_row + 1 <= end_row:
                debug_print(f"[DEBUG] ヘッダーがほとんど空のため、header_rowを1行下にシフトして再試行します (from {header_row} -> {header_row+1})")
                shifted_row = header_row + 1
                shifted_headers = []
                shifted_positions = []
                for col in range(start_col, end_col + 1):
                    text_val = ''
                    for r in range(shifted_row, min(shifted_row + header_height, end_row + 1)):
                        key = f"{r}_{col}"
                        if key in merged_info:
                            mi = merged_info[key]
                            master_cell = sheet.cell(mi['master_row'], mi['master_col'])
                            v = (str(master_cell.value).strip() if master_cell.value is not None else '')
                        else:
                            cell = sheet.cell(r, col)
                            v = (str(cell.value).strip() if cell.value is not None else '')
                        if v:
                            text_val = v
                            break
                    shifted_headers.append(text_val)
                    if text_val:
                        shifted_positions.append(col)

                if any(shifted_headers):
                    # シフトされたヘッダーを保守的に採用: シフトされた行で
                    # 空でないヘッダートークンが見つかった場合のみ置換
                    headers = shifted_headers
                    header_positions = shifted_positions
                    header_row = shifted_row
                    debug_print(f"[DEBUG] シフト後ヘッダー採用: headers={headers}, positions={header_positions}, header_row={header_row}")
        except (ValueError, TypeError) as e:
            debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")

        # 空のヘッダー列が混入していると余分な空列が出力されるため除去する
        try:
            filtered = [(h, p) for h, p in zip(headers, header_positions) if h and str(h).strip()]
            if len(filtered) != len(headers):
                if filtered:
                    headers, header_positions = [list(x) for x in zip(*filtered)]
                else:
                    headers, header_positions = [], []
                debug_print(f"[DEBUG] 空ヘッダー列を削除: headers={headers}, positions={header_positions}")
        except (ValueError, TypeError) as e:
            debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")

        # ヘッダー行が実は記述的データ（ファイルパス/XML/長文など）である場合は
        # ヘッダー扱いをやめ、結合セルを考慮した従来のテーブル構築へフォールバックする
        # ただし、検出されたヘッダー領域に結合セル（マスター/スパン情報）が含まれる
        # 場合は多段ヘッダーの可能性が高いため、データ寄り判定でスキップしないようにする。
        data_like_count = 0
        for h in headers:
            if not h:
                continue
            if ('\\' in h or '/' in h or '<' in h or '>' in h or 'xml' in h.lower()) or len(h) > 80:
                data_like_count += 1

        # チェック対象ヘッダーのいずれかに結合セル情報があるかを確認する
        header_height = int(getattr(self, '_detected_header_height', 1) or 1)
        has_merged_header = False
        try:
            for col in header_positions:
                for hr in range(0, header_height):
                    key = f"{header_row + hr}_{col}"
                    mi = merged_info.get(key)
                    if mi:
                        span_r = int(mi.get('span_rows', 1) or 1)
                        span_c = int(mi.get('span_cols', 1) or 1)
                        # スパンが2以上、またはマスターが別のセルなら結合として扱う
                        if span_r > 1 or span_c > 1 or (mi.get('master_row') != (header_row + hr)) or (mi.get('master_col') != col):
                            has_merged_header = True
                            break
                if has_merged_header:
                    break
        except (ValueError, TypeError):
            has_merged_header = False

        # データ寄り判定: 結合ヘッダーが無い場合のみフォールバックを許可する
        if headers and (data_like_count / len(headers)) > 0.4 and not has_merged_header:
            debug_print(f"[DEBUG] ヘッダーっぽい行がデータ寄りと判断({data_like_count}/{len(headers)})、ヘッダー処理をスキップします")
            return self._trim_edge_empty_columns(self._build_table_data_with_merges(sheet, region, merged_info))
        if has_merged_header:
            debug_print(f"[DEBUG] 結合セルを含むヘッダーが検出されたため、データ寄り判定を無視してヘッダー処理を継続します")

        if len(headers) < 2:
            debug_print(f"[DEBUG] ヘッダーが不十分、従来方式を使用")
            return self._trim_edge_empty_columns(self._build_table_data_with_merges(sheet, region, merged_info))

    # --- ヘッダーの連続重複を圧縮する ---
        # ヘッダーの正規化: 複数行ヘッダーの場合、タイトル行を除去する
        # 
        # 判定基準:
        # 1. 同じ文字列の繰り返し (例: "表示<br>表示") → 1つに圧縮
        # 2. 複数列で共通のプレフィックスがある場合 (例: 列4-7が "表示<br>...") → プレフィックスを除去
        # 3. 単一列のみの<br> → セル内改行として保持
        
        normalized_headers = []
        
        # まず、各ヘッダーの<br>を分割して重複を除去
        for h in headers:
            if h and '<br>' in h:
                parts = h.split('<br>')
                # 連続する重複パーツを除去
                unique_parts = []
                for p in parts:
                    p_stripped = p.strip()
                    if p_stripped and (not unique_parts or unique_parts[-1] != p_stripped):
                        unique_parts.append(p_stripped)
                
                if len(unique_parts) == 1:
                    # 重複が除去されて1つになった場合
                    normalized_headers.append(unique_parts[0])
                else:
                    # 複数の異なるパーツがある場合は、とりあえず結合して保持
                    normalized_headers.append('<br>'.join(unique_parts))
            else:
                normalized_headers.append(h)
        
        # 次に、複数列で共通のプレフィックスがあるか検出
        # (連続する3列以上で同じプレフィックスを持つ場合、それはタイトル行とみなす)
        if header_height and header_height > 1:
            # 各ヘッダーの最初のパート(タイトル候補)を取得
            # <br>を含む場合は最初のパート、含まない場合はそのまま使用
            first_parts = []
            last_parts = []
            for nh in normalized_headers:
                if nh and '<br>' in nh:
                    parts = nh.split('<br>')
                    first_parts.append(parts[0].strip())
                    last_parts.append(parts[-1].strip())
                elif nh:
                    first_parts.append(nh.strip())
                    last_parts.append(nh.strip())
                else:
                    first_parts.append(None)
                    last_parts.append(None)
            
            # タイトル行を検出: 単一列のみが異なるパターンを探す
            # 例: 列0が 'OnlineQC<br>名前'、列1-10が '名前' → 列0のみ異常(タイトル行を含む)
            # 例: 列3-6が '表示<br>...' → 正常(意味のあるカテゴリ名)
            prefix_ranges = []  # [(col_idx, prefix)]
            
            # 各列について、他の列と比較して孤立しているか判定
            for i in range(len(last_parts)):
                if first_parts[i] and first_parts[i] != last_parts[i]:
                    # この列は<br>を含む
                    # 同じlast_partを持つ他の列が3列以上あるか確認
                    same_last_count = sum(1 for lp in last_parts if lp == last_parts[i])
                    
                    if same_last_count >= 3:
                        # 同じlast_partを持つ列が3列以上ある
                        # この列だけがfirst_partを持つ場合、孤立している(タイトル行)
                        same_first_in_group = sum(1 for j in range(len(last_parts)) 
                                                  if last_parts[j] == last_parts[i] 
                                                  and first_parts[j] and first_parts[j] != last_parts[j])
                        
                        if same_first_in_group == 1:
                            # この列だけが異なるfirst_partを持つ = タイトル行
                            prefix_ranges.append((i, first_parts[i]))
            
            # タイトル行を除去
            for (col_idx, prefix) in prefix_ranges:
                if normalized_headers[col_idx] and '<br>' in normalized_headers[col_idx]:
                    parts = normalized_headers[col_idx].split('<br>')
                    if len(parts) > 1 and parts[0].strip() == prefix:
                        # プレフィックス(タイトル行)を除去して残りを結合
                        normalized_headers[col_idx] = '<br>'.join(parts[1:])
        
        # 最後に、空になったヘッダーを元に戻す
        for i in range(len(normalized_headers)):
            if not normalized_headers[i] or not normalized_headers[i].strip():
                normalized_headers[i] = headers[i]
        
        # ただし、罫線で明確に区切られている列は圧縮しない（罫線がある = 別列）
        groups = []  # (start_idx, end_idx)のリスト、排他的終端（headersインデックスベース）
        i = 0
        while i < len(normalized_headers):
            j = i + 1
            while j < len(normalized_headers) and normalized_headers[j] == normalized_headers[i]:
                j += 1
            groups.append((i, j))
            i = j

        # 明示的な罫線を尊重するためにグループを展開: グループ内の列がheader_rowで縦罫線で区切られている場合、
        # そのグループを単一列グループに分割して圧縮されないようにする
        final_groups = []
        for (a, b) in groups:
            if b - a <= 1:
                final_groups.append((a, b))
                continue

            # このグループのheader_positionsインデックス
            cols = [header_positions[k] for k in range(a, b) if k < len(header_positions)]
            # シート内の隣接する列境界にheader_rowで右罫線がある場合、その境界を越えて圧縮しない
            split_points = [a]
            for idx in range(len(cols) - 1):
                col_left = cols[idx]
                # ヘッダー行だけでなく、ヘッダー→データ行全体で縦罫線の存在を確認。
                right_count = 0
                total_check = 0
                for rr in range(header_row, end_row + 1):
                    try:
                        cell_l = sheet.cell(rr, col_left)
                        total_check += 1
                        if cell_l.border and cell_l.border.right and cell_l.border.right.style:
                            right_count += 1
                    except Exception:
                        # 無視して続行
                        pass

                has_strong_right = (total_check > 0 and (right_count / total_check) >= 0.5)

                # 追加の厳密チェック: ヘッダー行自体に右罫線が存在する場合、
                # それを決定的な列区切りとして扱い、分割を強制する。これにより、
                # 明示的な縦罫線が列を区切っている場合に同一のヘッダーラベルが圧縮されることを防ぐ。
                try:
                    hdr_cell = sheet.cell(header_row, col_left)
                    if hdr_cell and hdr_cell.border and hdr_cell.border.right and getattr(hdr_cell.border.right, 'style', None):
                        has_strong_right = True
                except Exception as e:
                    pass  # XML解析エラーは無視

                # ヘッダー行間で結合セルのマスターの違いもチェック
                masters_differ = False
                try:
                    header_height = int(getattr(self, '_detected_header_height', 1) or 1)
                    for hr in range(0, header_height):
                        left_key = f"{header_row + hr}_{cols[idx]}"
                        right_key = f"{header_row + hr}_{cols[idx+1]}"
                        left_master = merged_info.get(left_key)
                        right_master = merged_info.get(right_key)
                        if left_master and right_master:
                            lm = (left_master.get('master_row'), left_master.get('master_col'))
                            rm = (right_master.get('master_row'), right_master.get('master_col'))
                            if lm != rm:
                                masters_differ = True
                                break
                        elif left_master and not right_master:
                            masters_differ = True
                            break
                        elif right_master and not left_master:
                            masters_differ = True
                            break
                except Exception:
                    masters_differ = False

                if has_strong_right:
                    # この列と次の列の間で分割を強制
                    split_points.append(a + idx + 1)
                elif masters_differ:
                    # ヘッダー行間でマスターが異なる場合も分割を強制
                    split_points.append(a + idx + 1)

            split_points.append(b)
            # split_pointsから範囲を構築する
            for si in range(len(split_points) - 1):
                final_groups.append((split_points[si], split_points[si+1]))

        # final_groupsの後処理: グループのヘッダーフラグメントがすべて空の場合、そのグループを圧縮しない。
        processed_groups = []
        for (a, b) in final_groups:
            # グループのヘッダー位置全体でサンプルヘッダーフラグメントを構築
            try:
                fragments = []
                for idx in range(a, b):
                    if idx < len(headers):
                        fragments.append(str(headers[idx] or '').strip())
                    else:
                        fragments.append('')
                # すべてのフラグメントが空かどうかを判定（つまり、実質的にヘッダーラベルがない）
                if all((not f) for f in fragments):
                    # すべての基礎となる列を保持するために単一列グループに展開
                    for col_idx in range(a, b):
                        processed_groups.append((col_idx, col_idx + 1))
                else:
                    processed_groups.append((a, b))
            except (ValueError, TypeError):
                processed_groups.append((a, b))

        final_groups = processed_groups

        # 圧縮後のヘッダーと、それぞれのグループで代表となる列位置（左端）を保持
        # 正規化されたヘッダーを使用
        compressed_headers = [normalized_headers[a] for (a, b) in final_groups]
        group_positions = [header_positions[a] for (a, b) in final_groups]
        
        deduplicated_headers = []
        deduplicated_positions = []
        deduplicated_groups = []
        
        i = 0
        while i < len(compressed_headers):
            current_header = compressed_headers[i]
            current_group_start = final_groups[i][0]
            current_group_end = final_groups[i][1]
            
            j = i + 1
            while j < len(compressed_headers) and compressed_headers[j] == current_header:
                current_group_end = final_groups[j][1]
                j += 1
            
            deduplicated_headers.append(current_header)
            deduplicated_positions.append(group_positions[i])
            deduplicated_groups.append((current_group_start, current_group_end))
            
            i = j
        
        compressed_headers = deduplicated_headers
        group_positions = deduplicated_positions
        final_groups = deduplicated_groups
        
        # 実際に使用された列位置を保存（_output_right_side_plain_textで使用）
        self._last_group_positions = group_positions

        debug_print(f"[DEBUG] ヘッダーグループ (元): {groups}")
        debug_print(f"[DEBUG] ヘッダーグループ (最終): {final_groups}")
        debug_print(f"[DEBUG] 圧縮後ヘッダー: {compressed_headers}")
        # 詳細ダンプ: デバッグ用にヘッダー周りの内部状態を出力
        try:
            debug_print(f"[DEBUG-DUMP] headers={headers}")
            debug_print(f"[DEBUG-DUMP] header_positions={header_positions}")
            debug_print(f"[DEBUG-DUMP] group_positions={group_positions}")
            debug_print(f"[DEBUG-DUMP] final_groups={final_groups}")
            debug_print(f"[DEBUG-DUMP] compressed_headers={compressed_headers}")
            debug_print(f"[DEBUG-DUMP] merged_info_keys_sample={list(merged_info.keys())[:20]}")
            # 生のセル値を最初の数行だけダンプしてヘッダー位置との対応を確認
            for rr in range(header_row + 1, min(header_row + 6, end_row + 1)):
                rowvals = []
                for idx, pos in enumerate(header_positions):
                    try:
                        v = sheet.cell(rr, pos).value
                    except (ValueError, TypeError):
                        v = None
                    rowvals.append((pos, v))
                debug_print(f"[DEBUG-DUMP] raw row {rr}: {rowvals}")
        except Exception as _e:
            debug_print(f"[DEBUG-DUMP] failed to dump internal state: {_e}")

        # テーブルデータ構築（圧縮ヘッダーを使用）
        table_data = [compressed_headers]

        # グループごとに実際の列範囲（header_positions 間）を使ってデータ列を扱う
        # これにより、ヘッダーが結合セルで左端に存在し、実際のデータがその右側に複数列に分散しているケースに対応
        group_column_ranges = []  # (col_start, col_end)のリスト、包含的
        for (a, b) in final_groups:
            if a < len(header_positions):
                col_start = header_positions[a]
            else:
                col_start = start_col
            if b < len(header_positions):
                col_end = header_positions[b] - 1
            else:
                col_end = end_col
            # 境界を正規化
            if col_start < start_col:
                col_start = start_col
            if col_end > end_col:
                col_end = end_col
            if col_end < col_start:
                col_end = col_start
            group_column_ranges.append((col_start, col_end))
        debug_print(f"[DEBUG] group_column_ranges={group_column_ranges}")

        # 結合セルを考慮してセル値を取得するヘルパーを構築
        def _get_cell_value(r, c):
            key = f"{r}_{c}"
            if key in merged_info and merged_info[key]['is_merged']:
                mi = merged_info[key]
                mc = sheet.cell(mi['master_row'], mi['master_col'])
                return self._format_cell_content(mc) if mc.value is not None else ''
            cell = sheet.cell(r, c)
            return self._format_cell_content(cell) if cell.value is not None else ''

        # 各グループについて、異なる非空値の数に基づいて列の優先度を計算
        group_column_priority = []
        for (col_start, col_end) in group_column_ranges:
            col_scores = []
            for c in range(col_start, col_end + 1):
                vals = []
                for rr in range(header_row + 1, end_row + 1):
                    v = _get_cell_value(rr, c)
                    if v and v.strip():
                        vals.append(v.strip())
                distinct = len(set(vals))
                nonempty = len(vals)
                # 支配度を計算: 最も一般的な値の頻度
                max_freq = 0
                if vals:
                    from collections import Counter
                    counts = Counter(vals)
                    max_freq = max(counts.values())
                dominance = (max_freq / nonempty) if nonempty > 0 else 0
                # スコアタプル: より多くの異なる値を優先、次に低い支配度（単一トークンによる支配が少ない）、
                # 次により多くの非空セル、最後に左端の列
                col_scores.append((c, distinct, dominance, nonempty))
            # ソート: distinct降順、dominance昇順、nonempty降順、col昇順
            col_scores.sort(key=lambda x: (-x[1], x[2], -x[3], x[0]))
            ordered_cols = [c for (c, _, _, _) in col_scores]
            group_column_priority.append(ordered_cols)
        debug_print(f"[DEBUG] group_column_priority={group_column_priority}")

        # オフライン分析用にコンパクトなグループ/優先度情報をデバッグファイルに書き込む
        sheet_name = getattr(sheet, 'title', None)

        # データ行を構築（ヘッダー行の次から）。各グループ内では行ごとに優先列順で最初の非空セルを参照して値を取得する
        # header_heightを考慮しヘッダー行をスキップ
        data_start_row = header_row + (header_height if header_height else 1)
        for row_num in range(data_start_row, end_row + 1):
            row_data = []
            has_valid_data = False

            for g_idx, cols_priority in enumerate(group_column_priority):
                chosen_content = ''
                chosen_col = None
                if row_num == 28 and g_idx == 1:  # 行28グループ1(初期値)を特別追跡
                    debug_print(f"[DEBUG] 行28グループ1候補: {cols_priority[:5]}")
                for col_candidate in cols_priority:
                    content = _get_cell_value(row_num, col_candidate)
                    if row_num == 28 and g_idx == 1 and col_candidate <= 15:  # 最初の数列のみ
                        debug_print(f"[DEBUG] 行28列{col_candidate}: content='{content}', bool={bool(content and content.strip())}")
                    if content and content.strip():
                        chosen_content = content
                        chosen_col = col_candidate
                        break
                # デバッグ: 選択状況を出力
                header_name = compressed_headers[g_idx] if g_idx < len(compressed_headers) else 'unknown'
                if row_num == 28 or row_num <= header_row + 3:  # 行28を特別に追跡
                    debug_print(f"[DEBUG] 行{row_num}列{chosen_col}({header_name}): -> '{chosen_content}'")

                merged_val = chosen_content.strip() if chosen_content else ''
                row_data.append(merged_val)
                if merged_val:
                    has_valid_data = True

            # すべての行を追加(空行も含める)
            if len(row_data) == len(compressed_headers):
                table_data.append(row_data)

        # _build_table_with_header_rowで既に正規化されたヘッダーを1行目に設定しているため、
        # _output_markdown_tableでは複数行ヘッダーとして扱わないように_detected_header_heightを1に設定
        self._detected_header_height = 1
        
        debug_print(f"[DEBUG] table_data構築完了: {len(table_data)}行")
        if table_data:
            debug_print(f"[DEBUG] table_data[0] (ヘッダー): {table_data[0]}")
            if len(table_data) > 1:
                debug_print(f"[DEBUG] table_data[1] (最初のデータ行): {table_data[1]}")
            if len(table_data) > 2:
                debug_print(f"[DEBUG] table_data[2] (2番目のデータ行): {table_data[2]}")
        
        # 2列最適化は無効化（3列テーブルは3列のまま出力する）
        # ユーザー要望: 「3列のものは3列で表示すべき」
        # 将来の拡張のため関数本体は残しておく
        debug_print(f"[DEBUG] 2列最適化は無効化されています（3列テーブルは3列のまま出力）")
        
        # 先頭/末尾の空列を削除して返す
        # --- ヒューリスティック：任意の列内で結合されている設定行を分割 ---
        # 例: "転送設定初期値(CF-60) IsEnabled 「有効」or 「無効」" を
        #      [親項目, プロパティ, 値] の3列に分割する
        try:
            import re
            if table_data and len(table_data) > 1:
                headers = table_data[0]
                data_rows = table_data[1:]

                cols_to_split = set()

                # 各列について分割が多く発生するか確認する
                col_details = []
                # 2つの正規表現パターンを準備: より厳密なプライマリと寛容なフォールバック
                # プライマリ: より厳密な分割パターンだが、明示的なUnicode範囲チェックを避ける。
                # 空白でなく、明らかなパス/XML文字を含まない中間トークンを受け入れる
                primary_re = re.compile(r'^(.*?)\s+([^\\\/<>:\"\s]{1,60})\s+(.+)$')
                # 寛容: 中間トークンに多くの文字を許可するが、後でパス/XML文字を除外
                # 正規化後に日本語の引用符と全角句読点も受け入れる
                permissive_re = re.compile(r'^(.*?)\s+([^\\\/<>:\\"]{1,60})\s+(.+)$')

                def _normalize_for_split(s: str) -> str:
                    # マッチングを改善するために全角スペース/引用符/括弧を正規化
                    if not s:
                        return ''
                    s = s.replace('\u3000', ' ')
                    s = s.replace('\uFF08', '(').replace('\uFF09', ')')  # 全角括弧
                    s = s.replace('（', '(').replace('）', ')')
                    s = s.replace('「', ' ').replace('」', ' ')
                    s = s.replace('”', '"').replace('“', '"')
                    # 複数のスペースを圧縮
                    import re as _re
                    s = _re.sub(r'\s+', ' ', s).strip()
                    return s

                for col_idx in range(len(headers)):
                    non_empty = 0
                    matches = 0
                    for row in data_rows:
                        if col_idx < len(row):
                            cell = row[col_idx] or ''
                        else:
                            cell = ''
                        if cell and cell.strip():
                            non_empty += 1
                            norm = _normalize_for_split(cell)
                            # 首にまず厳密パターンでマッチを試みる
                            if primary_re.match(norm):
                                matches += 1
                            else:
                                # 次に寛容なパターンを試みるが、パスやXML等のトークンを含む場合は除外して誤検出を抑制
                                m2 = permissive_re.match(norm)
                                if m2:
                                    mid = m2.group(2)
                                    # 明らかなパスのようなまたはXMLのようなトークンを除外
                                    if ('\\' not in mid and '/' not in mid and '<' not in mid and '>' not in mid and ':' not in mid):
                                        matches += 1
                    # 非空行が一定数以上かつマッチ率が高ければ分割候補とする
                    ratio = (matches / non_empty) if non_empty > 0 else 0
                    col_details.append((col_idx, non_empty, matches, ratio))
                    # 分割の誤検出を減らすために閾値を上げる
                    if non_empty >= 2 and ratio >= 0.40:
                        cols_to_split.add(col_idx)

                # デバッグ出力: 列ごとのマッチ状況
                debug_print(f"[DEBUG] 列分割判定: headers={headers}")
                for d in col_details:
                    debug_print(f"[DEBUG] 列{d[0]}: non_empty={d[1]}, matches={d[2]}, ratio={d[3]:.2f}")
                debug_print(f"[DEBUG] 分割候補の列: {sorted(list(cols_to_split))}")

                if cols_to_split:
                    new_headers = []
                    for idx, h in enumerate(headers):
                        if idx in cols_to_split:
                            # 元のヘッダーを保持しつつ Property/Value 列を追加
                            new_headers.extend([h, 'Property', 'Value'])
                        else:
                            new_headers.append(h)

                    new_rows = []
                    for row in data_rows:
                        new_row = []
                        for idx in range(len(headers)):
                            cell = row[idx] if idx < len(row) else ''
                            if idx in cols_to_split:
                                # 正規化された形式に対してマッチするが、可能な限り元のピースを保持
                                norm_cell = _normalize_for_split(cell or '')
                                m = primary_re.match(norm_cell)
                                used_a = used_b = used_c = None
                                if not m:
                                    m = permissive_re.match(norm_cell)
                                    if m:
                                        mid = m.group(2)
                                        if ('\\' in mid or '/' in mid or '<' in mid or '>' in mid or ':' in mid):
                                            m = None

                                if m:
                                    a, b, c = m.groups()
                                    # 元のセルから対応する部分文字列を緩く抽出しようとする
                                    new_row.extend([a.strip(), b.strip(), c.strip()])
                                else:
                                    # マッチしない場合はオリジナルを維持し、Property/Value は空にする
                                    new_row.extend([cell or '', '', ''])
                            else:
                                new_row.append(cell)
                        new_rows.append(new_row)

                    table_data = [new_headers] + new_rows
        except Exception:
            # ここで失敗しても元のtable_dataを返す
            pass

        if len(table_data) > 1:
            headers = table_data[0]
            data_rows = table_data[1:]
            data_start_row = header_row + (header_height if header_height else 1)
            consolidated_data_rows = self._consolidate_merged_rows(data_rows, merged_info, data_start_row, start_col, end_col)
            
            deduplicated_rows = []
            for row in consolidated_data_rows:
                if not deduplicated_rows or row != deduplicated_rows[-1]:
                    deduplicated_rows.append(row)
            
            trimmed_rows = []
            for row in deduplicated_rows:
                is_empty = all(not cell or str(cell).strip() == '' for cell in row)
                if not is_empty:
                    trimmed_rows.append(row)
                else:
                    has_content_after = False
                    current_idx = deduplicated_rows.index(row)
                    for future_row in deduplicated_rows[current_idx + 1:]:
                        if any(cell and str(cell).strip() != '' for cell in future_row):
                            has_content_after = True
                            break
                    if has_content_after:
                        trimmed_rows.append(row)
            
            table_data = [headers] + trimmed_rows
        
        return self._trim_edge_empty_columns(table_data)

    
    def _find_table_title_in_region(self, sheet, region: Tuple[int, int, int, int], 
                                     strict_column_bounds: bool = False) -> Optional[str]:
        """テーブル領域内からタイトルを検出（汎用版: 特定キーワードには依存しない）
        
        Args:
            strict_column_bounds: Trueの場合、列範囲の拡張を制限（離散データ領域検出用）
        """
        start_row, end_row, start_col, end_col = region

        # テーブル領域の前後でタイトルを探す（より広い範囲）
        search_start = max(1, start_row - 10)
        search_end = min(start_row + 5, end_row + 1)

        # 最適なタイトル候補を探す
        title_candidates = []

        # strict_column_boundsがTrueの場合は列範囲を制限
        if strict_column_bounds:
            col_search_start = start_col
            col_search_end = end_col + 1
        else:
            col_search_start = max(1, start_col - 5)
            col_search_end = min(start_col + 15, end_col + 5)

        for row in range(search_start, search_end):
            for col in range(col_search_start, col_search_end):
                cell = sheet.cell(row, col)
                if cell.value:
                    text = str(cell.value).strip()
                    # ファイルパスや長すぎる文字列はタイトル候補から除外
                    if any(x in text for x in ['\\', '/', 'xml']) or len(text) > 120:
                        continue

                    # 太字やMarkdown強調は優先的にタイトル候補とする
                    if cell.font and cell.font.bold:
                        distance = abs(row - start_row)
                        # 太字/高優先度としてマーク
                        row_relation = 0 if row < start_row else (1 if row == start_row else 2)
                        title_candidates.append((text, distance, row, col, 'bold', row_relation))
                        continue
                    if text.startswith('**') and text.endswith('**') and len(text) > 4:
                        clean_text = text.replace('**', '')
                        distance = abs(row - start_row)
                        row_relation = 0 if row < start_row else (1 if row == start_row else 2)
                        title_candidates.append((clean_text, distance, row, col, 'markdown', row_relation))
                        continue

                    # その他は短めのテキストを候補として追加
                    if len(text) <= 80 and len(text.split()) <= 8:
                        # テーブル領域内の行の場合のみ、列チェックを実施
                        if start_row <= row <= end_row:
                            # ヘッダーと間違えやすい列（例: 備考欄のようにほとんど空の列）はタイトル候補から除外
                            col_ratio = self._column_nonempty_fraction(sheet, start_row, end_row, col)
                            if col_ratio < 0.2:
                                debug_print(f"[DEBUG] タイトル候補除外(注記っぽい列): '{text}' at 行{row}列{col} (col_nonempty={col_ratio:.2f})")
                                continue

                        distance = abs(row - start_row)
                        row_relation = 0 if row < start_row else (1 if row == start_row else 2)
                        title_candidates.append((text, distance, row, col, 'general', row_relation))
                        debug_print("[DEBUG] タイトル候補: '{}' at 行{}列{}, 距離{}".format(text, row, col, distance))

        # 最も適切なタイトルを選択
        if title_candidates:
            # 優先順位: (1) テーブル直前(1-2行前)のテキスト, (2) 太字/markdown > general, (3) 表の上方にある候補、(4) 長さ（直前の場合は長い方、それ以外は短い方）、(5) 距離
            def _title_key(x):
                text, distance, row, col, kind, row_relation = x
                is_immediately_before = (start_row - row) in (1, 2) and row < start_row
                immediate_priority = 0 if is_immediately_before else 1
                kind_priority = 0 if kind in ('bold', 'markdown') else 1
                length_priority = -len(text) if is_immediately_before else len(text)
                # row_relation: 0: テーブルの上、1: 同じ行、2: テーブルの下
                return (immediate_priority, kind_priority, row_relation, length_priority, distance)

            best_title = min(title_candidates, key=_title_key)
            # 検出されたタイトル行を記録し、呼び出し元がアンカーとして使用できるようにする
            try:
                self._last_table_title_row = int(best_title[2])
            except (ValueError, TypeError):
                self._last_table_title_row = None
            
            best_row = best_title[2]
            same_row_candidates = [c for c in title_candidates if c[2] == best_row]
            if len(same_row_candidates) > 1:
                same_row_candidates.sort(key=lambda x: x[3])
                combined_title = ' '.join([c[0] for c in same_row_candidates])
                debug_print("[DEBUG] タイトル選択（結合）: '{}' (type={}, row={})".format(combined_title, best_title[4], best_title[2]))
                return combined_title
            
            debug_print("[DEBUG] タイトル選択: '{}' (type={}, row={})".format(best_title[0], best_title[4], best_title[2]))
            return best_title[0]

        # タイトルが見つからない場合は以前のタイトル行をクリア
        self._last_table_title_row = None
        debug_print("[DEBUG] テーブルタイトルが見つかりませんでした")
        return None
    
    def _find_table_header_row(self, sheet, region: Tuple[int, int, int, int]) -> Optional[Tuple[int, int]]:
        """テーブルのヘッダー行を検出（結合セルでのヘッダーも考慮）
        
        Returns:
            Optional[Tuple[int, int]]: (header_row, header_height) または None
        """
        start_row, end_row, start_col, end_col = region
        
        debug_print(f"[DEBUG] ヘッダー検索: 行{start_row}〜{min(start_row + 5, end_row + 1)}")
        debug_print(f"[DEBUG][_find_table_header_row_entry] sheet={sheet.title} region={start_row}-{end_row},{start_col}-{end_col}")
        
        # 結合セル情報を取得して行ごとに評価（結合により上位行が単一ラベルで下位が分割されるケースを区別）
        merged_info = self._get_merged_cell_info(sheet, region)

        candidate_rows = list(range(max(1, start_row - 2), min(start_row + 3, end_row + 1)))
        best_row = None
        best_group_count = -1

        # 単一行および複数行のヘッダー候補を評価（最大高さ3）。
        for row in candidate_rows:
            for height in (1, 2, 3):
                if row + height - 1 > end_row:
                    break

                header_values = []
                # `height`行にわたって列ごとに結合されたヘッダーテキストを構築
                for col in range(start_col, min(start_col + 20, end_col + 1)):
                    parts = []
                    contributors = set()
                    for r2 in range(row, row + height):
                        key = f"{r2}_{col}"
                        if key in merged_info:
                            m = merged_info[key]
                            master_cell = sheet.cell(m['master_row'], m['master_col'])
                            text = (str(master_cell.value).strip() if master_cell.value is not None else '')
                            if text:
                                contributors.add(r2)
                        else:
                            cell = sheet.cell(r2, col)
                            text = (str(cell.value).strip() if cell.value is not None else '')
                            if text:
                                contributors.add(r2)

                        if text and len(text) <= 120:
                            # 行をスタックする際の重複を避ける
                            if not parts or parts[-1] != text:
                                parts.append(text)

                    combined = '<br>'.join(parts) if parts else ''
                    header_values.append(combined)

                # グループをカウント
                group_count = 0
                prev = None
                nonempty = 0
                for v in header_values:
                    if v:
                        nonempty += 1
                    if v and v != prev:
                        group_count += 1
                    prev = v or prev

                # 実際に複数の物理行からヘッダーフラグメントを取得している列の数を判定。
                # 複数行からフラグメントを取得している列が少数（<25%）の場合、
                # '<br>'は真の複数行ヘッダーではなく単一セル内のものである可能性が高い。
                # 複数行候補をスキップ。
                try:
                    multirow_cols = 0
                    total_columns = min(start_col + 20, end_col + 1) - start_col
                    # 利用可能な場合はheader_valuesと一緒にcontributors_per_colを記録
                    # この候補のための軽量なcontributors検出を再構築
                    for col in range(start_col, start_col + total_columns):
                        contribs = set()
                        for r2 in range(row, row + height):
                            key = f"{r2}_{col}"
                            if key in merged_info:
                                m = merged_info[key]
                                master_cell = sheet.cell(m['master_row'], m['master_col'])
                                if master_cell and master_cell.value is not None and str(master_cell.value).strip():
                                    contribs.add(r2)
                            else:
                                cell = sheet.cell(r2, col)
                                if cell and cell.value is not None and str(cell.value).strip():
                                    contribs.add(r2)
                        if len(contribs) > 1:
                            multirow_cols += 1
                    multirow_frac = (multirow_cols / total_columns) if total_columns > 0 else 0
                except (ValueError, TypeError):
                    multirow_frac = 1.0

                # 複数行ヘッダーの判定を改善:
                # multirow_fracが低くても全体として多くの非空セルがあれば有効なヘッダー
                # (行3-4のような「上段と下段で異なる列にテキストがある」構造に対応)
                if height > 1 and multirow_frac < 0.25:
                    # nonemptyが多ければ（全体の50%以上）有効な複数行ヘッダーとして扱う
                    if nonempty >= total_columns * 0.5:
                        debug_print(f"[DEBUG] 複数行ヘッダ候補を維持（非空セルが多い）: row={row}, height={height}, multirow_frac={multirow_frac:.2f}, nonempty={nonempty}/{total_columns}")
                    else:
                        debug_print(f"[DEBUG] 複数行ヘッダ候補をスキップ（実際には単一セル内改行が多い）: row={row}, height={height}, multirow_frac={multirow_frac:.2f}, nonempty={nonempty}/{total_columns}")
                        continue

                # 最上行で複数列結合マスターの一部である列の割合を計算
                top_row = row
                top_merged_count = 0
                total_columns = min(start_col + 20, end_col + 1) - start_col
                for col in range(start_col, start_col + total_columns):
                    key = f"{top_row}_{col}"
                    if key in merged_info and merged_info[key].get('span_cols', 1) > 1:
                        top_merged_count += 1
                top_merged_fraction = (top_merged_count / total_columns) if total_columns > 0 else 0

                # デバッグ
                debug_print(f"[DEBUG] 行{row}..{row+height-1} combined header_values (first16): {header_values[:16]}")
                debug_print(f"[DEBUG] 行{row} height={height} group_count={group_count}, nonempty_cols={nonempty}")

                # より大きなgroup_countを優先。タイブレーカー: 全列カバレッジを優先、次に強い下罫線の整列、
                # 次に最上行の複数列結合が少ないもの（下位レベルの分割を優先）、次により深い最下行、次により大きな高さ
                bottom_row = row + height - 1
                total_columns = min(start_col + 20, end_col + 1) - start_col
                full_coverage = (nonempty == total_columns)

                # ヘッダーの下罫線整列を計算: bottom_rowに下罫線があるか、次の行に上罫線がある列の割合
                border_hits = 0
                border_total = 0
                for c in range(start_col, start_col + total_columns):
                    try:
                        br_cell = sheet.cell(bottom_row, c)
                        border_total += 1
                        if (br_cell.border and br_cell.border.bottom and getattr(br_cell.border.bottom, 'style', None)):
                            border_hits += 1
                        else:
                            # 利用可能な場合は次の行の上罫線をチェック
                            if bottom_row + 1 <= end_row:
                                nx = sheet.cell(bottom_row + 1, c)
                                if nx.border and nx.border.top and getattr(nx.border.top, 'style', None):
                                    border_hits += 1
                    except Exception:
                        continue
                header_border_fraction = (border_hits / border_total) if border_total > 0 else 0.0

                # 結合マスター整列スコアを計算: 結合マスターがヘッダー行内に位置するヘッダー列の割合
                master_aligned = 0
                master_total = 0
                for c in range(start_col, start_col + total_columns):
                    key = f"{row}_{c}"
                    mi = merged_info.get(key)
                    if mi:
                        master_total += 1
                        mr = int(mi.get('master_row', row))
                        # master_rowがヘッダー候補行内にある場合整列
                        if row <= mr <= bottom_row:
                            master_aligned += 1
                masters_alignment_frac = (master_aligned / master_total) if master_total > 0 else 0.0

                # この候補の最下行に対して軽量な「ヘッダーらしさ」スコアを計算
                # 高いスコアはラベルのように見える行を優先（短いトークン、少ない<br>、パスのようでない、太字）
                def _row_header_likeness(rnum):
                    try:
                        total_nonempty = 0
                        short_count = 0
                        nobr_count = 0
                        path_like_count = 0
                        bold_count = 0
                        for c in range(start_col, start_col + total_columns):
                            keyc = f"{rnum}_{c}"
                            txt = ''
                            cell_obj = None
                            if keyc in merged_info:
                                mi = merged_info[keyc]
                                cell_obj = sheet.cell(mi.get('master_row', rnum), mi.get('master_col', c))
                                txt = str(cell_obj.value).strip() if cell_obj and cell_obj.value is not None else ''
                            else:
                                cell_obj = sheet.cell(rnum, c)
                                txt = str(cell_obj.value).strip() if cell_obj and cell_obj.value is not None else ''

                            if not txt:
                                continue
                            total_nonempty += 1
                            if len(txt) <= 40:
                                short_count += 1
                            if '<br>' not in txt:
                                nobr_count += 1
                            low = txt.lower()
                            if ('\\' in txt and ':' in txt) or '/' in txt or low.startswith('http') or 'xml' in low or '<' in txt or '>' in txt:
                                path_like_count += 1
                            try:
                                if cell_obj and cell_obj.font and getattr(cell_obj.font, 'bold', False):
                                    bold_count += 1
                            except Exception as e:
                                pass  # XML解析エラーは無視

                        if total_nonempty == 0:
                            return 0.0
                        short_frac = short_count / total_nonempty
                        nobr_frac = nobr_count / total_nonempty
                        path_frac = path_like_count / total_nonempty
                        bold_frac = bold_count / total_nonempty
                        # 過学習を避けるために保守的に選択された重み
                        score = short_frac + 0.45 * nobr_frac + 0.35 * bold_frac - 0.9 * path_frac
                        return max(0.0, score)
                    except Exception:
                        return 0.0

                likeness_score = _row_header_likeness(bottom_row)
                # デバッグ print for inspection
                debug_print(f"[DEBUG] header-likeness(bottom_row={bottom_row})={likeness_score:.3f}")
                debug_print(f"[DEBUG] header_border_fraction(bottom_row={bottom_row})={header_border_fraction:.3f}")

                # 拡張列範囲でのグループ数を計算（テーブル範囲外のヘッダーも考慮）
                # height>1の場合全行を走査し結合した値でグループをカウント
                extended_group_count = 0
                prev_val = None
                for c in range(max(1, start_col - 5), min(start_col + 30, end_col + 10)):
                    try:
                        # 複数行ヘッダーの場合は全行を結合
                        parts = []
                        for r2 in range(row, row + height):
                            key = f"{r2}_{c}"
                            if key in merged_info:
                                m = merged_info[key]
                                master_cell = sheet.cell(m['master_row'], m['master_col'])
                                text = (str(master_cell.value).strip() if master_cell.value is not None else '')
                            else:
                                cell_val = sheet.cell(r2, c).value
                                text = str(cell_val).strip() if cell_val is not None else ''
                            if text and len(text) <= 120:
                                if not parts or parts[-1] != text:
                                    parts.append(text)
                        val_str = '<br>'.join(parts) if parts else ''
                        
                        if val_str and val_str != prev_val:
                            extended_group_count += 1
                        if val_str:
                            prev_val = val_str
                    except (ValueError, TypeError) as e:
                        debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
                
                debug_print(f"[DEBUG] extended_group_count(row={row})={extended_group_count} (original group_count={group_count})")

                first_row_bonus = 0.5 if row == start_row else 0.0
                adjusted_border_fraction = header_border_fraction + first_row_bonus
                
                if first_row_bonus > 0:
                    debug_print(f"[DEBUG] 最初の行ボーナス適用: row={row}, border_fraction={header_border_fraction:.3f} -> {adjusted_border_fraction:.3f}")

                # この候補のメトリクスタプルを構築
                # 罫線を最優先、次に拡張グループ数を考慮
                # 同等の場合は上部の行を優先（-rowで小さい行番号が大きい値になる）
                # heightは小さい方を優先（保守的なヘッダー検出）
                metrics = (
                    adjusted_border_fraction,  # 1st: 罫線が最も重要な判断基準（start_rowにボーナス）
                    extended_group_count,    # 2nd: 拡張範囲でのグループ数（範囲外ヘッダー対応）
                    -row,                    # 3rd: より上の行を優先（負の値で小さい行番号が大きくなる）
                    group_count,             # 4th: テーブル範囲内のグループ数
                    1 if full_coverage else 0,  # 5th: 全列カバレッジ
                    -height,                 # 6th: より小さいheightを優先（保守的）
                    likeness_score,          # 7th: ヘッダーらしさ
                    -top_merged_fraction,    # 8th: トップ結合セル割合(小さい方が良い)
                    masters_alignment_frac,  # 9th: マスターセルアライメント
                    bottom_row               # 10th: 下部行
                )


                # 最良のメトリクスと比較
                # 拡張範囲で2つ以上のグループがある、またはテーブル範囲内でgroup_count>=2の場合に候補とする
                if extended_group_count >= 2 or group_count >= 2:
                    if best_group_count < 0:
                        best_group_count = group_count
                        best_row = row
                        best_height = height
                        best_metrics = metrics
                        try:
                            self._best_top_merged_fraction = top_merged_fraction
                        except Exception:
                            pass  # エラーは無視
                    else:
                        # 辞書順で比較
                        try:
                            if metrics > best_metrics:
                                best_group_count = group_count
                                best_row = row
                                best_height = height
                                best_metrics = metrics
                                try:
                                    self._best_top_merged_fraction = top_merged_fraction
                                except Exception:
                                    pass  # エラーは無視
                        except Exception:
                            # 以前のタイブレーカーにフォールバック
                            if group_count > best_group_count:
                                best_group_count = group_count
                                best_row = row
                                best_height = height

        if best_row:
            # 下流のヒューリスティック用に検出されたヘッダー開始と高さを保存
            self._detected_header_start = best_row
            self._detected_header_height = best_height
            # ガード: 複数行への昇格がほとんど利益をもたらさない場合は単一行ヘッダーを優先。
            try:
                if best_height and best_height > 1:
                    # 選択された開始行でheight=1のgroup_countを再計算
                    single_vals = []
                    for col in range(start_col, min(start_col + 20, end_col + 1)):
                        parts = []
                        key = f"{best_row}_{col}"
                        if key in merged_info:
                            m = merged_info[key]
                            master_cell = sheet.cell(m['master_row'], m['master_col'])
                            text = (str(master_cell.value).strip() if master_cell.value is not None else '')
                        else:
                            cell = sheet.cell(best_row, col)
                            text = (str(cell.value).strip() if cell.value is not None else '')
                        if text:
                            parts.append(text)
                        combined_one = '<br>'.join(parts) if parts else ''
                        single_vals.append(combined_one)

                    group_count_one = 0
                    prev = None
                    nonempty_one = 0
                    for v in single_vals:
                        if v:
                            nonempty_one += 1
                        if v and v != prev:
                            group_count_one += 1
                        prev = v or prev

                    # 複数行ヘッダーを維持するには意味のある利益が必要。小さな利益は
                    # 下位行がデータのようで吸収すべきでないことを示すことが多い。閾値は1に設定。
                    if (best_group_count - group_count_one) <= 1:
                        debug_print(f"[DEBUG] ヘッダー高さの見直し: 複数行によるグループ増分が小さいため単一行を優先します (row={best_row}, before_height={best_height}, groups_before={best_group_count}, groups_one={group_count_one})")
                        best_height = 1
                        self._detected_header_height = best_height
            except (ValueError, TypeError) as e:
                debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
            # 追加のガード: 選択された複数行ヘッダーの最下行が
            # 単独で同等以上のグループカバレッジを提供する場合、
            # 単一行ヘッダーとして優先（最初のデータ行をヘッダーに引き込むことを避ける）。
            try:
                if best_height and best_height > 1:
                    bottom_row = best_row + best_height - 1
                    bottom_vals = []
                    for col in range(start_col, min(start_col + 20, end_col + 1)):
                        parts = []
                        key = f"{bottom_row}_{col}"
                        if key in merged_info:
                            m = merged_info[key]
                            master_cell = sheet.cell(m['master_row'], m['master_col'])
                            text = (str(master_cell.value).strip() if master_cell.value is not None else '')
                        else:
                            cell = sheet.cell(bottom_row, col)
                            text = (str(cell.value).strip() if cell.value is not None else '')
                        if text:
                            parts.append(text)
                        bottom_vals.append('<br>'.join(parts) if parts else '')

                    group_count_bottom = 0
                    prev = None
                    nonempty_bottom = 0
                    for v in bottom_vals:
                        if v:
                            nonempty_bottom += 1
                        if v and v != prev:
                            group_count_bottom += 1
                        prev = v or prev

                    total_columns = min(start_col + 20, end_col + 1) - start_col
                    # 最下行のグループが複数行グループと等しく、多くの列をカバーする場合は最下行を優先
                    if group_count_bottom >= best_group_count and nonempty_bottom >= max(2, int(total_columns * 0.6)):
                        debug_print(f"[DEBUG] ヘッダー行選択の調整: 下端行が十分に代表的なヘッダーのため下端行を単一行ヘッダーにします (from row={best_row}, height={best_height} -> row={bottom_row}, height=1)")
                        best_row = bottom_row
                        best_height = 1
                        try:
                            self._detected_header_start = best_row
                            self._detected_header_height = best_height
                        except (ValueError, TypeError) as e:
                            debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
            except (ValueError, TypeError) as e:
                debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")

            debug_print(f"[DEBUG] ヘッダー行選択 (罫線優先): 行{best_row} (groups={best_group_count}, height={best_height})")
            return (best_row, best_height)

        # ヘッダー行が見つからなかった場合
        debug_print(f"[DEBUG] ヘッダー行が見つかりませんでした")
        return None
    
    def _get_merged_cell_info(self, sheet, region: Tuple[int, int, int, int]) -> Dict[str, Any]:
        """結合セル情報を取得"""
        start_row, end_row, start_col, end_col = region
        merged_info = {}
        
        try:
            debug_print(f"[DEBUG] 結合セル情報取得開始: region={region}")
            for merged_range in sheet.merged_cells.ranges:
                debug_print(f"[DEBUG] 結合セル範囲発見: 行{merged_range.min_row}〜{merged_range.max_row}, 列{merged_range.min_col}〜{merged_range.max_col}")
                
                # 結合セルがテーブル領域と重なっているかチェック（部分的な重なりも含む）
                if (merged_range.max_row >= start_row and merged_range.min_row <= end_row and
                    merged_range.max_col >= start_col and merged_range.min_col <= end_col):
                    
                    debug_print(f"[DEBUG] 結合セルが領域に重なる: 行{merged_range.min_row}〜{merged_range.max_row}, 列{merged_range.min_col}〜{merged_range.max_col}")
                    
                    # テーブル領域内の範囲のみで結合セル情報を記録
                    actual_start_row = max(merged_range.min_row, start_row)
                    actual_end_row = min(merged_range.max_row, end_row)
                    actual_start_col = max(merged_range.min_col, start_col)
                    actual_end_col = min(merged_range.max_col, end_col)
                    
                    # 結合セルの情報を記録
                    for row in range(actual_start_row, actual_end_row + 1):
                        for col in range(actual_start_col, actual_end_col + 1):
                            key = f"{row}_{col}"
                            merged_info[key] = {
                                'is_merged': True,
                                'master_row': merged_range.min_row,
                                'master_col': merged_range.min_col,
                                'span_rows': merged_range.max_row - merged_range.min_row + 1,
                                'span_cols': merged_range.max_col - merged_range.min_col + 1
                            }
                            debug_print(f"[DEBUG] 結合セル登録: {key} -> master({merged_range.min_row}, {merged_range.min_col})")
                else:
                    debug_print(f"[DEBUG] 結合セルが領域外: 行{merged_range.min_row}〜{merged_range.max_row}, 列{merged_range.min_col}〜{merged_range.max_col}")
        except Exception as e:
            debug_print(f"[DEBUG] 結合セル情報取得エラー: {e}")
        
        return merged_info

    def _column_nonempty_fraction(self, sheet, start_row: int, end_row: int, col: int) -> float:
        """指定列の start_row..end_row における非空セル割合を返す（0.0-1.0）。"""
        total = 0
        nonempty = 0
        for r in range(start_row, end_row + 1):
            total += 1
            cell = sheet.cell(r, col)
            if cell.value is not None and str(cell.value).strip() != "":
                nonempty += 1
        if total == 0:
            return 0.0
        return nonempty / total
    
    def _consolidate_merged_rows(self, table_data: List[List[str]], merged_info: Dict[str, Any],
                                 start_row: int, start_col: int, end_col: int) -> List[List[str]]:
        """マージセルを含む行を統合して重複を削除"""
        if not table_data or len(table_data) <= 1:
            return table_data
        
        debug_print(f"[DEBUG] _consolidate_merged_rows called: table_data rows={len(table_data)}, start_row={start_row}, start_col={start_col}, end_col={end_col}")
        debug_print(f"[DEBUG] merged_info keys sample: {list(merged_info.keys())[:10]}")
        
        rows_to_keep = []
        rows_to_skip = set()
        
        for row_idx in range(len(table_data)):
            if row_idx in rows_to_skip:
                continue
            
            actual_row_num = start_row + row_idx
            current_row = table_data[row_idx]
            debug_print(f"[DEBUG] Processing row_idx={row_idx}, actual_row_num={actual_row_num}, current_row={current_row}")
            
            multi_row_merges = []
            for col_idx in range(len(current_row)):
                actual_col_num = start_col + col_idx
                key = f"{actual_row_num}_{actual_col_num}"
                
                if key in merged_info and merged_info[key]['is_merged']:
                    merge_info = merged_info[key]
                    if (merge_info['master_row'] == actual_row_num and 
                        merge_info['master_col'] == actual_col_num and
                        merge_info['span_rows'] > 1):
                        multi_row_merges.append(merge_info)
            
            if multi_row_merges:
                max_span = max(m['span_rows'] for m in multi_row_merges)
                debug_print(f"[DEBUG] Row {actual_row_num} has {len(multi_row_merges)} multi-row merges, max_span={max_span}")
                
                for next_row_offset in range(1, max_span):
                    next_row_idx = row_idx + next_row_offset
                    if next_row_idx >= len(table_data):
                        break
                    
                    next_row = table_data[next_row_idx]
                    next_actual_row_num = start_row + next_row_idx
                    
                    has_non_merged_data = False
                    for next_col_idx in range(len(next_row)):
                        next_cell = next_row[next_col_idx]
                        if next_cell and str(next_cell).strip():
                            next_actual_col_num = start_col + next_col_idx
                            next_key = f"{next_actual_row_num}_{next_actual_col_num}"
                            
                            if next_key in merged_info and merged_info[next_key]['is_merged']:
                                merge_info = merged_info[next_key]
                                if merge_info['master_row'] < next_actual_row_num:
                                    continue
                            
                            has_non_merged_data = True
                            debug_print(f"[DEBUG] Row {next_actual_row_num} has non-merged data at col {next_col_idx}: {next_cell}")
                            break
                    
                    if has_non_merged_data:
                        debug_print(f"[DEBUG] Row {next_actual_row_num} has non-merged data, not consolidating")
                    else:
                        debug_print(f"[DEBUG] Row {next_actual_row_num} is empty (merged cell only), consolidating")
                        for next_col_idx in range(len(next_row)):
                            if next_col_idx < len(current_row):
                                next_cell = next_row[next_col_idx]
                                if next_cell and str(next_cell).strip():
                                    current_cell = current_row[next_col_idx]
                                    if not (current_cell and str(current_cell).strip()):
                                        current_row[next_col_idx] = next_cell
                        
                        rows_to_skip.add(next_row_idx)
            
            rows_to_keep.append(current_row)
        
        debug_print(f"[DEBUG] _consolidate_merged_rows: {len(table_data)} rows -> {len(rows_to_keep)} rows (skipped {len(rows_to_skip)} rows)")
        return rows_to_keep
    
    def _build_table_data_with_merges(self, sheet, region: Tuple[int, int, int, int], 
                                     merged_info: Dict[str, Any],
                                     strict_column_bounds: bool = False) -> List[List[str]]:
        """結合セルを考慮してテーブルデータを構築（ヘッダー行の検出とテーブル構造改善）
        
        Args:
            strict_column_bounds: Trueの場合、列範囲の拡張を制限（離散データ領域検出用）
        """
        start_row, end_row, start_col, end_col = region
        debug_print(f"[DEBUG] _build_table_data_with_merges実行: region={region}")
        
        # ヘッダー行を検出
        header_info = self._find_table_header_row(sheet, region)
        if header_info:
            header_row, header_height = header_info
            debug_print(f"[DEBUG] ヘッダー行発見: {header_row}, height={header_height}, テーブルをヘッダー行から開始")
            # ヘッダー行が見つかった場合、そこからテーブルを開始
            actual_start_row = header_row
            
            # ヘッダー行の実際の列範囲を確認し、start_col/end_colを拡張
            # (「名前」など、テーブル範囲外のヘッダーを含めるため)
            # strict_column_boundsがTrueの場合は列範囲の拡張を制限
            header_min_col = start_col
            header_max_col = end_col
            if not strict_column_bounds:
                for col_num in range(1, sheet.max_column + 1):
                    cell = sheet.cell(header_row, col_num)
                    if cell.value is not None and str(cell.value).strip():
                        header_min_col = min(header_min_col, col_num)
                        header_max_col = max(header_max_col, col_num)
            
            if header_min_col < start_col or header_max_col > end_col:
                debug_print(f"[DEBUG] ヘッダー行により列範囲を拡張: {start_col}-{end_col} → {header_min_col}-{header_max_col}")
                start_col = header_min_col
                end_col = header_max_col
        else:
            header_row = None
            header_height = 1
            debug_print(f"[DEBUG] ヘッダー行なし、最初の行から開始")
            actual_start_row = start_row
        
        # 最初にすべてのデータを取得
        raw_table_data = []
        for row_num in range(actual_start_row, end_row + 1):
            row_data = []
            for col_num in range(start_col, end_col + 1):
                cell = sheet.cell(row_num, col_num)
                
                # 結合セル情報を使用してセル結合の値を取得・展開
                key = f"{row_num}_{col_num}"
                if key in merged_info and merged_info[key]['is_merged']:
                    merge_info = merged_info[key]
                    # マスターセルから値を取得
                    master_cell = sheet.cell(merge_info['master_row'], merge_info['master_col'])
                    content = self._format_cell_content(master_cell)
                    
                    # 結合セルのデバッグ情報
                    if (row_num == merge_info['master_row'] and 
                        col_num == merge_info['master_col']):
                        if merge_info['span_rows'] > 1 or merge_info['span_cols'] > 1:
                            debug_print(f"[DEBUG] 結合セル検出: '{content}' を範囲 (行:{merge_info['span_rows']}, 列:{merge_info['span_cols']}) に展開")
                else:
                    content = self._format_cell_content(cell)
                
                row_data.append(content)
            raw_table_data.append(row_data)

        # dump raw_table_data sample for diagnostics
        debug_print(f"[DEBUG-DUMP] raw_table_data rows={len(raw_table_data)} sample (first 6):")
        for i, r in enumerate(raw_table_data[:6]):
            debug_print(f"[DEBUG-DUMP] raw row {i+actual_start_row}: cols={len(r)} -> {r}")
        
        # 空行も含めてすべての行を保持(罫線で囲まれた空行もテーブルの一部)
        filtered_table_data = raw_table_data

        # dump filtered_table_data sample for diagnostics
        debug_print(f"[DEBUG-DUMP] filtered_table_data rows={len(filtered_table_data)} sample (first 6):")
        for i, r in enumerate(filtered_table_data[:6]):
            debug_print(f"[DEBUG-DUMP] filtered row {i+actual_start_row}: cols={len(r)} -> {r}")
        
        filtered_table_data = self._consolidate_merged_rows(
            filtered_table_data, merged_info, actual_start_row, start_col, end_col
        )

        # 空列の検出と除去
        useful_columns = self._identify_useful_columns(filtered_table_data)
        # 診断用に初期の有用な列をキャプチャ
        initial_useful_columns = list(useful_columns)

        # 重要: ヘッダー行に値がある列は必ず保持する。
        # これにより、ヘッダーはあるがデータが少ない列が誤って削除される問題を防止する。
        try:
            if filtered_table_data:
                header_row_vals = filtered_table_data[0]
                for idx, v in enumerate(header_row_vals):
                    if v and str(v).strip() and idx not in useful_columns:
                        useful_columns.append(idx)
                useful_columns = sorted(set(useful_columns))
        except (ValueError, TypeError):
            # ロギングのみ行い、処理を継続
            debug_print('[TRACE-USE-HEADER]', str(region), f'useful_columns_before={useful_columns}')

        # 追加ガード: useful_columns 選定後に、各列のデータ非空割合を計算して
        # 一定のしきい値を満たす列は保持する（誤って削除されるのを防ぐ）
        try:
            total_rows_for_data = max(1, max(0, len(filtered_table_data) - 1))
            col_counts = []
            num_cols_all = max(len(r) for r in filtered_table_data) if filtered_table_data else 0
            for ci in range(num_cols_all):
                cnt = 0
                # データを持つセルをカウントする際にヘッダー行（インデックス0）をスキップ
                for r in filtered_table_data[1:]:
                    if ci < len(r) and r[ci] and str(r[ci]).strip():
                        cnt += 1
                col_counts.append(cnt)

            # 判定: cnt >= 1 または fraction >= 0.05 (5%) なら保持
            kept_by_guard = []
            for ci, cnt in enumerate(col_counts):
                frac = cnt / total_rows_for_data if total_rows_for_data > 0 else 0
                if ci not in useful_columns and (cnt >= 1 or frac >= 0.05):
                    useful_columns.append(ci)
                    kept_by_guard.append((ci, cnt, frac))

            useful_columns = sorted(set(useful_columns))
            debug_print(f"[TRACE-USEFUL-DECISION] region={region} initial_counts={col_counts} kept_by_guard={kept_by_guard} final_useful={useful_columns}")
        except (ValueError, TypeError) as e:
            debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")

        # 診断: 領域内のすべての元のシート列について、なぜそれが
        # 保持または削除されたかを記録。これはどのブランチが列を圧縮したかを追跡するのに役立つ。
        try:
            per_column_diag = []
            num_all_cols = end_col - start_col + 1
            # 元の相対インデックス -> initial_useful、header_keep、guard_keepに含まれるかどうかをマップ
            for rel in range(0, num_all_cols):
                abs_col = start_col + rel
                in_initial = rel in initial_useful_columns
                in_final = rel in useful_columns
                header_present = False
                header_texts = []
                # 利用可能な場合は検出されたヘッダー行からヘッダーフラグメントを取得しようとする
                try:
                    hdr_row = header_row or actual_start_row
                    detected_h = int(getattr(self, '_detected_header_height') or 1)
                    for hr in range(hdr_row, min(hdr_row + detected_h, end_row + 1)):
                        val = sheet.cell(hr, abs_col).value
                        if val is not None and str(val).strip():
                            header_present = True
                            header_texts.append(str(val).strip())
                except (ValueError, TypeError):
                    # フォールバック: filtered_table_dataヘッダーが存在する場合は使用
                    try:
                        if filtered_table_data and rel < len(filtered_table_data[0]):
                            hv = filtered_table_data[0][rel]
                            if hv and str(hv).strip():
                                header_present = True
                                header_texts.append(str(hv).strip())
                    except (ValueError, TypeError) as e:
                        debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")

                # この元のrel列のデータ非空カウント
                data_count = 0
                for r in filtered_table_data[1:]:
                    if rel < len(r) and r[rel] and str(r[rel]).strip():
                        data_count += 1

                reason = 'kept' if in_final else 'dropped'
                per_column_diag.append((abs_col, rel, in_initial, header_present, header_texts, data_count, reason))

            # 診断を出力
            debug_print('[COLUMN-MAP] region_abs_cols={} ->'.format((start_col, end_col)))
            for t in per_column_diag:
                abs_col, rel, in_initial, header_present, header_texts, data_count, reason = t
                debug_print(f"[COLUMN-MAP] col={abs_col} rel={rel} initial={in_initial} header_present={header_present} header_texts={header_texts} data_count={data_count} -> {reason}")
        except (ValueError, TypeError) as e:
            debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
        
        # 有用な列のみでテーブルを再構築
        table_data = []
        for row_data in filtered_table_data:
            compressed_row = [row_data[i] if i < len(row_data) else "" for i in useful_columns]
            table_data.append(compressed_row)

        # 列圧縮ステップ後にダンプ
        debug_print(f"[DEBUG-DUMP] after useful_columns compression: useful_columns={useful_columns}, table_rows={len(table_data)} sample (first 6):")
        for i, r in enumerate(table_data[:6]):
            debug_print(f"[DEBUG-DUMP] compressed row {i}: cols={len(r)} -> {r}")
        
        # --- 追加: ヘッダーに同一テキストが連続している場合、それらの列をまとめる ---
        if table_data:
            header = table_data[0]
            # 列を保持するかどうかを決定するために列ごとの非空カウントを計算
            col_nonempty_counts = []
            for ci in range(len(header)):
                cnt = 0
                for r in table_data[1:]:
                    if ci < len(r) and r[ci] and str(r[ci]).strip():
                        cnt += 1
                col_nonempty_counts.append(cnt)
            debug_print(f"[DEBUG-DUMP] per-column nonempty counts (after useful_columns): {col_nonempty_counts}")
            # groups: 連続する同一ヘッダーの(start_idx,end_idx)リスト
            groups = []
            i = 0
            while i < len(header):
                j = i + 1
                # ヘッダー値が空でない場合のみ連続する同一ヘッダーをマージ。
                # ヘッダーが空の場合、ヘッダー行にラベルがないときに複数のデータ列が
                # 1つに誤って圧縮されるのを避けるため、各列を個別に扱う。
                while j < len(header) and header[j] == header[i] and (header[i] and str(header[i]).strip()):
                    j += 1
                groups.append((i, j))
                i = j

            # グループの後処理: グループ（ヘッダーが空かどうかに関わらず）に
            # 非空データを持つ複数の列が含まれる場合、そのグループの圧縮を避ける。
            final_groups = []
            for (a, b) in groups:
                # このグループ内で非空データを持つ列の数をカウント
                nonempty_cols_in_group = sum(1 for k in range(a, b) if k < len(col_nonempty_counts) and col_nonempty_counts[k] > 0)
                # データを持つ列が複数ある場合、列を保持するために単一列に分割
                if nonempty_cols_in_group > 1:
                    for k in range(a, b):
                        final_groups.append((k, k+1))
                    debug_print(f"[HEADER-GROUP-SKIP-MULTIDATA] expanded group {(a,b)} into singletons because nonempty_cols={nonempty_cols_in_group}")
                else:
                    final_groups.append((a, b))

            # 圧縮が必要なとき（グループ数が列数と異なる）
            if len(final_groups) != len(header):
                new_table = []
                for ri, row in enumerate(table_data):
                    new_row = []
                    for (a, b) in final_groups:
                        # グループ内のセルを結合（空セルは無視）
                        merged_cells = [c for c in row[a:b] if c and str(c).strip()]
                        if ri == 0:
                            # ヘッダー行は重複を繰り返さないよう先頭の値を採用
                            merged_val = merged_cells[0] if merged_cells else ""
                        else:
                            # データ行は空でないセルをスペースで連結
                            merged_val = " ".join(merged_cells).strip()
                        new_row.append(merged_val)
                    new_table.append(new_row)
                table_data = new_table

        # 2列最適化チェック（_build_table_data_with_mergesでも実行）
        if len(table_data) > 1 and len(table_data[0]) == 3:
            # 簡易ヘッダー検出（最初の行をヘッダーとみなす）
            headers = table_data[0]
            debug_print(f"[DEBUG] _build_table_data_with_merges内で2列最適化チェック: headers={headers}")
            
            # 2列最適化を試行（簡易版）
            # このパスでは上位でtable_dataのみが利用可能。テーブルデータベースのチェックを維持
            if self._is_setting_item_pattern_tabledata(table_data, 1, 2):
                # 第1列と第3列を保持して2列テーブルを作る
                debug_print(f"[DEBUG] 2列最適化実行: {headers[0]} | {headers[2]}")
                optimized_table = [[headers[0], headers[2]]]  # ヘッダー行

                # データ行を処理: 第1列と第3列を採用
                matched = 0
                for i in range(1, len(table_data)):
                    row = table_data[i]
                    if len(row) >= 3 and row[0].strip() and row[2].strip():
                        optimized_table.append([row[0], row[2]])
                        matched += 1

                # 圧縮前に合理的な割合の行がマッチすることを要求
                total_data_rows = max(1, len(table_data) - 1)
                required = max(1, int(total_data_rows * 0.5))  # 少なくとも行の50%
                # 2列最適化決定のための追加診断ダンプ
                debug_print(f"[DEBUG-DUMP] 2col optimization: total_data_rows={total_data_rows}, matched={matched}, required={required}")
                # 決定に使用された行のサンプルを表示
                for j, r in enumerate(table_data[1: min(len(table_data), 1+10) ]):
                    debug_print(f"[DEBUG-DUMP] data row sample {j+1}: {r}")

                if matched >= required:
                    debug_print(f"[DEBUG] 2列最適化成功、{len(optimized_table)}行のテーブルを返す (matched={matched}/{total_data_rows})")
                    return optimized_table
                else:
                    debug_print(f"[DEBUG] 2列最適化スキップ（マッチ行不足: {matched}/{total_data_rows}、必要={required}）")
            else:
                debug_print(f"[DEBUG] パターンマッチせず（_build_table_data_with_merges内）")
        
        return table_data
    
    def _identify_useful_columns(self, table_data: List[List[str]]) -> List[int]:
        """テーブルから有用な列を特定"""
        if not table_data:
            return []
        
        num_cols = len(table_data[0]) if table_data else 0
        useful_columns = []
        
        for col_idx in range(num_cols):
            # 列に有意義な内容があるかチェック
            has_content = False
            for row_data in table_data:
                if col_idx < len(row_data) and row_data[col_idx].strip():
                    has_content = True
                    break
            
            if has_content:
                useful_columns.append(col_idx)
        
        # 少なくとも2列は保持する（最低限のテーブル構造）
        if len(useful_columns) < 2 and num_cols >= 2:
            useful_columns = [0, min(1, num_cols - 1)]
        
        return useful_columns

    def _trim_edge_empty_columns(self, table_data: List[List[str]]) -> List[List[str]]:
        """先頭および末尾の完全に空の列を削除して列ずれを防止する"""
        if not table_data:
            return table_data

        # 正規化: 各行を同じ列数に揃える
        num_cols = max(len(row) for row in table_data)
        for row in table_data:
            while len(row) < num_cols:
                row.append("")

        left = 0
        right = num_cols - 1

        # 左端から最初の非空列を見つける
        while left <= right:
            if any(r[left].strip() for r in table_data):
                break
            left += 1

        # 右端から最後の非空列を見つける
        while right >= left:
            if any(r[right].strip() for r in table_data):
                break
            right -= 1

        # 少なくとも2列は保持する（既存の方針に合わせる）
        if right - left + 1 < 2 and num_cols >= 2:
            left = 0
            right = min(1, num_cols - 1)

        # スライスして新しいテーブルを返す
        new_table = []
        for r in table_data:
            new_table.append(r[left:right+1])

        return new_table
    
    def _format_cell_content(self, cell) -> str:
        """セルの内容をフォーマット"""
        if cell.value is None:
            return ""

        # 値を文字列に変換
        cell_text = str(cell.value).strip()

        # 改行を統一して<br>に変換
        cell_text = cell_text.replace('\r\n', '\n')
        cell_text = cell_text.replace('\r', '\n')
        cell_text = cell_text.replace('\n', '<br>')

        # 複数の連続する<br>を整理
        import re
        cell_text = re.sub(r'(<br>\s*){2,}', '<br><br>', cell_text)

        # Markdownテーブル内の問題文字をエスケープ
        # '|' はテーブル区切りになるためエスケープ
        cell_text = cell_text.replace('|', '\\|')

        # '&' は既存のエンティティを壊さないように一律に変換
        cell_text = cell_text.replace('&', '&amp;')
        # 角括弧はユーザ要件によりそのまま保持する (< and > are preserved)

        # 書式設定を適用
        cell_text = self._apply_cell_formatting(cell, cell_text)

        return cell_text

    def _collapse_repeated_sequence(self, parts: List[str]) -> List[str]:
        """partsが繰り返しシーケンス（[A,B,A,B,A,B]のような）かどうかを検出し、最小の繰り返しパターンを1回返す。

        完全な繰り返しが見つからない場合は、partsをそのまま返す。
        """
        try:
            if not parts:
                return parts
            n = len(parts)
            # n//2までのすべての可能なパターン長を試す
            for plen in range(1, n // 2 + 1):
                if n % plen != 0:
                    continue
                pattern = parts[0:plen]
                if pattern * (n // plen) == parts:
                    return pattern
            return parts
        except Exception:
            return parts
    
    def _apply_cell_formatting(self, cell, text: str) -> str:
        """セルの書式設定をMarkdownに適用"""
        try:
            if not text:
                return text
            
            # フォントスタイル
            if cell.font:
                if cell.font.bold:
                    text = f"**{text}**"
                if cell.font.italic:
                    text = f"*{text}*"
            
            return text
            
        except Exception as e:
            print(f"[WARNING] セル書式適用エラー: {e}")
            return text

    def _escape_cell_for_table(self, text: str) -> str:
        """テーブル出力用にセル内の特殊文字を安全にエスケープする。

        既にエスケープされたHTMLエンティティ（例: &lt;）を二重にエスケープしないように、
        '&' は既存のエンティティでない場合のみ '&amp;' に変換する。
        次に '<' と '>' を '&lt;' '&gt;' に変換し、Markdownテーブルの区切り文字 '|' をエスケープする。
        """
        try:
            import re
            if text is None:
                return ''
            t = str(text)

            # プログラムで挿入された<br>（および一般的なバリアント）を保持し、
            # エスケープされないようにする。Excel由来の'<' '>'は引き続きエスケープする。
            # 許可されたタグをプレースホルダーに置き換え、汎用エスケープを実行し、
            # プレースホルダーをリテラルタグに戻す。
            allowed_tags = []
            # 保持したいタグのバリアントを正規化（小文字）
            for m in re.finditer(r'(?i)<br\s*/?>', t):
                allowed_tags.append(m.group(0))

            placeholders = {}
            for i, tag in enumerate(allowed_tags):
                ph = f'___BR_TAG_PLACEHOLDER_{i}___'
                # マッピングを維持するため毎回最初の出現のみを置換
                t = t.replace(tag, ph, 1)
                placeholders[ph] = tag

            # 既存のHTMLエンティティを保護: エンティティの一部でない'&'を変換
            t = re.sub(r'&(?![A-Za-z]+;|#\d+;)', '&amp;', t)

            # 残りの角括弧をエスケープ（これらはExcelセルコンテンツから来る）
            t = t.replace('<', '&lt;').replace('>', '&gt;')

            # Markdownテーブルのパイプ文字をエスケープ
            t = t.replace('|', '\\|')

            # 許可されたタグ（プレースホルダー）をリテラル形式に戻す
            for ph, tag in placeholders.items():
                # 正規化された'<br>'形式を使用
                t = t.replace(ph, '<br>')

            return t
        except Exception:
            return str(text)
    
    def _output_markdown_table(self, table_data: List[List[str]], source_rows: Optional[List[int]] = None, sheet_title: Optional[str] = None):
        """Markdownテーブルとして出力"""
        if not table_data:
            return
        
        if source_rows and len(source_rows) >= 2 and source_rows[0] <= 4:
            import traceback
            stack = traceback.extract_stack()
            caller_info = []
            for frame in stack[-6:-1]:
                if 'x2md.py' in frame.filename:
                    caller_info.append(f"{frame.name}:{frame.lineno}")
            debug_print(f"[DEBUG][{sheet_title}] テーブル生成: rows={source_rows[:5]}, cols={len(table_data[0]) if table_data else 0}, caller={' <- '.join(caller_info)}")

        # 行の長さを正規化
        max_cols = max(len(row) for row in table_data)
        for row in table_data:
            while len(row) < max_cols:
                row.append("")

        num_cols = max_cols
        max_header_rows = min(3, len(table_data))

        def _combined_header_nonempty_count(nrows: int):
            nonempty_total = 0
            length_acc = 0
            path_like_total = 0
            for col in range(num_cols):
                parts = []
                for ri in range(0, nrows):
                    v = table_data[ri][col]
                    if v is not None and str(v).strip():
                        parts.append(str(v).strip())
                joined = '<br>'.join(parts) if parts else ''
                if joined:
                    nonempty_total += 1
                    length_acc += len(joined)
                    if ('\\' in joined and ':' in joined) or '/' in joined or '<' in joined or '>' in joined or 'xml' in joined.lower():
                        path_like_total += 1
            avg_len = (length_acc / nonempty_total) if nonempty_total else 0
            path_like_frac = (path_like_total / nonempty_total) if nonempty_total else 0
            return nonempty_total, avg_len, path_like_frac

        # ヒューリスティックにヘッダー行数を選択
        best_candidate = 1
        best_metrics = None
        for candidate in range(1, max_header_rows + 1):
            nonempty_cnt, avg_len, path_like_frac = _combined_header_nonempty_count(candidate)
            metrics = (nonempty_cnt, -avg_len, -path_like_frac)
            if best_metrics is None or metrics > best_metrics:
                best_metrics = metrics
                best_candidate = candidate

        chosen_nonempty, _, _ = _combined_header_nonempty_count(best_candidate)
        chosen_coverage = chosen_nonempty / max(1, num_cols)
        header_rows_count = 1 if chosen_coverage < 0.10 else best_candidate

        # 利用可能な場合は以前に検出されたヘッダー高さを尊重
        if hasattr(self, '_detected_header_height'):
            detected = int(getattr(self, '_detected_header_height') or 0)
            if 1 <= detected <= max_header_rows:
                # 最初のヘッダー行が既に結合されたピース（多くの'<br>'）を含む場合、
                # フラグメントを重複させる可能性のある追加行の再結合を避ける。
                try:
                    first_row = table_data[0]
                    # 非空のヘッダーセルと既に'<br>'を含むものをカウント
                    nonempty_hdr = sum(1 for h in first_row if h and str(h).strip())
                    combined_hdr = sum(1 for h in first_row if h and '<br>' in str(h))
                    # ヘッダーセルに既に'<br>'が含まれている場合、単一の結合ヘッダーとして扱う
                    # （_build_table_with_header_rowが既に複数行ヘッダーを'<br>'でマージしているため）
                    if combined_hdr > 0:
                        header_rows_count = 1
                    else:
                        header_rows_count = detected
                except (ValueError, TypeError):
                    header_rows_count = detected

        # 以前はフィルダウンが最後に見た値を後のデータ行にコピーしていたが、
        # これにより真に空のデータセルが上からの無関係な値で埋められる可能性があった。
        # フィルダウンをヘッダー領域の補完のみに制限: 複数のヘッダー行があり、
        # 一部のヘッダー行セルが空の場合、データ領域に伝播するのではなく、
        # 近くのヘッダー行から埋める。
        try:
            if header_rows_count > 1:
                for col_idx in range(num_cols):
                    # この列のヘッダー行の値を収集
                    hdr_vals = [table_data[r][col_idx] if col_idx < len(table_data[r]) else '' for r in range(0, header_rows_count)]
                    # ヘッダー行内でのみ前方フィル
                    last = None
                    for ri in range(0, header_rows_count):
                        v = hdr_vals[ri]
                        if v and str(v).strip():
                            last = v
                        else:
                            if last is not None:
                                # table_dataのヘッダーセルに書き戻す
                                if col_idx < len(table_data[ri]):
                                    table_data[ri][col_idx] = last
        except (ValueError, TypeError) as e:
            debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")

        # ヘッダー行を結合してヘッダーセルを構築
        # 注: 各ヘッダー行セルは既に'<br>'シーケンスを含んでいる可能性がある（マージ/ヘッダー組み立てから）。
        # 複数のヘッダー行を結合する際にサブパーツの重複を避けるため、'<br>'で分割し、
        # 行間で連続するサブパーツを重複排除する。
        header_cells = []
        for col in range(num_cols):
            subparts = []
            for ri in range(0, header_rows_count):
                v = table_data[ri][col]
                if v is not None and str(v).strip():
                    # ほぼ重複するパーツを避けるためにヘッダービルダーと同様に正規化
                    try:
                        import re as _re
                        vv = str(v).replace('\r\n', '\n').replace('\r', '\n').replace('\n', '<br>')
                        vv = _re.sub(r'(<br>\s*){2,}', '<br>', vv)
                        vv = _re.sub(r'^(?:<br>\s*)+', '', vv)
                        vv = _re.sub(r'(?:\s*<br>)+$', '', vv)
                        vv = vv.strip()
                    except (ValueError, TypeError):
                        vv = str(v).replace('\n', '<br>').strip()
                    # 既に正規化されたvvをアトミックパーツに分割して拡張
                    for part in [p.strip() for p in vv.split('<br>') if p.strip()]:
                        subparts.append(part)

            # 行ごとの結合セルによって導入された繰り返しフラグメントを避けるために連続するサブパーツを重複排除
            dedup = []
            for p in subparts:
                if not dedup or dedup[-1] != p:
                    dedup.append(p)
            # [A,B,A,B,A,B] -> [A,B]のような完全な繰り返しシーケンスを圧縮
            try:
                collapsed = self._collapse_repeated_sequence(dedup)
            except Exception:
                collapsed = dedup
            header_cells.append('<br>'.join(collapsed) if collapsed else '')

        # 連続する同一ヘッダーを保守的に圧縮
        try:
            groups = []
            i = 0
            while i < len(header_cells):
                j = i + 1
                while j < len(header_cells) and header_cells[j] == header_cells[i]:
                    j += 1
                groups.append((i, j))
                i = j

            # ヘッダー値が空でない場合のみ連続する同一ヘッダーの保守的な圧縮を実行。
            # header_cellsが空文字列の場合（ヘッダー行が存在しないか位置ずれの場合）
            # 複数のデータ列が単一の列にマージされるのを避けるために圧縮をスキップ。
            collapse_needed = any((b - a > 1 and header_cells[a] and str(header_cells[a]).strip()) for (a, b) in groups)
            if collapse_needed:
                new_header = []
                new_table = []
                for (a, b) in groups:
                    new_header.append(header_cells[a])
                for row in table_data:
                    new_row = []
                    for (a, b) in groups:
                        vals = [row[k] for k in range(a, b) if k < len(row) and row[k] and str(row[k]).strip()]
                        new_row.append(' '.join(vals).strip())
                    new_table.append(new_row)
                table_data = new_table
                header_cells = new_header
                num_cols = len(header_cells)
        except (ValueError, TypeError) as e:
            debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")

    # ヘッダーを出力
        safe_header = [self._escape_cell_for_table(h) for h in header_cells]
        self.markdown_lines.append("| " + " | ".join(safe_header) + " |")
        self.markdown_lines.append("| " + " | ".join(["---"] * len(header_cells)) + " |")

        # 実際にはヘッダーのフラグメントである直後の行をフィルタリング
        # （一部のテーブルビルダーは生のヘッダー行の後に結合されたヘッダーを返す。
        # データ内の重複したヘッダーフラグメントを避けるためにこれらを検出してスキップ）
        try:
            skip_count = 0
            max_check = min(max_header_rows, len(table_data) - header_rows_count)
            # 列ごとのヘッダーパーツを準備
            header_parts = [ [p.strip() for p in (hc or '').split('<br>') if p.strip()] for hc in header_cells ]

            def _is_row_header_like(row) -> bool:
                if not row:
                    return False
                matched = 0
                nonempty = 0
                for ci in range(len(header_parts)):
                    if ci >= len(row):
                        continue
                    cell = (row[ci] or '').strip()
                    if not cell:
                        continue
                    nonempty += 1
                    parts = header_parts[ci]
                    # ヘッダーフラグメントである可能性が低い過度に長いセルをスキップ
                    if len(cell) > 200:
                        # データのように扱う
                        continue

                    # セルがその列のヘッダーパーツのいずれかと等しい場合、マッチとしてカウント
                    # ただし保守的に: 正確な短いマッチか、ヘッダーパーツとセルの両方が
                    # 適度に短く、'<br>'の不一致を含まないことを要求。
                    part_match = False
                    for hp in parts:
                        if not hp:
                            continue
                        # 一方に<br>が含まれ、もう一方に含まれない場合、マッチを避ける
                        if ('<br>' in hp) != ('<br>' in cell):
                            continue
                        # 完全一致を許可
                        if cell == hp:
                            part_match = True
                            break
                        # 短いファジーマッチを許可: 両方とも短く、長さが類似
                        if len(cell) <= max(60, int(len(hp) * 1.2)) and len(hp) <= 120 and abs(len(cell) - len(hp)) <= max(10, int(len(hp) * 0.2)):
                            if cell in hp or hp in cell:
                                part_match = True
                                break

                    if part_match:
                        matched += 1
                if nonempty == 0:
                    return False
                return (matched / nonempty) >= 0.6

            for i in range(max_check):
                candidate_row = table_data[header_rows_count + i]
                if _is_row_header_like(candidate_row):
                    skip_count += 1
                else:
                    break
        except Exception:
            skip_count = 0

        # データ行を出力（検出されたヘッダーのような行をスキップ）
        start_idx = header_rows_count + skip_count
        # source_rows提供時はマッピングを準備
        sheet_map = None
        # 既存の権威的マッピングのみを取得。ここでは作成しない。
        if source_rows and sheet_title:
            try:
                sheet_map = self._cell_to_md_index.get(sheet_title, {})
            except Exception:
                sheet_map = None

        for idx, row in enumerate(table_data[start_idx:], start=start_idx):
            while len(row) < len(header_cells):
                row.append("")
            row = row[:len(header_cells)]
            safe_row = [self._escape_cell_for_table(c) for c in row]
            # 利用可能な場合はソース行からmarkdownインデックスへのマッピングを記録
            if source_rows and idx < len(source_rows) and sheet_map is not None:
                src = source_rows[idx]
                # 行を追加し、src -> indexをマップ（ガード付き）
                self.markdown_lines.append("| " + " | ".join(safe_row) + " |")
                try:
                    self._mark_sheet_map(sheet_title, src, len(self.markdown_lines) - 1)
                except Exception:
                    pass  # データ構造操作失敗は無視
            else:
                self.markdown_lines.append("| " + " | ".join(safe_row) + " |")

        self.markdown_lines.append("")

    def _is_fully_bordered_row(self, sheet, row_num: int, min_col: int, max_col: int) -> bool:
        """行全体のすべてのセルが上下左右罫線で囲まれている場合のみTrue"""
        for col_num in range(min_col, max_col + 1):
            cell = sheet.cell(row_num, column=col_num)
            if not (cell.border and cell.border.left and cell.border.left.style and
                    cell.border.right and cell.border.right.style and
                    cell.border.top and cell.border.top.style and
                    cell.border.bottom and cell.border.bottom.style):
                return False
        return True
    
    def _is_table_row(self, sheet, row_num: int, min_col: int, max_col: int) -> bool:
        """行内のいずれかのセルが上下左右罫線で囲まれていればテーブルとみなす"""
        for col_num in range(min_col, max_col + 1):
            cell = sheet.cell(row=row_num, column=col_num)
            if (cell.border and cell.border.left and cell.border.left.style and
                cell.border.right and cell.border.right.style and
                cell.border.top and cell.border.top.style and
                cell.border.bottom and cell.border.bottom.style):
                return True
        return False
