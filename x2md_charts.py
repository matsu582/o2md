#!/usr/bin/env python3
"""
Excelチャートデータ抽出モジュール

openpyxlを使用してExcelファイルからチャートデータを抽出し、
ChartDataオブジェクトに変換する。

対応チャートタイプ:
- 棒グラフ (BarChart)
- 折れ線グラフ (LineChart)
- 円グラフ (PieChart)
- 散布図 (ScatterChart)
"""

import re
from typing import List, Optional, Any

from chart_utils import ChartData, SeriesData


CHART_TYPE_MAP = {
    'BarChart': 'bar',
    'LineChart': 'line',
    'PieChart': 'pie',
    'ScatterChart': 'scatter',
    'AreaChart': 'area',
    'RadarChart': 'radar',
    'DoughnutChart': 'doughnut',
}


def extract_charts_from_worksheet(worksheet, workbook) -> List[ChartData]:
    """
    ワークシートからチャートデータを抽出する
    
    Args:
        worksheet: openpyxlのワークシートオブジェクト
        workbook: openpyxlのワークブックオブジェクト
        
    Returns:
        ChartDataオブジェクトのリスト
    """
    charts = []
    
    ws_charts = getattr(worksheet, "_charts", [])
    
    for chart in ws_charts:
        chart_data = _extract_chart_data(chart, workbook)
        if chart_data:
            charts.append(chart_data)
    
    return charts


def _extract_chart_data(chart, workbook) -> Optional[ChartData]:
    """
    単一のチャートからデータを抽出する
    
    Args:
        chart: openpyxlのチャートオブジェクト
        workbook: openpyxlのワークブックオブジェクト
        
    Returns:
        ChartDataオブジェクト、または抽出できない場合はNone
    """
    chart_type_name = type(chart).__name__
    chart_type = CHART_TYPE_MAP.get(chart_type_name)
    
    if not chart_type:
        print(f"[WARNING] 未対応のチャートタイプ: {chart_type_name}")
        return None
    
    title = _extract_chart_title(chart)
    
    series_list = []
    categories = None
    
    chart_series = getattr(chart, 'series', [])
    
    for series in chart_series:
        series_data = _extract_series_data(series, workbook, chart_type)
        if series_data:
            series_list.append(series_data)
            
            if categories is None and chart_type != 'scatter':
                categories = _extract_categories(series, workbook)
    
    return ChartData(
        chart_type=chart_type,
        title=title,
        categories=categories,
        series=series_list
    )


def _extract_chart_title(chart) -> Optional[str]:
    """
    チャートタイトルを抽出する
    
    Args:
        chart: openpyxlのチャートオブジェクト
        
    Returns:
        タイトル文字列、またはNone
    """
    title_obj = getattr(chart, 'title', None)
    if not title_obj:
        return None
    
    tx = getattr(title_obj, 'tx', None)
    if tx:
        rich = getattr(tx, 'rich', None)
        if rich:
            p_list = getattr(rich, 'p', [])
            if p_list:
                for p in p_list:
                    r_list = getattr(p, 'r', [])
                    for r in r_list:
                        t = getattr(r, 't', None)
                        if t:
                            return str(t)
    
    return None


def _extract_series_data(series, workbook, chart_type: str) -> Optional[SeriesData]:
    """
    シリーズからデータを抽出する
    
    Args:
        series: openpyxlのシリーズオブジェクト
        workbook: openpyxlのワークブックオブジェクト
        chart_type: チャートタイプ
        
    Returns:
        SeriesDataオブジェクト、または抽出できない場合はNone
    """
    name = _extract_series_name(series, workbook)
    
    if chart_type == 'scatter':
        x_values = _extract_scatter_x_values(series, workbook)
        y_values = _extract_scatter_y_values(series, workbook)
        
        return SeriesData(
            name=name or "",
            values=y_values,
            x_values=x_values
        )
    else:
        values = _extract_series_values(series, workbook)
        
        return SeriesData(
            name=name or "",
            values=values
        )


def _extract_series_name(series, workbook) -> Optional[str]:
    """
    シリーズ名を抽出する
    
    Args:
        series: openpyxlのシリーズオブジェクト
        workbook: openpyxlのワークブックオブジェクト
        
    Returns:
        シリーズ名、またはNone
    """
    title_obj = getattr(series, 'title', None)
    if not title_obj:
        return None
    
    v = getattr(title_obj, 'v', None)
    if v:
        return str(v)
    
    str_ref = getattr(title_obj, 'strRef', None)
    if str_ref:
        f = getattr(str_ref, 'f', None)
        if f:
            values = _resolve_cell_reference(f, workbook)
            if values:
                return str(values[0])
    
    return None


def _extract_series_values(series, workbook) -> List[Any]:
    """
    シリーズの値を抽出する
    
    Args:
        series: openpyxlのシリーズオブジェクト
        workbook: openpyxlのワークブックオブジェクト
        
    Returns:
        値のリスト
    """
    val = getattr(series, 'val', None)
    if not val:
        return []
    
    num_ref = getattr(val, 'numRef', None)
    if num_ref:
        f = getattr(num_ref, 'f', None)
        if f:
            return _resolve_cell_reference(f, workbook)
    
    return []


def _extract_categories(series, workbook) -> Optional[List[str]]:
    """
    シリーズからカテゴリ（X軸ラベル）を抽出する
    
    Args:
        series: openpyxlのシリーズオブジェクト
        workbook: openpyxlのワークブックオブジェクト
        
    Returns:
        カテゴリのリスト、またはNone
    """
    cat = getattr(series, 'cat', None)
    if not cat:
        return None
    
    num_ref = getattr(cat, 'numRef', None)
    if num_ref:
        f = getattr(num_ref, 'f', None)
        if f:
            values = _resolve_cell_reference(f, workbook)
            return [str(v) for v in values]
    
    str_ref = getattr(cat, 'strRef', None)
    if str_ref:
        f = getattr(str_ref, 'f', None)
        if f:
            values = _resolve_cell_reference(f, workbook)
            return [str(v) for v in values]
    
    return None


def _extract_scatter_x_values(series, workbook) -> List[Any]:
    """
    散布図のX値を抽出する
    
    Args:
        series: openpyxlのシリーズオブジェクト
        workbook: openpyxlのワークブックオブジェクト
        
    Returns:
        X値のリスト
    """
    x_val = getattr(series, 'xVal', None)
    if not x_val:
        return []
    
    num_ref = getattr(x_val, 'numRef', None)
    if num_ref:
        f = getattr(num_ref, 'f', None)
        if f:
            return _resolve_cell_reference(f, workbook)
    
    return []


def _extract_scatter_y_values(series, workbook) -> List[Any]:
    """
    散布図のY値を抽出する
    
    Args:
        series: openpyxlのシリーズオブジェクト
        workbook: openpyxlのワークブックオブジェクト
        
    Returns:
        Y値のリスト
    """
    y_val = getattr(series, 'yVal', None)
    if not y_val:
        return []
    
    num_ref = getattr(y_val, 'numRef', None)
    if num_ref:
        f = getattr(num_ref, 'f', None)
        if f:
            return _resolve_cell_reference(f, workbook)
    
    return []


def _resolve_cell_reference(ref_str: str, workbook) -> List[Any]:
    """
    セル参照文字列を解決して値を取得する
    
    Args:
        ref_str: セル参照文字列（例: "'Sheet1'!$A$1:$A$5"）
        workbook: openpyxlのワークブックオブジェクト
        
    Returns:
        セル値のリスト
    """
    if not ref_str:
        return []
    
    pattern = r"'?([^'!]+)'?!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?"
    match = re.match(pattern, ref_str)
    
    if not match:
        return []
    
    sheet_name = match.group(1)
    col1 = match.group(2)
    row1 = int(match.group(3))
    col2 = match.group(4) or col1
    row2 = int(match.group(5)) if match.group(5) else row1
    
    try:
        ws = workbook[sheet_name]
    except KeyError:
        print(f"[WARNING] シートが見つかりません: {sheet_name}")
        return []
    
    values = []
    
    col1_idx = _col_to_idx(col1)
    col2_idx = _col_to_idx(col2)
    
    for row in range(row1, row2 + 1):
        for col in range(col1_idx, col2_idx + 1):
            cell_value = ws.cell(row=row, column=col).value
            values.append(cell_value)
    
    return values


def _col_to_idx(col_str: str) -> int:
    """
    列文字を列インデックスに変換する
    
    Args:
        col_str: 列文字（例: "A", "AA"）
        
    Returns:
        列インデックス（1始まり）
    """
    result = 0
    for char in col_str:
        result = result * 26 + (ord(char.upper()) - ord('A') + 1)
    return result
