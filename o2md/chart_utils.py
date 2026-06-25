#!/usr/bin/env python3
"""
チャートデータ抽出ユーティリティ

Excel/Wordのチャートからデータを抽出し、Markdownテーブルに変換するための
共通モジュール。

対応チャートタイプ:
- 棒グラフ (bar)
- 折れ線グラフ (line)
- 円グラフ (pie)
- 散布図 (scatter)
"""

from dataclasses import dataclass, field
from typing import List, Optional, Union


@dataclass
class SeriesData:
    """チャートシリーズのデータを保持するクラス"""
    name: str
    values: List[Union[float, int, str]]
    x_values: Optional[List[Union[float, int, str]]] = None


@dataclass
class ChartData:
    """チャート全体のデータを保持するクラス"""
    chart_type: str
    title: Optional[str]
    categories: Optional[List[str]]
    series: List[SeriesData] = field(default_factory=list)


def chart_data_to_markdown(chart_data: ChartData) -> str:
    """
    ChartDataをMarkdownテーブルに変換する
    
    Args:
        chart_data: 変換するチャートデータ
        
    Returns:
        Markdown形式のテーブル文字列
    """
    md_lines = []
    
    md_lines.append("\n### Chart")
    if chart_data.title:
        md_lines[-1] += f": {chart_data.title}"
    md_lines.append("")
    
    if not chart_data.series:
        md_lines.append("[データなし]")
        md_lines.append("")
        return "\n".join(md_lines)
    
    if chart_data.chart_type == 'scatter':
        table_md = _build_scatter_table(chart_data)
    else:
        table_md = _build_standard_table(chart_data)
    
    md_lines.append(table_md)
    md_lines.append("")
    
    return "\n".join(md_lines)


def _build_standard_table(chart_data: ChartData) -> str:
    """
    棒グラフ、折れ線グラフ、円グラフ用のMarkdownテーブルを構築する
    
    Args:
        chart_data: チャートデータ
        
    Returns:
        Markdownテーブル文字列
    """
    series_names = []
    for i, s in enumerate(chart_data.series):
        if s.name:
            series_names.append(s.name)
        else:
            series_names.append(f"Series{i+1}")
    
    header = ["Category"] + series_names
    
    rows = []
    categories = chart_data.categories
    if not categories and chart_data.series:
        first_series = chart_data.series[0]
        if first_series.values:
            categories = [str(i+1) for i in range(len(first_series.values))]
    
    num_categories = len(categories) if categories else 0
    
    for idx in range(num_categories):
        if categories:
            row = [str(categories[idx])]
        else:
            row = [str(idx + 1)]
        
        for series in chart_data.series:
            if idx < len(series.values):
                val = series.values[idx]
                row.append(_format_value(val))
            else:
                row.append("")
        rows.append(row)
    
    return _build_markdown_table(header, rows)


def _build_scatter_table(chart_data: ChartData) -> str:
    """
    散布図用のMarkdownテーブルを構築する
    
    Args:
        chart_data: チャートデータ
        
    Returns:
        Markdownテーブル文字列
    """
    series_names = []
    for i, s in enumerate(chart_data.series):
        if s.name:
            series_names.append(s.name)
        else:
            series_names.append(f"Series{i+1}")
    
    header = ["X"] + series_names
    
    all_x = set()
    for series in chart_data.series:
        if series.x_values:
            all_x.update(series.x_values)
    x_values = sorted(all_x, key=lambda x: (isinstance(x, str), x))
    
    rows = []
    for x in x_values:
        row = [_format_value(x)]
        for series in chart_data.series:
            if series.x_values and x in series.x_values:
                idx = series.x_values.index(x)
                if idx < len(series.values):
                    row.append(_format_value(series.values[idx]))
                else:
                    row.append("")
            else:
                row.append("")
        rows.append(row)
    
    return _build_markdown_table(header, rows)


def _build_markdown_table(header: List[str], rows: List[List[str]]) -> str:
    """
    ヘッダーと行データからMarkdownテーブルを構築する
    
    Args:
        header: ヘッダー行のリスト
        rows: データ行のリスト
        
    Returns:
        Markdownテーブル文字列
    """
    lines = []
    
    header_line = "| " + " | ".join(header) + " |"
    lines.append(header_line)
    
    separator = "|" + "|".join(["---"] * len(header)) + "|"
    lines.append(separator)
    
    for row in rows:
        row_line = "| " + " | ".join(row) + " |"
        lines.append(row_line)
    
    return "\n".join(lines)


def _format_value(val: Union[float, int, str, None]) -> str:
    """
    値を表示用にフォーマットする
    
    Args:
        val: フォーマットする値
        
    Returns:
        フォーマットされた文字列
    """
    if val is None:
        return ""
    if isinstance(val, float):
        if val == int(val):
            return str(int(val))
        return f"{val:.2f}"
    return str(val)
