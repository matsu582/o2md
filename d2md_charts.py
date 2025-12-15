#!/usr/bin/env python3
"""
Wordチャートデータ抽出モジュール

Word文書（.docx）からチャートデータを抽出し、
ChartDataオブジェクトに変換する。

chart.xmlのキャッシュ値から取得する方式を採用。

対応チャートタイプ:
- 棒グラフ (barChart)
- 折れ線グラフ (lineChart)
- 円グラフ (pieChart)
- 散布図 (scatterChart)
"""

import zipfile
import xml.etree.ElementTree as ET
from typing import List, Optional, Any

from chart_utils import ChartData, SeriesData


CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

CHART_TYPE_MAP = {
    'barChart': 'bar',
    'bar3DChart': 'bar',
    'lineChart': 'line',
    'line3DChart': 'line',
    'pieChart': 'pie',
    'pie3DChart': 'pie',
    'scatterChart': 'scatter',
    'areaChart': 'area',
    'area3DChart': 'area',
    'radarChart': 'radar',
    'doughnutChart': 'doughnut',
}


def extract_charts_from_docx(docx_path: str) -> List[ChartData]:
    """
    Word文書からチャートデータを抽出する
    
    Args:
        docx_path: Word文書のパス
        
    Returns:
        ChartDataオブジェクトのリスト
    """
    charts = []
    
    try:
        with zipfile.ZipFile(docx_path, 'r') as zf:
            chart_files = [
                f for f in zf.namelist() 
                if f.startswith('word/charts/chart') and f.endswith('.xml')
            ]
            
            for chart_file in chart_files:
                try:
                    with zf.open(chart_file) as f:
                        tree = ET.parse(f)
                        root = tree.getroot()
                        
                        chart_data = _parse_chart_xml(root)
                        if chart_data:
                            charts.append(chart_data)
                except Exception as e:
                    print(f"[WARNING] チャートファイル解析エラー: {chart_file} - {e}")
                    
    except zipfile.BadZipFile:
        print(f"[WARNING] 無効なZIPファイル: {docx_path}")
    except Exception as e:
        print(f"[WARNING] チャート抽出エラー: {e}")
    
    return charts


def _parse_chart_xml(root: ET.Element) -> Optional[ChartData]:
    """
    chart.xmlのルート要素からチャートデータを抽出する
    
    Args:
        root: XMLルート要素
        
    Returns:
        ChartDataオブジェクト、または抽出できない場合はNone
    """
    chart_type = None
    chart_type_elem = None
    
    for ct_name in CHART_TYPE_MAP.keys():
        elem = root.find(f".//{{{CHART_NS}}}{ct_name}")
        if elem is not None:
            chart_type = CHART_TYPE_MAP[ct_name]
            chart_type_elem = elem
            break
    
    if chart_type is None:
        return None
    
    title = _extract_title(root)
    
    series_list = []
    categories = None
    
    for ser_elem in chart_type_elem.findall(f"{{{CHART_NS}}}ser"):
        series_data = _extract_series_from_xml(ser_elem, chart_type)
        if series_data:
            series_list.append(series_data)
            
            if categories is None and chart_type != 'scatter':
                categories = _extract_categories_from_xml(ser_elem)
    
    return ChartData(
        chart_type=chart_type,
        title=title,
        categories=categories,
        series=series_list
    )


def _extract_title(root: ET.Element) -> Optional[str]:
    """
    チャートタイトルを抽出する
    
    Args:
        root: XMLルート要素
        
    Returns:
        タイトル文字列、またはNone
    """
    title_elem = root.find(f".//{{{CHART_NS}}}title")
    if title_elem is None:
        return None
    
    t_elems = title_elem.findall(f".//{{{DRAWING_NS}}}t")
    if t_elems:
        texts = [t.text for t in t_elems if t.text]
        if texts:
            return "".join(texts)
    
    return None


def _extract_series_from_xml(ser_elem: ET.Element, chart_type: str) -> Optional[SeriesData]:
    """
    <c:ser>要素からシリーズデータを抽出する
    
    Args:
        ser_elem: シリーズXML要素
        chart_type: チャートタイプ
        
    Returns:
        SeriesDataオブジェクト、または抽出できない場合はNone
    """
    name = _extract_series_name_from_xml(ser_elem)
    
    if chart_type == 'scatter':
        x_values = _extract_scatter_x_from_xml(ser_elem)
        y_values = _extract_scatter_y_from_xml(ser_elem)
        
        return SeriesData(
            name=name or "",
            values=y_values,
            x_values=x_values
        )
    else:
        values = _extract_values_from_xml(ser_elem)
        
        return SeriesData(
            name=name or "",
            values=values
        )


def _extract_series_name_from_xml(ser_elem: ET.Element) -> Optional[str]:
    """
    シリーズ名を抽出する
    
    Args:
        ser_elem: シリーズXML要素
        
    Returns:
        シリーズ名、またはNone
    """
    tx_elem = ser_elem.find(f"{{{CHART_NS}}}tx")
    if tx_elem is None:
        return None
    
    v_elem = tx_elem.find(f".//{{{CHART_NS}}}v")
    if v_elem is not None and v_elem.text:
        return v_elem.text
    
    str_cache = tx_elem.find(f".//{{{CHART_NS}}}strCache")
    if str_cache is not None:
        pt_elem = str_cache.find(f"{{{CHART_NS}}}pt")
        if pt_elem is not None:
            v_elem = pt_elem.find(f"{{{CHART_NS}}}v")
            if v_elem is not None and v_elem.text:
                return v_elem.text
    
    return None


def _extract_values_from_xml(ser_elem: ET.Element) -> List[Any]:
    """
    シリーズの値を抽出する
    
    Args:
        ser_elem: シリーズXML要素
        
    Returns:
        値のリスト
    """
    val_elem = ser_elem.find(f"{{{CHART_NS}}}val")
    if val_elem is None:
        return []
    
    return _extract_cached_values(val_elem)


def _extract_categories_from_xml(ser_elem: ET.Element) -> Optional[List[str]]:
    """
    カテゴリ（X軸ラベル）を抽出する
    
    Args:
        ser_elem: シリーズXML要素
        
    Returns:
        カテゴリのリスト、またはNone
    """
    cat_elem = ser_elem.find(f"{{{CHART_NS}}}cat")
    if cat_elem is None:
        return None
    
    values = _extract_cached_values(cat_elem)
    if values:
        return [str(v) for v in values]
    
    return None


def _extract_scatter_x_from_xml(ser_elem: ET.Element) -> List[Any]:
    """
    散布図のX値を抽出する
    
    Args:
        ser_elem: シリーズXML要素
        
    Returns:
        X値のリスト
    """
    x_val_elem = ser_elem.find(f"{{{CHART_NS}}}xVal")
    if x_val_elem is None:
        return []
    
    return _extract_cached_values(x_val_elem)


def _extract_scatter_y_from_xml(ser_elem: ET.Element) -> List[Any]:
    """
    散布図のY値を抽出する
    
    Args:
        ser_elem: シリーズXML要素
        
    Returns:
        Y値のリスト
    """
    y_val_elem = ser_elem.find(f"{{{CHART_NS}}}yVal")
    if y_val_elem is None:
        return []
    
    return _extract_cached_values(y_val_elem)


def _extract_cached_values(parent_elem: ET.Element) -> List[Any]:
    """
    キャッシュ値を抽出する
    
    Args:
        parent_elem: 親XML要素
        
    Returns:
        値のリスト
    """
    num_cache = parent_elem.find(f".//{{{CHART_NS}}}numCache")
    if num_cache is not None:
        values = []
        for pt in num_cache.findall(f"{{{CHART_NS}}}pt"):
            v_elem = pt.find(f"{{{CHART_NS}}}v")
            if v_elem is not None and v_elem.text:
                try:
                    values.append(float(v_elem.text))
                except ValueError:
                    values.append(v_elem.text)
        return values
    
    str_cache = parent_elem.find(f".//{{{CHART_NS}}}strCache")
    if str_cache is not None:
        values = []
        for pt in str_cache.findall(f"{{{CHART_NS}}}pt"):
            v_elem = pt.find(f"{{{CHART_NS}}}v")
            if v_elem is not None and v_elem.text:
                values.append(v_elem.text)
        return values
    
    return []
