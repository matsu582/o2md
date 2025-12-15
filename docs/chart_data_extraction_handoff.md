# Excel/Word チャートデータ抽出 引き継ぎ資料

## 概要

本資料は、o2mdにExcel/Wordのチャートデータ抽出機能を追加するための設計メモと実装方針をまとめたものです。

### 対応チャートタイプ
- 棒グラフ (BarChart)
- 折れ線グラフ (LineChart)
- 円グラフ (PieChart)
- 散布図 (ScatterChart)

### 目標
チャートからカテゴリ（X軸）とシリーズ（Y軸）のデータを抽出し、Markdownテーブルとして出力する。

---

## 1. 参考実装: markitdownのPowerPointチャート抽出

markitdownの `_pptx_converter.py` にあるチャート抽出コードが参考になります。

```python
# markitdown/converters/_pptx_converter.py より
def _convert_chart_to_markdown(self, chart):
    try:
        md = "\n\n### Chart"
        if chart.has_title:
            md += f": {chart.chart_title.text_frame.text}"
        md += "\n\n"
        data = []
        category_names = [c.label for c in chart.plots[0].categories]
        series_names = [s.name for s in chart.series]
        data.append(["Category"] + series_names)

        for idx, category in enumerate(category_names):
            row = [category]
            for series in chart.series:
                row.append(series.values[idx])
            data.append(row)

        markdown_table = []
        for row in data:
            markdown_table.append("| " + " | ".join(map(str, row)) + " |")
        header = markdown_table[0]
        separator = "|" + "|".join(["---"] * len(data[0])) + "|"
        return md + "\n".join([header, separator] + markdown_table[1:])
    except ValueError as e:
        if "unsupported plot type" in str(e):
            return "\n\n[unsupported chart]\n\n"
    except Exception:
        return "\n\n[unsupported chart]\n\n"
```

**ポイント:**
- `chart.plots[0].categories` からカテゴリ（X軸ラベル）を取得
- `chart.series` から各シリーズを取得
- `series.values[idx]` で各カテゴリに対応する値を取得
- 対応していないチャートタイプは `[unsupported chart]` を返す

---

## 2. Excelチャートデータ抽出

### 2.1 チャートオブジェクトへのアクセス

openpyxlでワークシートに紐づくチャートにアクセスする方法（要検証）:

```python
from openpyxl import load_workbook

wb = load_workbook("example.xlsx", data_only=True)
for ws in wb.worksheets:
    # チャートは非公開属性 _charts に格納されている可能性が高い
    charts = getattr(ws, "_charts", [])  # または getattr(ws, "charts", [])
    for chart in charts:
        print(ws.title, type(chart), getattr(chart, "type", None))
```

### 2.2 最初の探索手順（次セッションで実施）

1. サンプルの.xlsxファイル（チャート入り）を用意
2. 以下のコードで構造を確認:

```python
from openpyxl import load_workbook

wb = load_workbook("sample_with_chart.xlsx")
for ws in wb.worksheets:
    charts = getattr(ws, "_charts", [])
    for i, chart in enumerate(charts):
        print(f"=== Chart {i} ===")
        print(f"Type: {type(chart)}")
        print(f"Dir: {[a for a in dir(chart) if not a.startswith('_')]}")
        
        # シリーズの確認
        if hasattr(chart, "series"):
            for j, s in enumerate(chart.series):
                print(f"  Series {j}: {type(s)}")
                print(f"    name: {getattr(s, 'title', None) or getattr(s, 'name', None)}")
                print(f"    values: {getattr(s, 'values', None)}")
        
        # カテゴリの確認
        if hasattr(chart, "categories"):
            print(f"  Categories: {chart.categories}")
```

### 2.3 データ抽出の擬似コード（要検証）

```python
def extract_excel_chart_data(chart, workbook):
    """
    chart: openpyxl.chart系のオブジェクト
    workbook: openpyxl.Workbook
    戻り値: {
        "title": str or None,
        "categories": [label1, label2, ...] or None,
        "series": [
            {"name": "Series1", "values": [v1, v2, ...]},
            ...
        ]
    }
    """
    info = {
        "title": None,
        "categories": None,
        "series": [],
    }

    # タイトル
    if getattr(chart, "title", None):
        info["title"] = str(chart.title)

    # 系列一覧
    for s in getattr(chart, "series", []):
        name = getattr(s, "title", None) or getattr(s, "name", None)
        values = resolve_data_ref(s.values, workbook)
        info["series"].append({
            "name": str(name) if name is not None else "",
            "values": values
        })

    # カテゴリ
    categories_attr = getattr(chart, "categories", None)
    if categories_attr:
        info["categories"] = resolve_data_ref(categories_attr, workbook)

    return info
```

### 2.4 セル参照の解決ヘルパー（要検証）

```python
import re

def resolve_data_ref(ref_obj, workbook):
    """
    ref_obj: series.values や chart.categories が持っている参照オブジェクト
    workbook: openpyxl.Workbook
    戻り値: [value1, value2, ...]
    """
    if ref_obj is None:
        return []

    # 1. イテラブルな場合はそのまま展開を試みる
    try:
        values = list(ref_obj)
        if values and hasattr(values[0], "value"):
            return [c.value for c in values]
        return values
    except TypeError:
        pass

    # 2. 文字列参照 "Sheet1!$A$2:$A$5" の場合
    ref_str = getattr(ref_obj, "range", None) or getattr(ref_obj, "ref", None) or str(ref_obj)
    
    # 参照文字列をパース
    m = re.match(r"'?(.+?)'?!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)", ref_str)
    if not m:
        return []

    sheet_name, col1, row1, col2, row2 = m.groups()
    ws = workbook[sheet_name]
    
    values = []
    for row in ws[f"{col1}{row1}":f"{col2}{row2}"]:
        for cell in row:
            values.append(cell.value)
    return values
```

### 2.5 チャートタイプごとの違い

| チャートタイプ | カテゴリ | 値 | 備考 |
|--------------|---------|-----|------|
| 棒グラフ (BarChart) | `chart.categories` | `series.values` | 標準的な構造 |
| 折れ線グラフ (LineChart) | `chart.categories` | `series.values` | 標準的な構造 |
| 円グラフ (PieChart) | `chart.categories` または `series.labels` | `series.values` | 通常1シリーズのみ |
| 散布図 (ScatterChart) | `series.xValues` | `series.yValues` | X/Y両方が数値 |

**散布図のMarkdownテーブル形式（要仕様決定）:**
```
| X | Series1 | Series2 |
|---|---------|---------|
| 1 | 10 | 20 |
| 2 | 15 | 25 |
```

---

## 3. Wordチャートデータ抽出

### 3.1 Wordチャートの内部構造

Wordのチャートは Office Open XML (OOXML) のDrawingML形式で保存されています。

**ファイル構造:**
```
document.docx (ZIP)
├── word/
│   ├── document.xml
│   └── charts/
│       ├── chart1.xml
│       └── chart2.xml
└── word/embeddings/
    └── Microsoft_Excel_Sheet1.xlsx  (埋め込みExcel)
```

**chart.xmlの構造:**
```xml
<c:chartSpace>
  <c:chart>
    <c:title>...</c:title>
    <c:plotArea>
      <c:barChart>  <!-- または lineChart, pieChart, scatterChart -->
        <c:ser>  <!-- シリーズ -->
          <c:tx>  <!-- シリーズ名 -->
            <c:strRef>
              <c:f>Sheet1!$B$1</c:f>
            </c:strRef>
          </c:tx>
          <c:cat>  <!-- カテゴリ -->
            <c:strRef>
              <c:f>Sheet1!$A$2:$A$5</c:f>
            </c:strRef>
          </c:cat>
          <c:val>  <!-- 値 -->
            <c:numRef>
              <c:f>Sheet1!$B$2:$B$5</c:f>
              <c:numCache>  <!-- キャッシュ値 -->
                <c:pt idx="0"><c:v>10</c:v></c:pt>
                <c:pt idx="1"><c:v>20</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>
```

### 3.2 データ取得の2つのパターン

1. **キャッシュ値から取得** (`<c:numCache>`, `<c:strCache>`)
   - chart.xml内に直接値がキャッシュされている
   - 実装が簡単、埋め込みExcelを開く必要なし
   - ただしキャッシュが古い可能性あり

2. **埋め込みExcelから取得** (`<c:numRef>`, `<c:strRef>`)
   - セル参照を解決して埋め込みExcelから値を読む
   - 正確だが実装が複雑

**MVP推奨:** まずはキャッシュ値からの取得のみ対応

### 3.3 最初の探索手順（次セッションで実施）

1. サンプルの.docxファイル（チャート入り）を用意
2. 以下のコードで構造を確認:

```python
import zipfile
import xml.etree.ElementTree as ET

# 名前空間の定義
CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"

def explore_word_charts(docx_path):
    with zipfile.ZipFile(docx_path, 'r') as zf:
        # チャートファイルの一覧
        chart_files = [f for f in zf.namelist() if f.startswith('word/charts/chart')]
        print(f"Found {len(chart_files)} chart(s)")
        
        for chart_file in chart_files:
            print(f"\n=== {chart_file} ===")
            with zf.open(chart_file) as f:
                tree = ET.parse(f)
                root = tree.getroot()
                
                # チャートタイプを検出
                for chart_type in ['barChart', 'lineChart', 'pieChart', 'scatterChart']:
                    elements = root.findall(f".//{{{CHART_NS}}}{chart_type}")
                    if elements:
                        print(f"Chart type: {chart_type}")
                        
                        # シリーズを確認
                        for ser in root.findall(f".//{{{CHART_NS}}}ser"):
                            print(f"  Series found")
                            
                            # カテゴリ参照
                            cat = ser.find(f".//{{{CHART_NS}}}cat")
                            if cat is not None:
                                ref = cat.find(f".//{{{CHART_NS}}}f")
                                if ref is not None:
                                    print(f"    Category ref: {ref.text}")
                            
                            # 値参照
                            val = ser.find(f".//{{{CHART_NS}}}val")
                            if val is not None:
                                ref = val.find(f".//{{{CHART_NS}}}f")
                                if ref is not None:
                                    print(f"    Value ref: {ref.text}")
                                
                                # キャッシュ値
                                cache = val.find(f".//{{{CHART_NS}}}numCache")
                                if cache is not None:
                                    pts = cache.findall(f".//{{{CHART_NS}}}pt")
                                    values = [pt.find(f"{{{CHART_NS}}}v").text for pt in pts]
                                    print(f"    Cached values: {values[:5]}...")

explore_word_charts("sample_with_chart.docx")
```

### 3.4 チャートXML解析の擬似コード（要検証）

```python
import zipfile
import xml.etree.ElementTree as ET

CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"

def extract_word_chart_data(docx_path):
    """
    Wordドキュメントからチャートデータを抽出
    戻り値: [ChartData, ...]
    """
    charts = []
    
    with zipfile.ZipFile(docx_path, 'r') as zf:
        chart_files = [f for f in zf.namelist() if f.startswith('word/charts/chart')]
        
        for chart_file in chart_files:
            with zf.open(chart_file) as f:
                tree = ET.parse(f)
                root = tree.getroot()
                
                chart_data = parse_chart_xml(root)
                if chart_data:
                    charts.append(chart_data)
    
    return charts

def parse_chart_xml(root):
    """
    chart.xmlのルート要素からチャートデータを抽出
    """
    # チャートタイプを検出
    chart_type = None
    for ct in ['barChart', 'lineChart', 'pieChart', 'scatterChart']:
        if root.find(f".//{{{CHART_NS}}}{ct}") is not None:
            chart_type = ct
            break
    
    if chart_type is None:
        return None
    
    # タイトル
    title = None
    title_elem = root.find(f".//{{{CHART_NS}}}title")
    if title_elem is not None:
        # タイトルテキストの取得（構造が複雑なので要調整）
        pass
    
    # シリーズを抽出
    series_list = []
    for ser in root.findall(f".//{{{CHART_NS}}}ser"):
        series_data = extract_series_data(ser, chart_type)
        if series_data:
            series_list.append(series_data)
    
    return {
        "type": chart_type,
        "title": title,
        "series": series_list
    }

def extract_series_data(ser_elem, chart_type):
    """
    <c:ser>要素からシリーズデータを抽出
    """
    # シリーズ名
    name = None
    tx = ser_elem.find(f".//{{{CHART_NS}}}tx")
    if tx is not None:
        v = tx.find(f".//{{{CHART_NS}}}v")
        if v is not None:
            name = v.text
    
    # カテゴリ（散布図以外）
    categories = None
    cat = ser_elem.find(f".//{{{CHART_NS}}}cat")
    if cat is not None:
        categories = extract_cached_values(cat)
    
    # 値
    values = None
    val = ser_elem.find(f".//{{{CHART_NS}}}val")
    if val is not None:
        values = extract_cached_values(val)
    
    # 散布図の場合はxVal/yVal
    if chart_type == 'scatterChart':
        x_val = ser_elem.find(f".//{{{CHART_NS}}}xVal")
        y_val = ser_elem.find(f".//{{{CHART_NS}}}yVal")
        if x_val is not None:
            categories = extract_cached_values(x_val)
        if y_val is not None:
            values = extract_cached_values(y_val)
    
    return {
        "name": name,
        "categories": categories,
        "values": values
    }

def extract_cached_values(parent_elem):
    """
    <c:numCache>または<c:strCache>からキャッシュ値を抽出
    """
    # 数値キャッシュ
    num_cache = parent_elem.find(f".//{{{CHART_NS}}}numCache")
    if num_cache is not None:
        pts = num_cache.findall(f".//{{{CHART_NS}}}pt")
        return [float(pt.find(f"{{{CHART_NS}}}v").text) for pt in pts]
    
    # 文字列キャッシュ
    str_cache = parent_elem.find(f".//{{{CHART_NS}}}strCache")
    if str_cache is not None:
        pts = str_cache.findall(f".//{{{CHART_NS}}}pt")
        return [pt.find(f"{{{CHART_NS}}}v").text for pt in pts]
    
    return None
```

---

## 4. 共通設計: ChartDataモデルとMarkdown変換

### 4.1 共通データモデル

```python
from dataclasses import dataclass
from typing import List, Optional, Union

@dataclass
class SeriesData:
    name: str
    values: List[Union[float, int, str]]
    x_values: Optional[List[Union[float, int, str]]] = None  # 散布図用

@dataclass
class ChartData:
    chart_type: str  # 'bar', 'line', 'pie', 'scatter'
    title: Optional[str]
    categories: Optional[List[str]]
    series: List[SeriesData]
```

### 4.2 Markdown変換（共通）

```python
def chart_data_to_markdown(chart_data: ChartData) -> str:
    """
    ChartDataをMarkdownテーブルに変換
    """
    md = "\n\n### Chart"
    if chart_data.title:
        md += f": {chart_data.title}"
    md += "\n\n"
    
    if not chart_data.series:
        return md + "[no data]\n\n"
    
    # 散布図の場合
    if chart_data.chart_type == 'scatter':
        return md + _scatter_to_markdown(chart_data)
    
    # 棒・折れ線・円グラフの場合
    return md + _standard_chart_to_markdown(chart_data)

def _standard_chart_to_markdown(chart_data: ChartData) -> str:
    """
    棒・折れ線・円グラフのMarkdown変換
    """
    # ヘッダー行
    series_names = [s.name or f"Series{i+1}" for i, s in enumerate(chart_data.series)]
    header = ["Category"] + series_names
    
    # データ行
    rows = []
    categories = chart_data.categories or chart_data.series[0].values
    num_categories = len(categories) if categories else 0
    
    for idx in range(num_categories):
        row = [str(categories[idx]) if categories else str(idx)]
        for series in chart_data.series:
            if idx < len(series.values):
                row.append(str(series.values[idx]))
            else:
                row.append("")
        rows.append(row)
    
    # Markdownテーブル生成
    lines = []
    lines.append("| " + " | ".join(header) + " |")
    lines.append("|" + "|".join(["---"] * len(header)) + "|")
    for row in rows:
        lines.append("| " + " | ".join(row) + " |")
    
    return "\n".join(lines)

def _scatter_to_markdown(chart_data: ChartData) -> str:
    """
    散布図のMarkdown変換
    """
    # ヘッダー行: X | Series1 | Series2 | ...
    series_names = [s.name or f"Series{i+1}" for i, s in enumerate(chart_data.series)]
    header = ["X"] + series_names
    
    # X値の統合（全シリーズのX値をマージ）
    all_x = set()
    for series in chart_data.series:
        if series.x_values:
            all_x.update(series.x_values)
    x_values = sorted(all_x)
    
    # データ行
    rows = []
    for x in x_values:
        row = [str(x)]
        for series in chart_data.series:
            if series.x_values and x in series.x_values:
                idx = series.x_values.index(x)
                row.append(str(series.values[idx]) if idx < len(series.values) else "")
            else:
                row.append("")
        rows.append(row)
    
    # Markdownテーブル生成
    lines = []
    lines.append("| " + " | ".join(header) + " |")
    lines.append("|" + "|".join(["---"] * len(header)) + "|")
    for row in rows:
        lines.append("| " + " | ".join(row) + " |")
    
    return "\n".join(lines)
```

---

## 5. 実装計画

### フェーズ1: PoC（概念実証）

1. **サンプルファイルの準備**
   - チャート入りの.xlsxファイル（棒・折れ線・円・散布図）
   - チャート入りの.docxファイル（同上）

2. **構造探索**
   - 上記の探索コードを実行して、実際のオブジェクト構造を確認
   - openpyxlのバージョンと`_charts`属性の有無を確認

3. **単体スクリプトでの検証**
   - `x2md_charts_poc.py`: Excelチャート抽出のPoC
   - `d2md_charts_poc.py`: Wordチャート抽出のPoC

### フェーズ2: モジュール実装

1. **共通モジュール**
   - `chart_utils.py`: ChartDataモデル、Markdown変換、セル参照解決

2. **Excel用モジュール**
   - `x2md_charts.py`: openpyxlからチャートデータを抽出

3. **Word用モジュール**
   - `d2md_charts.py`: chart.xmlからチャートデータを抽出

### フェーズ3: 統合

1. **x2md.pyへの統合**
   - シート処理時にチャートを検出
   - チャートデータをMarkdownテーブルとして出力
   - 既存の画像ベース処理との共存

2. **d2md.pyへの統合**
   - ドキュメント処理時にチャートを検出
   - チャートデータをMarkdownテーブルとして出力

---

## 6. 注意事項

### 技術的な注意点

1. **openpyxlのバージョン依存**
   - `ws._charts`は非公開属性のため、将来変更される可能性あり
   - 現在のバージョンを確認: `pip show openpyxl`

2. **チャートタイプの制限**
   - 3D、積み上げ、複合グラフは対象外（MVPでは）
   - 対応外のタイプは`[unsupported chart type]`を返す

3. **データ参照の複雑さ**
   - 別シート参照、名前付き範囲、非連続範囲は複雑
   - MVPではシンプルな連続範囲のみ対応

4. **既存処理との共存**
   - 既存の画像ベース処理（LibreOffice経由）を壊さない
   - チャートデータ抽出は「追加の出力」として扱う

### テスト要件

1. **Excelテスト**
   - 各チャートタイプでデータが正しく抽出されるか
   - 既存のdebug_workbooks出力に影響がないか

2. **Wordテスト**
   - 各チャートタイプでデータが正しく抽出されるか
   - 既存の図形処理に影響がないか

---

## 7. 参考リンク

- [openpyxl Charts Documentation](https://openpyxl.readthedocs.io/en/stable/charts/introduction.html)
- [python-docx Documentation](https://python-docx.readthedocs.io/)
- [Office Open XML (OOXML) Chart Specification](https://docs.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/)
- [markitdown _pptx_converter.py](https://github.com/microsoft/markitdown) - チャート抽出の参考実装
