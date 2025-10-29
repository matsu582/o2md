#!/usr/bin/env python3
"""
Excel to Markdown Converter
Excelファイルをシートごとに詳細なMarkdown形式に変換するツール

特徴:
- セル内の改行を<br>タグで表現
- 罫線を意識した表の作成
- 図形や埋め込み画像を抽出してMarkdownに挿入

このファイルはxlsxファイルを解析してシートごとにMarkdownを生成します。
デバッグ情報や一時ファイルの出力が含まれます。
"""

import os
import sys
import tempfile
import subprocess
import shutil
from pathlib import Path
from typing import List, Dict, Tuple, Optional, Any, Set
import io
import zipfile
import xml.etree.ElementTree as ET

from utils import get_libreoffice_path, col_letter, normalize_excel_path, get_xml_from_zip, xml_exists_in_zip, extract_anchor_id, anchor_has_drawable as utils_anchor_has_drawable
from isolated_group_renderer import IsolatedGroupRenderer

try:
    import openpyxl
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    from openpyxl.drawing.image import Image as XLImage
except ImportError as e:
    raise ImportError(
        "openpyxlライブラリが必要です: pip install openpyxl または uv sync を実行してください"
    ) from e

try:
    from PIL import Image
except ImportError as e:
    raise ImportError(
        "Pillowライブラリが必要です: pip install pillow または uv sync を実行してください"
    ) from e

try:
    import fitz
except ImportError as e:
    raise ImportError(
        "PyMuPDFライブラリが必要です: pip install PyMuPDF または uv sync を実行してください"
    ) from e

# 設定定数
LIBREOFFICE_PATH = get_libreoffice_path()

# DPI設定
DEFAULT_DPI = 600
IMAGE_QUALITY = 100
IMAGE_BORDER_SIZE = 8

# スキャン設定
MAX_HEAD_SCAN_ROWS = 12
MAX_SCAN_COLUMNS = 60




# グローバルverboseフラグ
_VERBOSE = False

def set_verbose(verbose: bool):
    """verboseモードを設定"""
    global _VERBOSE
    _VERBOSE = verbose

def is_verbose() -> bool:
    """verboseモードかどうかを返す"""
    return _VERBOSE

def debug_print(*args, **kwargs):
    """verboseモード時のみ出力するデバッグ用print"""
    if _VERBOSE:
        print(*args, **kwargs)

class ExcelToMarkdownConverter:
    class _LoggingList(list):
        """デバッグ用にappend/insert操作をログ出力するlistのラッパー

        標準出力にログを出力し、可能であればコンバータのデバッグログにも書き込みます。
        """
        def __init__(self, owner, *args):
            super().__init__(*args)
            self._owner = owner

        def append(self, item):
            debug_print(f"[MD_APPEND] {repr(item)}")
            try:
                if isinstance(item, str) and item.strip() == '---' and len(self) and isinstance(self[-1], str) and self[-1].strip() == '---':
                    return
            except (ValueError, TypeError):
                pass

            return super().append(item)

    def __init__(self, excel_file_path: str, output_dir=None, debug_mode=False, shape_metadata=False):
        """コンバータインスタンスの初期化

        CLIから使用できるように、最小限で安全なコンストラクタを提供します。
        意図的に保守的な初期化を維持し、メソッド間で使用される共通のシート毎の
        一時的な状態を準備します。
        """
        self.excel_file = excel_file_path
        self.base_name = Path(excel_file_path).stem
        if output_dir:
            self.output_dir = output_dir
        else:
            self.output_dir = os.path.join(os.getcwd(), "output")
        self.images_dir = os.path.join(self.output_dir, "images")
        
        self.debug_mode = debug_mode
        self.shape_metadata = shape_metadata

        os.makedirs(self.output_dir, exist_ok=True)
        os.makedirs(self.images_dir, exist_ok=True)

        self.markdown_lines = self._LoggingList(self)
        self.image_counter = 0

        self._init_per_sheet_state()

        self.workbook = load_workbook(excel_file_path, data_only=True)
        print(f"[INFO] Excelワークブック読み込み完了: {excel_file_path}")

    def _init_per_sheet_state(self):
        """シート毎の状態変数を初期化"""
        self._cell_to_md_index = {}
        self._sheet_shape_images = {}
        self._sheet_shape_next_idx = {}
        self._sheet_shapes_generated = set()
        self._sheet_shape_image_start_rows = {}
        self._sheet_deferred_texts = {}
        self._sheet_deferred_tables = {}
        self._sheet_emitted_texts = {}
        self._sheet_emitted_rows = {}
        self._sheet_emitted_table_titles = {}
        self._emitted_images = set()
        self._embedded_image_cid_by_name = {}
        self._in_canonical_emit = False
        self._global_iso_preserved_ids = set()
        self._image_shape_ids = {}
        self._last_iso_preserved_ids = set()

    def _clear_sheet_state(self, sheet_name: str):
        """Clear state for a specific sheet."""
        for dict_attr in ['_cell_to_md_index', '_sheet_shape_images', '_sheet_shape_next_idx',
                          '_sheet_shape_image_start_rows', '_sheet_deferred_texts',
                          '_sheet_deferred_tables', '_sheet_emitted_texts', '_sheet_emitted_rows',
                          '_sheet_emitted_table_titles', '_embedded_image_cid_by_name']:
            getattr(self, dict_attr, {}).pop(sheet_name, None)
        
        self._sheet_shapes_generated.discard(sheet_name)
        self._global_iso_preserved_ids.clear()
        self._last_iso_preserved_ids.clear()

    def _is_canonical_emit(self) -> bool:
        """Check if currently in canonical emission mode."""
        return getattr(self, '_in_canonical_emit', False)

    def _safe_get_cell_value(self, sheet, row: int, col: int) -> Any:
        """Safely get cell value, return None if error."""
        try:
            return sheet.cell(row, col).value
        except Exception:
            return None

    def convert(self) -> str:
        """トップレベルの変換処理 (軽量ラッパ)

        既存のコードベースには複数の補助メソッドがあるため、ここでは
        最小限のフローを提供して CLI から呼べるようにします。
        - ドキュメント見出しの追加
        - 目次生成 (存在すれば呼ぶ)
        - 各シートを順に変換
        - Markdown ファイルを書き出してパスを返す
        """
        print(f"[INFO] Excel文書変換開始: {self.excel_file}")

        # ドキュメントタイトルを先頭に追加
        self.markdown_lines.append(f"# {self.base_name}")
        self.markdown_lines.append("")

        # ヘルパーが存在する場合は目次を生成
        if hasattr(self, '_generate_toc') and callable(getattr(self, '_generate_toc')):
            try:
                self._generate_toc()
            except Exception as e:
                print(f"[WARNING] 目次生成失敗: {e}")

        # シートを変換
        for sheet_name in self.workbook.sheetnames:
            try:
                sheet = self.workbook[sheet_name]
                
                # 非表示シートをスキップ
                if sheet.sheet_state == 'hidden':
                    print(f"[INFO] シートをスキップ（非表示）: {sheet_name}")
                    continue
                
                print(f"[INFO] シート変換中: {sheet_name}")
                self._convert_sheet(sheet)
            except Exception as e:
                print(f"[WARNING] シート処理中にエラーが発生しました: {sheet_name} -> {e}")
                import traceback
                traceback.print_exc()
                continue

        # Markdown出力を書き込み
        output_file = os.path.join(self.output_dir, f"{self.base_name}.md")
        content = "\n".join(str(x) for x in self.markdown_lines)
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"[SUCCESS] 変換完了: {output_file}")
        return output_file

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
        
        # Detect merged cells
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
        # デッドコード削除: self.image_counter = 0
        # マッピング: sheet.title -> 辞書(行番号 -> その行の出力後のmarkdown行インデックス)
        # 行/領域のmarkdownを出力する際にこれを設定し、描画を
        # セル(行)に固定された描画を対応する
        # 段落/テーブル出力の直後に挿入できるようにします。
        self._cell_to_md_index = {}
        # マッピング: sheet.title -> 生成された図形画像ファイル名のリスト(images_dir内)
        self._sheet_shape_images = {}
        # マッピング: sheet.title -> 次の挿入インデックス
        self._sheet_shape_next_idx = {}
        # 図形が生成されたシートタイトルの集合
        self._sheet_shapes_generated = set()
        # 過去のコードは永続化されたstart_mapを使用して画像の
        # 挿入位置を実行間で記憶していました。この動作により別々のグループ
        # 画像が一部の実行で単一の挿入バケットに集約されていました。
        # このような永続化されたマップはデフォルトで無効にし、新たに
        # 計算された代表的なstart_row値が正式なものとなるようにします。
        self._sheet_shape_image_start_rows = {}
        # 初期ヘッダースキャン中に収集された延期された自由形式テキスト
        # sheet.title -> (行, テキスト)のリスト
        self._sheet_deferred_texts = {}
        # プレスキャン中に収集された延期テーブル: sheet.title -> (アンカー行, テーブルデータ, ソース行)のリスト
        self._sheet_deferred_tables = {}
        # 重複する自由形式テキストを避けるためシート毎の出力済みテキストコンテンツ(正規化済み)を追跡
        self._sheet_emitted_texts = {}
        # 行の再出力を避けるためシート毎の出力済み行番号を追跡
        self._sheet_emitted_rows = {}
        # 重複画像挿入を避けるため出力済み画像ファイル名(ベース名)を追跡
        self._emitted_images = set()
        # マッピング: sheet.title -> { 画像ベース名: cNvPr_id }
        # 描画XMLを解析する際に設定され、どの埋め込み
        # 画像がどの描画アンカーID(cNvPr)に対応するかを判別できます。これにより
        # クラスタ化/グループレンダリングが既に
        # 同じcNvPr IDを含む画像を生成している場合、埋め込み画像を抑制できます。
        self._embedded_image_cid_by_name = {}
        # (削除済み) シート毎の特殊ケースフラグなし（特定のセルテキストによる処理制御を行わない）

        # Excelファイルを読み込み
        try:
            self.workbook = load_workbook(excel_file_path, data_only=True)
            print(f"[INFO] Excelワークブック読み込み完了: {excel_file_path}")
        except Exception as e:
            print(f"[ERROR] Excelファイル読み込み失敗: {e}")
            sys.exit(1)

    def _mark_image_emitted(self, img_name: str):
        """Mark an image as emitted only during the canonical emission pass."""
        if self._is_canonical_emit():
            self._emitted_images.add(str(img_name))
        else:
            debug_print(f"[TRACE] Skipping _emitted_images.add({img_name}) in non-canonical pass")

    def _mark_sheet_map(self, sheet_title: str, src_row: int, md_index: int):
        """Record a source-row -> markdown index mapping only during canonical emission."""
        if self._is_canonical_emit():
            self._cell_to_md_index.setdefault(sheet_title, {})[src_row] = int(md_index)
        else:
            debug_print(f"[TRACE] Skipping authoritative sheet_map[{sheet_title}][{src_row}] assignment in non-canonical pass")

    def _mark_emitted_row(self, sheet_title: str, row: int):
        """Mark a row as emitted only during canonical emission."""
        if self._is_canonical_emit():
            self._sheet_emitted_rows.setdefault(sheet_title, set()).add(int(row))
        else:
            debug_print(f"[TRACE] Skipping emitted_rows.add({sheet_title},{row}) in non-canonical pass")

    def _mark_emitted_text(self, sheet_title: str, norm_text: str):
        """Record a normalized emitted text only during canonical emission."""
        if self._is_canonical_emit():
            self._sheet_emitted_texts.setdefault(sheet_title, set()).add(str(norm_text))
        else:
            debug_print(f"[TRACE] Skipping emitted_texts.add({sheet_title},...) in non-canonical pass")
        

    def _escape_angle_brackets(self, text: str) -> str:
        """表示用に角括弧をエスケープして、MarkdownでHTMLタグと解釈されないようにする。

        例: '<Tag>' -> '&lt;Tag&gt;'
        """
        try:
            if text is None:
                return ''
            t = str(text)
            # 安全な表示のためリテラル山括弧をHTMLエンティティに置換
            t = t.replace('<', '&lt;').replace('>', '&gt;')
            return t
        except (ValueError, TypeError):
            return str(text)

    def _normalize_text(self, text: str) -> str:
        """Normalize text for duplicate-detection: collapse whitespace and strip."""
        try:
            if text is None:
                return ''
            import re
            s = str(text)
            s = s.strip()
            s = re.sub(r'\s+', ' ', s)
            return s
        except (ValueError, TypeError):
            return str(text).strip()

    def _add_separator(self):
        """Insert a blank, a Markdown thematic break '---', and a blank.

        Returns True if inserted, False if skipped due to dedupe or non-canonical mode.
        """
        if not self._is_canonical_emit():
            return False

        # 重複セパレータの出力を避けるため最後の数行をチェック
        tail = [x for x in self.markdown_lines[-6:] if isinstance(x, str)]
        for t in reversed(tail):
            if t.strip() == '':
                continue
            if t.strip() == '---':
                debug_print("[DEBUG][_add_separator] skipping duplicate separator '---'")
                return False
            break

        self.markdown_lines.append("")
        self.markdown_lines.append('---')
        self.markdown_lines.append("")
        return True

    def _emit_free_text(self, sheet, src_row: Optional[int], text: str):
        """Append a free-form text line if its normalized form hasn't been emitted for this sheet.

        - sheet: worksheet object
        - src_row: source row number (or None)
        - text: raw text to emit

        Returns True if emitted, False if skipped as duplicate.
        """
        try:
            if text is None:
                return False
            norm = self._normalize_text(text)
            # ここで正式なemitted_textsエントリを作成しない; getを使用して
            # 正規エミッタの外で正式なストアを変更することを避けます。
            emitted_texts = self._sheet_emitted_texts.get(sheet.title, set())
            if norm in emitted_texts:
                return False

            if self._is_canonical_emit():
                # 正規の出力: markdownバッファに追加
                self.markdown_lines.append(self._escape_angle_brackets(text) + "  ")
                
                # ソース行をmarkdownインデックスにマップ
                if src_row is not None:
                    md_index = len(self.markdown_lines) - 1
                    self._mark_sheet_map(sheet.title, src_row, md_index)
                    debug_print(f"[DEBUG][_text_emit] sheet={sheet.title} src_row={src_row} md_index={md_index} text_norm='{norm}'")
                
                # 出力済みとしてマーク
                if src_row is not None:
                    self._mark_emitted_row(sheet.title, src_row)
                self._mark_emitted_text(sheet.title, norm)
                return True
            else:
                # 後の正規パスのため出力を延期
                lst = self._sheet_deferred_texts.setdefault(sheet.title, [])
                
                # 重複する延期テキストをチェック
                already_deferred = any(
                    dt is not None and self._normalize_text(dt) == norm
                    for _, dt in lst
                )
                
                if not already_deferred:
                    lst.append((src_row, text))
                return True
        except Exception as e:
            print(f"[ERROR] _emit_free_text failed: {e}")
            return False
    
    def _insert_markdown_image(self, insert_at: Optional[int], md_line: str, img_name: str, sheet=None):
        """Insert or append an image markdown line and immediately mark it as emitted.

        Returns the new insert index (one past the inserted block) when inserted,
        or the current length of markdown_lines when appended.
        """
        try:
            import traceback
            stk = traceback.extract_stack()
            caller = stk[-3] if len(stk) >= 3 else None
            caller_info = f"{caller.filename}:{caller.lineno}:{caller.name}" if caller else 'unknown'
            debug_print(f"[DEBUG][_insert_markdown_image_called] insert_at={insert_at} img_name={img_name} caller={caller_info}")
            # 正規の出力パス中でなく、即座の画像
            # 挿入が明示的に許可されていない場合、このリクエストを
            # 延期登録に変換し、正規エミッタが配置を制御するようにします。
            if not getattr(self, '_in_canonical_emit', False) and not getattr(self, '_allow_immediate_image_inserts', False):
                try:
                    # markdownのaltテキストからシートタイトルを推測: '![<title>](images/...)'
                    import re
                    m = re.search(r'!\[(.*?)\]', md_line or "")
                    sheet_title = None
                    if m:
                        sheet_title = m.group(1)
                        # 末尾の'の図'が存在する場合は削除（一般的なaltテキストパターン）
                        if sheet_title.endswith('の図'):
                            sheet_title = sheet_title[:-2]
                except Exception:
                    sheet_title = None
                key = sheet_title if sheet_title is not None else 'unknown'
                # sheet_shape_imagesは延期された非正式なコレクションで
                # ここで安全に変更できます。
                lst = self._sheet_shape_images.setdefault(key, [])
                # 不明な場合は安全なデフォルトとして代表的な行=1を使用します。
                # 同じシートに対して同じ画像を複数回登録することを避けます。
                already = any((isinstance(it, (list, tuple)) and len(it) >= 2 and it[1] == img_name) or (str(it) == img_name) for it in lst)
                if not already:
                    lst.append((1, img_name))
                    debug_print(f"[DEBUG][_insert_markdown_image_deferred] img_name={img_name} sheet={key}")

                # 変更は実行されません; 挿入を期待する呼び出し側は
                # 追加されたかのように現在のmarkdown長を受け取ります。
                return len(self.markdown_lines)

            if insert_at is None:
                self.markdown_lines.append(md_line)
                self.markdown_lines.append("")
                
                if sheet is not None:
                    try:
                        filter_ids = self._image_shape_ids.get(img_name)
                        shapes_metadata = self._extract_all_shapes_metadata(sheet, filter_ids=filter_ids)
                        if shapes_metadata:
                            text_metadata = self._format_shape_metadata_as_text(shapes_metadata)
                            if text_metadata:
                                self.markdown_lines.append("")
                                for line in text_metadata.split('\n'):
                                    self.markdown_lines.append(line)
                                self.markdown_lines.append("")
                            
                            json_metadata = self._format_shape_metadata_as_json(shapes_metadata)
                            if json_metadata and json_metadata != "{}":
                                self.markdown_lines.append("<details>")
                                self.markdown_lines.append("<summary>JSON形式の図形情報</summary>")
                                self.markdown_lines.append("")
                                self.markdown_lines.append("```json")
                                for line in json_metadata.split('\n'):
                                    self.markdown_lines.append(line)
                                self.markdown_lines.append("```")
                                self.markdown_lines.append("")
                                self.markdown_lines.append("</details>")
                                self.markdown_lines.append("")
                    except Exception as e:
                        print(f"[WARNING] Failed to add shape metadata: {e}")
                
                self._mark_image_emitted(img_name)
                return len(self.markdown_lines)

            # insert_atをクランプ
            try:
                if insert_at < 0:
                    insert_at = 0
            except Exception:
                insert_at = 0
            if insert_at > len(self.markdown_lines):
                insert_at = len(self.markdown_lines)

            # 複数挿入の相対順序を保持するため空行とmd行を挿入
            self.markdown_lines.insert(insert_at, "")
            self.markdown_lines.insert(insert_at, md_line)
            
            lines_added = 2
            
            if sheet is not None:
                try:
                    filter_ids = self._image_shape_ids.get(img_name)
                    shapes_metadata = self._extract_all_shapes_metadata(sheet, filter_ids=filter_ids)
                    if shapes_metadata:
                        text_metadata = self._format_shape_metadata_as_text(shapes_metadata)
                        if text_metadata:
                            self.markdown_lines.insert(insert_at + lines_added, "")
                            lines_added += 1
                            for line in text_metadata.split('\n'):
                                self.markdown_lines.insert(insert_at + lines_added, line)
                                lines_added += 1
                            self.markdown_lines.insert(insert_at + lines_added, "")
                            lines_added += 1
                        
                        json_metadata = self._format_shape_metadata_as_json(shapes_metadata)
                        if json_metadata and json_metadata != "{}":
                            self.markdown_lines.insert(insert_at + lines_added, "<details>")
                            lines_added += 1
                            self.markdown_lines.insert(insert_at + lines_added, "<summary>JSON形式の図形情報</summary>")
                            lines_added += 1
                            self.markdown_lines.insert(insert_at + lines_added, "")
                            lines_added += 1
                            self.markdown_lines.insert(insert_at + lines_added, "```json")
                            lines_added += 1
                            for line in json_metadata.split('\n'):
                                self.markdown_lines.insert(insert_at + lines_added, line)
                                lines_added += 1
                            self.markdown_lines.insert(insert_at + lines_added, "```")
                            lines_added += 1
                            self.markdown_lines.insert(insert_at + lines_added, "")
                            lines_added += 1
                            self.markdown_lines.insert(insert_at + lines_added, "</details>")
                            lines_added += 1
                            self.markdown_lines.insert(insert_at + lines_added, "")
                            lines_added += 1
                except Exception as e:
                    print(f"[WARNING] Failed to add shape metadata: {e}")
            
            self._mark_image_emitted(img_name)
            return insert_at + lines_added
        except Exception:
            # フォールバック: 追加
            self.markdown_lines.append(md_line)
            self.markdown_lines.append("")
            self._mark_image_emitted(img_name)
            return len(self.markdown_lines)
    
    def _set_excel_fit_to_one_page(self, xlsx_path: str) -> bool:
        """ExcelファイルのpageSetupを縦横1ページに設定
        
        Args:
            xlsx_path: 対象のExcelファイルパス
            
        Returns:
            設定成功時True、失敗時False
        """
        try:
            import zipfile
            import tempfile
            import shutil
            import xml.etree.ElementTree as ET
            
            # 一時ディレクトリに解凍
            tmpdir = tempfile.mkdtemp(prefix='xls2md_fitpage_')
            try:
                with zipfile.ZipFile(xlsx_path, 'r') as zin:
                    zin.extractall(tmpdir)
                
                # 全シートのpageSetupを設定
                xl_worksheets = os.path.join(tmpdir, 'xl', 'worksheets')
                if os.path.exists(xl_worksheets):
                    for fname in os.listdir(xl_worksheets):
                        if fname.endswith('.xml') and fname.startswith('sheet'):
                            sheet_path = os.path.join(xl_worksheets, fname)
                            try:
                                tree = ET.parse(sheet_path)
                                root = tree.getroot()
                                ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
                                
                                # 既存のpageSetupを削除
                                for ps in root.findall('.//{%s}pageSetup' % ns):
                                    root.remove(ps)
                                
                                # 新しいpageSetupを追加（scaleで縮小）
                                # LibreOfficeはfitToWidth/fitToHeightを無視するため、scaleを使用
                                ps = ET.Element('{%s}pageSetup' % ns)
                                # 8ページ→1ページなので大幅縮小が必要
                                # まずは25%で試行（後で調整可能）
                                ps.set('scale', '25')
                                ps.set('orientation', 'landscape')  # 横向きで大きな図形に対応
                                ps.set('paperSize', '9')  # A4サイズ
                                ps.set('useFirstPageNumber', '1')
                                root.append(ps)
                                
                                # pageMargins を調整（余白を最小化）
                                for pm in root.findall('.//{%s}pageMargins' % ns):
                                    root.remove(pm)
                                pm = ET.Element('{%s}pageMargins' % ns)
                                pm.set('left', '0.25')
                                pm.set('right', '0.25')
                                pm.set('top', '0.25')
                                pm.set('bottom', '0.25')
                                pm.set('header', '0.0')
                                pm.set('footer', '0.0')
                                root.append(pm)
                                
                                # ファイルに書き戻し
                                tree.write(sheet_path, encoding='utf-8', xml_declaration=True)
                                debug_print(f"[DEBUG] {fname} のpageSetupを縦横1ページに設定")
                            except Exception as e:
                                print(f"[WARNING] {fname} のpageSetup設定に失敗: {e}")
                
                # 変更を元のファイルに上書き保存
                with zipfile.ZipFile(xlsx_path, 'w', zipfile.ZIP_DEFLATED) as zout:
                    for root_dir, dirs, files in os.walk(tmpdir):
                        for file in files:
                            file_path = os.path.join(root_dir, file)
                            arcname = os.path.relpath(file_path, tmpdir)
                            zout.write(file_path, arcname)
                
                return True
            finally:
                try:
                    shutil.rmtree(tmpdir)
                except Exception:
                    pass  # 一時ファイルの削除失敗は無視
        except Exception as e:
            print(f"[ERROR] pageSetup設定に失敗: {e}")
            return False
    
    def _convert_sheet(self, sheet):
        """シートを変換"""
        sheet_name = sheet.title
        
        # 前のシート毎の状態をクリアしてデフォルトを初期化
        self._clear_sheet_state(sheet_name)
        
        # このシートのデフォルトを初期化
        self._cell_to_md_index.setdefault(sheet_name, {})
        self._sheet_shape_images.setdefault(sheet_name, [])
        self._sheet_shape_next_idx.setdefault(sheet_name, 0)
        self._sheet_deferred_texts.setdefault(sheet_name, [])
        self._sheet_deferred_tables.setdefault(sheet_name, [])
        self._sheet_emitted_texts.setdefault(sheet_name, set())
        self._sheet_emitted_rows.setdefault(sheet_name, set())
        self._embedded_image_cid_by_name.setdefault(sheet_name, {})
        # 描画アンカーcNvPr IDの軽量マッピングを構築（順序付き）することで
        # クラスタループが候補クラスタが以前の分離レンダリングで
        # 既に保持されたアンカーを含むかどうかを迅速に判定できます。
        anchors_cid_list = []
        try:
            try:
                ztmp = zipfile.ZipFile(self.excel_file)
                sheet_index_tmp = self.workbook.sheetnames.index(sheet.title)
                rels_path_tmp = f"xl/worksheets/_rels/sheet{sheet_index_tmp+1}.xml.rels"
                if rels_path_tmp in ztmp.namelist():
                    rels_xml_tmp = ET.fromstring(ztmp.read(rels_path_tmp))
                    drawing_target_tmp = None
                    for rel in rels_xml_tmp.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                        if rel.attrib.get('Type','').endswith('/drawing'):
                            drawing_target_tmp = rel.attrib.get('Target')
                            break
                    if drawing_target_tmp:
                        drawing_path_tmp = drawing_target_tmp
                        if drawing_path_tmp.startswith('..'):
                            drawing_path_tmp = drawing_path_tmp.replace('../', 'xl/')
                        if drawing_path_tmp.startswith('/'):
                            drawing_path_tmp = drawing_path_tmp.lstrip('/')
                        if drawing_path_tmp not in ztmp.namelist():
                            drawing_path_tmp = drawing_path_tmp.replace('worksheets', 'drawings')
                        if drawing_path_tmp in ztmp.namelist():
                            drawing_xml_tmp = ET.fromstring(ztmp.read(drawing_path_tmp))
                            for node_tmp in drawing_xml_tmp:
                                lname_tmp = node_tmp.tag.split('}')[-1].lower()
                                if lname_tmp in ('twocellanchor', 'onecellanchor'):
                                    cid_tmp = None
                                    for sub_tmp in node_tmp.iter():
                                        if sub_tmp.tag.split('}')[-1].lower() == 'cnvpr':
                                            cid_tmp = sub_tmp.attrib.get('id') or sub_tmp.attrib.get('idx')
                                            break
                                    anchors_cid_list.append(str(cid_tmp) if cid_tmp is not None else None)
            except (ET.ParseError, KeyError, AttributeError):
                anchors_cid_list = []
        except (ValueError, TypeError):
            anchors_cid_list = []
        
        # シート見出し（番号とアンカーIDを削除）
        self.markdown_lines.append(f"## {sheet_name} (Sheet Data)")
        self.markdown_lines.append("")
        # シート先頭の説明文を、データ範囲の手前まで（最大12行）スキャンして
        # 表示順どおりに出力する。これにより「MailBoxより先の処理は」のような
        # 任意の説明文が欠落したり順序が入れ替わる問題を防止する。
        try:
            # 検出されたdata_range開始の直前にある連続した非空行のみを出力します。
            # これによりシート上の上から下への順序を保持しつつ、
            # データの遥か上に表示される可能性のある無関係なヘッダーブロックの
            # 出力を回避します。シートにまだdata_rangeがない場合は、
            # 保守的なヒューリスティックとして最初の12行をスキャンすることにフォールバックします。
            # 現在のmarkdown挿入インデックスに設定され、延期された
            # 画像処理を要求する際に使用されます。以前はinsert_posが
            # 代入前に参照されてUnboundLocalErrorを引き起こしていました。
            insert_pos = len(self.markdown_lines)
            max_head_scan = min(12, sheet.max_row)
            data_range = self._get_data_range(sheet)  # Initialize here to avoid UnboundLocalError
            head_rows = []
            # max_head_scanまで各行をスキャンし、行毎に結合された非空
            # セルテキストを収集します。ここではmarkdownを出力しません --- 出力
            # （「このシートには表示可能なデータがありません」メッセージを含む）
            # は正規の出力パス中に行う必要があります。
            # プレスキャン中の出力により、同じメッセージが全ての非空セルに対して
            # 繰り返し追加されていました; 代わりに、今は行テキストのみを収集し
            # 挿入を延期します。
            for r in range(1, max_head_scan + 1):
                row_texts = []
                for c in range(1, min(20, sheet.max_column) + 1):
                    try:
                        v = sheet.cell(r, c).value
                    except Exception:
                        v = None
                    if v is not None:
                        s = str(v).strip()
                        if s:
                            row_texts.append(s)
                # この行のセル値を結合; 空行にはNoneを保持
                if row_texts:
                    combined = ' '.join(row_texts)
                else:
                    combined = None
                head_rows.append(combined)

            emitted_any = False
            # 収集されたhead_rowsを延期テキストとして登録し、
            # 正規の出力パス中にのみ出力されるようにします
            # （同じメッセージのセル毎の繰り返し出力を防ぎます）。
            for idx_row, combined in enumerate(head_rows, start=1):
                if not combined:
                    continue
                # 隣接する同一の延期テキストの重複を回避
                lst = self._sheet_deferred_texts.setdefault(sheet.title, [])
                if len(lst) > 0 and lst[-1][1].strip() == combined.strip():
                    continue
                lst.append((idx_row, combined))

            if data_range:
                start_row = data_range[0]
                # start_rowの直前の連続した非空行を出力（スキャンされたhead_rows内）
                # 1..min(max_head_scan, start_row-1)の中から候補行を検索
                cand_end = min(max_head_scan, start_row - 1)
                if cand_end >= 1:
                    # cand_endから逆方向に歩いて連続した非空ブロックを検索
                    block = []
                    for rr in range(cand_end, 0, -1):
                        content = head_rows[rr-1]
                        if content is None:
                            # 最初の空白に遭遇したら停止
                            break
                        block.insert(0, (rr, content))
                    # ヘッダーブロックを後でソート出力するため延期バッファに収集
                    for (rnum, combined) in block:
                        if len(self.markdown_lines) > 0 and self.markdown_lines[-1].strip() == combined:
                            continue
                        self._sheet_deferred_texts.setdefault(sheet.title, []).append((rnum, combined))
                        emitted_any = True
            else:
                # まだdata_rangeが検出されていません: 元の保守的な動作にフォールバック
                for r in range(1, max_head_scan + 1):
                    combined = head_rows[r-1]
                    if not combined:
                        continue
                    if len(self.markdown_lines) > 0 and self.markdown_lines[-1].strip() == combined:
                        continue
                    # 出力を延期: 後でソート出力するためヘッダー行を収集
                    self._sheet_deferred_texts.setdefault(sheet.title, []).append((r, combined))
                    emitted_any = True

        except Exception as e:
            print(f"[WARNING] シートヘッダー処理でエラー: {e}")
            # エラー時はdata_rangeを再取得
            data_range = self._get_data_range(sheet)

        # シート内のデータ範囲を確認
        if not data_range:
            # シートにデータが無い場合でも、描画が存在するなら図を出力する
            try:
                insert_pos = len(self.markdown_lines)
                # 描画がある場合、_process_sheet_imagesは挿入を延期し
                # 正規エミッタ（_reorder_sheet_output_by_row_order）が
                # 画像を決定論的に配置できるようにします。延期挿入を要求します。
                self._process_sheet_images(sheet, insert_index=insert_pos, insert_images=False)
                if not self._sheet_has_drawings(sheet):
                    # 複数の
                    # 異なるブランチによる同一メッセージの追加を避けるため正規対応の自由テキストエミッタを使用します。
                    self._emit_free_text(sheet, None, "*このシートには表示可能なデータがありません*")
                    # 正規の出力パス中のみ末尾の空行を追加
                    if getattr(self, '_in_canonical_emit', False):
                        self.markdown_lines.append("")
                else:
                    # 描画が処理されました; セパレータを追加して続行
                    self._add_separator()
                return
            except Exception:
                self._emit_free_text(sheet, None, "*このシートには表示可能なデータがありません*")
                if getattr(self, '_in_canonical_emit', False):
                    self.markdown_lines.append("")
                return
        else:
            # data_range が存在しても、セルの実体（テキスト等）が無い場合がある
            # （罫線や書式のみで範囲が検出されるケース）。その場合は空の表を出力しない。
            try:
                r1, r2, c1, c2 = data_range
                has_content = False
                for rr in range(r1, r2 + 1):
                    for cc in range(c1, c2 + 1):
                        try:
                            v = sheet.cell(row=rr, column=cc).value
                        except Exception:
                            v = None
                        if v is not None and str(v).strip():
                            has_content = True
                            break
                    if has_content:
                        break
                if not has_content:
                    # セル内容が無いため、図のみを挿入して終了する
                    insert_pos = len(self.markdown_lines)
                    # 正規パス中に画像が出力されるよう挿入を延期
                    self._process_sheet_images(sheet, insert_index=insert_pos, insert_images=False)
                    if not self._sheet_has_drawings(sheet):
                        self._emit_free_text(sheet, None, "*このシートには表示可能なデータがありません*")
                        if getattr(self, '_in_canonical_emit', False):
                            self.markdown_lines.append("")
                    else:
                        self._add_separator()
                    return
            except Exception:
                # 何か失敗した場合は従来の処理にフォールバック
                pass
        # まずデータをテーブルとして変換（図は表の下に出力したいので後で処理）
        # 注意: ここではreturnしません — 画像の処理を続け、
        # 延期されたテキストと画像が決定論的に出力されるよう
        # 以下の正規の行順序出力パスを実行します。以前は早期returnが
        # 正規出力をバイパスし、テーブル/段落出力の欠落を引き起こしていました。
        self._convert_sheet_data(sheet, data_range)

        # テーブル出力後、図形を生成し（即座の挿入なし）、その後
        # シートテキストと画像を厳密に行番号の昇順で出力し、
        # MarkdownがExcelの上から下への順序と一致するようにします。これは各グループの
        # self._sheet_shape_imagesに格納された代表的なstart_rowを使用します。
        try:
            # 図形が生成され記録されることを確認; 延期挿入を要求
            insert_pos = len(self.markdown_lines)
            self._process_sheet_images(sheet, insert_index=insert_pos, insert_images=False)

            # 厳密な行順序出力を実行: 全てのテキスト行を収集
            # （既に出力されたものを除く）と全ての画像start_rowsを収集し、その後
            # 1..max_rowの行を歩き、テキストまたは画像が存在する場合に出力します。
            self._reorder_sheet_output_by_row_order(sheet)
        except Exception:
            # 保守的な動作にフォールバック: 保留中の画像を挿入しセパレータを追加
            # フォールバックの場合でも、正規
            # エミッタが配置と重複抑制を制御するよう延期挿入を優先します。
            self._process_sheet_images(sheet, insert_index=len(self.markdown_lines), insert_images=False)
        finally:
            # シート処理後の最終セパレータ
            self._add_separator()

    def _reorder_sheet_output_by_row_order(self, sheet):
        """Emit sheet content (text and deferred images) strictly by source row order.

        - Uses self._sheet_shape_images[sheet.title] which is a list of (start_row, filename)
        - Uses self._emit_free_text to avoid duplicates
        - Updates self._cell_to_md_index so images can anchor to emitted md indices
        """
        try:
            # 計測: reorderルーチンへのエントリをマーク
            debug_print(f"[DEBUG][_reorder_entry] sheet={sheet.title}")
            # 実行毎のディスク上マーカーを削除; デバッグトレースのみを保持。
            debug_print(f"[DEBUG][_reorder_entry_marker] sheet={sheet.title}")
            max_row = sheet.max_row
            # avoid creating the per-sheet emitted rows set here; only the
            # canonical emitter should mutate _sheet_emitted_rows via helpers.
            emitted = self._sheet_emitted_rows.get(sheet.title, set())
            # マッピングを構築: 行 -> 画像ファイル名のリスト（_sheet_shape_imagesから）
            img_map = {}
            pairs = self._sheet_shape_images.get(sheet.title, []) or []
            # pairs may be either list of filenames or list of (row, filename)
            normalized_pairs = []
            for item in pairs:
                if isinstance(item, (list, tuple)) and len(item) >= 2:
                    try:
                        r = int(item[0]) if item[0] is not None else 1
                    except (ValueError, TypeError):
                        r = 1
                    normalized_pairs.append((r, item[1]))
                else:
                    # treat as filename with start_row=1
                    normalized_pairs.append((1, str(item)))
            # normalized_pairsから代表的なstart_rowを直接使用し、
            # 出力が元のExcel行順序に従うようにします。これにより
            # テキスト内容に基づくシート固有のヒューリスティックを回避します。
            for r, fn in normalized_pairs:
                img_map.setdefault(r, []).append(fn)

            # また、現在のsheet_map（行->mdインデックス）を出力して比較できるようにします
            sheet_map = self._cell_to_md_index.get(sheet.title, {})
            debug_print(f"[DEBUG][_img_insertion_debug] sheet={sheet.title} sheet_map={sheet_map}")

            # 注意: eventsログを永続化できるようmarkdownダンプを後に移動
            # 現在のmarkdown状態のデバッグ出力の前に。

            # 正規出力モードに入る: 延期テキストは今
            # _emit_free_textによって実際にmarkdownバッファに追加されるべきです。
            self._in_canonical_emit = True

            # デバッグ: 出力済み行と延期テキストを
            # 正規出力に入った直後にダンプして、以前にマークされたものを確認できるようにします。
            emitted_rows = self._sheet_emitted_rows.get(sheet.title, set()) if hasattr(self, '_sheet_emitted_rows') else set()
            deferred_texts = self._sheet_deferred_texts.get(sheet.title, []) if hasattr(self, '_sheet_deferred_texts') else []
            try:
                debug_print(f"[DEBUG][_canonical_enter] sheet={sheet.title} emitted_rows={sorted(list(emitted_rows))[:50]} deferred_texts_count={len(deferred_texts)}")
            except (ValueError, TypeError):
                debug_print(f"[DEBUG][_canonical_enter] sheet={getattr(sheet, 'title', None)} emitted_rows=<error> deferred_texts_count=<error>")

            # 全ての項目がstart_row==1に集約された場合（保存されたリストがファイル名のみを含む場合に一般的）、
            # 描画セル範囲から代表的なstart_rowsを再計算し、
            # それらの計算された行に順番に画像を再分配することを試みます。これにより
            # _render_sheet_fallbackがstart rowsを永続化しなかった場合により正確な配置が生成されます。
            try:
                all_rows = [r for r, _ in normalized_pairs]
                filenames_only = all(r == 1 for r in all_rows) and len(normalized_pairs) > 0
            except (ValueError, TypeError):
                filenames_only = False
            if filenames_only:
                cell_ranges = self._extract_drawing_cell_ranges(sheet) or []
                if cell_ranges:
                    # 各アンカーインデックスをそのstart_rowにマップ
                    start_rows = [cr[2] for cr in cell_ranges]
                    # start_rowでアンカーインデックスをソート
                    idxs = list(range(len(start_rows)))
                    idxs.sort(key=lambda i: start_rows[i])
                    # 割り当てる画像グループの数
                    nimgs = len(normalized_pairs)
                    # インデックスをnimgsの連続したバケットに分割（カウントで）
                    buckets = [[] for _ in range(nimgs)]
                    for i, idx in enumerate(idxs):
                        buckets[i % nimgs].append(idx)
                    # 各バケットの代表的な行を計算し、順番にファイル名を割り当て
                    new_img_map = {}
                    for bi, bucket in enumerate(buckets):
                        insert_r = 1
                        try:
                            if bucket:
                                vals = [start_rows[i] for i in bucket if isinstance(i, int) and i < len(start_rows)]
                                if vals:
                                    insert_r = int(min(vals))
                        except (ValueError, TypeError):
                            insert_r = 1
                        # インデックスをnimgsで割った余りがこのバケットインデックスと等しいファイル名を割り当て
                        try:
                            for j, (_, fn) in enumerate(normalized_pairs):
                                if j % nimgs == bi:
                                    new_img_map.setdefault(insert_r, []).append(fn)
                        except (ValueError, TypeError) as e:
                            debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
                    img_map = new_img_map
                    # より簡単な
                    # 実行後の検査のため、stdoutとloggerの両方にログ出力（利用可能な場合）。
                    msg = f"[DEBUG][_img_fallback_row] sheet={sheet.title} assigned_images_row={insert_r} images={list(img_map.get(insert_r,[]))}"
                    debug_print(msg)
                    # 正規の
                    # 以下の出力ループが調整された行アンカーを使用するようimg_mapからnormalized_pairsを再構築します。
                    new_normalized = []
                    for rr in sorted(img_map.keys()):
                        for fn in img_map.get(rr, []):
                            new_normalized.append((int(rr), fn))
                    if new_normalized:
                        normalized_pairs = new_normalized

            # 上記で実行された調整（例: フォールバック再アンカー）を
            # 反映するようnormalized_pairsからimg_mapを再構築します。
            try:
                img_map = {}
                for r, fn in normalized_pairs:
                    img_map.setdefault(r, []).append(fn)
            except (ValueError, TypeError) as e:
                debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")

            # 同じ行にiso_group（トリミング/グループ）画像が存在する場合、
            # それを優先し、その行の個別の埋め込み画像を抑制します。
            # これによりグループ化された
            # レンダリングが埋め込み画像を1つの合成PNGにキャプチャした場合、同じビジュアルコンテンツを2回出力することを回避します。
            try:
                for rr, fns in list(img_map.items()):
                    try:
                        has_group = any((('iso_group' in (fn or '')) or ('.fixed' in (fn or '')))
                                        for fn in fns)
                    except Exception:
                        has_group = False
                    if has_group:
                        kept = [fn for fn in fns if (('iso_group' in (fn or '')) or ('.fixed' in (fn or '')))]
                        suppressed = [fn for fn in fns if fn not in kept]
                        if suppressed:
                            debug_print(f"[DEBUG][_img_suppress] sheet={sheet.title} row={rr} suppressed={suppressed} kept={kept}")
                        img_map[rr] = kept
                    
            except (ValueError, TypeError) as e:
                debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")

            # フィルタされた可能性のあるimg_mapからnormalized_pairsを再構築し、
            # 以下の正規出力ループが更新されたセットを使用するようにします。
            try:
                new_normalized = []
                for rr in sorted(img_map.keys()):
                    for fn in img_map.get(rr, []):
                        new_normalized.append((int(rr), fn))
                if new_normalized:
                    normalized_pairs = new_normalized
            except (ValueError, TypeError) as e:
                debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")

            # 各非空ソース行のテキストを収集（既に出力された行はスキップ）
            # これは画像アンカーを決定する前に行う必要があり、新しく検出された
            # ヘッダー/テキスト行（まだself._cell_to_md_indexに存在しない可能性がある）が
            # 近くの画像のアンカーとして使用できるようにします。また、ヘッダースキャン中に
            # 以前に延期されたヘッダー/テキスト行をマージし、それらが
            # 正規のソート済み出力パスでのみ出力されるようにします（早期書き込みを防ぐ）。
            texts_by_row = {}
            try:
                # 以前に収集された延期ヘッダー/テキスト行を取得
                deferred = []
                if hasattr(self, '_sheet_deferred_texts'):
                    try:
                        deferred = self._sheet_deferred_texts.pop(sheet.title, []) or []
                    except Exception:
                        deferred = []
                if deferred:
                    try:
                        # 延期テキストをtexts_by_rowに統合し、出力済みセットを尊重
                        # ここでは行を出力済みとしてマークしないでください; 実際のマーキングは
                        # テキストが正規のmarkdownバッファに正常に書き込まれたときにのみ
                        # 出力ループ中に行われるべきです。ここで
                        # 行をマークすると、正式な出力済みセットが
                        # 早期に設定され、正当なテーブル行の刈り込みにつながっていました。
                        for dr, dtxt in deferred:
                            try:
                                rr = int(dr) if dr is not None else 1
                            except (ValueError, TypeError):
                                rr = 1
                            if rr in emitted:
                                continue
                            if dtxt:
                                texts_by_row[rr] = dtxt
                    except (ValueError, TypeError) as e:
                        debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
            except (ValueError, TypeError) as e:
                debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
            for r in range(1, max_row + 1):
                if r in emitted:
                    continue
                row_texts = []
                for c in range(1, min(60, sheet.max_column) + 1):
                    try:
                        v = sheet.cell(r, c).value
                    except Exception:
                        v = None
                    if v is not None:
                        s = str(v).strip()
                        if s:
                            row_texts.append(s)
                if row_texts:
                    texts_by_row[r] = " ".join(row_texts)
                    # ここでは行を出力済みとしてマークしないでください。実際の正式な
                    # マーキングは正規の出力パス中に行う必要があります
                    # （_emit_free_text内、または画像/テキストが書き込まれたとき）
                    # テーブル行の早期刈り込みを避けるため。

            # よりシンプルな決定論的出力: 統一されたイベントリストを構築
            # テキスト項目（src_row -> コンテンツ）と画像項目（start_row -> ファイル名）の、
            # その後行でソートし順番に出力します。同一行の場合、テキストを
            # 画像の前に出力し、画像がそのアンカーテキストの直後に表示されるようにします。
            try:
                # Build the list of events that will actually be emitted.
                # We avoid mutating the logging list while constructing the
                # emission list to prevent double-maintenance. Instead, build
                # `events_emit` first and then synthesize `events_log` from the
                # existing sheet mappings plus the finalized emit list.
                events_emit = []

                # Add freshly-collected text rows that haven't been emitted yet
                for r, txt in texts_by_row.items():
                    try:
                        events_emit.append((int(r), 0, 'text', txt))
                    except (ValueError, TypeError):
                        events_emit.append((1, 0, 'text', txt))

                # Add image events from the normalized_pairs (order 1)
                for start_row, fn in normalized_pairs:
                    try:
                        r = int(start_row) if start_row is not None else 1
                    except (ValueError, TypeError):
                        r = 1
                    events_emit.append((r, 1, 'image', str(fn)))

                # Add deferred table emissions as events (order 0.5 so they come after text but before images)
                try:
                    deferred_tables = self._sheet_deferred_tables.get(sheet.title, []) if hasattr(self, '_sheet_deferred_tables') else []
                    for entry in deferred_tables:
                        try:
                            # deferred_tables entries may be (anchor, table_data, src_rows)
                            # or (anchor, table_data, src_rows, meta). Normalize both.
                            if isinstance(entry, (list, tuple)) and len(entry) >= 3:
                                anchor_row = entry[0]
                                t_data = entry[1]
                                src_rows = entry[2]
                                meta = entry[3] if len(entry) >= 4 else None
                            else:
                                # unexpected shape: skip
                                continue
                            # use order 0.5 to place tables after text at same row
                            events_emit.append((int(anchor_row) if anchor_row else 1, 0.5, 'table', (t_data, src_rows, meta)))
                        except (ValueError, TypeError):
                            try:
                                events_emit.append((int(anchor_row) if anchor_row else 1, 0.5, 'table', (t_data, src_rows, meta)))
                            except (ValueError, TypeError):
                                # give up on this entry
                                continue
                except (ValueError, TypeError) as e:
                    debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")

                # Remove text events that overlap with deferred table source rows
                try:
                    table_src_rows = set()
                    try:
                        deferred_tables = self._sheet_deferred_tables.get(sheet.title, []) if hasattr(self, '_sheet_deferred_tables') else []
                    except Exception:
                        deferred_tables = []

                    for entry in deferred_tables:
                        try:
                            # Normalize possible entry shapes: (anchor, table_data, src_rows) or
                            # (anchor, table_data, src_rows, meta)
                            anchor_row = None
                            src_rows = None
                            meta = None
                            tdata = None
                            if isinstance(entry, (list, tuple)) and len(entry) >= 3:
                                anchor_row = entry[0]
                                tdata = entry[1]
                                src_rows = entry[2]
                                meta = entry[3] if len(entry) >= 4 else None
                            else:
                                # unexpected shape: try best-effort unpack
                                try:
                                    anchor_row = entry[0]
                                except Exception:
                                    anchor_row = None
                                try:
                                    tdata = entry[1]
                                except Exception:
                                    tdata = None
                                src_rows = None

                            # If explicit src_rows provided, sanitize and add them
                            added_any = False
                            if src_rows:
                                try:
                                    for rr in src_rows:
                                        try:
                                            if rr is None:
                                                continue
                                            table_src_rows.add(int(rr))
                                            added_any = True
                                        except (ValueError, TypeError):
                                            # skip non-int-like entries
                                            continue
                                except (ValueError, TypeError) as e:
                                    debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")

                            # Heuristic: sometimes src_rows can be off-by-one and
                            # the actual table contains one additional row immediately
                            # after the listed src_rows. If the next row (max+1)
                            # contains text (we collected texts_by_row earlier) or
                            # has non-empty cells, conservatively include it.
                            try:
                                if added_any:
                                    try:
                                        mx = max(int(x) for x in src_rows if x is not None)
                                    except (ValueError, TypeError):
                                        mx = None
                                    if mx is not None:
                                        cand = mx + 1
                                        if cand not in table_src_rows:
                                            has_text = False
                                            try:
                                                # texts_by_row was built earlier in this scope
                                                if isinstance(texts_by_row, dict) and texts_by_row.get(cand):
                                                    has_text = True
                                            except Exception:
                                                has_text = False
                                            if not has_text:
                                                # fallback: inspect sheet for any non-empty cell in the first 60 cols
                                                try:
                                                    for cc in range(1, min(60, sheet.max_column) + 1):
                                                        v = sheet.cell(cand, cc).value
                                                        if v is not None and str(v).strip():
                                                            has_text = True
                                                            break
                                                except (ValueError, TypeError):
                                                    has_text = False
                                            if has_text:
                                                try:
                                                    table_src_rows.add(int(cand))
                                                except (ValueError, TypeError) as e:
                                                    debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
                            except (ValueError, TypeError) as e:
                                debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")

                            # Fallback: when no explicit src_rows, but we have an anchor and
                            # table data, conservatively add the anchor..anchor+len(table)-1
                            if (not added_any) and anchor_row and tdata and isinstance(tdata, list) and len(tdata) > 0:
                                start = int(anchor_row)
                                cnt = len(tdata)
                                for rr in range(start, start + cnt):
                                    table_src_rows.add(int(rr))

                            # If this deferred table carries a detected title (meta.title),
                            # also treat the table anchor row as overlapping text so any
                            # free-text on the same row (e.g. the raw title) is suppressed
                            # because the canonical emitter will output the title via
                            # the table metadata path.
                            if meta and isinstance(meta, dict) and meta.get('title') and anchor_row:
                                table_src_rows.add(int(anchor_row))
                        except (ValueError, TypeError):
                            continue

                    if table_src_rows:
                        filtered = []
                        for r, order, kind, payload in events_emit:
                            try:
                                if kind == 'text' and int(r) in table_src_rows:
                                    # Titles are emitted via table metadata; skip any free-text
                                    # on the same source rows to avoid duplicate output.
                                    continue
                            except (ValueError, TypeError):
                                # on error, be conservative and skip the text
                                continue
                            filtered.append((r, order, kind, payload))
                        events_emit = filtered
                except (ValueError, TypeError) as e:
                    debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")

                # Sort deterministically by (row, order) and preserve original relative order
                events_emit.sort(key=lambda e: (e[0], e[1]))
                # Log the finalized, sorted emit list for deterministic tracing.
                # Emit one line per event to make scanning/grabbing rows/kinds easy.
                try:
                    for e in events_emit:
                        try:
                            row = int(e[0])
                        except (ValueError, TypeError):
                            row = e[0]
                        order = e[1]
                        kind = e[2]
                        payload = e[3]
                        # Build a concise payload summary for readability.
                        try:
                            if kind == 'table':
                                tdata, src_rows, meta = payload
                                p_summary = f"rows={len(tdata)} src_rows={src_rows} meta={meta}"
                            elif kind == 'image':
                                p_summary = os.path.basename(str(payload))
                            else:
                                p_summary = str(payload)
                        except (ValueError, TypeError):
                            p_summary = str(payload)
                        log_line = f"row={row} order={order} kind={kind} payload={p_summary}"
                        debug_print(f"[DEBUG][_events_emit_sorted] sheet={sheet.title} {log_line}")
                except (ValueError, TypeError) as e:
                    debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")

                # Now synthesize events_log from the authoritative sheet mapping
                # (previously-emitted lines) followed by the finalized emit list.
                events_log = []
                try:
                    sheet_map_all = self._cell_to_md_index.get(sheet.title, {})
                    for r, md_idx in sheet_map_all.items():
                        try:
                            md_idx_i = int(md_idx)
                        except (ValueError, TypeError):
                            continue
                        try:
                            if 0 <= md_idx_i < len(self.markdown_lines):
                                text_line = (self.markdown_lines[md_idx_i] or "").rstrip()
                                if text_line.endswith("  "):
                                    text_line = text_line[:-2]
                                events_log.append((int(r), 0, 'text', text_line))
                        except (ValueError, TypeError):
                            continue
                except (ValueError, TypeError):
                    pass  # データ構造操作失敗は無視

                # Append a logging representation of each event_emit item so
                # callers can see the full canonical sequence.
                for r, order, kind, payload in events_emit:
                    try:
                        if kind == 'text':
                            events_log.append((int(r), order, 'text', payload))
                        elif kind == 'table':
                            # payload is (table_data, src_rows, meta)
                            try:
                                tdata = payload[0]
                                rows_count = len(tdata) if isinstance(tdata, list) else 0
                            except (ValueError, TypeError):
                                rows_count = 0
                            events_log.append((int(r), order, 'table', f"rows={rows_count}"))
                        else:
                            events_log.append((int(r), order, 'image', os.path.basename(str(payload))))
                    except (ValueError, TypeError):
                        events_log.append((int(r) if r else 1, order, kind, payload))

                # Ensure deterministic ordering for the log as well
                events_log.sort(key=lambda e: (e[0], e[1]))

                # Pre-scan the current in-memory markdown for any already-inserted
                # image references and mark them as emitted so the canonical
                # emission pass does not insert duplicates. This handles cases
                # where an earlier codepath inserted images (insert_images=True)
                # before the deferred, deterministic emission ran.
                try:
                    for ln in list(self.markdown_lines):
                        try:
                            m = re.search(r"!\[.*?\]\(images/([^\)]+)\)", ln or "")
                            if m:
                                imgnm = m.group(1)
                                self._mark_image_emitted(imgnm)
                        except Exception:
                            continue
                except Exception as e:
                    print(f"[WARNING] ファイル操作エラー: {e}")

                # Strong debug: always emit events count so we can detect empty/filled
                debug_print(f"[DEBUG][_events_sorted] sheet={sheet.title} events_count_emit={len(events_emit)}")

                # Before mutating markdown, emit a canonical, deterministic
                # logging snapshot derived directly from events_emit. This avoids
                # maintaining a separate events_log list and ensures the log
                # matches the actual emission sequence.
                try:
                    debug_print(f"[DEBUG][_sorted_events_block] sheet={sheet.title} events_count={len(events_emit)}")

                    # Diagnostic/log-only pass: emit a deterministic, human-readable
                    # snapshot of the events_emit sequence for debugging. THIS PASS
                    # MUST NOT mutate self.markdown_lines or call emission helpers
                    # that create side-effects (files, mappings). The canonical
                    # emission loop below is responsible for authoritative writes.
                    for row, _, kind, payload in events_emit:
                        try:
                            if kind == 'text':
                                print(f"  [LOG] text @{row}: {payload}")
                            elif kind == 'table':
                                # payload may be (table_data, src_rows, meta)
                                tdata = None
                                src_rows = None
                                meta = None
                                if isinstance(payload, (list, tuple)):
                                    tdata = payload[0] if len(payload) >= 1 else None
                                    src_rows = payload[1] if len(payload) >= 2 else None
                                    meta = payload[2] if len(payload) >= 3 else None
                                title = None
                                try:
                                    if isinstance(meta, dict):
                                        title = meta.get('title')
                                except Exception:
                                    title = None
                                if title:
                                    print(f"  [LOG] table @{row} title: {title} rows={len(tdata) if isinstance(tdata, list) else 'N/A'} src_rows={src_rows}")
                                else:
                                    print(f"  [LOG] table @{row} rows={len(tdata) if isinstance(tdata, list) else 'N/A'} src_rows={src_rows}")
                            else:  # image
                                print(f"  [LOG] image @{row}: {payload}")
                        except (ValueError, TypeError):
                            # be robust in diagnostic pass; do not raise
                            print(f"  [LOG] event @{row} kind={kind} (payload unstable)")
                except (ValueError, TypeError) as e:
                    debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")

                # Now emit events in deterministic order and record positions.
                # This mutates self.markdown_lines (canonical output path).
                for row, _, kind, payload in events_emit:
                    if kind == 'text':
                        try:
                            emitted_ok = self._emit_free_text(sheet, row, payload)
                        except (ValueError, TypeError):
                            emitted_ok = False
                        if not emitted_ok:
                            # Best-effort fallback: append escaped/normalized text
                            try:
                                txt = self._escape_angle_brackets(payload) + "  "
                                debug_print(f"[DEBUG][_emit_fallback] row={row} text={txt} >>")
                                self.markdown_lines.append(txt)
                                # Only assign authoritative mappings during canonical pass
                                # Only record authoritative mappings during canonical pass
                                md_idx = len(self.markdown_lines) - 1
                                self._mark_sheet_map(sheet.title, row, md_idx)
                                self._mark_emitted_row(sheet.title, row)
                                self._mark_emitted_text(sheet.title, self._normalize_text(payload))
                                self.markdown_lines.append("")
                            except (ValueError, TypeError) as e:
                                debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
                    elif kind == 'table':
                        try:
                            # payload may be either (table_data, src_rows) or
                            # (table_data, src_rows, meta). Normalize both shapes.
                            table_data = None
                            src_rows = None
                            meta = None
                            try:
                                if isinstance(payload, (list, tuple)):
                                    if len(payload) == 2:
                                        table_data, src_rows = payload
                                    elif len(payload) >= 3:
                                        table_data, src_rows, meta = payload[0], payload[1], payload[2]
                                    else:
                                        # unexpected small shape
                                        table_data = payload[0] if len(payload) >= 1 else None
                                else:
                                    # unexpected: treat whole payload as table_data
                                    table_data = payload
                            except Exception:
                                table_data = payload

                            debug_print(f"[DEBUG][_emit_table] row={row} table_rows={len(table_data) if isinstance(table_data, list) else 0} src_rows={src_rows} meta={meta} >>")

                            # If a title/meta is present, emit it first (canonical pass)
                            try:
                                title = None
                                if isinstance(meta, dict):
                                    title = meta.get('title')
                                if title:
                                    normalized_title = ' '.join(str(title).strip().split())
                                    
                                    if sheet.title not in self._sheet_emitted_table_titles:
                                        self._sheet_emitted_table_titles[sheet.title] = set()
                                    
                                    if normalized_title in self._sheet_emitted_table_titles[sheet.title]:
                                        debug_print(f"[DEBUG] タイトル '{title}' は既に出力済みのため抑制")
                                        should_emit_title = False
                                    else:
                                        should_emit_title = True
                                        try:
                                            if table_data and len(table_data) > 0:
                                                header_row = table_data[0]
                                                if isinstance(header_row, (list, tuple)):
                                                    header_text = ' '.join(str(cell).strip() for cell in header_row if cell)
                                                    normalized_header = ' '.join(header_text.split())
                                                    
                                                    if normalized_title and normalized_header:
                                                        if normalized_title in normalized_header or normalized_header.startswith(normalized_title):
                                                            should_emit_title = False
                                                            debug_print(f"[DEBUG] タイトル '{title}' はヘッダー行と重複しているため出力を抑制")
                                        except Exception as e:
                                            debug_print(f"[DEBUG] タイトル冗長性チェックエラー（無視）: {e}")
                                            should_emit_title = True
                                    
                                    self._sheet_emitted_table_titles[sheet.title].add(normalized_title)
                                    
                                    if should_emit_title:
                                        try:
                                            h = f"{self._escape_angle_brackets(title)}  "
                                            self.markdown_lines.append(h)
                                            md_idx = len(self.markdown_lines) - 1
                                            self._mark_sheet_map(sheet.title, row, md_idx)
                                            self._mark_emitted_text(sheet.title, self._normalize_text(title))
                                            self._mark_emitted_row(sheet.title, row)
                                        except Exception:
                                            self._emit_free_text(sheet, row, title)
                            except Exception as e:
                                pass  # XML解析エラーは無視

                            # In canonical emission: output table and record mappings
                            try:
                                if table_data is not None:
                                    # call output which will use guarded helpers
                                    self._output_markdown_table(table_data, source_rows=src_rows, sheet_title=sheet.title)
                                    self.markdown_lines.append("")
                            except Exception:
                                if table_data is not None:
                                    self._output_markdown_table(table_data)

                            # mark source rows as emitted for pruning logic
                            if src_rows:
                                for rr in src_rows:
                                    self._mark_emitted_row(sheet.title, rr)
                        except (ValueError, TypeError) as e:
                            debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
                    else:  # image
                        img_fn = payload
                        debug_print(f"[DEBUG][_emit_image] row={row} image={img_fn} >> ")
                        ref = f"images/{img_fn}"
                        # Gate duplicate images deterministically
                        if ref in self._emitted_images or img_fn in self._emitted_images:
                            continue
                        md = f"![{sheet.title}](images/{img_fn})"
                        # Insert image with metadata using helper method
                        try:
                            if self.shape_metadata:
                                filter_ids = self._image_shape_ids.get(img_fn)
                                shapes_metadata = self._extract_all_shapes_metadata(sheet, filter_ids=filter_ids)
                            else:
                                shapes_metadata = []
                            
                            if shapes_metadata:
                                debug_print(f"[DEBUG] 図形メタデータ抽出成功: {img_fn} -> {len(shapes_metadata)} shapes")
                                text_metadata = self._format_shape_metadata_as_text(shapes_metadata)
                                json_metadata = self._format_shape_metadata_as_json(shapes_metadata)
                                
                                self.markdown_lines.append(md)
                                self.markdown_lines.append("")
                                
                                if text_metadata:
                                    self.markdown_lines.append("")
                                    for line in text_metadata.split('\n'):
                                        self.markdown_lines.append(line)
                                    self.markdown_lines.append("")
                                
                                if json_metadata and json_metadata != "{}":
                                    self.markdown_lines.append("<details>")
                                    self.markdown_lines.append("<summary>JSON形式の図形情報</summary>")
                                    self.markdown_lines.append("")
                                    self.markdown_lines.append("```json")
                                    for line in json_metadata.split('\n'):
                                        self.markdown_lines.append(line)
                                    self.markdown_lines.append("```")
                                    self.markdown_lines.append("")
                                    self.markdown_lines.append("</details>")
                                    self.markdown_lines.append("")
                            else:
                                self.markdown_lines.append(md)
                                self.markdown_lines.append("")
                        except Exception as e:
                            print(f"[WARNING] 図形メタデータ追加失敗: {e}")
                            self.markdown_lines.append(md)
                            self.markdown_lines.append("")
                        # record authoritative mapping only via helper
                        try:
                            md_idx = len(self.markdown_lines) - 2
                            self._mark_sheet_map(sheet.title, row, md_idx)
                        except (ValueError, TypeError):
                            debug_print(f"WARNING self._mark_sheet_map({sheet.title}, {row}, {md_idx})")
                        try:
                            # Mark emitted_images regardless to prevent duplicates; it is safe
                            # because emitted_images only tracks filenames and does not affect pruning.
                            self._mark_image_emitted(img_fn)
                        except (ValueError, TypeError):
                            debug_print(f"WARNING self._mark_image_emitted({img_fn})")
            except (ValueError, TypeError):
                # If anything goes wrong in the simplified flow, fall back to the
                # original complex insertion path by re-raising and letting the
                # outer exception handler perform a conservative insertion later.
                raise

            # We intentionally skip the subsequent anchor-based/fallback insertion
            # pass because the canonical events emission above already placed
            # images deterministically and recorded them in self._emitted_images.
            # This avoids double-insertions caused by multiple insertion codepaths.
            # Exit canonical emission mode so future calls to _emit_free_text
            # will defer again until the next canonical pass.
            self._in_canonical_emit = False

            # Finally, ensure any images with start_row out of bounds or start_row==1
            # that weren't inserted are appended at end (safety net).
            remaining_imgs = []
            for r, imgs in img_map.items():
                if r > max_row or r < 1:
                    remaining_imgs.extend(imgs)
            for r, imgs in img_map.items():
                if r == 1:
                    for i in imgs:
                        ref = f"images/{i}"
                        if ref not in '\n'.join(self.markdown_lines):
                            remaining_imgs.append(i)
            for img in remaining_imgs:
                ref = f"images/{img}"
                if ref in self._emitted_images or img in self._emitted_images:
                    continue
                md = f"![{sheet.title}](images/{img})"
                self.markdown_lines.append(md)
                self.markdown_lines.append("")
                try:
                    self._mark_sheet_map(sheet.title, 1, len(self.markdown_lines) - 2)
                except Exception:
                    debug_print(f"WARNING: Exception self._mark_sheet_map({sheet.title}, {1}, {len(self.markdown_lines) - 2})")
                try:
                    self._mark_image_emitted(img)
                except (ValueError, TypeError):
                    debug_print(f"WARNING: Exception self._mark_image_emitted({img})")
            # このシートの延期テーブルをクリア（既に出力済み）
            if hasattr(self, '_sheet_deferred_tables') and sheet.title in self._sheet_deferred_tables:
                del self._sheet_deferred_tables[sheet.title]
            # final sorted-events fallback removed: no additional logging here.
        except Exception as _exc:
            # デバッグ: 簡略化フローが失敗した理由を確認するため例外情報を出力
            debug_print(f"[DEBUG][_reorder_exception] sheet={sheet.title} exc={_exc!r}")
            import traceback
            traceback.print_exc()
            # エラー時は、全ての延期画像の即座挿入にフォールバック
            for item in self._sheet_shape_images.get(sheet.title, []) or []:
                fn = item[1] if isinstance(item, (list, tuple)) and len(item) >= 2 else str(item)
                md = f"![{sheet.title}](images/{fn})"
                self.markdown_lines.append(md)
                self.markdown_lines.append("")
    
    def _get_data_range(self, sheet) -> Optional[Tuple[int, int, int, int]]:
        """シート内のデータ範囲を取得 (min_row, max_row, min_col, max_col)"""
        try:
            # 実際にデータが存在する範囲を取得
            if self._is_empty_sheet(sheet):
                return None
                
            # 実際に値のあるセルの範囲を計算
            return self._calculate_data_bounds(sheet)
            
        except Exception as e:
            print(f"[WARNING] データ範囲取得エラー: {e}")
            return None

    def _prune_emitted_rows(self, sheet_title: str, table_data: List[List[str]], source_rows: Optional[List[int]]):
        """既に事前出力された行があれば、table_data と source_rows からその行を除去する。

        戻り値: (pruned_table_data, pruned_source_rows)
        """
        try:
            emitted = self._sheet_emitted_rows.get(sheet_title, set()) if hasattr(self, '_sheet_emitted_rows') else set()
        except (ValueError, TypeError):
            emitted = set()

        # 記録された出力済み行のうち、実際にmarkdown
        # mapping. Some code paths may have added rows to the emitted set
        # conservatively (or erroneously) before a canonical write occurred.
        # Only prune rows that both appear in _sheet_emitted_rows AND have a
        # concrete mapping in _cell_to_md_index (i.e. were written to
        # self.markdown_lines). This avoids removing rows that were merely
        # registered but not yet emitted.
        try:
            sheet_map = self._cell_to_md_index.get(sheet_title, {}) if hasattr(self, '_cell_to_md_index') else {}
        except Exception:
            sheet_map = {}

        try:
            # 両方の構造に存在する行のみを正式なものとして扱う
            authoritative_emitted = set(r for r in emitted if r in sheet_map)
            sample_emitted = sorted(list(authoritative_emitted))[:20]
            debug_print(f"[TRACE][_prune_emitted_rows_entry] sheet={sheet_title} emitted_count_total={len(emitted)} emitted_count_auth={len(authoritative_emitted)} emitted_sample={sample_emitted} source_rows_count={len(source_rows) if source_rows else 0}")
        except (ValueError, TypeError):
            debug_print(f"[TRACE][_prune_emitted_rows_entry] sheet={sheet_title} unable to snapshot emitted set")

        if not authoritative_emitted or not source_rows:
            return table_data, source_rows

        pruned_table = []
        pruned_src = []
        for row, src in zip(table_data, source_rows):
            try:
                # ソース行が実際にmarkdownに出力された場合のみ刈り込み
                # (present in authoritative_emitted). Rows that are only listed in
                # the broader emitted set but lack a markdown mapping will be
                # preserved here.
                if src not in authoritative_emitted:
                    pruned_table.append(row)
                    pruned_src.append(src)
                else:
                    # debug: note that this source row was removed due to prior authoritative emission
                    debug_print(f"[TRACE][_prune_emitted_rows_removed] sheet={sheet_title} removed_src_row={src}")
            except (ValueError, TypeError):
                pruned_table.append(row)
                pruned_src.append(src)

        debug_print(f"[TRACE][_prune_emitted_rows_exit] sheet={sheet_title} in={len(source_rows)} out={len(pruned_src)}")

        return pruned_table, pruned_src
    
    def _is_empty_sheet(self, sheet) -> bool:
        """シートが空かどうかをチェック"""
        return (sheet.max_row == 1 and 
                sheet.max_column == 1 and 
                not sheet.cell(1, 1).value)
    
    def _calculate_data_bounds(self, sheet) -> Optional[Tuple[int, int, int, int]]:
        """データの境界を計算"""
        min_row, max_row = None, None
        min_col, max_col = None, None
        
        for row in range(1, sheet.max_row + 1):
            for col in range(1, sheet.max_column + 1):
                cell = sheet.cell(row, col)
                if cell.value is not None or self._has_cell_formatting(cell):
                    min_row, max_row = self._update_row_bounds(min_row, max_row, row)
                    min_col, max_col = self._update_col_bounds(min_col, max_col, col)
        
        return (min_row, max_row, min_col, max_col) if min_row is not None else None
    
    def _update_row_bounds(self, min_row: Optional[int], _max_row: Optional[int], row: int) -> Tuple[int, int]:
        """行の境界を更新"""
        new_min_row = row if min_row is None else min_row
        new_max_row = row
        return new_min_row, new_max_row
    
    def _update_col_bounds(self, min_col: Optional[int], max_col: Optional[int], col: int) -> Tuple[int, int]:
        """列の境界を更新"""
        new_min_col = col if min_col is None or col < min_col else min_col
        new_max_col = col if max_col is None or col > max_col else max_col
        return new_min_col, new_max_col
    
    def _has_cell_formatting(self, cell) -> bool:
        """セルに特別な書式設定があるかチェック"""
        try:
            # セルの背景色、罫線、フォントスタイルなどをチェック
            if (cell.fill and cell.fill.fgColor and hasattr(cell.fill.fgColor, 'rgb') and 
                cell.fill.fgColor.rgb and cell.fill.fgColor.rgb != 'FFFFFFFF'):
                return True
            
            if cell.border and any([
                cell.border.left.style, cell.border.right.style, 
                cell.border.top.style, cell.border.bottom.style
            ]):
                return True
                
            if cell.font and (cell.font.bold or cell.font.italic):
                return True
                
            return False
        except Exception:
            return False
    
    def _process_sheet_images(self, sheet, insert_index: Optional[int] = None, insert_images: bool = True):
        """シート内の画像を処理"""
        try:
            debug_print(f"[DEBUG][_process_sheet_images_entry] sheet={sheet.title} insert_index={insert_index} insert_images={insert_images}")
            debug_print(f"[DEBUG][_process_sheet_images_entry] sheet={sheet.title} insert_index={insert_index} insert_images={insert_images}")
            # 重複した重い処理を防止: 図形が既に生成されている場合
            # for this sheet earlier in the run, skip processing to avoid
            # repeatedly creating tmp_xlsx and invoking external converters.
            if sheet.title in self._sheet_shapes_generated:
                debug_print(f"[DEBUG][_process_sheet_images] sheet={sheet.title} shapes already generated; skipping repeated processing")
                return False
            images_found = False
            # 描画図形（ベクトル図形、コネクタなど）を確認
            # of whether embedded images were found. This ensures that sheets with
            # only vector shapes (no embedded images) are still processed correctly.
            # Phase 2-D修正: images_found=Trueの時だけでなく、常に描画図形を確認
            if True:  # Always execute isolated-group processing for drawing shapes
                    debug_print(f"[DEBUG] {len(sheet._images)} 個の埋め込み画像が検出されました。描画要素を調査中...")
                    # 埋め込み画像が1つ（またはゼロ）の場合、
                    # use that image directly rather than performing costly
                    # isolated-group clustering and trimmed workbook rendering.
                    # これによりtmp_xlsx/.fixed.xlsxの作成と
                    # external converters when unnecessary (common for simple
                    # sheets like input_files/three_sheet_.xlsx).
                    try:
                        emb_count = len(getattr(sheet, '_images', []) or [])
                        # 埋め込み画像がちょうど1つ存在する場合、その画像を優先
                        # directly and skip heavy isolated-group/fallback rendering.
                        # これはクラスタリングを避けるユーザーのリクエストを尊重
                        # a single embedded graphic is present.
                        if emb_count == 1:
                            debug_print(f"[DEBUG][_process_sheet_images_shortcircuit] sheet={sheet.title} single embedded image detected; using embedded image without clustering")
                            for image in sheet._images:
                                img_name = self._process_excel_image(image, f"{sheet.title} (Image)")
                                if img_name:
                                    start_row = 1
                                    try:
                                        pos = self._get_image_position(image)
                                        if pos and isinstance(pos, dict) and 'row' in pos:
                                            start_row = pos['row']
                                    except Exception:
                                        start_row = 1
                                    self._sheet_shape_images.setdefault(sheet.title, [])
                                    self._sheet_shape_images[sheet.title].append((start_row, img_name))
                            try:
                                self._sheet_shapes_generated.add(sheet.title)
                            except (ValueError, TypeError):
                                pass  # データ構造操作失敗は無視
                            return True
                        # 埋め込み画像がゼロの場合、描画チェックにフォールスルー
                        # and possibly run isolated-group or full-sheet fallback.
                    except (ValueError, TypeError):
                        pass  # データ構造操作失敗は無視
                    try:
                        z = zipfile.ZipFile(self.excel_file)
                        sheet_index = self.workbook.sheetnames.index(sheet.title)
                        rels_path = f"xl/worksheets/_rels/sheet{sheet_index+1}.xml.rels"
                        if rels_path in z.namelist():
                            rels_xml = ET.fromstring(z.read(rels_path))
                            drawing_target = None
                            for rel in rels_xml.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                                t = rel.attrib.get('Type','')
                                if t.endswith('/drawing'):
                                    drawing_target = rel.attrib.get('Target')
                                    break
                            if drawing_target:
                                drawing_path = drawing_target
                                if drawing_path.startswith('..'):
                                    drawing_path = drawing_path.replace('../', 'xl/')
                                if drawing_path.startswith('/'):
                                    drawing_path = drawing_path.lstrip('/')
                                if drawing_path not in z.namelist():
                                    drawing_path = drawing_path.replace('worksheets', 'drawings')
                                if drawing_path in z.namelist():
                                    drawing_xml = ET.fromstring(z.read(drawing_path))
                                    # 簡素化されバランスの取れた解析: アンカーIDを収集
                                    # and count pic/sp anchors. Also attempt to map any
                                    # embedded image filenames to their cNvPr ids. Keep
                                    # errors non-fatal and avoid deep nesting of try/except.
                                    anchors_cid_list = []
                                    total_anchors = 0
                                    pic_anchors = 0
                                    sp_anchors = 0
                                    try:
                                        # collect anchor ids and basic counts
                                        for node in list(drawing_xml):
                                            lname = node.tag.split('}')[-1].lower()
                                            if lname not in ('twocellanchor', 'onecellanchor'):
                                                continue
                                            total_anchors += 1
                                            for sub in node.iter():
                                                t = sub.tag.split('}')[-1].lower()
                                                if t == 'pic':
                                                    pic_anchors += 1
                                                if t == 'sp':
                                                    sp_anchors += 1
                                                if t == 'cnvpr':
                                                    cid_val = sub.attrib.get('id') or sub.attrib.get('idx')
                                                    anchors_cid_list.append(str(cid_val) if cid_val is not None else None)

                                        # attempt to read drawing relationships and map embedded images
                                        self._embedded_image_cid_by_name.setdefault(sheet.title, {})
                                        drawing_rels_path = os.path.dirname(drawing_path) + '/_rels/' + os.path.basename(drawing_path) + '.rels'
                                        if drawing_rels_path in z.namelist():
                                            rels_xml = ET.fromstring(z.read(drawing_rels_path))
                                            rid_to_target = {}
                                            for rel in rels_xml.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                                                rid = rel.attrib.get('Id') or rel.attrib.get('Id')
                                                tgt = rel.attrib.get('Target')
                                                if rid and tgt:
                                                    tgtp = tgt
                                                    if tgtp.startswith('..'):
                                                        tgtp = tgtp.replace('../', 'xl/')
                                                    if tgtp.startswith('/'):
                                                        tgtp = tgtp.lstrip('/')
                                                    rid_to_target[rid] = tgtp

                                            for node_c in list(drawing_xml):
                                                lname_c = node_c.tag.split('}')[-1].lower()
                                                if lname_c not in ('twocellanchor', 'onecellanchor'):
                                                    continue
                                                cid_val = None
                                                for sub_c in node_c.iter():
                                                    if sub_c.tag.split('}')[-1].lower() == 'cnvpr':
                                                        cid_val = sub_c.attrib.get('id') or sub_c.attrib.get('idx')
                                                        break
                                                for sub in node_c.iter():
                                                    if sub.tag.split('}')[-1].lower() == 'blip':
                                                        rid = sub.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed') or sub.attrib.get('embed')
                                                        if rid and rid in rid_to_target:
                                                            target = rid_to_target[rid]
                                                            fname = os.path.basename(target)
                                                            try:
                                                                self._embedded_image_cid_by_name[sheet.title][fname] = str(cid_val) if cid_val is not None else None
                                                            except (ValueError, TypeError) as e:
                                                                debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
                                                            break
                                    except (ValueError, TypeError):
                                        # non-fatal: ensure we have defaults
                                        anchors_cid_list = anchors_cid_list if 'anchors_cid_list' in locals() else []
                                        total_anchors = total_anchors if 'total_anchors' in locals() else 0
                                        pic_anchors = pic_anchors if 'pic_anchors' in locals() else 0
                                        sp_anchors = sp_anchors if 'sp_anchors' in locals() else 0
                                    
                                    debug_print(f"[DEBUG] Sheet '{sheet.title}': total_anchors={total_anchors}, pic_anchors={pic_anchors}, sp_anchors={sp_anchors}, sheet._images={len(sheet._images)}")
                                    debug_print(f"[DEBUG] Sheet '{sheet.title}': total_anchors={total_anchors}, pic_anchors={pic_anchors}, sp_anchors={sp_anchors}, sheet._images={len(sheet._images)}")
                                    
                                    # 埋め込み画像よりアンカーが多く、少なくとも1つの図形がある場合、
                                    # attempt isolated-group rendering to capture vector shapes
                                    debug_print(f"[DEBUG] Checking condition: total_anchors({total_anchors}) > len(sheet._images)({len(sheet._images)}) = {total_anchors > len(sheet._images)} AND sp_anchors({sp_anchors}) > 0 = {sp_anchors > 0}")
                                    if total_anchors > len(sheet._images) and sp_anchors > 0:
                                        debug_print(f"[DEBUG] Condition TRUE - entering isolated group rendering block for sheet '{sheet.title}'")
                                        debug_print(f"[DEBUG] Detected additional drawing shapes (anchors={total_anchors}, pics={pic_anchors}, sps={sp_anchors}) - attempting isolated-group rendering")
                                        try:
                                            # 図形のバウンディングボックスを抽出
                                            shapes = None
                                            try:
                                                debug_print(f"[DEBUG] Calling _extract_drawing_shapes for sheet '{sheet.title}'")
                                                shapes = self._extract_drawing_shapes(sheet)
                                                debug_print(f"[DEBUG] _extract_drawing_shapes returned {len(shapes) if shapes else 0} shapes")
                                            except Exception as shape_ex:
                                                print(f"[WARNING] _extract_drawing_shapes failed: {shape_ex}")
                                                import traceback
                                                traceback.print_exc()
                                            
                                            debug_print(f"[DEBUG] _extract_drawing_shapes returned: {len(shapes) if shapes else 'None'} shapes")
                                            debug_print(f"[DEBUG] Checking shapes: shapes={'Not None' if shapes else 'None'}, len={len(shapes) if shapes else 0}")
                                            if shapes and len(shapes) > 0:
                                                debug_print(f"[DEBUG] Shapes condition TRUE - entering clustering block")
                                                # 適切なクラスタリングロジックを使用して図形をクラスタリング
                                                # 行ベースのギャップ分割のためセル範囲を抽出
                                                try:
                                                    cell_ranges_all = self._extract_drawing_cell_ranges(sheet)
                                                except (ValueError, TypeError):
                                                    cell_ranges_all = []
                                                
                                                # 適切なクラスタリングのため_cluster_shapes_commonを使用
                                                # max_groups=1 means cluster into 1 group if possible (no splitting)
                                                # ただし、このメソッドは大きなギャップがある場合は分割します
                                                debug_print(f"[DEBUG] Calling _cluster_shapes_common with {len(shapes)} shapes")
                                                clusters, debug_info = self._cluster_shapes_common(
                                                    sheet, shapes, cell_ranges=cell_ranges_all, max_groups=1
                                                )
                                                debug_print(f"[DEBUG] _cluster_shapes_common returned {len(clusters)} clusters")
                                                debug_print(f"[DEBUG] clustered into {len(clusters)} groups: sizes={[len(c) for c in clusters]}")
                                                debug_print(f"[DEBUG] clustering debug_info: {debug_info}")
                                                
                                                # 各クラスタを分離グループとしてレンダリング
                                                # Using stable _render_sheet_isolated_group method (not v2)
                                                # v2 is experimental and incomplete (missing connector cosmetic processing)
                                                isolated_produced = False
                                                isolated_images = []  # List of (filename, row) tuples
                                                debug_print(f"[DEBUG] Starting to render {len(clusters)} clusters for sheet '{sheet.title}'")
                                                for idx, cluster in enumerate(clusters):
                                                    if len(cluster) > 0:
                                                        debug_print(f"[DEBUG] Rendering cluster {idx+1}/{len(clusters)} with {len(cluster)} shapes")
                                                        result = self._render_sheet_isolated_group(sheet, cluster)
                                                        debug_print(f"[DEBUG] Cluster {idx+1} rendering result: {result}")
                                                        if result:
                                                            if isinstance(result, tuple) and len(result) == 2:
                                                                img_name, cluster_row = result
                                                            else:
                                                                img_name = result
                                                                cluster_row = 1
                                                            
                                                            isolated_produced = True
                                                            isolated_images.append((cluster_row, img_name))
                                                            print(f"[INFO] シート '{sheet.title}' のクラスタ {idx+1} をisolated groupとして出力: {img_name} (row={cluster_row})")
                                                
                                                if isolated_produced:
                                                    print(f"[INFO] シート '{sheet.title}' の図形をisolated groupとして出力しました")
                                                    debug_print(f"[DEBUG] isolated_images count: {len(isolated_images)}")
                                                    # isolated group画像をMarkdownに追加するため、images_foundをTrueに設定
                                                    images_found = True
                                                    # 各画像を登録（row情報を使用）
                                                    for cluster_row, img_name in isolated_images:
                                                        debug_print(f"[DEBUG] Processing isolated group image: {img_name} at row={cluster_row}")
                                                        try:
                                                            self._mark_image_emitted(img_name)
                                                            debug_print(f"[DEBUG] _mark_image_emitted succeeded for: {img_name}")
                                                        except Exception as e:
                                                            print(f"[WARNING] _mark_image_emitted failed: {e}")
                                                        
                                                        try:
                                                            # _sheet_shape_images に追加（クラスタの最小行を使用）
                                                            if not hasattr(self, '_sheet_shape_images'):
                                                                self._sheet_shape_images = {}
                                                            self._sheet_shape_images.setdefault(sheet.title, [])
                                                            # クラスタの最小行に配置
                                                            self._sheet_shape_images[sheet.title].append((cluster_row, img_name))
                                                            debug_print(f"[DEBUG] isolated group画像を_sheet_shape_imagesに追加: {img_name} at row={cluster_row}")
                                                            
                                                            if hasattr(self, '_last_iso_preserved_ids') and self._last_iso_preserved_ids:
                                                                self._image_shape_ids[img_name] = set(self._last_iso_preserved_ids)
                                                                debug_print(f"[DEBUG] 図形IDマッピングを保存: {img_name} -> {len(self._last_iso_preserved_ids)} shapes")
                                                        except Exception as e:
                                                            print(f"[WARNING] Failed to add to _sheet_shape_images: {e}")
                                                            import traceback
                                                            traceback.print_exc()
                                            else:
                                                isolated_produced = False
                                        except Exception as e:
                                            print(f"[WARNING] isolated-group rendering failed: {e}")
                                            import traceback
                                            traceback.print_exc()
                                            isolated_produced = False
                                        
                                        # end of drawing parsing block
                    except (ValueError, TypeError) as e:
                        debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
            # If no parser-detected images were found, attempt a conservative
            # fallback: render the sheet to PDF via LibreOffice and rasterize the
            # corresponding PDF page to PNG using ImageMagick. This captures vector
            # shapes and drawings that openpyxl doesn't expose as images.
            if hasattr(sheet, '_images') and sheet._images:
                print(f"[INFO] シート '{sheet.title}' 内の画像を処理中...")
                images_found = True
                # 埋め込みメディアからのマッピングを事前に設定（描画relsから）
                # to cNvPr ids so that when we process embedded images below we
                # can decide whether to suppress them if a clustered/group
                # render already preserved the same drawing anchor.
                try:
                    z = zipfile.ZipFile(self.excel_file)
                    sheet_index = self.workbook.sheetnames.index(sheet.title)
                    rels_path = f"xl/worksheets/_rels/sheet{sheet_index+1}.xml.rels"
                    if rels_path in z.namelist():
                        rels_xml = ET.fromstring(z.read(rels_path))
                        drawing_target = None
                        for rel in rels_xml.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                            t = rel.attrib.get('Type','')
                            if t.endswith('/drawing'):
                                drawing_target = rel.attrib.get('Target')
                                break
                        if drawing_target:
                            drawing_path = drawing_target
                            if drawing_path.startswith('..'):
                                drawing_path = drawing_path.replace('../', 'xl/')
                            if drawing_path.startswith('/'):
                                drawing_path = drawing_path.lstrip('/')
                            if drawing_path not in z.namelist():
                                drawing_path = drawing_path.replace('worksheets', 'drawings')
                            if drawing_path in z.namelist():
                                drawing_xml = ET.fromstring(z.read(drawing_path))
                                # ensure map exists
                                self._embedded_image_cid_by_name.setdefault(sheet.title, {})
                                # attempt to read drawing rels if present and map rId -> target
                                drawing_rels_path = os.path.dirname(drawing_path) + '/_rels/' + os.path.basename(drawing_path) + '.rels'
                                if drawing_rels_path in z.namelist():
                                    try:
                                        rels_xml2 = ET.fromstring(z.read(drawing_rels_path))
                                        rid_to_target = {}
                                        for rel2 in rels_xml2.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                                            rid = rel2.attrib.get('Id') or rel2.attrib.get('Id')
                                            tgt = rel2.attrib.get('Target')
                                            if rid and tgt:
                                                tgtp = tgt
                                                if tgtp.startswith('..'):
                                                    tgtp = tgtp.replace('../', 'xl/')
                                                if tgtp.startswith('/'):
                                                    tgtp = tgtp.lstrip('/')
                                                rid_to_target[rid] = tgtp
                                        # iterate anchors and map both media basename and media SHA8 -> cNvPr
                                        import hashlib as _hashlib
                                        for node_c in list(drawing_xml):
                                            lname_c = node_c.tag.split('}')[-1].lower()
                                            if lname_c not in ('twocellanchor', 'onecellanchor'):
                                                continue
                                            cid_val = None
                                            for sub_c in node_c.iter():
                                                if sub_c.tag.split('}')[-1].lower() == 'cnvpr':
                                                    cid_val = sub_c.attrib.get('id') or sub_c.attrib.get('idx')
                                                    break
                                            for sub in node_c.iter():
                                                if sub.tag.split('}')[-1].lower() == 'blip':
                                                    rid = sub.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed') or sub.attrib.get('embed')
                                                    if rid and rid in rid_to_target:
                                                        target = rid_to_target[rid]
                                                        # normalize path
                                                        tgtp = target
                                                        if tgtp.startswith('..'):
                                                            tgtp = tgtp.replace('../', 'xl/')
                                                        if tgtp.startswith('/'):
                                                            tgtp = tgtp.lstrip('/')
                                                        # extract basename
                                                        fname = os.path.basename(tgtp)
                                                        try:
                                                            media_bytes = z.read(tgtp) if tgtp in z.namelist() else None
                                                        except Exception:
                                                            media_bytes = None
                                                        sha8 = None
                                                        if media_bytes:
                                                            try:
                                                                sha8 = _hashlib.sha1(media_bytes).hexdigest()[:8]
                                                            except Exception:
                                                                sha8 = None
                                                        if cid_val is not None:
                                                            try:
                                                                # map by original basename
                                                                self._embedded_image_cid_by_name[sheet.title][fname] = str(cid_val)
                                                            except (ValueError, TypeError):
                                                                pass  # データ構造操作失敗は無視
                                                            try:
                                                                # map by short sha if available
                                                                if sha8:
                                                                    self._embedded_image_cid_by_name[sheet.title][sha8] = str(cid_val)
                                                            except (ValueError, TypeError):
                                                                pass  # データ構造操作失敗は無視
                                                        else:
                                                            try:
                                                                self._embedded_image_cid_by_name[sheet.title][fname] = None
                                                            except (ValueError, TypeError):
                                                                pass  # データ構造操作失敗は無視
                                                            try:
                                                                if sha8:
                                                                    self._embedded_image_cid_by_name[sheet.title][sha8] = None
                                                            except Exception as e:
                                                                print(f"[WARNING] ファイル操作エラー: {e}")
                                    except Exception as e:
                                        print(f"[WARNING] ファイル操作エラー: {e}")
                except Exception as e:
                    print(f"[WARNING] ファイル操作エラー: {e}")
                md_lines = []
                for image in sheet._images:
                    # _process_excel_image now returns the saved image filename (basename)
                    img_name = self._process_excel_image(image, f"{sheet.title} (Image)")
                    if img_name:
                            # この画像の代表的なstart_rowを決定（利用可能な場合）
                            start_row = 1
                            try:
                                pos = None
                                try:
                                    pos = self._get_image_position(image)
                                except Exception:
                                    pos = None
                                if pos and isinstance(pos, dict):
                                    # prefer explicit row if provided
                                    if 'row' in pos and isinstance(pos['row'], int):
                                        start_row = pos['row']
                            except Exception:
                                start_row = 1

                            # 正規出力パス中の場合、即座に挿入し
                            # the image appears inline with emitted text. Otherwise,
                            # defer by registering into self._sheet_shape_images so the
                            # canonical emission will place it deterministically.
                            if getattr(self, '_in_canonical_emit', False):
                                md_line = f"![{sheet.title}の図](images/{img_name})"
                                ref = f"images/{img_name}"
                                # この埋め込み画像が描画アンカーに対応する場合
                                # that has already been preserved by a grouped render,
                                # skip emitting it to avoid duplicate presentation.
                                try:
                                    cid_map = self._embedded_image_cid_by_name.get(sheet.title, {}) if hasattr(self, '_embedded_image_cid_by_name') else {}
                                    mapped_cid = cid_map.get(img_name)
                                    # ファイル名に_<sha8>.extのような短いハッシュサフィックスが含まれる場合、それを抽出してキーとして試行
                                    if mapped_cid is None:
                                        try:
                                            # try extracting trailing 8-hex from filename
                                            import re
                                            m = re.search(r'([0-9a-f]{8})', img_name)
                                            if m:
                                                maybe = m.group(1)
                                                mapped_cid = cid_map.get(maybe)
                                        except Exception as e:
                                            print(f"[WARNING] ファイル操作エラー: {e}")
                                    # まだ不明な場合、ディスク上の既存ファイルから短いshaを計算して試行
                                    if mapped_cid is None:
                                        try:
                                            fp = os.path.join(self.images_dir, img_name)
                                            if os.path.exists(fp):
                                                import hashlib as _hashlib
                                                with open(fp, 'rb') as _f:
                                                    d = _f.read()
                                                sha8 = _hashlib.sha1(d).hexdigest()[:8]
                                                mapped_cid = cid_map.get(sha8)
                                        except (OSError, IOError, FileNotFoundError):
                                            print(f"[WARNING] ファイル操作エラー: {e if 'e' in locals() else '不明'}")
                                    global_iso_preserved_ids = getattr(self, '_global_iso_preserved_ids', set()) or set()
                                    if mapped_cid and str(mapped_cid) in global_iso_preserved_ids:
                                        debug_print(f"[DEBUG][_emit_image_skip] sheet={sheet.title} embedded image {img_name} suppressed (cid={mapped_cid} already preserved)")
                                        continue
                                except (OSError, IOError, FileNotFoundError):
                                    print(f"[WARNING] ファイル操作エラー: {e if 'e' in locals() else '不明'}")
                                if ref in self._emitted_images or img_name in self._emitted_images:
                                    continue
                                try:
                                    new_idx = self._insert_markdown_image(insert_index, md_line, img_name, sheet=sheet)
                                    try:
                                        if insert_index is not None:
                                            insert_index = new_idx
                                    except Exception as e:
                                        print(f"[WARNING] ファイル操作エラー: {e}")
                                except Exception:
                                    try:
                                        self.markdown_lines.append(md_line)
                                        self.markdown_lines.append("")
                                        try:
                                            self._mark_image_emitted(img_name)
                                        except Exception as e:
                                            print(f"[WARNING] ファイル操作エラー: {e}")
                                    except Exception as e:
                                        print(f"[WARNING] ファイル操作エラー: {e}")
                                # 挿入を延期: 正規の行ソート済み出力のため登録
                                try:
                                    # check mapped cNvPr for this embedded image and
                                    # skip deferral if already preserved by a group render
                                    cid_map = self._embedded_image_cid_by_name.get(sheet.title, {}) if hasattr(self, '_embedded_image_cid_by_name') else {}
                                    mapped_cid = cid_map.get(img_name)
                                    global_iso_preserved_ids = getattr(self, '_global_iso_preserved_ids', set()) or set()
                                    if mapped_cid and str(mapped_cid) in global_iso_preserved_ids:
                                        debug_print(f"[DEBUG][_defer_image_skip] sheet={sheet.title} embedded image {img_name} suppressed on defer (cid={mapped_cid} already preserved)")
                                    else:
                                        self._sheet_shape_images.setdefault(sheet.title, [])
                                        self._sheet_shape_images[sheet.title].append((start_row, img_name))
                                except (ValueError, TypeError) as e:
                                    debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
                            else:
                                # 非正規コンテキスト: 画像を登録/延期し
                                # canonical emitter will place it deterministically.
                                try:
                                    cid_map = self._embedded_image_cid_by_name.get(sheet.title, {}) if hasattr(self, '_embedded_image_cid_by_name') else {}
                                    mapped_cid = cid_map.get(img_name)
                                    if mapped_cid is None:
                                        try:
                                            import re
                                            m = re.search(r'([0-9a-f]{8})', img_name)
                                            if m:
                                                maybe = m.group(1)
                                                mapped_cid = cid_map.get(maybe)
                                        except Exception as e:
                                            print(f"[WARNING] ファイル操作エラー: {e}")
                                    if mapped_cid is None:
                                        try:
                                            fp = os.path.join(self.images_dir, img_name)
                                            if os.path.exists(fp):
                                                import hashlib as _hashlib
                                                with open(fp, 'rb') as _f:
                                                    d = _f.read()
                                                sha8 = _hashlib.sha1(d).hexdigest()[:8]
                                                mapped_cid = cid_map.get(sha8)
                                        except (OSError, IOError, FileNotFoundError):
                                            print(f"[WARNING] ファイル操作エラー: {e if 'e' in locals() else '不明'}")
                                    global_iso_preserved_ids = getattr(self, '_global_iso_preserved_ids', set()) or set()
                                    if mapped_cid and str(mapped_cid) in global_iso_preserved_ids:
                                        debug_print(f"[DEBUG][_noncanonical_image_skip] sheet={sheet.title} embedded image {img_name} suppressed (cid={mapped_cid} already preserved)")
                                        continue
                                    
                                    md_line = f"![{sheet.title}の図](images/{img_name})"
                                    new_idx = self._insert_markdown_image(insert_index, md_line, img_name, sheet=sheet)
                                    try:
                                        if insert_index is not None:
                                            insert_index = new_idx
                                    except Exception as e:
                                        print(f"[WARNING] ファイル操作エラー: {e}")
                                except Exception:
                                    # フォールバック: sheet_shape_imagesに直接登録
                                    try:
                                        self._sheet_shape_images.setdefault(sheet.title, [])
                                        self._sheet_shape_images[sheet.title].append((start_row, img_name))
                                    except Exception as e:
                                        print(f"[WARNING] ファイル操作エラー: {e}")

            if not images_found:
                debug_print(f"[DEBUG] イメージが見つかりませんでした。")
                # Avoid rendering sheets that contain only cell text; only fallback
                # when the sheet has drawing elements.
                if not self._sheet_has_drawings(sheet):
                    return False
                # Before launching the heavy PDF->PNG pipeline, try to extract
                # drawing bounding boxes from the drawing XML. If XML parsing
                # returns an empty list it is likely there are no visible shapes
                # to render and we should skip producing a full-page image.
                shapes = None
                try:
                    shapes = self._extract_drawing_shapes(sheet)
                except Exception:
                    shapes = None

                # If extraction succeeded and returned an empty list, skip fallback
                # to avoid inserting a full-sheet raster when no drawable elements
                # are present. If extraction errored (shapes is None) or returned
                # non-empty, proceed with rendering as before.
                if shapes == []:
                    print(f"[INFO] シート '{sheet.title}' に描画要素が見つかりませんでした（XML解析結果）。フォールバックレンダリングをスキップします。")
                    return False

                print(f"[INFO] シート '{sheet.title}' に検出されたラスタ画像がありません。フォールバックレンダリングを試行します...")
                try:
                    # Generate sheet-level shape images (will be saved into images_dir)
                    rendered = self._render_sheet_fallback(sheet, insert_index=insert_index, insert_images=insert_images)
                    if rendered:
                        # mark shapes as generated for this sheet
                        self._sheet_shapes_generated.add(sheet.title)
                        # initialize next index
                        if sheet.title not in self._sheet_shape_next_idx:
                            self._sheet_shape_next_idx[sheet.title] = 0
                        # If shapes were created, insert them at insert_index (table end) in markdown.
                        try:
                            imgs = self._sheet_shape_images.get(sheet.title, [])
                            if imgs:
                                # Prefer to insert shapes based on a row-ordered merge so
                                # text and images appear in the final Markdown in the
                                # same top-to-bottom order as the Excel sheet.
                                imgs_by_row = {}
                                assigned = self._sheet_shape_images.get(sheet.title, []) or []

                                # assigned may be list of pairs or list of filenames (backcompat)
                                normalized = []
                                for item in assigned:
                                    if isinstance(item, (list, tuple)) and len(item) >= 2:
                                        try:
                                            row_key = int(item[0]) if item[0] is not None else 1
                                        except (ValueError, TypeError):
                                            row_key = 1
                                        normalized.append((row_key, item[1]))
                                    else:
                                        # fallback: treat as filename with default row=1
                                        try:
                                            normalized.append((1, str(item)))
                                        except (ValueError, TypeError) as e:
                                            debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")

                                # allow small adjustments: if a representative start_row
                                # is near an existing textual anchor (within SNAP_DIST rows),
                                # snap the image to that anchor to keep group images
                                # adjacent to nearby headers. Prefer earlier rows on ties.
                                SNAP_DIST = getattr(self, '_anchor_snap_distance', 3)
                                # Ensure we have a sheet->md mapping available for snapping
                                sheet_map = self._cell_to_md_index.get(sheet.title, {}) or {}
                                # precompute sorted text rows for snapping
                                try:
                                    text_rows_sorted = sorted(list(sheet_map.keys()))
                                except Exception:
                                    text_rows_sorted = []
                                for r, img in normalized:
                                    adjusted_row = r
                                    try:
                                        if text_rows_sorted:
                                            # find nearest textual row and snap if within SNAP_DIST
                                            nearest = min(text_rows_sorted, key=lambda tr: (abs(tr - r), tr))
                                            if abs(nearest - r) <= SNAP_DIST:
                                                adjusted_row = nearest
                                    except Exception:
                                        pass  # データ構造操作失敗は無視
                                    imgs_by_row.setdefault(adjusted_row, []).append(img)

                                # get existing text->md mapping for this sheet
                                # sheet_map is already defined above; reuse it (or fetch fresh)
                                sheet_map = self._cell_to_md_index.get(sheet.title, {}) or sheet_map

                                # NOTE: legacy code used a persisted start_map (self._sheet_shape_image_start_rows)
                                # to reassign image insertion rows across runs. That logic could collapse
                                # multiple distinct group images into a single insertion bucket. Prefer the
                                # freshly computed representative start_row values stored in normalized
                                # and do NOT consult persisted start_map here.
                                if hasattr(self, '_sheet_shape_image_start_rows') and self._sheet_shape_image_start_rows.get(sheet.title):
                                    # clear any persisted hints for this sheet to avoid overriding
                                    # the computed start_row pairs we just generated.
                                    try:
                                        # log for diagnostics but do not use it
                                        debug_print(f"[DEBUG] Ignoring persisted start_map for sheet={sheet.title}")
                                    except (ValueError, TypeError) as e:
                                        debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")

                                # determine a sensible set of rows to iterate (union of text rows and image rows)
                                rows = sorted(set(list(sheet_map.keys()) + list(imgs_by_row.keys())))

                                # Diagnostic debug: emit the current mapping of source rows -> markdown indices
                                try:
                                    debug_print(f"[DEBUG][_img_insertion_debug] sheet={sheet.title} sheet_map={sheet_map}")
                                    debug_print(f"[DEBUG][_img_insertion_debug] imgs_by_row={imgs_by_row}")
                                    debug_print(f"[DEBUG][_img_insertion_debug] normalized_pairs={normalized}")
                                except (ValueError, TypeError) as e:
                                    debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")

                                # base insertion index when no mapped text exists for a row
                                insert_base = insert_index if insert_index is not None else len(self.markdown_lines)

                                # precompute sorted text rows for nearest-neighbor mapping
                                text_rows_sorted = sorted(list(sheet_map.keys()))

                                # For each image row, choose an insertion point that best reflects
                                # the representative start_row: prefer the first textual row >= start_row
                                # (so the image appears just before that text block). If no such row
                                # exists, prefer the last textual row < start_row (insert after it).
                                # If the sheet has no textual mapping at all, fall back to inserting
                                # sequentially at insert_base in ascending start_row order. This keeps
                                # group images at their representative start_rows without collapsing
                                # distinct groups into the same insertion bucket unless they truly
                                # map to the same textual anchor.
                                # Collect final insertion mapping for verification: md_index -> [filenames]
                                md_index_map = {}
                                for row_num in sorted(imgs_by_row.keys()):
                                    imgs_for_row = imgs_by_row.get(row_num, [])
                                    # determine candidate text anchor
                                    md_pos = None
                                    if row_num in sheet_map:
                                        md_pos = sheet_map.get(row_num)
                                        insert_at = md_pos + 1 if md_pos is not None else insert_base
                                    else:
                                        # Choose the nearest textual anchor to this image start_row.
                                        # Using the nearest anchor avoids binding to a distant
                                        # later block (e.g. row26) when the logical anchor is
                                        # the nearby header (e.g. row3). Prefer earlier rows on ties.
                                        if text_rows_sorted:
                                            try:
                                                nearest = min(text_rows_sorted, key=lambda tr: (abs(tr - row_num), tr))
                                                md_pos = sheet_map.get(nearest)
                                                insert_at = (md_pos + 1) if md_pos is not None else insert_base
                                            except Exception:
                                                insert_at = insert_base
                                        else:
                                            # no textual mapping; append sequentially at insert_base
                                            insert_at = insert_base

                                    # insert_atをクランプ to valid markdown range
                                    if insert_at < 0:
                                        insert_at = 0
                                    if insert_at > len(self.markdown_lines):
                                        insert_at = len(self.markdown_lines)

                                    # Ensure we do not insert images before the textual anchor
                                    # we specifically chose for this group (md_pos). The
                                    # previous global clamp (using the latest anchor index)
                                    # could move images after unrelated later anchors,
                                    # causing images to appear far from their logical
                                    # textual context. Only enforce a minimum relative to
                                    # the anchor we used for this image (if any).
                                    try:
                                        if md_pos is not None:
                                            # md_pos is the markdown index of the chosen anchor
                                            # insert_at should be at least one line after it.
                                            if insert_at <= md_pos:
                                                insert_at = md_pos + 1
                                    except Exception:
                                        # conservative fallback: leave insert_at unchanged
                                        pass

                                    # insert each image for this row, preserving original relative order
                                    for img in imgs_for_row:
                                        if not insert_images:
                                            # if caller requested deferred insertion, just record mapping
                                            md_index_map.setdefault(row_num, []).append(img)
                                            continue
                                        ref = f"images/{img}"
                                        already = any(ref in (ln or '') for ln in self.markdown_lines)
                                        if already:
                                            continue
                                        md = f"![{sheet.title}](images/{img})"
                                        # Use helper to insert and mark emitted
                                        try:
                                            new_at = self._insert_markdown_image(insert_at, md, img, sheet=sheet)
                                            md_index_map.setdefault(insert_at, []).append(img)
                                            insert_at = new_at
                                        except Exception:
                                            # fallback
                                            try:
                                                self.markdown_lines.append(md)
                                                self.markdown_lines.append("")
                                                self._mark_image_emitted(img)
                                            except Exception as e:
                                                print(f"[WARNING] ファイル操作エラー: {e}")

                                    # if we inserted at the global insert_base position, advance it
                                    if (row_num not in sheet_map) and insert_at > insert_base:
                                        insert_base = insert_at

                                    # update sheet_map offsets for subsequent insertions that rely on it
                                    if sheet_map and md_pos is not None:
                                        # Only update existing sheet_map offsets during canonical emission
                                        if getattr(self, '_in_canonical_emit', False):
                                            for k, v in list(sheet_map.items()):
                                                try:
                                                    if v > (md_pos if md_pos is not None else -1):
                                                        # don't update the anchor we just used
                                                        if k != (row_num if row_num in sheet_map else None):
                                                            # update mapping to new index
                                                            self._mark_sheet_map(sheet.title, k, v + 2 * len(imgs_for_row))
                                                except Exception as e:
                                                    pass  # XML解析エラーは無視
                                        else:
                                            debug_print(f"[TRACE] Skipping sheet_map offset updates in non-canonical pass for sheet={sheet.title}")

                                # mark all images used
                                self._sheet_shape_next_idx[sheet.title] = len(imgs)
                                # Log final insertion mapping for this sheet (if any)
                                try:
                                        if md_index_map:
                                            print(f"[INFO][_final_img_map] sheet={sheet.title} insert_mappings={md_index_map}")
                                except (ValueError, TypeError) as e:
                                    debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
                        except (ValueError, TypeError):
                            # fallback to previous simple insertion if anything fails
                            try:
                                if insert_index is not None:
                                    insert_at = insert_index
                                    for item in imgs:
                                        # item may be a filename (str) or a (row, filename) pair
                                        if isinstance(item, (list, tuple)) and len(item) >= 2:
                                            img_fn = str(item[1])
                                        else:
                                            img_fn = str(item)
                                        ref = f"images/{img_fn}"
                                        already = any(ref in (ln or '') for ln in self.markdown_lines)
                                        if already:
                                            continue
                                        md = f"![{sheet.title}](images/{img_fn})"
                                        try:
                                            new_at = self._insert_markdown_image(insert_at, md, img_fn)
                                            try:
                                                insert_at = new_at
                                            except (ValueError, TypeError) as e:
                                                debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
                                        except Exception:
                                            try:
                                                self.markdown_lines.append(md)
                                                self.markdown_lines.append("")
                                                self._mark_image_emitted(img_fn)
                                            except Exception as e:
                                                print(f"[WARNING] ファイル操作エラー: {e}")
                                    # record next idx as number of saved images (filenames)
                                    try:
                                        self._sheet_shape_next_idx[sheet.title] = len(imgs)
                                    except Exception as e:
                                        print(f"[WARNING] ファイル操作エラー: {e}")
                            except Exception:
                                self._sheet_shape_next_idx[sheet.title] = len(imgs)
                    else:
                        print(f"[WARNING] フォールバックレンダリングが実行されませんでした（外部ツール未検出など）。")
                except Exception as e:
                    print(f"[WARNING] フォールバックレンダリング中にエラーが発生しました: {e}")
                    
        except Exception as e:
            print(f"[WARNING] 画像処理エラー: {e}")
            return False
        return True
    
    def _process_excel_image(self, image, sheet_name: str) -> Optional[str]:
        """Excel画像を処理"""
        try:
            self.image_counter += 1
            
            # 画像データを取得
            image_data = None
            # Prefer image._data() when available. However, openpyxl may have
            # created a ZipExtFile which becomes invalid if its parent ZipFile
            # was closed; PIL then raises ValueError: I/O operation on closed file.
            # Detect that and fall back to reading the media bytes directly from
            # the XLSX zip by using image.ref (path or basename) when possible.
            if hasattr(image, '_data') and callable(getattr(image, '_data')):
                try:
                    image_data = image._data()
                    debug_print(f"[DEBUG] image._data() succeeded for image #{self.image_counter} on sheet '{sheet_name}'")
                except ValueError:
                    # Likely a closed ZipExtFile. Fall through to zip-based fallback.
                    image_data = None
                except (ValueError, TypeError):
                    image_data = None

            if image_data is None:
                # Try to load from the workbook ZIP using image.ref if it looks like a path
                try:
                    ref = getattr(image, 'ref', None)
                    if isinstance(ref, bytes):
                        try:
                            ref = ref.decode('utf-8')
                        except Exception:
                            ref = None
                    if isinstance(ref, str) and ref:
                        ref_path = ref.lstrip('/') if ref.startswith('/') else ref
                        try:
                            z = zipfile.ZipFile(self.excel_file, 'r')
                            try:
                                # direct match first
                                if ref_path in z.namelist():
                                    image_data = z.read(ref_path)
                                else:
                                    # try to match by basename
                                    import os as _os
                                    b = _os.path.basename(ref_path)
                                    for nm in z.namelist():
                                        if nm.endswith('/' + b) or nm == b:
                                            image_data = z.read(nm)
                                            break
                            finally:
                                try:
                                    z.close()
                                except Exception as e:
                                    print(f"[WARNING] ファイル操作エラー: {e}")
                        except Exception:
                            image_data = None
                    # refがstrでない場合はimage_dataはNone
                    debug_print(f"[DEBUG] image.ref-based extraction succeeded for image #{self.image_counter} on sheet '{sheet_name}'")
                except (ValueError, TypeError):
                    image_data = None

            if not image_data:
                print("[WARNING] 画像データを取得できませんでした")
                return
            
            # 画像形式を判定
            extension = self._detect_image_format(image_data)
            
            # ファイル名を生成（安全化）
            safe_sheet_name = self._sanitize_filename(sheet_name)
            # Use a deterministic filename based on the image bytes so repeated
            # conversions of the same workbook won't produce new files.
            # Compute a short SHA1 of the image bytes.
            try:
                import hashlib
                h = hashlib.sha1(image_data).hexdigest()[:8]
                image_filename = f"{self.base_name}_{safe_sheet_name}_image_{h}{extension}"
            except Exception:
                # fallback to sheet-level stable name
                image_filename = f"{self.base_name}_{safe_sheet_name}_image{extension}"
            image_path = os.path.join(self.images_dir, image_filename)
            
            # 画像を保存
            # If a file with this content-hash already exists, avoid rewriting it.
            try:
                if os.path.exists(image_path):
                    # Verify existing file content matches; if so, reuse.
                    try:
                        with open(image_path, 'rb') as ef:
                            existing = ef.read()
                        if existing == image_data:
                            # reuse
                            debug_print(f"[DEBUG] 既存の画像ファイルを再利用: {image_filename}")
                        else:
                            # collision unlikely, fallback to unique suffix
                            import time
                            alt = f"_{int(time.time())}"
                            image_filename = f"{self.base_name}_{safe_sheet_name}_image_{h}{alt}{extension}"
                            image_path = os.path.join(self.images_dir, image_filename)
                            with open(image_path, 'wb') as f:
                                f.write(image_data)
                    except (OSError, IOError, FileNotFoundError):
                        with open(image_path, 'wb') as f:
                            f.write(image_data)
                else:
                    with open(image_path, 'wb') as f:
                        f.write(image_data)
            except (OSError, IOError, FileNotFoundError):
                # last-resort write
                with open(image_path, 'wb') as f:
                    f.write(image_data)
            
            # 画像位置情報を取得
            position_info = self._get_image_position(image)
            
            # return the saved image filename (basename). Caller will generate
            # the markdown using this concrete filename so that links always
            # point to an existing file on disk.
            print(f"[SUCCESS] 画像を処理: {image_filename}")
            return os.path.basename(image_filename)
        except Exception as e:
            try:
                import traceback
                tb = traceback.format_exc()
                print(f"[ERROR] Excel画像処理エラー: {e}\n{tb}")
            except (ValueError, TypeError):
                print(f"[ERROR] Excel画像処理エラー: {e}")
            return None

    def _deduplicate_image_files(self):
        """output/images のファイルを内容でグループ化し、同一内容の複数ファイルを1つに集約する。

        - 同一ハッシュのファイル群から最も短い名前（またはアルファベット順先頭）を正規名とする
        - markdown_lines 内の参照を正規名に置換する
        - output/sorted_events.txt 内の image 行を正規名で置換する
        - 重複ファイルは削除する
        """
        try:
            import hashlib
            import collections
            img_dir = self.images_dir
            if not os.path.isdir(img_dir):
                return

            # Debug: log image filenames and their SHA256 before building groups
            try:
                import hashlib as _hashlib
                debug_print('[DEBUG][_dedupe] listing images and computing sha256 before dedupe:')
                for _fn in sorted(os.listdir(img_dir)):
                    _fp = os.path.join(img_dir, _fn)
                    if not os.path.isfile(_fp):
                        continue
                    try:
                        _h = _hashlib.sha256()
                        with open(_fp, 'rb') as _f:
                            for _chunk in iter(lambda: _f.read(8192), b''):
                                _h.update(_chunk)
                        debug_print(f"[DEBUG][_dedupe] pre-sha {_fn} = {_h.hexdigest()}")
                    except Exception as _e:
                        debug_print(f"[DEBUG][_dedupe] pre-sha {_fn} FAILED: {_e}")
            except (OSError, IOError, FileNotFoundError):
                # non-fatal; continue with normal dedupe
                pass

            # build hash -> [files]
            groups = collections.defaultdict(list)
            for fn in os.listdir(img_dir):
                fp = os.path.join(img_dir, fn)
                if not os.path.isfile(fp):
                    continue
                try:
                    with open(fp, 'rb') as f:
                        data = f.read()
                    h = hashlib.sha256(data).hexdigest()
                    groups[h].append((fn, data))
                except (OSError, IOError, FileNotFoundError):
                    continue

            # For each group with >1 file, choose canonical and update references
            for h, items in groups.items():
                if len(items) <= 1:
                    continue
                # choose canonical filename: prefer shortest, then lexicographic
                items_sorted = sorted(items, key=lambda it: (len(it[0]), it[0]))
                canonical = items_sorted[0][0]
                duplicate_names = [it[0] for it in items_sorted[1:]]
                if not duplicate_names:
                    continue

                # Determine whether all files in this hash-group originate from
                # the same workbook (self.base_name). If not, skip dedupe for
                # this group to respect the user's requirement that images
                # from different Excel files be treated as distinct.
                try:
                    bases = set([fn.split('_', 1)[0] if '_' in fn else fn for fn, _ in items_sorted])
                    if len(bases) != 1 or (self.base_name not in bases):
                        debug_print(f"[DEBUG][_dedupe] skipping cross-workbook dedupe for hash {h}: bases={bases}")
                        # Do not remove any files in this group; leave as-is
                        continue
                except (ValueError, TypeError):
                    # If any failure determining origin, be conservative and skip
                    debug_print(f"[DEBUG][_dedupe] skipping dedupe for hash {h} due to error determining origins")
                    continue

                # Update markdown_lines references (only for files belonging to this workbook)
                try:
                    import re
                    new_lines = []
                    for ln in self.markdown_lines:
                        if not isinstance(ln, str):
                            new_lines.append(ln)
                            continue
                        s = ln
                        for dup in duplicate_names:
                            s = re.sub(r"!\[(.*?)\]\(images/" + re.escape(dup) + r"\)", r"![\1](images/" + canonical + r")", s)
                        new_lines.append(s)
                    self.markdown_lines = ExcelToMarkdownConverter._LoggingList(self)
                    self.markdown_lines += new_lines
                except Exception as e:
                    print(f"[WARNING] ファイル操作エラー: {e}")

                # Remove duplicate files (keep canonical)
                for dup in duplicate_names:
                    try:
                        p = os.path.join(img_dir, dup)
                        if os.path.exists(p):
                            os.remove(p)
                            debug_print(f"[DEBUG][_dedupe] removed duplicate image: {dup} -> canonical: {canonical}")
                    except (ValueError, TypeError):
                        pass  # ファイル操作失敗は無視

            # Also rebuild emitted images set to reflect final filenames
            try:
                self._emitted_images = set()
                for ln in self.markdown_lines:
                    try:
                        import re
                        m = re.search(r"!\[.*?\]\(images/([^\)]+)\)", ln or "")
                        if m:
                            self._emitted_images.add(m.group(1))
                    except Exception:
                        continue
            except Exception as e:
                print(f"[WARNING] ファイル操作エラー: {e}")
        except Exception as e:
            print(f"[WARNING] ファイル操作エラー: {e}")

    # ========================================================================
    # Phase 1: 画像・図形処理の共通基盤メソッド群
    # ========================================================================

    def _get_drawing_xml_and_metadata(self, sheet) -> Optional[Dict[str, Any]]:
        """シートのdrawing.xmlとメタデータを取得
        
        ExcelファイルをZIPとして開き、指定シートのdrawing.xmlを取得します。
        3つのレンダリングメソッドで重複していた処理を統合。
        
        Args:
            sheet: 対象シート
        
        Returns:
            {
                'zip': ZipFile,
                'drawing_xml': ET.Element,
                'drawing_path': str,
                'sheet_index': int
            }
            または None (drawing.xmlが存在しない場合)
        
        Note:
            呼び出し側で返されたzip_fileをcloseする責任があります
        """
        try:
            zpath = self.excel_file
            z = zipfile.ZipFile(zpath, 'r')
            sheet_index = self.workbook.sheetnames.index(sheet.title)
            rels_path = f"xl/worksheets/_rels/sheet{sheet_index+1}.xml.rels"
            
            if rels_path not in z.namelist():
                z.close()
                return None
            
            rels_xml = ET.fromstring(z.read(rels_path))
            drawing_target = None
            for rel in rels_xml.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                if rel.attrib.get('Type', '').endswith('/drawing'):
                    drawing_target = rel.attrib.get('Target')
                    break
            
            if not drawing_target:
                z.close()
                return None
            
            # drawing_pathの正規化
            drawing_path = drawing_target
            if drawing_path.startswith('..'):
                drawing_path = drawing_path.replace('../', 'xl/')
            if drawing_path.startswith('/'):
                drawing_path = drawing_path.lstrip('/')
            if drawing_path not in z.namelist():
                drawing_path = drawing_path.replace('worksheets', 'drawings')
                if drawing_path not in z.namelist():
                    z.close()
                    return None
            
            drawing_xml_bytes = z.read(drawing_path)
            drawing_xml = ET.fromstring(drawing_xml_bytes)
            
            return {
                'zip': z,
                'drawing_xml': drawing_xml,
                'drawing_path': drawing_path,
                'sheet_index': sheet_index
            }
        
        except Exception as e:
            print(f"[WARNING] Drawing XML取得失敗: {e}")
            try:
                z.close()
            except:
                pass  # データ構造操作失敗は無視
            return None

    def _parse_theme_colors(self, zip_file: zipfile.ZipFile) -> Tuple[Dict[str, str], Dict[str, Any]]:
        """theme1.xmlからカラースキームとline参照を抽出
        
        ExcelのテーマXMLを解析し、色スキーム(schemeClr -> srgbClr)と
        lnRef(線のスタイル参照)のマッピングを取得します。
        2つのレンダリングメソッドで重複していた処理を統合。
        
        Args:
            zip_file: 開かれたZipFileオブジェクト
        
        Returns:
            (theme_color_map, ln_ref_map)
            - theme_color_map: {color_name: hex_value}
            - ln_ref_map: {index: ln_element}
        """
        theme_color_map = {}
        ln_ref_map = {}
        
        try:
            if 'xl/theme/theme1.xml' not in zip_file.namelist():
                return theme_color_map, ln_ref_map
            
            theme_xml = ET.fromstring(zip_file.read('xl/theme/theme1.xml'))
            a_ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
            
            # カラースキームの抽出
            clr = None
            for node in theme_xml.iter():
                if node.tag.split('}')[-1].lower() == 'clrscheme':
                    clr = node
                    break
            
            if clr is not None:
                for child in list(clr):
                    name = child.tag.split('}')[-1]
                    hexval = None
                    for sub in child.iter():
                        tag_name = sub.tag.split('}')[-1].lower()
                        if tag_name == 'srgbclr':
                            hexval = sub.attrib.get('val')
                            break
                        if tag_name == 'sysclr':
                            hexval = sub.attrib.get('lastclr') or sub.attrib.get('lastClr')
                            break
                    if hexval:
                        theme_color_map[name.lower()] = hexval
            
            # lnStyleLstの抽出
            try:
                ns = {'a': a_ns}
                ln_style_lst = theme_xml.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}lnStyleLst')
                if ln_style_lst is None:
                    ln_style_lst = theme_xml.find('.//a:lnStyleLst', ns)
                
                if ln_style_lst is not None:
                    import copy as _copy
                    for i, ln_el in enumerate(list(ln_style_lst)):
                        try:
                            # 属性も含めてディープコピー
                            ln_ref_map[str(i)] = _copy.deepcopy(ln_el)
                        except (ValueError, TypeError):
                            ln_ref_map[str(i)] = None
            except (ValueError, TypeError):
                pass  # データ構造操作失敗は無視
        
        except Exception as e:
            print(f"[WARNING] テーマカラー解析失敗: {e}")
        
        return theme_color_map, ln_ref_map

    def _resolve_connector_references(
        self,
        drawing_xml: ET.Element,
        anchors: List[ET.Element],
        keep_cnvpr_ids: Set[str]
    ) -> Set[str]:
        """
        Resolve connector references using BFS to determine complete set of anchor IDs to preserve.
        
        Args:
            drawing_xml: Root element of the drawing XML
            anchors: List of filtered anchor elements
            keep_cnvpr_ids: Initial set of cNvPr IDs to keep
        
        Returns:
            Complete set of cNvPr IDs to preserve (including connectors and endpoints)
        """
        from collections import deque
        
        # Build reference mappings
        refs = {}
        reverse_refs = {}
        
        for orig in list(drawing_xml):
            lname = orig.tag.split('}')[-1].lower()
            if lname not in ('twocellanchor', 'onecellanchor'):
                continue
            cid = None
            for sub in orig.iter():
                if sub.tag.split('}')[-1].lower() == 'cnvpr':
                    cid = str(sub.attrib.get('id'))
                    break
            if cid is None:
                continue
            
            # Find referenced IDs
            rset = set()
            for sub in orig.iter():
                st = sub.tag.split('}')[-1].lower()
                if st in ('stcxn', 'endcxn', 'stcxnpr', 'endcxnpr'):
                    vid = sub.attrib.get('id') or sub.attrib.get('idx')
                    if vid is not None:
                        rset.add(str(vid))
            if rset:
                refs[cid] = rset
                for rid in rset:
                    reverse_refs.setdefault(rid, set()).add(cid)
        
        # Build row mappings
        id_to_row = {}
        all_id_to_row = {}
        ns_xdr = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
        
        for an in anchors:
            a_cid = None
            for sub in an.iter():
                if sub.tag.split('}')[-1].lower() == 'cnvpr':
                    a_cid = sub.attrib.get('id') or sub.attrib.get('idx')
                    break
            if a_cid is None:
                continue
            fr = an.find('{%s}from' % ns_xdr)
            if fr is not None:
                r = fr.find('{%s}row' % ns_xdr)
                if r is not None and r.text is not None:
                    try:
                        id_to_row[str(a_cid)] = int(r.text)
                    except (ValueError, TypeError) as e:
                        debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
        
        # Fallback mapping from ALL anchors
        for orig_an in list(drawing_xml):
            lname2 = orig_an.tag.split('}')[-1].lower()
            if lname2 not in ('twocellanchor', 'onecellanchor'):
                continue
            a_cid2 = None
            for sub2 in orig_an.iter():
                if sub2.tag.split('}')[-1].lower() == 'cnvpr':
                    a_cid2 = sub2.attrib.get('id') or sub2.attrib.get('idx')
                    break
            if a_cid2 is None:
                continue
            fr2 = orig_an.find('{%s}from' % ns_xdr)
            if fr2 is not None:
                r2 = fr2.find('{%s}row' % ns_xdr)
                if r2 is not None and r2.text is not None:
                    try:
                        all_id_to_row[str(a_cid2)] = int(r2.text)
                    except (ValueError, TypeError) as e:
                        debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
        
        # Determine group's row span
        group_rows = set()
        for cid in keep_cnvpr_ids:
            rowval = id_to_row.get(str(cid))
            if rowval is not None:
                group_rows.add(int(rowval))
        
        # BFS expansion
        preserve = set(keep_cnvpr_ids)
        q = deque(keep_cnvpr_ids)
        
        while q:
            current = q.popleft()
            for fwd in refs.get(current, set()):
                if fwd not in preserve:
                    preserve.add(fwd)
                    q.append(fwd)
            for rev in reverse_refs.get(current, set()):
                if rev not in preserve:
                    preserve.add(rev)
                    q.append(rev)
        
        # Include connector-only anchors on group rows (±1 tolerance)
        for cid in list(all_id_to_row.keys()):
            scid = str(cid)
            if scid in preserve:
                continue
            rowc = id_to_row.get(scid) or all_id_to_row.get(scid)
            if rowc is not None and group_rows:
                if rowc in group_rows or any(abs(int(rowc) - int(gr)) <= 1 for gr in group_rows):
                    preserve.add(scid)
        
        debug_print(f"[DEBUG][_resolve_connector] keep={len(keep_cnvpr_ids)} → preserve={len(preserve)} (rows={sorted(list(group_rows))})")
        return preserve

    def _prune_drawing_anchors(
        self,
        drawing_relpath: str,
        keep_cnvpr_ids: Set[str],
        referenced_ids: Set[str],
        cell_range: Optional[Tuple[int, int, int, int]],
        group_rows: Set[int]
    ) -> None:
        """
        Prune drawing XML to keep only specified anchors.
        
        Args:
            drawing_relpath: Path to the drawing XML file
            keep_cnvpr_ids: Set of cNvPr IDs to preserve
            referenced_ids: Set of IDs referenced by connectors
            cell_range: Optional cell range (s_col, e_col, s_row, e_row)
            group_rows: Set of row numbers within the group's range
        """
        try:
            def node_contains_referenced_id(n):
                try:
                    for sub in n.iter():
                        lname = sub.tag.split('}')[-1].lower()
                        if lname == 'cnvpr' or lname.endswith('cnvpr'):
                            vid = sub.attrib.get('id') or sub.attrib.get('idx')
                            if vid is not None and str(vid) in referenced_ids:
                                return True
                        if lname in ('stcxn', 'endcxn', 'stcxnpr', 'endcxnpr'):
                            vid = sub.attrib.get('id') or sub.attrib.get('idx')
                            if vid is not None and str(vid) in referenced_ids:
                                return True
                except (ValueError, TypeError) as e:
                    debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
                return False
            
            tree = ET.parse(drawing_relpath)
            root = tree.getroot()
            
            removed_count = 0
            kept_count = 0
            
            for node in list(root):
                lname = node.tag.split('}')[-1].lower()
                if lname in ('twocellanchor', 'onecellanchor'):
                    this_cid = None
                    for sub in node.iter():
                        if sub.tag.split('}')[-1].lower() == 'cnvpr':
                            this_cid = sub.attrib.get('id') or sub.attrib.get('idx')
                            break
                    
                    if this_cid is not None and str(this_cid) in keep_cnvpr_ids:
                        kept_count += 1
                        debug_print(f"[DEBUG][_prune] KEEP anchor id={this_cid}")
                        continue
                    
                    try:
                        if node_contains_referenced_id(node):
                            continue
                    except (ValueError, TypeError) as e:
                        debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
                    
                    try:
                        if (not keep_cnvpr_ids) and group_rows:
                            ns_xdr = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
                            fr = node.find('{%s}from' % ns_xdr)
                            if fr is not None:
                                r = fr.find('{%s}row' % ns_xdr)
                                if r is not None and r.text is not None:
                                    try:
                                        from_row = int(r.text)
                                        if from_row in group_rows or any(abs(from_row - gr) <= 1 for gr in group_rows):
                                            continue
                                    except (ValueError, TypeError) as e:
                                        debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
                    except (ValueError, TypeError) as e:
                        debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
                    
                    try:
                        root.remove(node)
                        removed_count += 1
                        debug_print(f"[DEBUG][_prune] REMOVE anchor id={this_cid}")
                    except (ValueError, TypeError):
                        try:
                            root.remove(node)
                            removed_count += 1
                            debug_print(f"[DEBUG][_prune] REMOVE anchor id={this_cid} (retry)")
                        except (ValueError, TypeError) as e:
                            debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
            
            debug_print(f"[DEBUG][_prune] Summary: kept={kept_count}, removed={removed_count}, total={kept_count+removed_count}")
            tree.write(drawing_relpath, encoding='utf-8', xml_declaration=True)
        except Exception as e:
            debug_print(f"[DEBUG][_prune_drawing_anchors] Error: {e}")

    def _convert_excel_to_pdf(self, xlsx_path: str, tmpdir: str, apply_fit_to_page: bool = True) -> Optional[str]:
        """ExcelファイルをPDFに変換
        
        LibreOfficeを使用してExcelファイルをPDF形式に変換します。
        2つのレンダリングメソッドで重複していた処理を統合。
        
        Args:
            xlsx_path: 変換するExcelファイルのパス
            tmpdir: PDF出力先ディレクトリ
            apply_fit_to_page: 1ページに収める設定を適用するか
        
        Returns:
            生成されたPDFファイルのパス または None
        """
        try:
            # 元のファイルを上書きしないように一時コピーを作成
            tmp_xlsx = os.path.join(tmpdir, os.path.basename(xlsx_path))
            shutil.copyfile(xlsx_path, tmp_xlsx)
            
            # PDF変換前に縦横1ページ設定を適用
            if apply_fit_to_page:
                self._set_excel_fit_to_one_page(tmp_xlsx)
            
            # LibreOfficeでPDF変換
            cmd = [LIBREOFFICE_PATH, '--headless', '--convert-to', 'pdf', '--outdir', tmpdir, tmp_xlsx]
            debug_print(f"[DEBUG] LibreOffice export command: {' '.join(cmd)}")
            proc = subprocess.run(cmd, capture_output=True, text=True, timeout=90)
            
            if proc.returncode != 0:
                print(f"[WARNING] LibreOffice PDF 変換失敗: {proc.stderr}")
                return None
            
            # 生成されたPDFを探す
            pdf_name = f"{self.base_name}.pdf"
            pdf_path = os.path.join(tmpdir, pdf_name)
            
            if not os.path.exists(pdf_path):
                # LibreOfficeが異なる名前で出力した可能性
                candidates = [os.path.join(tmpdir, f) for f in os.listdir(tmpdir) if f.lower().endswith('.pdf')]
                if not candidates:
                    print("[WARNING] LibreOffice がPDFを出力しませんでした")
                    return None
                pdf_path = candidates[0]
            
            return pdf_path
        
        except Exception as e:
            print(f"[WARNING] Excel→PDF変換失敗: {e}")
            return None

    def _convert_pdf_page_to_png(self, pdf_path: str, page_index: int, dpi: int,
                                  output_dir: str, filename_prefix: str) -> Optional[str]:
        """PDFの指定ページをPNG画像に変換（PyMuPDF使用）
        
        PyMuPDFを使用してPDFの特定ページをPNG形式に変換します。
        複数のレンダリングメソッドで重複していた処理を統合。
        
        Args:
            pdf_path: 変換するPDFファイルのパス
            page_index: ページ番号(0始まり)
            dpi: 解像度
            output_dir: PNG出力先ディレクトリ
            filename_prefix: 出力ファイル名のプレフィックス
        
        Returns:
            生成されたPNGファイル名(相対パス) または None
        """
        try:
            png_filename = f"{filename_prefix}.png"
            png_path = os.path.join(output_dir, png_filename)
            
            debug_print(f"[DEBUG] PyMuPDFでPDF→PNG変換実行 (ページ {page_index}, DPI: {dpi})...")
            
            doc = fitz.open(pdf_path)
            if page_index >= len(doc):
                print(f"[WARNING] ページ{page_index}が存在しません（全{len(doc)}ページ）")
                doc.close()
                return None
            
            page = doc[page_index]
            
            mat = fitz.Matrix(dpi/72, dpi/72)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            
            img_data = pix.tobytes("png")
            pix = None
            doc.close()
            
            img = Image.open(io.BytesIO(img_data))
            
            if img.mode == 'RGBA':
                background = Image.new('RGB', img.size, (255, 255, 255))
                background.paste(img, mask=img.split()[3] if len(img.split()) > 3 else None)
                img = background
            elif img.mode != 'RGB':
                img = img.convert('RGB')
            
            img.save(png_path, 'PNG', quality=100)
            
            print(f"[INFO] PNG変換完了: {png_path} (サイズ: {img.size[0]}x{img.size[1]})")
            
            return png_filename
        
        except Exception as e:
            print(f"[WARNING] PDF→PNG変換失敗: {e}")
            import traceback
            traceback.print_exc()
            return None

    # ========================================================================

    def _render_sheet_fallback(self, sheet, dpi: int = 600, insert_index: Optional[int] = None, insert_images: bool = True) -> bool:
        """シート全体を1枚のPNG画像にレンダリング(真のフォールバック)
        
        isolated-groupレンダリングが行われない場合、または失敗した場合の最終手段として、
        シート全体を1枚のPNG画像として出力します。
        
        注意:
            isolated-groupレンダリングは_process_sheet_imagesで実行されるため、
            このメソッドでは単純にシート全体をPNG化するのみです。
        
        Args:
            sheet: 対象シート
            dpi: DPI設定(デフォルト: 600)
            insert_index: Markdown挿入位置(None=末尾)
            insert_images: True=即座に挿入、False=登録のみ
        
        Returns:
            成功時True、失敗時False
        """
        tmpdir = None
        try:
            # 一時ディレクトリを作成
            tmpdir = tempfile.mkdtemp(prefix='xls2md_render_')
            
            # 1. Excel→PDF変換 (Phase 1メソッド)
            debug_print(f"[DEBUG] Fallback rendering for sheet: {sheet.title}")
            pdf_path = self._convert_excel_to_pdf(self.excel_file, tmpdir, apply_fit_to_page=True)
            if pdf_path is None:
                print("[WARNING] LibreOffice がPDFを出力しませんでした")
                return False
            
            # 2. シートのページインデックスを取得
            try:
                page_index = int(self.workbook.sheetnames.index(sheet.title))
            except (ValueError, TypeError):
                page_index = 0
            
            # 3. PDF→PNG変換 (Phase 1メソッド)
            safe_sheet = self._sanitize_filename(sheet.title)
            result_filename = self._convert_pdf_page_to_png(
                pdf_path,
                page_index,
                dpi,
                self.images_dir,
                f"{self.base_name}_{safe_sheet}_sheet"
            )
            
            if result_filename is None:
                print("[WARNING] ImageMagick による PNG 変換が失敗しました")
                return False
            
            # 4. 画像をMarkdownに登録または挿入
            if insert_images:
                # 即座にMarkdownに挿入
                md_line = f"![{sheet.title}](images/{result_filename})"
                try:
                    self._insert_markdown_image(insert_index, md_line, result_filename, sheet=sheet)
                    print(f"[SUCCESS] シート全体の画像を挿入: {result_filename}")
                except Exception as e:
                    print(f"[WARNING] 画像挿入失敗: {e}")
                    # フォールバック: markdown_linesに直接追加
                    self.markdown_lines.append(md_line)
                    self.markdown_lines.append("")
                    self._mark_image_emitted(result_filename)
            else:
                # 後で挿入するために登録のみ
                self._sheet_shape_images.setdefault(sheet.title, [])
                self._sheet_shape_images[sheet.title].append((1, result_filename))
            
            return True
            
        except Exception as e:
            print(f"[WARNING] フォールバックレンダリングエラー: {e}")
            import traceback
            traceback.print_exc()
            return False
            
        finally:
            # 一時ディレクトリをクリーンアップ
            if tmpdir and os.path.isdir(tmpdir):
                try:
                    shutil.rmtree(tmpdir, ignore_errors=True)
                except (ValueError, TypeError):
                    pass  # 一時ディレクトリ削除失敗は無視

    def _detect_image_format(self, image_data: bytes) -> str:
        """Detect common image formats from initial bytes and return extension.

        Falls back to .png if unknown.
        """
        try:
            if not image_data or len(image_data) < 4:
                return '.png'
            # JPEG
            if image_data.startswith(b'\xff\xd8'):
                return '.jpg'
            # PNG
            if image_data.startswith(b'\x89PNG'):
                return '.png'
            # GIF
            if image_data.startswith(b'GIF87a') or image_data.startswith(b'GIF89a'):
                return '.gif'
            # BMP
            if image_data.startswith(b'BM'):
                return '.bmp'
            return '.png'
        except Exception:
            return '.png'

    def _sanitize_filename(self, s: str) -> str:
        """ファイルシステム上で安全なファイル名に正規化する。

        - Unicode 正規化 (NFKC)
        - 連続空白はアンダースコアに置換
        - ファイル名に使えない記号 (/\\:*?"<>|) を除去
        - 複数アンダースコアを単一に、先頭/末尾のアンダースコアを除去
        """
        import unicodedata, re
        if s is None:
            return 'image'
        txt = unicodedata.normalize('NFKC', str(s))
        # replace whitespace with underscore
        txt = re.sub(r"\s+", '_', txt)
        # remove characters that are problematic in filenames
        txt = re.sub(r'[/\\:*?"<>|]', '', txt)
        # collapse multiple underscores
        txt = re.sub(r'_+', '_', txt)
        # remove leading/trailing underscores
        txt = txt.strip('_')
        if not txt:
            return 'image'
        return txt

    def _get_drawing_max_col_row(self, sheet):
        """図形が参照する最大の列・行番号を取得する。
        
        Returns:
            (max_col, max_row): 図形が参照する最大の列・行番号のタプル。
                               図形が存在しない場合は (None, None) を返す。
        """
        try:
            metadata = self._get_drawing_xml_and_metadata(sheet)
            if metadata is None:
                return None, None
            
            drawing_xml = metadata['drawing_xml']
            ns = {'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'}
            
            max_col = None
            max_row = None
            
            for node in drawing_xml:
                lname = node.tag.split('}')[-1].lower()
                if lname not in ('twocellanchor', 'onecellanchor'):
                    continue
                
                if lname == 'twocellanchor':
                    fr = node.find('xdr:from', ns)
                    to = node.find('xdr:to', ns)
                    if fr is not None:
                        try:
                            col = int(fr.find('xdr:col', ns).text)
                            row = int(fr.find('xdr:row', ns).text)
                            if max_col is None or col > max_col:
                                max_col = col
                            if max_row is None or row > max_row:
                                max_row = row
                        except Exception:
                            pass
                    if to is not None:
                        try:
                            col = int(to.find('xdr:col', ns).text)
                            row = int(to.find('xdr:row', ns).text)
                            if max_col is None or col > max_col:
                                max_col = col
                            if max_row is None or row > max_row:
                                max_row = row
                        except Exception:
                            pass
                elif lname == 'onecellanchor':
                    fr = node.find('xdr:from', ns)
                    if fr is not None:
                        try:
                            col = int(fr.find('xdr:col', ns).text)
                            row = int(fr.find('xdr:row', ns).text)
                            if max_col is None or col > max_col:
                                max_col = col
                            if max_row is None or row > max_row:
                                max_row = row
                        except Exception:
                            pass
            
            return max_col, max_row
        except Exception:
            return None, None

    def _compute_sheet_cell_pixel_map(self, sheet, DPI=300, min_cols=None, min_rows=None):
        """Compute approximate pixel positions for column right-edges and row bottom-edges.

        Returns (col_x, row_y) where col_x[0] == 0 and col_x[i] is the right edge
        of column i (1-based). row_y similar for rows.
        
        図形が参照する列・行がシートの max_column/max_row より大きい場合は、
        図形の範囲まで計算を拡張します。
        """
        try:
            max_col = sheet.max_column
            max_row = sheet.max_row
            
            drawing_max_col, drawing_max_row = self._get_drawing_max_col_row(sheet)
            if drawing_max_col is not None:
                drawing_max_col_1based = drawing_max_col + 1
                if drawing_max_col_1based > max_col:
                    max_col = drawing_max_col_1based
            if drawing_max_row is not None:
                drawing_max_row_1based = drawing_max_row + 1
                if drawing_max_row_1based > max_row:
                    max_row = drawing_max_row_1based
            
            if min_cols is not None:
                max_col = max(max_col, min_cols)
            if min_rows is not None:
                max_row = max(max_row, min_rows)
            
            col_pixels = []
            from openpyxl.utils import get_column_letter
            for c in range(1, max_col+1):
                cd = sheet.column_dimensions.get(get_column_letter(c))
                # Excel column width is in character units. Use a more
                # accurate conversion based on Microsoft's documented
                # equation: pixels = floor(((256*W + floor(128/MAX_DIGIT_WIDTH))/256) * MAX_DIGIT_WIDTH)
                # where MAX_DIGIT_WIDTH approximates the maximum digit width
                # in pixels for the workbook's default font. We use 7 as a
                # conservative default (common for Calibri/Arial at default size).
                width = getattr(cd, 'width', None) if cd is not None else None
                if width is None:
                    try:
                        from openpyxl.utils import units as _units
                        width = getattr(sheet.sheet_format, 'defaultColWidth', None) or _units.DEFAULT_COLUMN_WIDTH
                    except Exception:
                        width = 8.43
                try:
                    import math
                    # Compute base pixel width at standard screen DPI (96). Then
                    # scale to the requested DPI so EMU offsets (which are later
                    # converted using the target rasterization DPI) align with the
                    # produced PDF/PNG pixels. This reduces mismatches between
                    # drawing EMU conversions and column pixel map used elsewhere.
                    MAX_DIGIT_WIDTH = 7
                    base_px = int(math.floor(((256.0 * float(width) + math.floor(128.0 / MAX_DIGIT_WIDTH)) / 256.0) * MAX_DIGIT_WIDTH))
                    if base_px < 1:
                        base_px = 1
                    # scale from 96 DPI (typical screen) to target DPI
                    scale = float(DPI) / 96.0 if DPI and DPI > 0 else 1.0
                    px = max(1, int(round(base_px * scale)))
                except (ValueError, TypeError):
                    # fallback heuristic, also scale by DPI
                    try:
                        base = max(1, int(float(width) * 7 + 5))
                        px = max(1, int(round(base * (float(DPI) / 96.0))))
                    except (ValueError, TypeError):
                        px = max(1, int(float(width) * 7 + 5))
                col_pixels.append(px)

            row_pixels = []
            for r in range(1, max_row+1):
                rd = sheet.row_dimensions.get(r)
                hpts = getattr(rd, 'height', None) if rd is not None else None
                if hpts is None:
                    try:
                        from openpyxl.utils import units as _units
                        hpts = _units.DEFAULT_ROW_HEIGHT
                    except Exception:
                        hpts = 15
                # Row heights are in points; convert to pixels at the target DPI
                try:
                    px = max(1, int(float(hpts) * DPI / 72.0))
                except (ValueError, TypeError):
                    px = max(1, int(hpts * DPI / 72))
                row_pixels.append(px)

            col_x = [0]
            for v in col_pixels:
                col_x.append(col_x[-1] + v)
            row_y = [0]
            for v in row_pixels:
                row_y.append(row_y[-1] + v)

            return col_x, row_y
        except Exception:
            return [0], [0]

    def _to_positive(self, value, orig_ext, orig_ch_ext, target_px):
        """Ensure the EMU extent is positive.

        Priority for choosing a positive extent:
        1. keep 'value' if it's already positive
        2. fall back to orig_ext (if provided and >0)
        3. fall back to orig_ch_ext (if provided and >0)
        4. fall back to converting target_px -> EMU (at least 1 px)
        5. finally return 1 EMU as absolute safe minimum

        This helper is defensive and avoids raising errors; it always
        returns an int > 0.
        """
        try:
            v = int(value) if value is not None else 0
        except (ValueError, TypeError):
            v = 0
        if v and v > 0:
            return v
        try:
            if orig_ext is not None:
                oe = int(orig_ext)
                if oe > 0:
                    return oe
        except (ValueError, TypeError):
            pass  # 型変換失敗は無視
        try:
            if orig_ch_ext is not None:
                oc = int(orig_ch_ext)
                if oc > 0:
                    return oc
        except (ValueError, TypeError):
            pass  # 型変換失敗は無視
        try:
            # target_px is in pixels; convert to EMU using object's dpi if available
            DPI = int(getattr(self, 'dpi', 300) or 300)
            EMU_PER_INCH = 914400
            emu_per_pixel = EMU_PER_INCH / float(DPI) if DPI and DPI > 0 else EMU_PER_INCH / 300.0
            px = float(target_px) if target_px is not None else 1.0
            emu = int(round(max(1.0, px) * emu_per_pixel))
            if emu and emu > 0:
                return emu
        except (ValueError, TypeError):
            pass  # 型変換失敗は無視
        # absolute fallback
        return 1

    def _snap_box_to_cell_bounds(self, box, col_x, row_y, DPI=300):
        """Snap a pixel box (l,t,r,b) to nearest enclosing cell boundaries using
        the provided col_x and row_y arrays. Returns integer pixel box.
        """
        try:
            l, t, r, btm = box
            # find start column: smallest c such that col_x[c] >= l (allow small tolerance)
            # tol scales with DPI to preserve previous behavior when DPI differs
            try:
                tol = max(1, int(DPI / 300.0 * 3))  # a few pixels tolerance dependent on DPI
            except (ValueError, TypeError):
                tol = 3
            start_col = None
            for c in range(1, len(col_x)):
                if col_x[c] >= l - tol:
                    start_col = c
                    break
            if start_col is None:
                start_col = max(1, len(col_x)-1)

            # find end column: smallest c such that col_x[c] >= r (allow small tolerance)
            end_col = None
            for c in range(1, len(col_x)):
                if col_x[c] >= r + tol:
                    end_col = c
                    break
            if end_col is None:
                end_col = max(1, len(col_x)-1)

            # rows
            start_row = None
            for rr in range(1, len(row_y)):
                if row_y[rr] >= t - tol:
                    start_row = rr
                    break
            if start_row is None:
                start_row = max(1, len(row_y)-1)

            end_row = None
            for rr in range(1, len(row_y)):
                if row_y[rr] >= btm + tol:
                    end_row = rr
                    break
            if end_row is None:
                end_row = max(1, len(row_y)-1)

            left_px = max(0, int(col_x[start_col-1]))
            top_px = max(0, int(row_y[start_row-1]))
            right_px = int(col_x[end_col]) if end_col < len(col_x) else int(col_x[-1])
            bottom_px = int(row_y[end_row]) if end_row < len(row_y) else int(row_y[-1])

            return left_px, top_px, right_px, bottom_px
        except (ValueError, TypeError):
            return int(box[0]), int(box[1]), int(box[2]), int(box[3])

    def _find_content_bbox(self, pil_image, white_thresh: int = 245):
        """Find bounding box of non-white content in a PIL Image.

        white_thresh: pixel brightness threshold (0-255). Pixels with all channels
        >= white_thresh are considered background/white. Returns (l,t,r,b) or None
        if no content detected.
        """
        try:
            if pil_image.mode not in ('RGB', 'RGBA'):
                img = pil_image.convert('RGB')
            else:
                img = pil_image
            pixels = img.load()
            w, h = img.size
            left = w; top = h; right = 0; bottom = 0
            found = False
            for y in range(h):
                for x in range(w):
                    r, g, b = pixels[x, y][:3]
                    if r < white_thresh or g < white_thresh or b < white_thresh:
                        found = True
                        if x < left: left = x
                        if x > right: right = x
                        if y < top: top = y
                        if y > bottom: bottom = y
            if not found:
                return None
            # make right/bottom inclusive -> convert to typical crop coords (r+1,b+1)
            return (left, top, right + 1, bottom + 1)
        except Exception:
            return None

    def _crop_image_preserving_connectors(self, image_path: str, dpi: int = 300, white_thresh: int = 245):
        """Open image at image_path, find non-white bbox and crop with padding.

        Adds small padding (dependent on DPI) to avoid cutting connector/arrow tips.
        Overwrites the original file with the cropped result.
        """
        try:
            from PIL import Image
            if not os.path.exists(image_path):
                return False
            im = Image.open(image_path)
            bbox = self._find_content_bbox(im, white_thresh=white_thresh)
            if not bbox:
                # nothing to crop
                im.close()
                return True
            l, t, r, b = bbox
            # padding to avoid cutting thin arrow tips; scale with DPI
            base_pad = max(6, int(dpi / 300.0 * 12))
            # Bias bottom padding slightly larger to avoid clipped tails/arrowheads
            pad_top = base_pad
            pad_left = base_pad
            pad_right = base_pad
            pad_bottom = max(base_pad, int(base_pad * 1.25))
            l = max(0, l - pad_left)
            t = max(0, t - pad_top)
            r = min(im.width, r + pad_right)
            b = min(im.height, b + pad_bottom)
            # perform crop and save (preserve mode)
            try:
                cropped = im.crop((l, t, r, b))
                cropped.save(image_path)
                cropped.close()
            except (ValueError, TypeError):
                # fallback: do not overwrite if crop fails
                pass
            im.close()
            return True
        except Exception:
            return False

    def _get_pdf_page_box_points(self, pdf_path: str):
        """Return (width_points, height_points) for the first page using CropBox or MediaBox in the PDF.

        This is a lightweight parser that searches for '/CropBox' or '/MediaBox' arrays in the PDF bytes.
        Returns None on failure.
        """
        try:
            with open(pdf_path, 'rb') as f:
                data = f.read()
            # search for CropBox first then MediaBox
            import re
            pat = re.compile(rb"/CropBox\s*\[\s*([0-9.+\-eE]+)\s+([0-9.+\-eE]+)\s+([0-9.+\-eE]+)\s+([0-9.+\-eE]+)\s*\]")
            m = pat.search(data)
            if not m:
                pat2 = re.compile(rb"/MediaBox\s*\[\s*([0-9.+\-eE]+)\s+([0-9.+\-eE]+)\s+([0-9.+\-eE]+)\s+([0-9.+\-eE]+)\s*\]")
                m = pat2.search(data)
            if not m:
                return None
            a = float(m.group(1))
            b = float(m.group(2))
            c = float(m.group(3))
            d = float(m.group(4))
            width_pts = abs(c - a)
            height_pts = abs(d - b)
            return (width_pts, height_pts)
        except (ValueError, TypeError):
            return None

    def _extract_drawing_cell_ranges(self, sheet) -> List[Tuple[int,int,int,int]]:
        """Extract drawing cell ranges (start_col, end_col, start_row, end_row) for each drawable anchor.

        Uses drawing XML when available. Returns list aligned with anchors order used in other extractors.
        """
        print(f"[INFO] シート図形セル範囲抽出: {sheet.title}")
        ranges = []
        try:
            # Use Phase 1 foundation method to get drawing XML
            metadata = self._get_drawing_xml_and_metadata(sheet)
            if metadata is None:
                return ranges
            
            z = metadata['zip']
            drawing_xml = metadata['drawing_xml']

            # prepare pixel map for oneCell ext conversions
            # Use a consistent DPI when converting EMU offsets to pixels.
            DPI = 300
            try:
                DPI = int(getattr(self, 'dpi', DPI) or DPI)
            except (ValueError, TypeError):
                pass  # データ構造操作失敗は無視
            try:
                DPI = int(getattr(self, 'dpi', DPI) or DPI)
            except (ValueError, TypeError):
                DPI = DPI
            col_x, row_y = self._compute_sheet_cell_pixel_map(sheet, DPI=DPI)
            EMU_PER_INCH = 914400
            try:
                EMU_PER_PIXEL = EMU_PER_INCH / float(DPI)
            except (ValueError, TypeError):
                EMU_PER_PIXEL = EMU_PER_INCH / float(DPI)

            ns = {'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'}
            for node in drawing_xml:
                lname = node.tag.split('}')[-1].lower()
                if lname not in ('twocellanchor', 'onecellanchor'):
                    continue
                # determine cell indices
                if lname == 'twocellanchor':
                    fr = node.find('xdr:from', ns)
                    to = node.find('xdr:to', ns)
                    if fr is None or to is None:
                        continue
                    try:
                        col = int(fr.find('xdr:col', ns).text)
                        row = int(fr.find('xdr:row', ns).text)
                        to_col = int(to.find('xdr:col', ns).text)
                        to_row = int(to.find('xdr:row', ns).text)
                    except (ValueError, TypeError):
                        continue
                    start_col = col + 1
                    end_col = to_col + 1
                    start_row = row + 1
                    end_row = to_row + 1

                    try:
                        ns_a = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
                        sp = node.find('.//xdr:sp', ns)
                        if sp is not None:
                            prst_geom = sp.find('.//a:prstGeom', ns_a)
                            if prst_geom is not None:
                                prst = prst_geom.get('prst', '')
                                if 'callout' in prst.lower():
                                    # This is a callout shape, check adjustment values
                                    av_lst = prst_geom.find('a:avLst', ns_a)
                                    if av_lst is not None:
                                        adj1_elem = av_lst.find('a:gd[@name="adj1"]', ns_a)
                                        adj2_elem = av_lst.find('a:gd[@name="adj2"]', ns_a)
                                        
                                        adj1 = 0
                                        adj2 = 0
                                        if adj1_elem is not None:
                                            fmla = adj1_elem.get('fmla', '')
                                            if fmla.startswith('val '):
                                                try:
                                                    adj1 = int(fmla.split()[1])
                                                except (ValueError, IndexError):
                                                    pass
                                        if adj2_elem is not None:
                                            fmla = adj2_elem.get('fmla', '')
                                            if fmla.startswith('val '):
                                                try:
                                                    adj2 = int(fmla.split()[1])
                                                except (ValueError, IndexError):
                                                    pass
                                        
                                        if adj1 < 0 or adj2 < 0:
                                            if start_row > 1:
                                                start_row -= 1
                    except Exception:
                        pass

                    ranges.append((start_col, end_col, start_row, end_row))
                else:
                    # oneCellAnchor: use from.col/from.row and ext cx/cy to derive end cell
                    fr = node.find('xdr:from', ns)
                    ext = node.find('xdr:ext', ns)
                    if fr is None or ext is None:
                        continue
                    try:
                        col = int(fr.find('xdr:col', ns).text)
                        row = int(fr.find('xdr:row', ns).text)
                        colOff = int(fr.find('xdr:colOff', ns).text)
                    except (ValueError, TypeError):
                        continue
                    cx = int(ext.attrib.get('cx', '0'))
                    cy = int(ext.attrib.get('cy', '0'))
                    left_px = col_x[col] + (colOff / EMU_PER_PIXEL) if col < len(col_x) else col_x[-1]
                    right_px = left_px + (cx / EMU_PER_PIXEL)
                    top_px = row_y[row] if row < len(row_y) else row_y[-1]
                    bottom_px = top_px + (cy / EMU_PER_PIXEL)
                    # map pixels to cell indices
                    # find start_col index
                    start_col = 1
                    for ci in range(1, len(col_x)):
                        if col_x[ci] >= left_px:
                            start_col = ci
                            break
                    end_col = len(col_x)-1
                    for ci in range(1, len(col_x)):
                        if col_x[ci] >= right_px:
                            end_col = ci
                            break
                    start_row = 1
                    for ri in range(1, len(row_y)):
                        if row_y[ri] >= top_px:
                            start_row = ri
                            break
                    end_row = len(row_y)-1
                    for ri in range(1, len(row_y)):
                        if row_y[ri] >= bottom_px:
                            end_row = ri
                            break
                    ranges.append((start_col, end_col, start_row, end_row))
        except Exception:
            pass  # データ構造操作失敗は無視
        print(f"[INFO] 抽出されたセル範囲: {ranges}")
        return ranges

    def _anchor_is_connector_only(self, sheet, anchor_idx) -> bool:
        """Return True if the anchor at anchor_idx in the sheet's drawing
        appears to be a connector-only anchor (i.e. contains connector endpoint
        references but no drawable pictorial/shape elements). Conservative:
        return False if information can't be determined.
        """
        try:
            z = zipfile.ZipFile(self.excel_file)
            sheet_index = self.workbook.sheetnames.index(sheet.title)
            rels_path = f"xl/worksheets/_rels/sheet{sheet_index+1}.xml.rels"
            if rels_path not in z.namelist():
                return False
            rels_xml = ET.fromstring(z.read(rels_path))
            drawing_target = None
            for rel in rels_xml.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                t = rel.attrib.get('Type','')
                if t.endswith('/drawing'):
                    drawing_target = rel.attrib.get('Target')
                    break
            if not drawing_target:
                return False
            drawing_path = drawing_target
            if drawing_path.startswith('..'):
                drawing_path = drawing_path.replace('../', 'xl/')
            if drawing_path.startswith('/'):
                drawing_path = drawing_path.lstrip('/')
            if drawing_path not in z.namelist():
                drawing_path = drawing_path.replace('worksheets', 'drawings')
                if drawing_path not in z.namelist():
                    return False
            drawing_xml = ET.fromstring(z.read(drawing_path))
            # locate the requested anchor node
            idx = 0
            for node in drawing_xml:
                lname = node.tag.split('}')[-1].lower()
                if lname not in ('twocellanchor', 'onecellanchor'):
                    continue
                if idx == anchor_idx:
                    # inspect node children for drawable types vs connector refs
                    has_drawable = False
                    has_connector_ref = False
                    for desc in node.iter():
                        t = desc.tag.split('}')[-1].lower()
                        if t in ('pic', 'sp', 'graphicframe', 'grpsp'):
                            has_drawable = True
                        if t in ('stcxn', 'endcxn', 'stcxnpr', 'endcxnpr'):
                            has_connector_ref = True
                        for k in desc.attrib.keys():
                            if k.lower() == 'id' and desc.tag.split('}')[-1].lower() != 'cnvpr':
                                has_connector_ref = True
                    return (has_connector_ref and not has_drawable)
                idx += 1
        except Exception:
            return False
        return False

    def _sheet_has_drawings(self, sheet) -> bool:
        """Return True if the sheet has drawing relationships pointing to drawing XML
        that contain drawable elements (pic/sp/graphicFrame)."""
        try:
            z = zipfile.ZipFile(self.excel_file)
            sheet_index = self.workbook.sheetnames.index(sheet.title)
            rels_path = f"xl/worksheets/_rels/sheet{sheet_index+1}.xml.rels"
            if rels_path not in z.namelist():
                return False
            rels_xml = ET.fromstring(z.read(rels_path))
            drawing_target = None
            for rel in rels_xml.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                t = rel.attrib.get('Type','')
                if t.endswith('/drawing'):
                    drawing_target = rel.attrib.get('Target')
                    break
            if not drawing_target:
                return False
            drawing_path = drawing_target
            if drawing_path.startswith('..'):
                drawing_path = drawing_path.replace('../', 'xl/')
            if drawing_path.startswith('/'):
                drawing_path = drawing_path.lstrip('/')
            if drawing_path not in z.namelist():
                drawing_path = drawing_path.replace('worksheets', 'drawings')
                if drawing_path not in z.namelist():
                    return False
            drawing_xml = ET.fromstring(z.read(drawing_path))
            # look for drawable descendants
            for node in drawing_xml.iter():
                lname = node.tag.split('}')[-1].lower()
                if lname in ('pic', 'sp', 'graphicframe', 'graphic', 'grpsp'):
                    return True
            return False
        except (ET.ParseError, KeyError, AttributeError):
            return False

    def _extract_shape_metadata_from_anchor(self, anchor, sheet) -> Dict[str, Any]:
        """drawing anchorから図形のメタデータを抽出
        
        Args:
            anchor: XML anchor element
            sheet: 対象シート
            
        Returns:
            図形のメタデータを含む辞書
        """
        metadata = {
            'type': 'unknown',
            'name': '',
            'position': {},
            'size': {},
            'text_content': [],
            'connector_type': None,
            'shape_properties': {}
        }
        
        try:
            ns_xdr = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
            ns_a = 'http://schemas.openxmlformats.org/drawingml/2006/main'
            
            anchor_type = anchor.tag.split('}')[-1].lower()
            metadata['anchor_type'] = anchor_type
            
            for sub in anchor.iter():
                if sub.tag.split('}')[-1].lower() == 'cnvpr':
                    metadata['name'] = sub.attrib.get('name', '')
                    metadata['id'] = sub.attrib.get('id', '')
                    metadata['description'] = sub.attrib.get('descr', '')
                    break
            
            if anchor_type == 'twocellanchor':
                fr = anchor.find('{%s}from' % ns_xdr)
                to = anchor.find('{%s}to' % ns_xdr)
                if fr is not None:
                    col_elem = fr.find('{%s}col' % ns_xdr)
                    row_elem = fr.find('{%s}row' % ns_xdr)
                    if col_elem is not None and row_elem is not None:
                        metadata['position']['from_col'] = int(col_elem.text) + 1
                        metadata['position']['from_row'] = int(row_elem.text) + 1
                if to is not None:
                    col_elem = to.find('{%s}col' % ns_xdr)
                    row_elem = to.find('{%s}row' % ns_xdr)
                    if col_elem is not None and row_elem is not None:
                        metadata['position']['to_col'] = int(col_elem.text) + 1
                        metadata['position']['to_row'] = int(row_elem.text) + 1
            
            elif anchor_type == 'onecellanchor':
                fr = anchor.find('{%s}from' % ns_xdr)
                ext = anchor.find('{%s}ext' % ns_xdr)
                if fr is not None:
                    col_elem = fr.find('{%s}col' % ns_xdr)
                    row_elem = fr.find('{%s}row' % ns_xdr)
                    if col_elem is not None and row_elem is not None:
                        metadata['position']['from_col'] = int(col_elem.text) + 1
                        metadata['position']['from_row'] = int(row_elem.text) + 1
                if ext is not None:
                    cx = ext.attrib.get('cx', '0')
                    cy = ext.attrib.get('cy', '0')
                    metadata['size']['width_emu'] = int(cx)
                    metadata['size']['height_emu'] = int(cy)
            
            for desc in anchor.iter():
                tag_name = desc.tag.split('}')[-1].lower()
                
                if tag_name == 'pic':
                    metadata['type'] = 'picture'
                elif tag_name == 'sp':
                    metadata['type'] = 'shape'
                    for prstgeom in desc.iter():
                        if prstgeom.tag.split('}')[-1].lower() == 'prstgeom':
                            prst = prstgeom.attrib.get('prst', '')
                            metadata['shape_properties']['preset'] = prst
                            break
                elif tag_name == 'cxnsp':
                    metadata['type'] = 'connector'
                    for prstgeom in desc.iter():
                        if prstgeom.tag.split('}')[-1].lower() == 'prstgeom':
                            prst = prstgeom.attrib.get('prst', '')
                            metadata['connector_type'] = prst
                            break
                elif tag_name == 'grpsp':
                    metadata['type'] = 'group'
                elif tag_name == 'graphicframe':
                    metadata['type'] = 'graphic_frame'
                
                if tag_name == 't' and desc.text and desc.text.strip():
                    metadata['text_content'].append(desc.text.strip())
            
            return metadata
            
        except Exception as e:
            print(f"[WARNING] Shape metadata extraction failed: {e}")
            return metadata

    def _extract_all_shapes_metadata(self, sheet, filter_ids: Optional[Set[str]] = None) -> List[Dict[str, Any]]:
        """シート内の全図形のメタデータを抽出
        
        Args:
            sheet: 対象シート
            filter_ids: 抽出対象の図形IDセット（Noneの場合は全て抽出）
            
        Returns:
            図形メタデータのリスト
        """
        shapes_metadata = []
        
        try:
            metadata = self._get_drawing_xml_and_metadata(sheet)
            if metadata is None:
                return shapes_metadata
            
            drawing_xml = metadata['drawing_xml']
            
            for anchor in drawing_xml:
                anchor_type = anchor.tag.split('}')[-1].lower()
                if anchor_type in ('twocellanchor', 'onecellanchor'):
                    if self._anchor_has_drawable(anchor):
                        if filter_ids is not None:
                            shape_id = None
                            for sub in anchor.iter():
                                if sub.tag.split('}')[-1].lower() == 'cnvpr':
                                    shape_id = sub.attrib.get('id', '')
                                    break
                            if shape_id and shape_id not in filter_ids:
                                continue
                        
                        shape_meta = self._extract_shape_metadata_from_anchor(anchor, sheet)
                        shapes_metadata.append(shape_meta)
            
            try:
                metadata['zip'].close()
            except:
                pass
                
        except Exception as e:
            print(f"[WARNING] Failed to extract shapes metadata: {e}")
        
        return shapes_metadata

    def _format_shape_metadata_as_text(self, shapes_metadata: List[Dict[str, Any]]) -> str:
        """図形メタデータを人間が読みやすいテキスト形式に整形
        
        Args:
            shapes_metadata: 図形メタデータのリスト
            
        Returns:
            整形されたテキスト
        """
        if not shapes_metadata:
            return ""
        
        lines = []
        lines.append("### 図形情報")
        lines.append("")
        
        for idx, meta in enumerate(shapes_metadata, 1):
            shape_type = meta.get('type', 'unknown')
            shape_name = meta.get('name', f'図形{idx}')
            
            type_map = {
                'picture': '画像',
                'shape': '図形',
                'connector': 'コネクタ',
                'group': 'グループ',
                'graphic_frame': 'グラフィックフレーム',
                'unknown': '不明'
            }
            type_ja = type_map.get(shape_type, shape_type)
            
            lines.append(f"**{shape_name}** ({type_ja})")
            
            pos = meta.get('position', {})
            if 'from_col' in pos and 'from_row' in pos:
                from_cell = f"{col_letter(pos['from_col'])}{pos['from_row']}"
                if 'to_col' in pos and 'to_row' in pos:
                    to_cell = f"{col_letter(pos['to_col'])}{pos['to_row']}"
                    lines.append(f"- 位置: {from_cell} ～ {to_cell}")
                else:
                    lines.append(f"- 位置: {from_cell} から")
            
            if shape_type == 'shape':
                preset = meta.get('shape_properties', {}).get('preset', '')
                if preset:
                    lines.append(f"- 図形タイプ: {preset}")
            
            if shape_type == 'connector':
                conn_type = meta.get('connector_type', '')
                if conn_type:
                    lines.append(f"- コネクタタイプ: {conn_type}")
            
            text_content = meta.get('text_content', [])
            if text_content:
                lines.append(f"- テキスト: {' / '.join(text_content)}")
            
            description = meta.get('description', '')
            if description:
                lines.append(f"- 説明: {description}")
            
            lines.append("")
        
        return '\n'.join(lines)

    def _format_shape_metadata_as_json(self, shapes_metadata: List[Dict[str, Any]]) -> str:
        """図形メタデータをJSON形式に整形
        
        Args:
            shapes_metadata: 図形メタデータのリスト
            
        Returns:
            JSON文字列
        """
        if not shapes_metadata:
            return "{}"
        
        import json
        
        output_data = {
            'shapes': [],
            'total_count': len(shapes_metadata)
        }
        
        for meta in shapes_metadata:
            shape_data = {
                'name': meta.get('name', ''),
                'type': meta.get('type', 'unknown'),
                'anchor_type': meta.get('anchor_type', ''),
                'position': meta.get('position', {}),
                'size': meta.get('size', {}),
                'text_content': meta.get('text_content', []),
                'properties': {}
            }
            
            if meta.get('type') == 'shape':
                shape_data['properties'] = meta.get('shape_properties', {})
            elif meta.get('type') == 'connector':
                shape_data['properties']['connector_type'] = meta.get('connector_type', '')
            
            if meta.get('description'):
                shape_data['description'] = meta.get('description')
            
            if meta.get('id'):
                shape_data['id'] = meta.get('id')
            
            output_data['shapes'].append(shape_data)
        
        return json.dumps(output_data, ensure_ascii=False, indent=2)

    def _anchor_has_drawable(self, a) -> bool:
        """Shared helper: determine whether a drawing anchor contains drawable
        content (pictures, shapes, graphicFrames or connector references).
        This central implementation keeps extraction and trimming logic
        consistent so clustering indices align with anchors.
        """
        try:
            # Create cache dict on the instance to avoid re-evaluating the same
            # anchor multiple times during a single conversion run. Use the
            # closest cNvPr/@id attribute as a stable key when available; fall
            # back to a short hash of the anchor XML when no id is present.
            try:
                cache = getattr(self, '_anchor_drawable_cache')
            except Exception:
                cache = {}
                try:
                    setattr(self, '_anchor_drawable_cache', cache)
                except Exception as e:
                    pass  # XML解析エラーは無視

            key = None
            try:
                for sub in a.iter():
                    if sub.tag.split('}')[-1].lower() == 'cnvpr':
                        cid = sub.attrib.get('id') or sub.attrib.get('idx')
                        if cid is not None:
                            key = f"cnvpr:{cid}"
                            break
            except Exception:
                key = None

            if key is None:
                try:
                    # fallback: small stable fingerprint of the anchor XML
                    import hashlib
                    raw = ET.tostring(a) if hasattr(ET, 'tostring') else None
                    if raw:
                        key = 'hash:' + hashlib.sha1(raw).hexdigest()[:8]
                    else:
                        key = 'anon'
                except Exception:
                    key = 'anon'

            # Return cached result if present
            try:
                if key in cache:
                    return cache[key]
            except Exception as e:
                pass  # XML解析エラーは無視

            drawable_types = []
            has_text = False
            has_connector_ref = False
            for desc in a.iter():
                lname = desc.tag.split('}')[-1].lower()
                # detect text content
                if lname == 't' and (desc.text and desc.text.strip()):
                    has_text = True
                # explicit pictorial/shape types (including connector shapes)
                if lname in ('pic', 'sp', 'graphicframe', 'grpsp', 'cxnsp'):
                    # check for hidden flag on closest cNvPr child
                    is_hidden = False
                    for sub in desc.iter():
                        if sub.tag.split('}')[-1].lower() == 'cnvpr':
                            if sub.attrib.get('hidden') in ('1', 'true'):
                                is_hidden = True
                                break
                    if is_hidden:
                        continue
                    drawable_types.append(lname)
                # detect connector endpoint references
                if lname in ('stcxn', 'endcxn', 'stcxnpr', 'endcxnpr'):
                    has_connector_ref = True
                # detect non-cNvPr elements exposing an id attribute (heuristic)
                for k in desc.attrib.keys():
                    if k.lower() == 'id' and desc.tag.split('}')[-1].lower() != 'cnvpr':
                        has_connector_ref = True

            result = False
            if drawable_types:
                debug_print(f"[DEBUG] Anchor has drawable elements: {drawable_types}")
                result = True
            elif has_connector_ref:
                debug_print(f"[DEBUG] Anchor has connector references; treating as drawable")
                result = True
            elif has_text:
                debug_print(f"[DEBUG] Anchor contains only text; treating as non-drawable")
                result = False
            else:
                debug_print(f"[DEBUG] Anchor has no drawable elements")
                result = False

            # cache and return
            cache[key] = result
            return result
        except (ValueError, TypeError):
            return False
    
    def _cluster_shapes_common(self, sheet, shapes, cell_ranges=None, max_groups=2):
        """Centralized clustering by integer row gaps when cell_ranges are available.

        Returns (clusters, debug_dict). clusters is a list of groups (lists of indices).
        debug_dict contains diagnostic information useful for tracing split decisions.
        If cell_ranges is not provided or insufficient, falls back to centroid clustering.
        """
        try:
            debug = {'method': 'row_gap', 'clusters': None, 'indices_sorted': None, 'chosen_split': None, 'reason': None}
            if not cell_ranges or len(cell_ranges) < len(shapes):
                debug['reason'] = 'no_cell_ranges'
                clusters = self._cluster_shape_indices(shapes, max_groups=max_groups)
                debug['clusters'] = clusters
                return clusters, debug

            # build centers by vertical midpoint and sort
            row_centers = [(((cr[2] + cr[3]) / 2.0) if (cr[2] is not None and cr[3] is not None) else 0.0, idx) for idx, cr in enumerate(cell_ranges)]
            row_centers.sort(key=lambda x: x[0])
            indices_sorted = [idx for _, idx in row_centers]
            debug['indices_sorted'] = indices_sorted

            # compute per-index start/end rows
            s_rows = []
            e_rows = []
            for idx in indices_sorted:
                try:
                    cr = cell_ranges[idx]
                    s_rows.append(int(cr[2]))
                    e_rows.append(int(cr[3]))
                except (ValueError, TypeError):
                    s_rows.append(None); e_rows.append(None)

            # build covered rows set
            all_covered = set()
            for cr in cell_ranges:
                try:
                    rs = int(cr[2]); re_ = int(cr[3])
                except (ValueError, TypeError):
                    continue
                for rr in range(rs, re_ + 1):
                    all_covered.add(rr)
            debug['all_covered_count'] = len(all_covered)

            # check for dominating large spans (relative)
            try:
                row_spans = []
                s_list = [int(cr[2]) for cr in cell_ranges if cr[2] is not None]
                e_list = [int(cr[3]) for cr in cell_ranges if cr[3] is not None]
                for _, idx in row_centers:
                    cr = cell_ranges[idx]
                    row_spans.append(int(cr[3]) - int(cr[2]))
                total_row_span = max(e_list) - min(s_list) if e_list and s_list else 0
                rel_row_span_thresh = 0.75
                dominating = []
                if total_row_span > 0:
                    dominating = [rs for rs in row_spans if (rs / float(total_row_span) >= rel_row_span_thresh)]
                if dominating:
                    debug['reason'] = 'dominating_span'
                    clusters = [indices_sorted]
                    debug['clusters'] = clusters
                    debug['chosen_split'] = None
                    return clusters, debug
            except (ValueError, TypeError) as e:
                debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")

            # try adjacent pair splits where integer empty rows exist
            split_at = None
            chosen_row = None
            total_rows = None
            try:
                # Calculate total_rows from cell_ranges instead of sheet.max_row
                e_list = [int(cr[3]) for cr in cell_ranges if cr[3] is not None]
                total_rows = max(e_list) if e_list else None
            except (ValueError, TypeError):
                total_rows = None

            for gi in range(len(indices_sorted) - 1):
                try:
                    left_max = max([int(cell_ranges[i][3]) for i in indices_sorted[:gi+1] if cell_ranges[i][3] is not None])
                    right_min = min([int(cell_ranges[i][2]) for i in indices_sorted[gi+1:] if cell_ranges[i][2] is not None])
                except (ValueError, TypeError):
                    left_max = None; right_min = None
                if left_max is None or right_min is None:
                    continue
                if right_min - left_max >= 2:
                    candidate = left_max + 1
                    if candidate not in all_covered:
                        split_at = gi + 1
                        chosen_row = candidate
                        break

            # fallback: find largest uncovered interior gap
            if split_at is None:
                try:
                    if total_rows:
                        uncovered = [r for r in range(1, total_rows+1) if r not in all_covered]
                        if uncovered:
                            # build contiguous gaps
                            gaps = []
                            start = uncovered[0]; prev = uncovered[0]
                            for r in uncovered[1:]:
                                if r == prev + 1:
                                    prev = r
                                else:
                                    gaps.append((start, prev)); start = r; prev = r
                            gaps.append((start, prev))
                            gaps.sort(key=lambda x: x[1] - x[0], reverse=True)
                            for gap_start, gap_end in gaps:
                                if gap_start == 1 or gap_end == total_rows:
                                    continue
                                gap_len = (gap_end - gap_start + 1)
                                if gap_len >= 2:
                                    left = []; right = []
                                    for idx in indices_sorted:
                                        try:
                                            s_r = int(cell_ranges[idx][2]); e_r = int(cell_ranges[idx][3])
                                        except (ValueError, TypeError):
                                            s_r = None; e_r = None
                                        if e_r is not None and e_r < gap_start:
                                            left.append(idx)
                                        elif s_r is not None and s_r > gap_end:
                                            right.append(idx)
                                    if left and right:
                                        clusters = [left, right]
                                        debug['chosen_split'] = ('gap', gap_start, gap_end)
                                        debug['clusters'] = clusters
                                        return clusters, debug
                except (ValueError, TypeError) as e:
                    debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")

            if split_at is not None:
                clusters = [indices_sorted[:split_at], indices_sorted[split_at:]]
                debug['chosen_split'] = ('adjacent', chosen_row)
                debug['clusters'] = clusters
                return clusters, debug

            # no valid integer row split found; return all shapes as single cluster
            clusters = [list(range(len(shapes)))]
            debug['reason'] = 'no_row_split'
            debug['clusters'] = clusters
            return clusters, debug
        except Exception as e:
            debug_print(f"[DEBUG] クラスタリングエラー: {e}")
            return [[i for i in range(len(shapes))]], {'reason': 'fatal'}
    
    def _get_image_position(self, image):
        """画像の位置情報を取得

        戻り値:
          - 成功時: {'col': int or None, 'row': int} の辞書（1-based 行/列インデックス）
          - 失敗時/不明: 文字列メッセージ（既存ログとの互換性維持）

        呼び出し側は dict を期待しており、dict の場合は 'row' を使って
        画像の代表開始行を決めるロジックがあります。ここで構造化して
        返すことで start_row が 1 固定になる問題を修正します。
        """
        try:
            if hasattr(image, 'anchor'):
                anchor = image.anchor
                # openpyxl anchor may expose a _from attribute with 0-based col/row
                if hasattr(anchor, '_from'):
                    try:
                        col_idx = getattr(anchor._from, 'col', None)
                        row_idx = getattr(anchor._from, 'row', None)
                        # convert to 1-based indices when present
                        col_val = int(col_idx) + 1 if col_idx is not None else None
                        row_val = int(row_idx) + 1 if row_idx is not None else None
                        if row_val is not None:
                            return {'col': col_val, 'row': row_val}
                    except (ValueError, TypeError):
                        # fall through to string fallback
                        pass
            return "位置情報なし"
        except (ValueError, TypeError):
            return "位置情報不明"

    def _extract_drawing_shapes(self, sheet) -> List[Tuple[int,int,int,int]]:
        """Extract shape bounding boxes from the workbook drawing XML and convert
        coordinates to pixel units matching the DPI used for rasterization.
        Returns list of (left, top, right, bottom) tuples.
        """
        try:
            # Use Phase 1 foundation method to get drawing XML
            metadata = self._get_drawing_xml_and_metadata(sheet)
            if metadata is None:
                return []
            
            drawing_xml = metadata['drawing_xml']
            # prepare simple column/row pixel mapping using runtime DPI
            DPI = 300
            try:
                DPI = int(getattr(self, 'dpi', DPI) or DPI)
            except (ValueError, TypeError):
                DPI = DPI
            EMU_PER_INCH = 914400
            try:
                EMU_PER_PIXEL = EMU_PER_INCH / float(DPI)
            except (ValueError, TypeError):
                EMU_PER_PIXEL = EMU_PER_INCH / float(DPI)

            max_col = max(sheet.max_column, 100)  # 図形が範囲外にある可能性を考慮
            max_row = max(sheet.max_row, 200)
            col_pixels = []
            for c in range(1, max_col+1):
                cd = sheet.column_dimensions.get(get_column_letter(c))
                width = getattr(cd, 'width', None) if cd is not None else None
                if width is None:
                    width = 8.43
                px = max(1, int(width * 7 + 5))
                col_pixels.append(px)
            row_pixels = []
            for r in range(1, max_row+1):
                rd = sheet.row_dimensions.get(r)
                hpts = getattr(rd, 'height', None) if rd is not None else None
                if hpts is None:
                    hpts = 15
                px = max(1, int(hpts * DPI / 72))
                row_pixels.append(px)
            col_x = [0]
            for v in col_pixels:
                col_x.append(col_x[-1] + v)
            row_y = [0]
            for v in row_pixels:
                row_y.append(row_y[-1] + v)

            ns = {'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'}
            bboxes = []
            # helper to check if anchor contains drawable element
            # delegate to centralized helper for consistency
            def anchor_has_drawable(a):
                return self._anchor_has_drawable(a)

            # Iterate top-level drawing children in document order so the
            # ordering matches the anchors list built by the isolated-shape
            # trimming path. This prevents index/anchor mismatches when
            # clustering by shape centers.
            for node in list(drawing_xml):
                lname = node.tag.split('}')[-1].lower()
                if lname not in ('twocellanchor', 'onecellanchor'):
                    continue
                # only consider anchors that have drawable content
                if not anchor_has_drawable(node):
                    continue
                if lname == 'twocellanchor':
                    fr = node.find('xdr:from', ns)
                    to = node.find('xdr:to', ns)
                    if fr is None or to is None:
                        continue
                    try:
                        col = int(fr.find('xdr:col', ns).text)
                        colOff = int(fr.find('xdr:colOff', ns).text)
                        row = int(fr.find('xdr:row', ns).text)
                        rowOff = int(fr.find('xdr:rowOff', ns).text)
                        to_col = int(to.find('xdr:col', ns).text)
                        to_colOff = int(to.find('xdr:colOff', ns).text)
                        to_row = int(to.find('xdr:row', ns).text)
                        to_rowOff = int(to.find('xdr:rowOff', ns).text)
                    except (ValueError, TypeError):
                        continue
                    # 安全な配列アクセス(範囲外チェック)
                    if col < 0 or col >= len(col_x) or row < 0 or row >= len(row_y):
                        continue
                    if to_col < 0 or to_col >= len(col_x) or to_row < 0 or to_row >= len(row_y):
                        continue
                    left = col_x[col] + (colOff / EMU_PER_PIXEL)
                    top = row_y[row] + (rowOff / EMU_PER_PIXEL)
                    right = col_x[to_col] + (to_colOff / EMU_PER_PIXEL)
                    bottom = row_y[to_row] + (to_rowOff / EMU_PER_PIXEL)
                else:
                    fr = node.find('xdr:from', ns)
                    ext = node.find('xdr:ext', ns)
                    if fr is None or ext is None:
                        continue
                    try:
                        col = int(fr.find('xdr:col', ns).text)
                        colOff = int(fr.find('xdr:colOff', ns).text)
                        row = int(fr.find('xdr:row', ns).text)
                        rowOff = int(fr.find('xdr:rowOff', ns).text)
                        cx = int(ext.attrib.get('cx', '0'))
                        cy = int(ext.attrib.get('cy', '0'))
                    except (ValueError, TypeError):
                        continue
                    # 安全な配列アクセス(範囲外チェック)
                    if col < 0 or col >= len(col_x) or row < 0 or row >= len(row_y):
                        continue
                    left = col_x[col] + (colOff / EMU_PER_PIXEL)
                    top = row_y[row] + (rowOff / EMU_PER_PIXEL)
                    right = left + (cx / EMU_PER_PIXEL)
                    bottom = top + (cy / EMU_PER_PIXEL)
                # filter out boxes that cover most of the page (likely not a small drawing)
                page_w = col_x[-1]
                page_h = row_y[-1]
                try:
                    box_area = max(0, right-left) * max(0, bottom-top)
                    page_area = max(1, page_w * page_h)
                    if box_area / page_area > 0.85:
                        continue
                except Exception as e:
                    print(f"[WARNING] ファイル操作エラー: {e}")
                bboxes.append((left, top, right, bottom))

            # return bboxes (list of (left, top, right, bottom) in pixel-ish units)
            debug_print(f"[DEBUG] _extract_drawing_shapes found {len(bboxes)} bboxes")
            return bboxes
        except Exception as e:
            print(f"[WARNING] _extract_drawing_shapes exception: {e}")
            import traceback
            traceback.print_exc()
            return []

    def _render_sheet_isolated_group_v2(self, sheet, shape_indices: List[int], dpi: int = 600, cell_range: Optional[Tuple[int,int,int,int]] = None) -> Optional[Tuple[str, int]]:
        """
        Render a group of shape indices as a single isolated image (refactored version).
        
        **EXPERIMENTAL - NOT RECOMMENDED FOR PRODUCTION USE**
        
        This is a streamlined implementation that uses extracted helper methods for:
        - Connector reference resolution (_resolve_connector_references)
        - Drawing anchor pruning (_prune_drawing_anchors)
        
        **MISSING**: Connector cosmetic processing (~600 lines)
        - connector_children_by_id construction
        - Theme color resolution for schemeClr -> srgbClr
        - Line style materialization from lnRef
        - Arrow head/tail preservation
        - Duplicate cosmetic element deduplication
        
        **Result**: Images generated by this method may have:
        - Missing or incorrect connector line styles
        - Wrong connector colors
        - Missing arrow heads/tails
        
        **RECOMMENDATION**: Use the original _render_sheet_isolated_group method for production.
        This v2 method is kept for educational purposes and as a foundation for future refactoring.
        
        Args:
            sheet: Worksheet to render
            shape_indices: List of shape indices to include
            dpi: DPI for rendering (default: 600)
            cell_range: Optional tuple (s_col, e_col, s_row, e_row) to constrain the output
        
        Returns:
            Generated filename (relative to images_dir) or None on failure
        """
        import warnings
        warnings.warn(
            "_render_sheet_isolated_group_v2 is experimental and missing connector cosmetic processing. "
            "Use the original _render_sheet_isolated_group for production.",
            FutureWarning,
            stacklevel=2
        )
        
        try:
            # Reset preserved IDs marker
            try:
                self._last_iso_preserved_ids = set()
            except Exception:
                pass  # ファイルクローズ失敗は無視
            
            # Open Excel file and locate drawing
            zpath = self.excel_file
            with zipfile.ZipFile(zpath, 'r') as z:
                sheet_index = self.workbook.sheetnames.index(sheet.title)
                rels_path = f"xl/worksheets/_rels/sheet{sheet_index+1}.xml.rels"
                
                if rels_path not in z.namelist():
                    debug_print(f"[DEBUG][_iso_v2] sheet={sheet.title} missing rels")
                    return None
                
                # Find drawing relationship
                rels_xml = ET.fromstring(z.read(rels_path))
                drawing_target = None
                for rel in rels_xml.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                    if rel.attrib.get('Type', '').endswith('/drawing'):
                        drawing_target = rel.attrib.get('Target')
                        break
                
                if not drawing_target:
                    debug_print(f"[DEBUG][_iso_v2] sheet={sheet.title} no drawing relationship")
                    return None
                
                # Normalize drawing path
                drawing_path = drawing_target
                if drawing_path.startswith('..'):
                    drawing_path = drawing_path.replace('../', 'xl/')
                drawing_path = drawing_path.lstrip('/')
                
                if drawing_path not in z.namelist():
                    drawing_path = drawing_path.replace('worksheets', 'drawings')
                    if drawing_path not in z.namelist():
                        debug_print(f"[DEBUG][_iso_v2] drawing_path not found: {drawing_path}")
                        return None
                
                # Parse drawing XML
                drawing_xml_bytes = z.read(drawing_path)
                drawing_xml = ET.fromstring(drawing_xml_bytes)
            
            # Filter anchors to drawable elements only
            anchors = []
            for node in drawing_xml:
                lname = node.tag.split('}')[-1].lower()
                if lname in ('twocellanchor', 'onecellanchor') and self._anchor_has_drawable(node):
                    anchors.append(node)
            
            if not anchors:
                debug_print(f"[DEBUG][_iso_v2] no drawable anchors found")
                return None
            
                # Compute cell_range if not provided
            # Also track the minimum row for this cluster (used for markdown ordering)
            cluster_min_row = 1  # Default fallback
            if cell_range is None and shape_indices:
                try:
                    all_ranges = self._extract_drawing_cell_ranges(sheet)
                    picked = [all_ranges[idx] for idx in shape_indices if 0 <= idx < len(all_ranges)]
                    if picked:
                        s_col = min(r[0] for r in picked)
                        e_col = max(r[1] for r in picked)
                        s_row = min(r[2] for r in picked)
                        e_row = max(r[3] for r in picked)
                        
                        # Store cluster minimum row for later use in markdown ordering
                        cluster_min_row = s_row
                        
                        # Add 10% padding to ensure shapes are fully visible
                        # Some shapes may have borders or connectors that extend beyond their anchor points
                        col_padding = max(2, int((e_col - s_col) * 0.1))
                        row_padding = max(2, int((e_row - s_row) * 0.1))
                        s_col = max(1, s_col - col_padding)
                        e_col = e_col + col_padding
                        s_row = max(1, s_row - row_padding)
                        e_row = e_row + row_padding
                        
                        cell_range = (s_col, e_col, s_row, e_row)
                        debug_print(f"[DEBUG][_iso_v2] Computed cell_range from shapes: cols {s_col}-{e_col}, rows {s_row}-{e_row} (with padding)")
                        debug_print(f"[DEBUG][_iso_v2] Original shape ranges: {picked}")
                except Exception as e:
                    debug_print(f"[DEBUG][_iso_v2] Failed to compute cell_range: {e}")            # Build keep_cnvpr_ids from shape_indices
            keep_cnvpr_ids = set()
            for si in shape_indices:
                if 0 <= si < len(anchors):
                    for sub in anchors[si].iter():
                        if sub.tag.split('}')[-1].lower() == 'cnvpr':
                            cid = sub.attrib.get('id')
                            if cid:
                                keep_cnvpr_ids.add(str(cid))
                            break
            
            debug_print(f"[DEBUG][_iso_v2] anchors={len(anchors)} keep_ids={sorted(list(keep_cnvpr_ids))}")
            
            # Use helper method to resolve connector references
            referenced_ids = self._resolve_connector_references(
                drawing_xml=drawing_xml,
                anchors=anchors,
                keep_cnvpr_ids=keep_cnvpr_ids
            )
            
            # Expose preserved IDs for callers
            try:
                self._last_iso_preserved_ids = set(referenced_ids)
            except Exception as e:
                print(f"[WARNING] ファイル操作エラー: {e}")
            
            # Create temp directory and extract workbook
            tmp_base = tempfile.mkdtemp(prefix='xls2md_iso_v2_base_')
            tmpdir = tempfile.mkdtemp(prefix='xls2md_iso_v2_', dir=tmp_base)
            try:
                with zipfile.ZipFile(zpath, 'r') as zin:
                    zin.extractall(tmpdir)
                
                # Remove all sheets except the target sheet to avoid including unrelated sheets in output
                # This ensures the generated Excel file contains only the trimmed target sheet
                # Keep the target sheet's drawing file to maintain proper references
                try:
                    # Get the target sheet's drawing file name (if any) to preserve it
                    target_sheet_drawing = None
                    target_sheet_rels_path = os.path.join(tmpdir, f'xl/worksheets/_rels/sheet{sheet_index+1}.xml.rels')
                    if os.path.exists(target_sheet_rels_path):
                        try:
                            rels_tree = ET.parse(target_sheet_rels_path)
                            rels_root = rels_tree.getroot()
                            for rel in rels_root:
                                rel_type = rel.attrib.get('Type', '')
                                if '/drawing' in rel_type:
                                    target_drawing = rel.attrib.get('Target', '')
                                    if target_drawing:
                                        # Normalize path: ../drawings/drawing1.xml -> drawing1.xml
                                        target_sheet_drawing = os.path.basename(target_drawing)
                                        break
                        except (ET.ParseError, KeyError, AttributeError) as e:
                            debug_print(f"[DEBUG] XML解析エラー（無視）: {type(e).__name__}")
                    
                    # Parse workbook.xml to get sheet relationships
                    wb_path = os.path.join(tmpdir, 'xl/workbook.xml')
                    wb_rels_path = os.path.join(tmpdir, 'xl/_rels/workbook.xml.rels')
                    
                    if os.path.exists(wb_path):
                        wb_tree = ET.parse(wb_path)
                        wb_root = wb_tree.getroot()
                        wb_ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
                        rel_ns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
                        
                        # Find all sheets and keep only the target sheet
                        sheets_el = wb_root.find(f'{{{wb_ns}}}sheets')
                        if sheets_el is not None:
                            target_sheet_rid = None
                            sheets_to_remove = []
                            
                            for idx, sheet_el in enumerate(list(sheets_el)):
                                if idx == sheet_index:
                                    # This is our target sheet - keep it and get its relationship ID
                                    target_sheet_rid = sheet_el.attrib.get(f'{{{rel_ns}}}id')
                                else:
                                    # Mark for removal
                                    sheets_to_remove.append((idx, sheet_el))
                            
                            # Remove non-target sheets from workbook.xml
                            for _, sheet_el in sheets_to_remove:
                                sheets_el.remove(sheet_el)
                            
                            # Renumber the target sheet to be sheet 1 (sheetId="1")
                            # This ensures Excel/LibreOffice properly recognize it as the first sheet
                            if sheets_el is not None:
                                for sheet_el in list(sheets_el):
                                    # Set sheetId to 1 (first sheet)
                                    sheet_el.set('sheetId', '1')
                                    # Update relationship ID to rId1
                                    sheet_el.set(f'{{{rel_ns}}}id', 'rId1')
                            
                            # Write back modified workbook.xml
                            wb_tree.write(wb_path, encoding='utf-8', xml_declaration=True)
                            
                            # Parse workbook.xml.rels to find which relationship IDs to keep
                            if os.path.exists(wb_rels_path):
                                rels_tree = ET.parse(wb_rels_path)
                                rels_root = rels_tree.getroot()
                                pkg_rel_ns = 'http://schemas.openxmlformats.org/package/2006/relationships'
                                
                                # Find sheet relationship targets to remove
                                rels_to_remove = []
                                target_sheet_rel = None
                                for rel in list(rels_root):
                                    rid = rel.attrib.get('Id')
                                    target = rel.attrib.get('Target')
                                    rel_type = rel.attrib.get('Type', '')
                                    
                                    # Keep target sheet relationship, remove others
                                    if rel_type.endswith('/worksheet'):
                                        if rid == target_sheet_rid:
                                            target_sheet_rel = rel
                                        else:
                                            rels_to_remove.append(rel)
                                
                                # Remove non-target sheet relationships
                                for rel in rels_to_remove:
                                    rels_root.remove(rel)
                                
                                # Renumber target sheet relationship to rId1
                                if target_sheet_rel is not None:
                                    target_sheet_rel.set('Id', 'rId1')
                                
                                # Write back modified rels
                                rels_tree.write(wb_rels_path, encoding='utf-8', xml_declaration=True)
                            
                            # Remove physical sheet files for non-target sheets
                            for idx, _ in sheets_to_remove:
                                # Remove sheet XML file
                                sheet_file = os.path.join(tmpdir, f'xl/worksheets/sheet{idx+1}.xml')
                                if os.path.exists(sheet_file):
                                    os.remove(sheet_file)
                                
                                # Remove sheet rels file
                                sheet_rels = os.path.join(tmpdir, f'xl/worksheets/_rels/sheet{idx+1}.xml.rels')
                                if os.path.exists(sheet_rels):
                                    os.remove(sheet_rels)
                            
                            # Remove ALL drawing files EXCEPT the target sheet's drawing
                            # This prevents orphaned drawing references from causing errors
                            drawings_dir = os.path.join(tmpdir, 'xl/drawings')
                            if os.path.exists(drawings_dir):
                                for fname in os.listdir(drawings_dir):
                                    # Skip the target sheet's drawing file
                                    if target_sheet_drawing and fname == target_sheet_drawing:
                                        continue
                                    
                                    # Remove other drawing XML files
                                    if fname.endswith('.xml') and not fname.startswith('_rels'):
                                        drawing_file = os.path.join(drawings_dir, fname)
                                        try:

                                            os.remove(p)

                                        except (OSError, FileNotFoundError):

                                            pass  # ファイル削除失敗は無視
                                
                                # Remove drawing rels that don't belong to target sheet
                                rels_dir = os.path.join(drawings_dir, '_rels')
                                if os.path.exists(rels_dir) and target_sheet_drawing:
                                    target_rels = target_sheet_drawing.replace('.xml', '.xml.rels')
                                    for fname in os.listdir(rels_dir):
                                        if fname != target_rels and fname.endswith('.rels'):
                                            try:
                                                os.remove(os.path.join(rels_dir, fname))
                                            except Exception:
                                                pass  # 一時ファイルの削除失敗は無視
                            
                            debug_print(f"[DEBUG][_iso_v2] Removed {len(sheets_to_remove)} non-target sheets from workbook (kept drawing: {target_sheet_drawing or 'none'})")
                
                except Exception as e:
                    if getattr(self, 'verbose', False):
                        print(f"[WARN][_iso_v2] Failed to remove non-target sheets: {e}")
                        import traceback
                        traceback.print_exc()
                
                # Compute group_rows from cell_range
                group_rows = set()
                if cell_range:
                    try:
                        s_col, e_col, s_row, e_row = cell_range
                        group_rows = set(range(s_row, e_row + 1))
                    except (ValueError, TypeError) as e:
                        debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
                
                # Use helper method to prune drawing anchors
                drawing_relpath = os.path.join(tmpdir, drawing_path)
                self._prune_drawing_anchors(
                    drawing_relpath=drawing_relpath,
                    keep_cnvpr_ids=keep_cnvpr_ids,
                    referenced_ids=referenced_ids,
                    cell_range=cell_range,
                    group_rows=group_rows
                )
                
                # CRITICAL: Do NOT adjust drawing coordinates
                # Keep original drawing.xml coordinates intact
                # LibreOffice needs the original coordinates to properly render shapes
                # We only trim the cell data, not the drawing positions
                debug_print(f"[DEBUG][_iso_v2] Preserving original drawing coordinates (no adjustment)")
                if cell_range:
                    s_col, e_col, s_row, e_row = cell_range
                    debug_print(f"[DEBUG][_iso_v2] Cell range for data trimming: cols {s_col}-{e_col}, rows {s_row}-{e_row}")
                
                # DO NOT reconstruct worksheet - keep all original data
                # This preserves the original cell positions so shapes can reference them correctly
                # Only prune the drawing anchors, not the cell data
                sheet_rel = os.path.join(tmpdir, f"xl/worksheets/sheet{sheet_index+1}.xml")
                
                # However, we MUST fix the pageSetup to prevent scale=25 shrinking
                # This is done separately from worksheet reconstruction
                if os.path.exists(sheet_rel):
                    try:
                        stree = ET.parse(sheet_rel)
                        sroot = stree.getroot()
                        ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
                        
                        # Configure pageSetup with proper scaling
                        # CRITICAL: Remove existing pageSetup and create new one with scale=100
                        # fitToHeight/fitToWidth can shrink shapes to microscopic sizes
                        # Remove all existing pageSetup elements
                        for old_ps in list(sroot.findall(f'.//{{{ns}}}pageSetup')):
                            sroot.remove(old_ps)
                        
                        # Create new pageSetup with normal 100% scale
                        ps = ET.Element(f'{{{ns}}}pageSetup')
                        ps.set('scale', '100')
                        ps.set('paperSize', '1')  # Letter (standard)
                        ps.set('orientation', 'portrait')
                        ps.set('pageOrder', 'downThenOver')
                        ps.set('blackAndWhite', 'false')
                        ps.set('draft', 'false')
                        ps.set('cellComments', 'none')
                        ps.set('horizontalDpi', '300')
                        ps.set('verticalDpi', '300')
                        ps.set('copies', '1')
                        # Append at the end of sheet
                        sroot.append(ps)
                        
                        # Write back the modified sheet
                        stree.write(sheet_rel, encoding='utf-8', xml_declaration=True)
                        debug_print(f"[DEBUG][_iso_v2] Set pageSetup to scale=100 (normal size) to preserve shapes")
                    except Exception as e:
                        if getattr(self, 'verbose', False):
                            print(f"[WARN][_iso_v2] Failed to fix pageSetup: {e}")
                
                # Worksheet reconstruction code (DISABLED - keep original sheet data)
                if False and os.path.exists(sheet_rel) and cell_range:
                    try:
                        s_col, e_col, s_row, e_row = cell_range
                        stree = ET.parse(sheet_rel)
                        sroot = stree.getroot()
                        ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
                        
                        # Read original sheet.xml from source Excel file to get cell values
                        with zipfile.ZipFile(self.excel_file, 'r') as src_z:
                            src_sheet_path = f"xl/worksheets/sheet{sheet_index+1}.xml"
                            if src_sheet_path in src_z.namelist():
                                src_sheet_xml = ET.fromstring(src_z.read(src_sheet_path))
                                src_sheet_data = src_sheet_xml.find(f'{{{ns}}}sheetData')
                            else:
                                src_sheet_data = None
                        
                        # Reconstruct sheetData to only include rows/columns in range
                        # KEEP ORIGINAL ROW/COLUMN NUMBERS (do not renumber from 1)
                        sheet_data_tag = f'{{{ns}}}sheetData'
                        sheet_data = sroot.find(sheet_data_tag)
                        if sheet_data is not None and src_sheet_data is not None:
                            new_sheet_data = ET.Element(sheet_data_tag)
                            src_rows = src_sheet_data.findall(f'{{{ns}}}row')
                            debug_print(f"[DEBUG][_iso_v2] Found {len(src_rows)} rows in source sheet.xml")
                            cells_copied = 0
                            
                            # Copy rows in range, keeping original row numbers
                            for row_el in src_rows:
                                try:
                                    rnum = int(row_el.attrib.get('r', '0'))
                                except (ValueError, TypeError):
                                    continue
                                if rnum < s_row or rnum > e_row:
                                    continue
                                
                                # Create new row with ORIGINAL row number
                                new_row = ET.Element(f'{{{ns}}}row')
                                new_row.set('r', str(rnum))  # Keep original row number
                                
                                # Copy row attributes
                                for attr in ('ht', 'hidden', 'customHeight'):
                                    if attr in row_el.attrib:
                                        new_row.set(attr, row_el.attrib.get(attr))
                                
                                # Copy cells in column range, keeping original column letters
                                for c in list(row_el):
                                    if c.tag.split('}')[-1] != 'c':
                                        continue
                                    cell_r = c.attrib.get('r', '')
                                    col_letters = ''.join([ch for ch in cell_r if ch.isalpha()]) if cell_r else None
                                    if not col_letters:
                                        continue
                                    
                                    # Convert column letters to index
                                    col_idx = 0
                                    for ch in col_letters:
                                        col_idx = col_idx * 26 + (ord(ch.upper()) - 64)
                                    if col_idx < s_col or col_idx > e_col:
                                        continue
                                    
                                    # Copy cell with ORIGINAL cell reference (e.g., "D17")
                                    import copy
                                    new_cell = copy.deepcopy(c)
                                    new_row.append(new_cell)
                                    cells_copied += 1
                                
                                if len(new_row) > 0:  # Only add row if it has cells
                                    new_sheet_data.append(new_row)
                            
                            debug_print(f"[DEBUG][_iso_v2] Copied {cells_copied} cells with original row/col numbers")
                            
                            # Replace old sheetData with new one
                            for child in list(sroot):
                                if child.tag == sheet_data_tag:
                                    sroot.remove(child)
                            sroot.append(new_sheet_data)
                            
                            # Update dimension element with ORIGINAL row/column numbers
                            dim_tag = f'{{{ns}}}dimension'
                            dim = sroot.find(dim_tag)
                            if dim is None:
                                dim = ET.Element(dim_tag)
                                sroot.insert(0, dim)
                            # Use original row/col numbers
                            start_addr = f"{col_letter(s_col)}{s_row}"
                            end_addr = f"{col_letter(e_col)}{e_row}"
                            dim.set('ref', f"{start_addr}:{end_addr}")
                        
                        # Rebuild cols element with ORIGINAL column numbers
                        cols_tag = f'{{{ns}}}cols'
                        col_tag = f'{{{ns}}}col'
                        for child in list(sroot):
                            if child.tag == cols_tag:
                                try:
                                    sroot.remove(child)
                                except Exception:
                                    pass  # 一時ファイルの削除失敗は無視
                        
                        cols_el = ET.Element(cols_tag)
                        try:
                            from openpyxl.utils import get_column_letter
                            default_col_w = getattr(sheet.sheet_format, 'defaultColWidth', None) or 8.43
                            for c in range(s_col, e_col + 1):
                                cd = sheet.column_dimensions.get(get_column_letter(c))
                                width = getattr(cd, 'width', None) if cd else None
                                hidden = getattr(cd, 'hidden', None) if cd else None
                                if width is None:
                                    width = default_col_w
                                
                                col_el = ET.Element(col_tag)
                                # Use ORIGINAL column indices (not renumbered)
                                col_el.set('min', str(c))
                                col_el.set('max', str(c))
                                col_el.set('width', str(float(width)))
                                if cd and getattr(cd, 'width', None) is not None:
                                    col_el.set('customWidth', '1')
                                if hidden:
                                    col_el.set('hidden', '1')
                                cols_el.append(col_el)
                        except (ValueError, TypeError):
                            # Fallback: set default widths with ORIGINAL column numbers
                            for c in range(s_col, e_col + 1):
                                col_el = ET.Element(col_tag)
                                col_el.set('min', str(c))
                                col_el.set('max', str(c))
                                col_el.set('width', '8.43')
                                cols_el.append(col_el)
                        
                        # Insert cols before sheetData
                        sd_idx = list(sroot).index(new_sheet_data)
                        sroot.insert(sd_idx, cols_el)
                        
                        # Set page margins to zero (same as original method)
                        # Add or modify sheetPr with pageSetupPr fitToPage attribute
                        sheet_pr = sroot.find(f'.//{{{ns}}}sheetPr')
                        if sheet_pr is None:
                            sheet_pr = ET.Element(f'{{{ns}}}sheetPr')
                            sroot.insert(0, sheet_pr)
                        page_setup_pr = sheet_pr.find(f'{{{ns}}}pageSetUpPr')
                        if page_setup_pr is None:
                            page_setup_pr = ET.SubElement(sheet_pr, f'{{{ns}}}pageSetUpPr')
                        page_setup_pr.set('fitToPage', '1')
                        
                        # Add or modify printOptions
                        print_opts = sroot.find(f'.//{{{ns}}}printOptions')
                        if print_opts is None:
                            print_opts = ET.Element(f'{{{ns}}}printOptions')
                            sroot.append(print_opts)
                        print_opts.set('horizontalCentered', '1')
                        print_opts.set('verticalCentered', '1')
                        
                        # Configure pageSetup with proper scaling
                        # CRITICAL: Remove existing pageSetup and create new one with scale=100
                        # fitToHeight/fitToWidth can shrink shapes to microscopic sizes
                        # Remove all existing pageSetup elements
                        for old_ps in list(sroot.findall(f'.//{{{ns}}}pageSetup')):
                            sroot.remove(old_ps)
                        
                        # Create new pageSetup with normal 100% scale
                        ps = ET.Element(f'{{{ns}}}pageSetup')
                        ps.set('scale', '100')
                        ps.set('paperSize', '1')  # Letter (standard)
                        ps.set('orientation', 'portrait')
                        ps.set('pageOrder', 'downThenOver')
                        ps.set('blackAndWhite', 'false')
                        ps.set('draft', 'false')
                        ps.set('cellComments', 'none')
                        ps.set('horizontalDpi', '300')
                        ps.set('verticalDpi', '300')
                        ps.set('copies', '1')
                        # Append at the end of sheet
                        sroot.append(ps)
                        debug_print(f"[DEBUG][_iso_v2] Set pageSetup to scale=100 (normal size) to preserve shapes")
                        
                        # Set page margins (as attributes, standard Excel format)
                        pm_tag = f'{{{ns}}}pageMargins'
                        pm = sroot.find(pm_tag)
                        if pm is None:
                            pm = ET.Element(pm_tag)
                            sroot.append(pm)
                        pm.set('left', '0.25')
                        pm.set('right', '0.25')
                        pm.set('top', '0.75')
                        pm.set('bottom', '0.75')
                        pm.set('header', '0.3')
                        pm.set('footer', '0.3')
                        
                        # Remove header/footer elements
                        hf_tag = f'{{{ns}}}headerFooter'
                        for hf in list(sroot.findall(hf_tag)):
                            sroot.remove(hf)
                        
                        stree.write(sheet_rel, encoding='utf-8', xml_declaration=True)
                        debug_print(f"[DEBUG][_iso_v2] Reconstructed sheet data: kept original rows {s_row}-{e_row}, cols {s_col}-{e_col}")
                    except Exception as e:
                        if getattr(self, 'verbose', False):
                            print(f"[WARN][_iso_v2] Failed to reconstruct worksheet: {e}")

                # CRITICAL: Remove Print_Area completely to ensure all shapes are visible
                # Print_Area restricts the visible area and can hide shapes outside the defined range
                # Since we're preserving the full sheet structure, we don't need Print_Area
                try:
                    wb_rel = os.path.join(tmpdir, 'xl/workbook.xml')
                    if os.path.exists(wb_rel):
                        wtree = ET.parse(wb_rel)
                        wroot = wtree.getroot()
                        ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
                        
                        # Find definedNames element
                        dn_tag = f'{{{ns}}}definedNames'
                        dn = wroot.find(dn_tag)
                        
                        # Remove ALL defined names (including Print_Area) to prevent display issues
                        if dn is not None:
                            wroot.remove(dn)
                            debug_print(f"[DEBUG][_iso_v2] Removed Print_Area and all defined names to ensure shapes are visible")
                        
                        wtree.write(wb_rel, encoding='utf-8', xml_declaration=True)
                except Exception as e:
                    if getattr(self, 'verbose', False):
                        print(f"[WARN][_iso_v2] Failed to remove Print_Area: {e}")

                # Create trimmed workbook ZIP for debugging (saved in output dir)
                debug_xlsx_filename = f"{self.base_name}_{sheet.title}_group_{shape_indices[0] if shape_indices else 0}_debug.xlsx"
                debug_xlsx_path = os.path.join(self.output_dir, debug_xlsx_filename)
                debug_zip_base = os.path.join(self.output_dir, f"{self.base_name}_{sheet.title}_group_{shape_indices[0] if shape_indices else 0}_debug")
                
                try:
                    # Remove old files if they exist
                    if os.path.exists(debug_xlsx_path):
                        os.remove(debug_xlsx_path)
                    if os.path.exists(debug_zip_base + '.zip'):
                        os.remove(debug_zip_base + '.zip')
                    
                    shutil.make_archive(debug_zip_base, 'zip', tmpdir)
                    shutil.move(debug_zip_base + '.zip', debug_xlsx_path)
                    debug_print(f"[DEBUG][_iso_v2] Saved debug workbook: {debug_xlsx_path}")
                except Exception as e:
                    if getattr(self, 'verbose', False):
                        print(f"[WARN][_iso_v2] Failed to create trimmed workbook: {e}")
                    return None

                # Convert to PDF and PNG (save PDF for debugging)
                try:
                    # DO NOT apply fit-to-page - it shrinks shapes to 25% making them invisible
                    # pageSetup is already configured properly in the worksheet XML above
                    # self._set_excel_fit_to_one_page(debug_xlsx_path)  # DISABLED
                    
                    # Convert to PDF (output to same directory as xlsx)
                    cmd = [LIBREOFFICE_PATH, '--headless', '--convert-to', 'pdf', '--outdir', self.output_dir, debug_xlsx_path]
                    debug_print(f"[DEBUG][_iso_v2] LibreOffice command: {' '.join(cmd)}")
                    proc = subprocess.run(cmd, capture_output=True, text=True, timeout=90)
                    
                    if proc.returncode != 0:
                        if getattr(self, 'verbose', False):
                            print(f"[WARN][_iso_v2] LibreOffice PDF conversion failed: {proc.stderr}")
                        return None
                    
                    # Find generated PDF
                    debug_pdf_filename = debug_xlsx_filename.replace('.xlsx', '.pdf')
                    pdf_path = os.path.join(self.output_dir, debug_pdf_filename)
                    
                    if not os.path.exists(pdf_path):
                        # Try to find any PDF that was created
                        pdf_candidates = [f for f in os.listdir(self.output_dir) 
                                        if f.lower().endswith('.pdf') and 'group' in f and sheet.title in f]
                        if not pdf_candidates:
                            if getattr(self, 'verbose', False):
                                print("[WARN][_iso_v2] PDF conversion failed - no output")
                            return None
                        pdf_path = os.path.join(self.output_dir, pdf_candidates[0])
                    
                    debug_print(f"[DEBUG][_iso_v2] Saved debug PDF: {pdf_path}")
                    
                    # Convert PDF to PNG (final output in images directory)
                    png_filename = f"{self.base_name}_{sheet.title}_group_{shape_indices[0] if shape_indices else 0}.png"
                    png_path = os.path.join(self.images_dir, png_filename)
                    
                    # Remove old PNG if exists
                    if os.path.exists(png_path):
                        os.remove(png_path)
                    
                    # ImageMagick: Use -background white -flatten to prevent transparent/black areas
                    # -flatten composites all layers onto a white background
                    cmd = ['convert', '-density', str(dpi), f'{pdf_path}[0]', 
                           '-background', 'white', '-flatten',
                           '-colorspace', 'sRGB', '-quality', str(IMAGE_QUALITY), png_path]
                    debug_print(f"[DEBUG][_iso_v2] ImageMagick command: {' '.join(cmd)}")
                    proc = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
                    
                    if proc.returncode != 0 or not os.path.exists(png_path):
                        if getattr(self, 'verbose', False):
                            print(f"[WARN][_iso_v2] ImageMagick PNG conversion failed: {proc.stderr}")
                        return None
                    
                    debug_print(f"[DEBUG][_iso_v2] Successfully rendered group: {png_filename}")
                    debug_print(f"[DEBUG][_iso_v2] Debug files: {debug_xlsx_filename}, {debug_pdf_filename}")
                    
                    # Crop image to remove excess whitespace while preserving connectors
                    # Use tighter cropping for isolated groups (white_thresh=250 for more aggressive)
                    try:
                        # More aggressive cropping with higher white threshold
                        from PIL import Image
                        if os.path.exists(png_path):
                            im = Image.open(png_path)
                            bbox = self._find_content_bbox(im, white_thresh=250)
                            if bbox:
                                l, t, r, b = bbox
                                # Minimal padding for isolated groups (shapes already have proper margins)
                                pad = max(4, int(dpi / 300.0 * 6))  # Half of normal padding
                                l = max(0, l - pad)
                                t = max(0, t - pad)
                                r = min(im.width, r + pad)
                                b = min(im.height, b + pad)
                                cropped = im.crop((l, t, r, b))
                                cropped.save(png_path)
                                cropped.close()
                                debug_print(f"[DEBUG][_iso_v2] Cropped image: {im.size} → {cropped.size}")
                            im.close()
                    except Exception as crop_err:
                        if getattr(self, 'verbose', False):
                            print(f"[WARN][_iso_v2] Failed to crop image: {crop_err}")
                    
                    # Return tuple: (filename, minimum_row_for_cluster)
                    debug_print(f"[DEBUG][_iso_v2] Returning: filename={png_filename}, cluster_min_row={cluster_min_row}")
                    return (png_filename, cluster_min_row)
                    
                except Exception as e:
                    if getattr(self, 'verbose', False):
                        print(f"[ERROR][_iso_v2] Conversion failed: {e}")
                    return None
                
            finally:
                try:
                    if tmpdir and os.path.exists(tmpdir):
                        shutil.rmtree(tmpdir, ignore_errors=True)
                    if tmp_base and os.path.exists(tmp_base):
                        shutil.rmtree(tmp_base, ignore_errors=True)
                except (ValueError, TypeError) as e:
                    debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
        
        except Exception as e:
            print(f"[ERROR][_iso_v2] Exception: {e}")
            import traceback
            traceback.print_exc()
            return None

    def _render_sheet_isolated_group(self, sheet, shape_indices: List[int], dpi: int = 600, cell_range: Optional[Tuple[int,int,int,int]] = None) -> Optional[Tuple[str, int]]:
        """Render a group of shape indices as a single isolated image.
        
        **PRODUCTION METHOD - RECOMMENDED FOR ALL USE CASES**
        
        This implementation delegates to IsolatedGroupRenderer class for better code organization.
        The original monolithic implementation has been refactored into separate phases.
        
        Features:
        - Complete connector cosmetic processing
        - Theme color resolution (schemeClr -> srgbClr)
        - Line style materialization from lnRef references
        - Arrow head/tail preservation
        - Robust handling of complex flowcharts
        
        Creates a temporary workbook containing only the specified drawing anchors
        and renders them together as a single PNG image using LibreOffice -> PDF -> ImageMagick.
        This preserves the spatial relationships between shapes, making it ideal for
        flowcharts and composite diagrams.
        
        Returns:
            Optional[Tuple[str, int]]: (filename, start_row) or None on failure
        """
        renderer = IsolatedGroupRenderer(self)
        return renderer.render(sheet, shape_indices, dpi, cell_range)
    
    def _convert_sheet_data(self, sheet, data_range: Tuple[int, int, int, int]):
        """シートデータをテーブルとして変換（複数テーブル対応）"""
        min_row, max_row, min_col, max_col = data_range
        
        print(f"[INFO] データ範囲: 行{min_row}〜{max_row}, 列{min_col}〜{max_col}")
        
        # 罫線で囲まれた矩形領域のみを表として抽出
        print("[INFO] 罫線で囲まれた領域によるテーブル抽出を開始...")
        table_regions = self._detect_bordered_tables(sheet, min_row, max_row, min_col, max_col)
        debug_print(f"[DEBUG][_convert_sheet_data] bordered_table_regions_count={len(table_regions)} sample={table_regions[:5]}")

        # If no bordered tables found, or if bordered tables don't include the top rows (1-4),
        # attempt a broader table-region detection that uses heuristics (merged cells, annotations, column separations).
        # This ensures that header rows at the top of sheets are properly detected even when
        top_region_in_bordered = any(r[0] == 1 and r[1] <= 4 for r in table_regions)
        if not table_regions or not top_region_in_bordered:
            try:
                if not table_regions:
                    debug_print("[DEBUG] no bordered tables found; trying heuristic _detect_table_regions fallback")
                else:
                    debug_print(f"[DEBUG] bordered tables found but no top region (rows 1-4); trying heuristic _detect_table_regions to find header rows")
                heur_tables, heur_annotations = self._detect_table_regions(sheet, min_row, max_row, min_col, max_col)
                try:
                    debug_print(f"[TRACE][_detect_table_regions_result] sheet={sheet.title} heur_tables_count={len(heur_tables) if heur_tables else 0} heur_annotations_count={len(heur_annotations) if heur_annotations else 0}")
                    if heur_tables:
                        debug_print(f"[TRACE][_detect_table_regions_result_sample] {heur_tables[:10]}")
                    if heur_annotations:
                        debug_print(f"[TRACE][_detect_table_regions_annotations_sample] {heur_annotations[:10]}")
                except (ValueError, TypeError) as e:
                    debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
                if heur_tables:
                    if not table_regions:
                        debug_print(f"[DEBUG] heuristic detection found {len(heur_tables)} table regions")
                        table_regions = heur_tables
                    else:
                        top_heur_tables = [r for r in heur_tables if r[0] == 1 and r[1] <= 4]
                        if top_heur_tables:
                            debug_print(f"[DEBUG] adding {len(top_heur_tables)} top regions from heuristic detection to bordered tables")
                            table_regions = top_heur_tables + table_regions
            except Exception as _e:
                debug_print(f"[DEBUG] heuristic table detection failed: {_e}")

        # 変更: 描画（図形/画像）が占有するセル領域を検出し、重複する表領域は
        # テーブルとして出力せずプレーンテキスト扱いで出力する（ユーザ要望）。
        try:
            drawing_cell_ranges = self._extract_drawing_cell_ranges(sheet)
        except (ValueError, TypeError):
            drawing_cell_ranges = []

        def region_overlaps_drawings(region, drawing_ranges, overlap_threshold=0.25):
            """
            テーブル領域(region)が描画領域と重複するか確認する（重複割合ベース）。
            - region: (start_row, end_row, start_col, end_col)
            - drawing_ranges: list of (start_col, end_col, start_row, end_row)
            - overlap_threshold: テーブルセル総数に対する重複セル数の割合の閾値。デフォルト0.25

            戻り値: Trueなら重複とみなす（テーブルを除外）。
            """
            if not drawing_ranges:
                return False

            r1, r2, c1, c2 = region if len(region) == 4 else (region[0], region[1], region[2], region[3])

            table_cells = max(0, (r2 - r1 + 1)) * max(0, (c2 - c1 + 1))
            if table_cells <= 0:
                return False

            overlap_cells = 0
            for dr in drawing_ranges:
                try:
                    d_c1, d_c2, d_r1, d_r2 = dr
                except Exception:
                    continue
                inter_r1 = max(r1, d_r1)
                inter_r2 = min(r2, d_r2)
                inter_c1 = max(c1, d_c1)
                inter_c2 = min(c2, d_c2)
                if inter_r1 <= inter_r2 and inter_c1 <= inter_c2:
                    overlap_cells += (inter_r2 - inter_r1 + 1) * (inter_c2 - inter_c1 + 1)

            frac = overlap_cells / table_cells if table_cells > 0 else 0.0
            
            num_rows = r2 - r1 + 1
            num_cols = c2 - c1 + 1
            is_very_large_table = num_rows > 50 and num_cols > 30
            
            if is_very_large_table and len(drawing_ranges) >= 10:
                adjusted_threshold = 0.02
                debug_print(f"[DEBUG] Large table with many drawings detected - using stricter threshold {adjusted_threshold}")
            else:
                adjusted_threshold = overlap_threshold
            
            debug_print(f"[DEBUG] table_region={region} overlap_cells={overlap_cells} table_cells={table_cells} frac={frac:.3f} threshold={adjusted_threshold:.3f}")

            return frac >= adjusted_threshold

        # Split table_regions into those to keep and those to exclude due to overlap
        kept_table_regions = []
        excluded_table_regions = []
        for tr in table_regions:
            # _detect_bordered_tables returns (r1, r2, c1, c2)
            if region_overlaps_drawings(tr, drawing_cell_ranges):
                print(f"[INFO] テーブル領域が描画と重複しているため除外: {tr}")
                excluded_table_regions.append(tr)
            else:
                kept_table_regions.append(tr)

        filtered_table_regions = []
        for i, region_a in enumerate(kept_table_regions):
            r1_a, r2_a, c1_a, c2_a = region_a
            width_a = c2_a - c1_a + 1
            height_a = r2_a - r1_a + 1
            
            is_nested = False
            for j, region_b in enumerate(kept_table_regions):
                if i == j:
                    continue
                r1_b, r2_b, c1_b, c2_b = region_b
                width_b = c2_b - c1_b + 1
                height_b = r2_b - r1_b + 1
                
                if (r1_b <= r1_a and r2_a <= r2_b and 
                    c1_b <= c1_a and c2_a <= c2_b):
                    if width_a == 1 and width_b >= 2:
                        overlap_ratio = height_a / height_b if height_b > 0 else 0
                        if overlap_ratio > 0.8:
                            debug_print(f"[DEBUG] ネストされた1列テーブルを除外: {region_a} (含まれる先: {region_b}, 幅: {width_a} vs {width_b}, 重複率: {overlap_ratio:.2f})")
                            is_nested = True
                            break
            
            if not is_nested:
                filtered_table_regions.append(region_a)

        table_regions = filtered_table_regions
        debug_print(f"[DEBUG][_convert_sheet_data] kept_table_regions_count={len(table_regions)} kept_sample={table_regions[:5]}")

        # 重複テーブルを除外する処理
        def regions_overlap(r1, r2, threshold=0.5):
            """2つのテーブル領域が重複しているかチェック"""
            row1_start, row1_end, col1_start, col1_end = r1
            row2_start, row2_end, col2_start, col2_end = r2
            
            overlap_row_start = max(row1_start, row2_start)
            overlap_row_end = min(row1_end, row2_end)
            overlap_col_start = max(col1_start, col2_start)
            overlap_col_end = min(col1_end, col2_end)
            
            if overlap_row_start > overlap_row_end or overlap_col_start > overlap_col_end:
                return False
            
            overlap_cells = (overlap_row_end - overlap_row_start + 1) * (overlap_col_end - overlap_col_start + 1)
            
            r1_cells = (row1_end - row1_start + 1) * (col1_end - col1_start + 1)
            r2_cells = (row2_end - row2_start + 1) * (col2_end - col2_start + 1)
            smaller_cells = min(r1_cells, r2_cells)
            
            overlap_ratio = overlap_cells / smaller_cells if smaller_cells > 0 else 0
            
            return overlap_ratio >= threshold
        
        # 重複テーブルを除外（大きいテーブルを優先）
        deduplicated_regions = []
        for i, region in enumerate(table_regions):
            is_duplicate = False
            for j, other_region in enumerate(table_regions):
                if i != j and regions_overlap(region, other_region):
                    r1_cells = (region[1] - region[0] + 1) * (region[3] - region[2] + 1)
                    r2_cells = (other_region[1] - other_region[0] + 1) * (other_region[3] - other_region[2] + 1)
                    
                    if r1_cells < r2_cells:
                        debug_print(f"[DEBUG] 重複テーブルを除外: {region} (重複先: {other_region})")
                        is_duplicate = True
                        break
                    elif r1_cells == r2_cells and i > j:
                        debug_print(f"[DEBUG] 重複テーブルを除外: {region} (重複先: {other_region})")
                        is_duplicate = True
                        break
            
            if not is_duplicate:
                deduplicated_regions.append(region)
        
        table_regions = deduplicated_regions
        debug_print(f"[DEBUG][_convert_sheet_data] deduplicated_table_regions_count={len(table_regions)} deduplicated_sample={table_regions[:5]}")

        processed_rows = set()
        # Emit detected table regions as actual tables (not just reserve rows).
        # We convert each detected table region into markdown here, then mark
        # the rows as processed so subsequent plain-text collection skips them.
        table_index = 0
        for region in table_regions:
            debug_print(f"[DEBUG] emitting detected table region: {region}")
            try:
                # Convert the detected region to a markdown table. Use a
                # monotonically increasing table_index so filenames/ids are
                # deterministic across runs.
                self._convert_table_region(sheet, region, table_number=table_index)
                table_index += 1
            except Exception as _e:
                debug_print(f"[DEBUG] _convert_table_region failed for region={region}: {_e}")
            # Mark rows as processed regardless of conversion success so
            # they won't be re-collected as plain text.
            for r in range(region[0], region[1]+1):
                processed_rows.add(r)

        # If there are excluded table regions (overlapping drawings), collect their
        # row texts but defer actual emission until we merge with other plain text
        # so that final output preserves strict sheet row ordering.
        excluded_blocks = []  # list of (start_row, end_row, [(row, text), ...])
        excluded_end_rows = set()
        if 'excluded_table_regions' in locals() and excluded_table_regions:
            for excl in excluded_table_regions:
                try:
                    print(f"[INFO] 描画重複のためプレーンテキストとして収集: {excl}")
                    srow, erow, sc, ec = excl
                    lines = []
                    for rr in range(srow, erow + 1):
                        # collect row text
                        row_texts = []
                        for col_num in range(sc, ec + 1):
                            if rr <= sheet.max_row and col_num <= sheet.max_column:
                                cell_value = sheet.cell(row=rr, column=col_num).value
                                if cell_value is not None:
                                    text = str(cell_value).strip()
                                    if text:
                                        row_texts.append(text)
                        if row_texts:
                            lines.append((rr, " ".join(row_texts)))
                    if lines:
                        excluded_blocks.append((srow, erow, lines))
                        excluded_end_rows.add(erow)
                        # mark rows as processed so they are not re-collected later
                        # Only mark when we actually captured text within the excluded
                        # region. If no text was captured (e.g. relevant text is in
                        # columns outside the excluded columns), do not mark the
                        # rows so that they can still be discovered as plain text
                        # later.
                        for (rr, _) in lines:
                            processed_rows.add(rr)
                except Exception:
                    pass  # データ構造操作失敗は無視

        # プレーンテキスト領域を先に走査して収集する
        # 変更点: プレーン判定でTrueにならない場合でも、非空の行を"説明文"として出力するフォールバックを追加
        plain_texts = []  # list of (row_num, text)
        for row_num in range(min_row, max_row + 1):
            if row_num in processed_rows:
                continue
            region = (row_num, row_num, min_col, max_col)
            # 行のテキストを結合
            row_texts = []
            for col_num in range(min_col, max_col + 1):
                if row_num <= sheet.max_row and col_num <= sheet.max_column:
                    v = sheet.cell(row=row_num, column=col_num).value
                    if v is not None:
                        s = str(v).strip()
                        if s:
                            row_texts.append(s)
            if not row_texts:
                continue
            line = " ".join(row_texts)
            # プレーンテキスト判定を試みるが、判定がFalseでも説明的な単一カラムの行などは
            # ユーザ期待の文書テキストである可能性が高いためフォールバックで出力対象にする
            is_plain = self._is_plain_text_region(sheet, region)
            if is_plain or len(row_texts) == 1:
                plain_texts.append((row_num, line))
            else:
                # フォールバック: 非表形式として出力する候補に含める
                plain_texts.append((row_num, line))
            # NOTE: Do NOT mark the row as processed here. Previously we
            # added the row immediately which prevented the subsequent
            # implicit-table detection from seeing contiguous multi-column
            # runs (candidate_rows became empty). We defer marking until
            # after implicit-table detection and actual emission.

        # プレーンテキストはシート上の行順でそのまま出力する。
        # Collect plain_texts (rows not in processed_rows) and merge with
        # excluded_blocks collected earlier, then sort by row number so final
        # emission follows sheet top-to-bottom order.
        # plain_texts は (row_num, line) のリストなので行番号でソートしておく。
        # merge excluded block lines into plain_texts container
        merged_texts = []
        for (srow, erow, lines) in excluded_blocks:
            for (r, txt) in lines:
                merged_texts.append((r, txt, True if r == erow else False))

        plain_texts.sort(key=lambda x: x[0])
        for (r, line) in plain_texts:
            merged_texts.append((r, line, False))

        # Ensure that any rows already processed as detected tables are not
        # present in merged_texts. This avoids emitting the same content both
        # as a converted table and as free-form text during the final sort
        # and output stage.
        if processed_rows:
            try:
                before_count = len(merged_texts)
                merged_texts = [t for t in merged_texts if t[0] not in processed_rows]
                debug_print(f"[DEBUG] filtered merged_texts: removed {before_count - len(merged_texts)} rows that were already processed as tables")
            except (ValueError, TypeError) as e:
                debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")

        # Before emitting merged_texts, attempt to detect implicit tables formed
        # by contiguous rows that have multiple non-empty columns. This recovers
        # cases where table detection heuristics missed a table but the data is
        # clearly tabular (e.g. many rows with >=2 non-empty columns).
        try:
            candidate_rows = [r for r in range(min_row, max_row + 1) if r not in processed_rows]
            # map row -> list of non-empty column indices
            row_cols = {}
            for r in candidate_rows:
                cols = [c for c in range(min_col, max_col + 1) if r <= sheet.max_row and c <= sheet.max_column and sheet.cell(r, c).value is not None and str(sheet.cell(r, c).value).strip()]
                row_cols[r] = cols

            # find contiguous runs where each row has at least 2 non-empty cols
            runs = []
            cur_run = []
            for r in candidate_rows:
                if len(row_cols.get(r, [])) >= 2:
                    cur_run.append(r)
                else:
                    if cur_run:
                        runs.append((cur_run[0], cur_run[-1]))
                        cur_run = []
            if cur_run:
                runs.append((cur_run[0], cur_run[-1]))

            # emit any sufficiently long runs as tables (threshold=3 rows)
            for (srow, erow) in runs:
                if (erow - srow + 1) >= 3:
                    # compute min/max cols across the run
                    cols_used = [c for r in range(srow, erow + 1) for c in row_cols.get(r, [])]
                    if cols_used:
                        smin = min(cols_used)
                        smax = max(cols_used)
                        debug_print(f"[DEBUG] implicit table detected rows={srow}-{erow} cols={smin}-{smax}")
                        
                        if self._is_colon_separated_list(sheet, srow, erow, smin, smax):
                            debug_print(f"[DEBUG] implicit table is colon-separated list; skipping rows={srow}-{erow}")
                            continue
                        # Strong guard: if the run is a two-column numbered/list style
                        # (left column is enumeration markers like ①, 1., a) and right
                        # column is descriptive text, skip converting to an implicit table
                        try:
                            content_cols_set = set(cols_used)
                            if len(content_cols_set) == 2:
                                sorted_cols = sorted(list(content_cols_set))
                                lcol, rcol = sorted_cols[0], sorted_cols[1]
                                l_texts = []
                                r_texts = []
                                for rr in range(srow, erow + 1):
                                    try:
                                        lv = sheet.cell(rr, lcol).value
                                    except Exception:
                                        lv = None
                                    try:
                                        rv = sheet.cell(rr, rcol).value
                                    except Exception:
                                        rv = None
                                    if lv is not None and str(lv).strip():
                                        l_texts.append(str(lv).strip())
                                    if rv is not None and str(rv).strip():
                                        r_texts.append(str(rv).strip())

                                if l_texts and r_texts and len(l_texts) >= 2:
                                    import re
                                    circled = '①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳'
                                    num_matches = 0
                                    for t in l_texts:
                                        tt = t.strip()
                                        if any(ch in circled for ch in tt):
                                            num_matches += 1
                                        elif re.match(r'^[0-9]+[\.)]?$|^[A-Za-z]$|^[IVXivx]+$', tt):
                                            num_matches += 1
                                        elif len(tt) <= 2:
                                            num_matches += 1

                                    ratio = num_matches / len(l_texts) if l_texts else 0.0
                                    r_avg = sum(len(x) for x in r_texts) / len(r_texts) if r_texts else 0
                                    if ratio >= 0.8 and r_avg >= 8:
                                        debug_print(f"[DEBUG] implicit run looks like enumerated list; skipping table conversion rows={srow}-{erow} cols={lcol}-{rcol} left_ratio={ratio:.2f} right_avg={r_avg:.1f}")
                                        continue
                        except (ValueError, TypeError) as e:
                            debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")

                        # convert region as a table (this will append to markdown_lines)
                        try:
                            self._convert_table_region(sheet, (srow, erow, smin, smax), table_number=0)
                            # mark these rows as processed so they won't be emitted as plain text
                            for rr in range(srow, erow + 1):
                                processed_rows.add(rr)
                        except Exception:
                            pass  # データ構造操作失敗は無視
        except Exception:
            pass  # データ構造操作失敗は無視

        # sort merged_texts by row number (ascending) to preserve sheet order
        merged_texts.sort(key=lambda x: x[0])
        # If implicit-table detection above converted any rows, ensure merged_texts
        # no longer contains those rows. This double-check prevents duplicates when
        # implicit tables were found after merged_texts was initially constructed.
        if processed_rows:
            try:
                before_count2 = len(merged_texts)
                merged_texts = [t for t in merged_texts if t[0] not in processed_rows]
                debug_print(f"[DEBUG] post-implicit-filter: removed {before_count2 - len(merged_texts)} rows processed by implicit-table conversion")
            except (ValueError, TypeError) as e:
                debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
        # Emit merged free-form text entries in ascending row order.
        last_emitted_row = None
        if merged_texts:
            debug_print(f"[DEBUG] merged_texts出力開始: {len(merged_texts)}件")
            for (r, txt, is_excl_end) in merged_texts:
                debug_print(f"[DEBUG] merged_texts出力: 行{r}, text='{txt[:50]}...' (is_excl_end={is_excl_end})")
                self._emit_free_text(sheet, r, txt)
                # if this row is the end of an excluded block, append a blank line
                # and mark emitted rows ONLY during the canonical emission pass.
                if is_excl_end and getattr(self, '_in_canonical_emit', False):
                    self.markdown_lines.append("")
                    # map the end_row to the blank line index and mark emitted rows
                    try:
                        self._mark_sheet_map(sheet.title, r, len(self.markdown_lines) - 1)
                    except (ValueError, TypeError) as e:
                        debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
                    try:
                        # mark all rows in the corresponding excluded block as emitted
                        for (srow, erow, lines) in excluded_blocks:
                            if erow == r:
                                for rr in range(srow, erow + 1):
                                    self._mark_emitted_row(sheet.title, rr)
                                break
                    except Exception as e:
                        pass  # XML解析エラーは無視
                last_emitted_row = r
            # Add a separating blank line after any merged free-text region (only when actually emitting)
            if getattr(self, '_in_canonical_emit', False):
                self.markdown_lines.append("")
    
    def _detect_and_process_plain_text_regions(self, sheet, min_row: int, max_row: int, min_col: int, max_col: int, processed_rows: set = None) -> set:
        """プレーンテキスト領域を検出して処理し、処理済み行のセットを返す"""
        if processed_rows is None:
            processed_rows = set()
        for row_num in range(min_row, max_row + 1):
            if row_num in processed_rows:
                continue
            # この行がプレーンテキスト行かチェック
            region = (row_num, row_num, min_col, max_col)
            if self._is_plain_text_region(sheet, region):
                debug_print(f"[DEBUG] プレーンテキスト行を検出: 行{row_num}")
                # 連続するプレーンテキスト行を検索
                text_end_row = row_num
                for next_row in range(row_num + 1, max_row + 1):
                    next_region = (next_row, next_row, min_col, max_col)
                    if self._is_plain_text_region(sheet, next_region):
                        text_end_row = next_row
                    else:
                        break
                # プレーンテキスト領域を出力
                self._output_plain_text_region(sheet, row_num, text_end_row, min_col, max_col)
                # 処理済み行を記録
                for r in range(row_num, text_end_row + 1):
                    processed_rows.add(r)
        return processed_rows
    
    def _process_excluded_region_as_text(self, sheet, region: Tuple[int, int, int, int]):
        """フィルタで除外されたテーブル領域をプレーンテキストとして処理"""
        start_row, end_row, min_col, max_col = region
        
        for row_num in range(start_row, end_row + 1):
            # 行のテキストを収集
            row_texts = []
            for col_num in range(min_col, max_col + 1):
                if row_num <= sheet.max_row and col_num <= sheet.max_column:
                    cell_value = sheet.cell(row=row_num, column=col_num).value
                    if cell_value is not None:
                        text = str(cell_value).strip()
                        if text:
                            row_texts.append(text)
            
            # 行にテキストがある場合は出力
            if row_texts:
                line_text = " ".join(row_texts)
                self._emit_free_text(sheet, row_num, line_text)
        
        if end_row >= start_row:  # 何らかのテキストが処理された場合
            # Only append separator and mark emitted rows during canonical emission
            if getattr(self, '_in_canonical_emit', False):
                self.markdown_lines.append("")  # 空行を追加
                # map the end_row to the blank line index and mark emitted rows (helper already registered normalized texts)
                try:
                    self._mark_sheet_map(sheet.title, end_row, len(self.markdown_lines) - 1)
                except Exception as e:
                    pass  # XML解析エラーは無視
                try:
                    for r in range(start_row, end_row + 1):
                        self._mark_emitted_row(sheet.title, r)
                except Exception as e:
                    pass  # XML解析エラーは無視
            else:
                debug_print(f"[TRACE] Skipping authoritative mapping for excluded_region rows {start_row}-{end_row} (non-canonical)")
    
    def _output_plain_text_region(self, sheet, start_row: int, end_row: int, min_col: int, max_col: int):
        """プレーンテキスト領域をMarkdownに出力"""
        text_content = []

        for row_num in range(start_row, end_row + 1):
            row_text = []
            for col_num in range(min_col, max_col + 1):
                cell = sheet.cell(row=row_num, column=col_num)
                if cell.value is not None:
                    text = str(cell.value).strip()
            if text:
                row_text.append(text)

            if row_text:
                text_content.append(" ".join(row_text))

        # 空でないテキストのみ出力
        if text_content:
            # Do not create a mapping here if it does not exist; mapping is authoritative
            # and should only be populated during canonical emission via _mark_sheet_map.
            sheet_map = self._cell_to_md_index.get(sheet.title, {})
            for i, text in enumerate(text_content):
                if text.strip():
                    src_row = start_row + i
                    emitted = self._emit_free_text(sheet, src_row, text)
                    # if emitted is False, it was a duplicate and skipped
            # append blank line and map the last source row to the blank separator index
            # Only actually append and mark emitted rows during the canonical emission
            if getattr(self, '_in_canonical_emit', False):
                self.markdown_lines.append("")  # 空行を追加
                try:
                    self._mark_sheet_map(sheet.title, end_row, len(self.markdown_lines) - 1)
                except Exception as e:
                    pass  # XML解析エラーは無視
                # mark all emitted rows
                try:
                    for r in range(start_row, end_row + 1):
                        self._mark_emitted_row(sheet.title, r)
                except Exception as e:
                    pass  # XML解析エラーは無視
            debug_print(f"[DEBUG] プレーンテキスト出力: {len(text_content)}行")

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
        
        # Trace summary of detected regions
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
                        # generic heuristics for descriptive content: file paths, urls, xml, very long text
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
        # Debug: basic sheet metrics
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

        # Postprocess: merge adjacent single-row regions that have identical non-empty column masks
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

        # Guard: ヘッダー数が3でない場合は最適化をスキップ
        # (正規化後のヘッダー数で判定、元の列数ではない)
        if len(headers) != 3:
            debug_print(f"[DEBUG] 2列最適化スキップ（ヘッダー数が3ではない: {len(headers)}列）")
            return None
        
        # header_positionsが3つ以上必要
        if len(header_positions) < 3:
            debug_print(f"[DEBUG] 2列最適化スキップ（ヘッダー位置が3未満: {len(header_positions)}位置）")
            return None
            
        # 3列で、1列目が冗長な場合を検出
        # 第1列と第2列の組み合わせが設定項目のパターンかチェック(名前|初期値のパターン)
        debug_print(f"[DEBUG] 3列テーブル検出、パターンチェック: '{headers[0]}' と '{headers[1]}' (列{header_positions[0]}, 列{header_positions[1]})")
        # Use column-range based check (inspect the data under header columns) instead
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
                # fallthrough: if not enough rows, try the original fallback below

            # original fallback: use headers[1] and headers[2]
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
            # if value column distinct is low relative to nonempty, it's likely flag-like
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
                    # Markdown強調や太字はタイトル候補として扱う（特定キーワードには依存しない）
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
            # dump table_data for debugging before output
            try:
                cols = max(len(r) for r in table_data) if table_data else 0
            except Exception:
                cols = 0
            debug_print(f"[DEBUG] _output_markdown_table called (single_table path): rows={len(table_data)}, max_cols={cols}")
            for i, r in enumerate(table_data[:10]):
                debug_print(f"[DEBUG] table_data row {i} cols={len(r)}: {r}")
            # build source_rows sequentially from min_row..max_row assumption
                try:
                    source_rows = list(range(min_row, max_row + 1))[:len(table_data)]
                except (ValueError, TypeError):
                    source_rows = None
                # prune rows already emitted earlier in the sheet (pre-data rows)
                try:
                    debug_print(f"[DEBUG][_prune_call_single] sheet={sheet.title} before_prune rows={len(table_data) if table_data else 0} source_rows_sample={source_rows[:10] if source_rows else None}")
                    table_data, source_rows = self._prune_emitted_rows(sheet.title, table_data, source_rows)
                    debug_print(f"[DEBUG][_prune_result_single] sheet={sheet.title} after_prune rows={len(table_data) if table_data else 0} source_rows_sample={source_rows[:10] if source_rows else None}")
                except (ValueError, TypeError) as e:
                    debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
                # Pre-output deterministic dump for debugging: capture small preview of table_data and source_rows
                try:
                    src_sample = source_rows[:10] if source_rows else None
                    rows_len = len(table_data) if table_data else 0
                    debug_print(f"[DEBUG][_pre_output_call] path=single_table sheet={sheet.title} rows={rows_len} source_rows_sample={src_sample}")
                except (ValueError, TypeError) as e:
                    debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
                # Defer table emission to canonical pass so authoritative mappings
                # are recorded only during that pass. Use the first source row as
                # the anchor. Include optional metadata (no title available in
                # this path) for backwards-compatible shape.
                try:
                    anchor = (source_rows[0] if source_rows else min_row)
                except (ValueError, TypeError):
                    anchor = min_row
                try:
                    meta = None
                    self._sheet_deferred_tables.setdefault(sheet.title, []).append((anchor, table_data, source_rows, meta))
                    debug_print(f"DEFER_TABLE single_table sheet={sheet.title} anchor={anchor} rows={len(table_data)}")
                except (ValueError, TypeError):
                    # On any failure, fallback to immediate output to avoid data loss
                    self._output_markdown_table(table_data, source_rows=source_rows, sheet_title=sheet.title)
    
    def _convert_table_region(self, sheet, region: Tuple[int, int, int, int], table_number: int):
        """指定された領域をテーブルとして変換（結合セル対応、ヘッダー行検出）"""
        start_row, end_row, start_col, end_col = region
        # Diagnostic entry log: print region and a small sample of raw cell values
        try:
            debug_print(f"[DEBUG][_convert_table_region_entry] sheet={getattr(sheet, 'title', None)} region={start_row}-{end_row},{start_col}-{end_col}")
            # Dump up to 5 rows of raw values to help identify whether a table was detected
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
            title_text = self._find_table_title_in_region(sheet, region)
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
        title_text = self._find_table_title_in_region(sheet, region)
        
        # If we detected a title text for this region, keep it locally and
        # attach it to the deferred table metadata later when the table_data
        # is constructed. This avoids emitting the title as a separate
        # deferred text entry (in _sheet_deferred_texts) which complicates
        # duplicate-suppression and ordering.
        safe_title = None
        if title_text:
            safe_title = self._escape_angle_brackets(str(title_text))

        # If the detected title is actually part of the region (appears in the top row),
        # skip that row when building the table so the title isn't misinterpreted as a header cell.
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
        # table_data may be assigned only in some conditional branches below; ensure it's defined
        # to avoid UnboundLocalError when later code checks `if table_data:`
        table_data = None

        # フォールバック: ヘッダー行が無い場合、領域内の非空セルが存在する列の集合を使って
        # テーブルを組み立てる（最大8列）。これにより、見た目上複数列に分かれている表を復元する。
        if not header_row:
            unique_cols = [c for c in range(start_col, end_col + 1)
                           if any((sheet.cell(r, c).value is not None and str(sheet.cell(r, c).value).strip())
                                  for r in range(start_row, end_row + 1))]
            # 限度: 2〜8列のときのみ適用
            if 1 < len(unique_cols) <= 8:
                # Heuristic: if the left-most unique col is a repeated section label (same value for many rows)
                # and the next column contains varying property names, drop the left-most column so that
                # property column is used as the first data column. This prevents the first output column
                # from being filled with a section name like 'TransferFileList'.
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
                        # if left has a single repeated non-empty value and right has multiple distinct non-empty values,
                        # and left is present in fewer than 95% of rows (to avoid eliminating true data columns), drop left
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
                        # first_row_vals will be treated as header row, but sometimes it contains
                        # empty entries (e.g. ['', '']) that should be merged into the left
                        # non-empty header (like '名前'). Merge such empty-header columns
                        # into their left neighbour to avoid producing empty header columns.
                        headers_candidate = list(first_row_vals)

                        # determine columns to merge: if a header is empty and left header exists
                        merge_into_left = set()
                        for idx in range(1, len(headers_candidate)):
                            if not headers_candidate[idx] and headers_candidate[idx-1]:
                                merge_into_left.add(idx)

                        if merge_into_left:
                            # build mapping from original unique_cols indices to new columns
                            new_unique_cols = []
                            merge_map = {}  # col_idx -> target_new_index
                            new_idx = 0
                            for idx, col in enumerate(unique_cols):
                                if idx in merge_into_left:
                                    # merge this column into previous new_idx-1
                                    merge_map[idx] = new_idx - 1
                                else:
                                    new_unique_cols.append(col)
                                    merge_map[idx] = new_idx
                                    new_idx += 1

                            # build header row for new columns by merging text from merged columns
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

                            # replace unique_cols and header row with merged versions
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
                # safe dump: some variables may not exist in this scope (like header_positions etc.)
                try:
                    ctx = {}
                    ctx['unique_cols'] = unique_cols
                    ctx['table_data_rows'] = len(table_data)
                    # header_positions/final_groups/compressed_headers may not be defined here
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
                # write compact machine-friendly trace to file (if debug log available)
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
                        # dump first_row_vals if present
                        if 'first_row_vals' in locals():
                            debug_print(f"[DEBUG-TRACE] first_row_vals={first_row_vals}")
                        if 'merge_into_left' in locals():
                            debug_print(f"[DEBUG-TRACE] merge_into_left={merge_into_left}")
                        if 'merge_map' in locals():
                            debug_print(f"[DEBUG-TRACE] merge_map={merge_map}")
                        # header-related structures if present
                        for name in ('header_positions', 'final_groups', 'compressed_headers', 'group_positions'):
                            if name in locals():
                                debug_print(f"[DEBUG-TRACE] {name}={locals()[name]}")
                        # dump a few raw cell values for the region to cross-check
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
                # dump table_data shape and first rows for debugging
                try:
                    cols = max(len(r) for r in table_data) if table_data else 0
                except (ValueError, TypeError):
                    cols = 0
                debug_print(f"[DEBUG] _output_markdown_table called (unique_cols path): rows={len(table_data)}, max_cols={cols}")
                for i, r in enumerate(table_data[:10]):
                    debug_print(f"[DEBUG] table_data row {i} cols={len(r)}: {r}")
                try:
                    # prune pre-emitted rows that may duplicate earlier lines
                    debug_print(f"[DEBUG][_prune_call_unique] sheet={sheet.title} before_prune rows={len(table_data) if table_data else 0} source_rows_sample={source_rows[:10] if source_rows else None}")
                    table_data, source_rows = self._prune_emitted_rows(sheet.title, table_data, source_rows)
                    debug_print(f"[DEBUG][_prune_result_unique] sheet={sheet.title} after_prune rows={len(table_data) if table_data else 0} source_rows_sample={source_rows[:10] if source_rows else None}")
                except (ValueError, TypeError) as e:
                    debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
                # Pre-output deterministic dump for debugging (unique_cols path)
                try:
                    src_sample = source_rows[:10] if source_rows else None
                    rows_len = len(table_data) if table_data else 0
                    debug_print(f"[DEBUG][_pre_output_call] path=unique_cols sheet={getattr(sheet, 'title', None)} rows={rows_len} source_rows_sample={src_sample}")
                except (ValueError, TypeError) as e:
                    debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
                try:
                    # Defer table emission to canonical pass. Use first source row
                    # as anchor when available and include no title meta here.
                    try:
                        anchor = (source_rows[0] if source_rows else start_row)
                    except (ValueError, TypeError):
                        anchor = start_row
                    try:
                        meta = None
                        self._sheet_deferred_tables.setdefault(sheet.title, []).append((anchor, table_data, source_rows, meta))
                        debug_print(f"DEFER_TABLE unique_cols sheet={sheet.title} anchor={anchor} rows={len(table_data)}")
                    except (ValueError, TypeError):
                        # fallback to immediate output on any failure to avoid data loss
                        try:
                            self._output_markdown_table(table_data, source_rows=source_rows)
                        except (ValueError, TypeError):
                            self._output_markdown_table(table_data)
                except (ValueError, TypeError):
                    # outer try - if anything else fails, try direct output
                    try:
                        self._output_markdown_table(table_data)
                    except Exception as e:
                        pass  # XML解析エラーは無視
                return

        # テーブルデータを結合セル考慮で構築
        if header_row:
            # ヘッダー行を考慮した構築
            table_data = self._build_table_with_header_row(sheet, region, header_row, merged_cells, header_height=header_height)
            # ヘッダー行から開始するため、approx_rowsもheader_rowから計算
            actual_start_row = header_row
        else:
            # 従来の方法
            table_data = self._build_table_data_with_merges(sheet, region, merged_cells)
            actual_start_row = start_row
        
        if table_data:
            debug_print(f"[DEBUG] 出力前テーブルプレビュー: rows={len(table_data)}, first_row={table_data[0] if table_data else None}")
            # dump table_data shape before output
            try:
                cols = max(len(r) for r in table_data) if table_data else 0
            except (ValueError, TypeError):
                cols = 0
            debug_print(f"[DEBUG] _output_markdown_table called (header/data path): rows={len(table_data)}, max_cols={cols}")
            for i, r in enumerate(table_data[:10]):
                debug_print(f"[DEBUG] table_data row {i} cols={len(r)}: {r}")
            try:
                # actual_start_rowから開始（header_rowまたはstart_row）
                # regionのend_rowを使用して、実際のテーブル範囲全体をカバー
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
            # Pre-output deterministic dump for debugging (header/data path)
            try:
                src_sample = approx_rows[:10] if approx_rows else None
                rows_len = len(table_data) if table_data else 0
                debug_print(f"[DEBUG][_pre_output_call] path=header_data sheet={sheet.title} rows={rows_len} source_rows_sample={src_sample}")
            except (ValueError, TypeError) as e:
                debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
            # Defer table emission until canonical pass so authoritative maps are
            # recorded only during that pass. Store anchor row = first source row
            try:
                # Prefer the detected title row as the table anchor when available.
                title_anchor = getattr(self, '_last_table_title_row', None) if safe_title else None
                if title_anchor and isinstance(title_anchor, int):
                    anchor = title_anchor
                else:
                    anchor = (approx_rows[0] if approx_rows else start_row)
            except Exception:
                anchor = start_row
            try:
                # Include optional metadata (title) with the deferred table so
                # the canonical emitter can output the title together with the
                # table in a single, atomic event. Backwards-compatible shape
                # for deferred tables: (anchor, table_data, approx_rows) ->
                # (anchor, table_data, approx_rows, meta_dict)
                meta = {'title': safe_title} if safe_title else None
                self._sheet_deferred_tables.setdefault(sheet.title, []).append((anchor, table_data, approx_rows, meta))
                # clear transient title row after deferring
                try:
                    self._last_table_title_row = None
                except Exception as e:
                    pass  # XML解析エラーは無視
                debug_print(f"DEFER_TABLE sheet={sheet.title} anchor={anchor} rows={len(table_data)} title_present={bool(safe_title)}")
            except (ValueError, TypeError):
                # fallback to immediate output if deferral fails
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
                # emit via centralized emitter so duplicates and emitted-rows are tracked
                for text in right_texts:
                    try:
                        self._emit_free_text(sheet, row_num, text)
                    except (ValueError, TypeError):
                        # fallback to direct append if emitter fails for some reason
                        self.markdown_lines.append(f"{text}  ")
        # テーブル右隣のテキストがあれば空行で区切る
        if any(sheet.cell(row=row_num, column=col_num).value for row_num in range(start_row, end_row + 1) for col_num in range(end_col + 1, max_col + 1)):
            self.markdown_lines.append("")
    
    def _is_plain_text_region(self, sheet, region: Tuple[int, int, int, int]) -> bool:
        """領域が通常のテキスト（非表形式）かどうかを判定"""
        start_row, end_row, start_col, end_col = region
        # Early debug: report entry and simple metrics
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
        # Debug: report computed heuristics for this region
        debug_print(f"PLAIN_ENTRY sheet={getattr(sheet,'title',None)} region={start_row}-{end_row},{start_col}-{end_col} non_empty={non_empty_cells} total={total_cells}")
        text_content = ' '.join(texts)
        
        avg_len = sum(len(t) for t in texts) / non_empty_cells if non_empty_cells > 0 else 0
        
        # token-based heuristic: a single row containing multiple short tokens
        # is likely a compact table header or data row (e.g. "名前 初期値 設定値").
        # Be conservative: require the average cell length to be not too large so
        # we don't misclassify descriptive sentences as tables.
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

        # Exception: two-column numbered-list pattern -> treat as plain text
        # If left column mostly contains numbering/markers (①, 1, A, i, etc.) and
        # right column contains longer descriptive text, prefer treating the
        # region as a numbered list / descriptive lines rather than a table.
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
                    # include common circled numbers explicitly (①〜⑳)
                    circled = '①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳'
                    for t in left_texts:
                        tt = t.strip()
                        # Normalize fullwidth digits/punctuation to their ASCII equivalents
                        try:
                            nn = unicodedata.normalize('NFKC', tt)
                        except Exception:
                            nn = tt

                        # Check circled numbers first (they don't normalize to ascii)
                        if any(ch in circled for ch in tt):
                            num_matches += 1
                            continue

                        # Accept patterns like:
                        #  - (1) / （1） / 1) / 1）
                        #  - 1. / 1．
                        #  - 1 / １ (fullwidth normalized by NFKC)
                        #  - (a) / a)
                        #  - roman numerals I, II, III optionally with punctuation
                        # Use normalized string for regex so fullwidth punctuation is handled
                        # Allow optional surrounding parentheses (both ASCII and fullwidth)
                        # and optional trailing punctuation like '.' or '．'
                        try:
                            if re.match(r'^[\(\（]?\s*(?:\d+|[IVXivx]+|[A-Za-z])\s*[\)\）]?[\.．]?$', nn):
                                num_matches += 1
                                continue
                        except Exception:
                            pass  # データ構造操作失敗は無視

                        # fallback: single-character markers (e.g. '-', 'a', '1')
                        try:
                            if len(nn.strip()) == 1 and re.match(r'^[A-Za-z0-9\-]$', nn.strip()):
                                num_matches += 1
                                continue
                        except Exception:
                            pass  # データ構造操作失敗は無視

                    ratio = (num_matches / len(left_texts)) if left_texts else 0.0
                    right_avg = sum(len(s) for s in right_texts) / len(right_texts) if right_texts else 0
                    # Heuristic thresholds: >=80% left are numbering-like and right avg length >=10
                    if ratio >= 0.8 and right_avg >= 10:
                        debug_print(f"[DEBUG] 番号付きリスト検出: 行{start_row}〜{end_row} 左番号率={num_matches}/{len(left_texts)} 右平均長={right_avg:.1f}")
                        return True
        except (ValueError, TypeError) as e:
            debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")

        # ルール1: ファイルパス/URL/XMLが多い場合はプレーンテキスト（説明的な列）
        if non_empty_cells > 0 and (path_like_count / non_empty_cells) > 0.25:
            # If there are strong vertical borders that indicate multiple columns, prefer table interpretation
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

        # Exception: single-row with multiple short columns likely represents a compact table
        # e.g. a single row of short labels like 'A  B  C' should be treated as a table
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
        
        # Emit each source row as a single combined line using the centralized emitter
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
                    # fallback to direct append if emitter fails
                    # fallback to direct append if emitter fails
                    # Do NOT mutate authoritative mappings unless we're in the
                    # canonical emission pass. Prematurely marking rows/texts as
                    # emitted caused pruning of legitimate table rows.
                    try:
                        self.markdown_lines.append(self._escape_angle_brackets(combined) + "  ")
                        if getattr(self, '_in_canonical_emit', False):
                            try:
                                # Only record authoritative mappings during canonical pass
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
                            # non-canonical context: canonical pass will assign indices
                            debug_print(f"[TRACE] Skipping authoritative mapping for plain-text fallback row={row_num} (non-canonical)")
                    except (ValueError, TypeError) as e:
                        debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")

        # add a separating blank line if any lines were emitted
        try:
            emitted = self._sheet_emitted_rows.get(sheet.title, set())
            any_emitted = any(r in emitted for r in range(start_row, end_row + 1))
        except (ValueError, TypeError):
            any_emitted = True
        if any_emitted:
            self.markdown_lines.append("")  # セクション区切りの空行を追加
    
    def _build_table_with_header_row(self, sheet, region: Tuple[int, int, int, int], 
                                   header_row: int, merged_info: Dict[str, Any], header_height: int = 1) -> List[List[str]]:
        """ヘッダー行を基にテーブルを正しく構築
        
        Args:
            header_height: ヘッダーの高さ（行数）。_find_table_header_rowから渡される
        """
        start_row, end_row, start_col, end_col = region
        
        debug_print(f"[DEBUG] ヘッダー行{header_row}でテーブルを構築中...")
        
        # ヘッダー行の実際の行・列範囲を確認し、regionを拡張
        # (header_rowがregion外の場合や、「名前」など範囲外のヘッダーを含めるため)
        actual_start_row = min(start_row, header_row)
        actual_end_row = max(end_row, header_row + header_height - 1)
        
        header_min_col = start_col
        header_max_col = end_col
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

                # normalize newlines to <br> and collapse/trim redundant <br> tokens
                try:
                    import re as _re
                    text = raw_text.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '<br>')
                    # collapse multiple consecutive <br> into one
                    text = _re.sub(r'(<br>\s*){2,}', '<br>', text)
                    # strip any leading/trailing <br>
                    text = _re.sub(r'^(?:<br>\s*)+', '', text)
                    text = _re.sub(r'(?:\s*<br>)+$', '', text)
                    text = text.strip()
                except Exception:
                    text = raw_text.replace('\n', '<br>').strip() if raw_text else ''

                if text and len(text) <= 120:
                    if not parts or parts[-1] != text:
                        parts.append(text)

            # remove consecutive duplicate parts to avoid repeated concatenation
            dedup_parts = []
            for p in parts:
                if not dedup_parts or dedup_parts[-1] != p:
                    dedup_parts.append(p)
            # filter out parts that likely belong to data rows (appear frequently below header)
            try:
                # Determine if this column contains any merged/master header cells within the header rows.
                is_master_col = False
                for r in range(header_row, min(header_row + header_height, end_row + 1)):
                    keym = f"{r}_{col}"
                    if keym in merged_info:
                        mi = merged_info[keym]
                        # If the master cell is within the header candidate rows and/or spans multiple rows/cols,
                        # treat this column as a master/header column and avoid aggressive data-token removal.
                        mr = int(mi.get('master_row', header_row))
                        mc = int(mi.get('master_col', col))
                        span_r = int(mi.get('span_rows', 1) or 1)
                        span_c = int(mi.get('span_cols', 1) or 1)
                        if (header_row <= mr < header_row + header_height) or span_r > 1 or span_c > 1 or (mc != col):
                            is_master_col = True
                            break

                filtered_parts = []
                # If this column is a master/header column, or if we have a multi-row header (height>1),
                # skip the sampling-based drop and keep dedup_parts intact
                if is_master_col or header_height > 1:
                    filtered_parts = list(dedup_parts)
                else:
                    for p in dedup_parts:
                        # keep very short tokens (e.g. '-', single char markers)
                        if not p or len(p.strip()) <= 2:
                            filtered_parts.append(p)
                            continue

                        # count occurrences of this token in rows below the header area
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
                            # consider exact match or contained match as evidence
                            if vv == p or vv == p.replace('<br>', '\n') or vv in p or p in vv:
                                cnt += 1

                        frac = (cnt / total) if total > 0 else 0.0
                        # Conservative rule: only drop tokens that are both relatively long and
                        # appear frequently in data rows. This avoids removing short header labels
                        # like '装置名' or 'PSコード' that may also reappear in data samples.
                        drop = False
                        try:
                            plen = len(p.strip()) if p else 0
                            # Strong evidence: very frequent (>=90%) -> drop only for tokens of reasonable length
                            # (avoid dropping short label-like tokens such as '装置名')
                            if frac >= 0.9 and plen >= 4:
                                drop = True
                            # Moderate evidence: frequent (>=60%) and token not very short -> drop
                            elif frac >= 0.6 and plen >= 8:
                                drop = True
                        except Exception:
                            drop = False

                        if drop:
                            debug_print(f"[DEBUG] ヘッダからデータトークン除外: '{p}' at 列{col} (occurrence_fraction={frac:.2f}, len={plen})")
                            continue
                        filtered_parts.append(p)

                combined = '<br>'.join(filtered_parts) if filtered_parts else ''
                # additionally remove repeated subparts while preserving order to avoid
                # patterns like 'A<br>B<br>A<br>B<br>A<br>B' appearing due to multi-row joins
                try:
                    if combined:
                        subs = [s.strip() for s in combined.split('<br>') if s.strip()]
                        # first remove consecutive duplicates while preserving order
                        seen = set()
                        uniq = []
                        for s in subs:
                            if not uniq or uniq[-1] != s:
                                uniq.append(s)
                        # then collapse perfect repeated sequences like [A,B,A,B,A,B] -> [A,B]
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

        # Fallback: if the detected headers are mostly empty, the real header
        # content may be shifted by one row (common when a title row was skipped).
        # Try a single one-line downward shift and re-extract simple header texts
        # conservatively. This is a small, low-risk heuristic to avoid losing
        # columns when header tokens actually appear on the next row.
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
                    # adopt shifted headers conservatively: only replace when we found
                    # any non-empty header tokens on the shifted row(s)
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
        groups = []  # list of (start_idx, end_idx) exclusive-end (based on headers index)
        i = 0
        while i < len(normalized_headers):
            j = i + 1
            while j < len(normalized_headers) and normalized_headers[j] == normalized_headers[i]:
                j += 1
            groups.append((i, j))
            i = j

        # expand groups to respect explicit borders: if columns within a group are separated by vertical borders in the header_row,
        # split that group into single-column groups so they won't be compressed
        final_groups = []
        for (a, b) in groups:
            if b - a <= 1:
                final_groups.append((a, b))
                continue

            # header_positions indices for this group
            cols = [header_positions[k] for k in range(a, b) if k < len(header_positions)]
            # if any adjacent column boundary in the sheet has a right border on header_row, do not compress across it
            split_points = [a]
            for idx in range(len(cols) - 1):
                col_left = cols[idx]
                # Check vertical border presence across the header->data rows, not only the header row.
                right_count = 0
                total_check = 0
                for rr in range(header_row, end_row + 1):
                    try:
                        cell_l = sheet.cell(rr, col_left)
                        total_check += 1
                        if cell_l.border and cell_l.border.right and cell_l.border.right.style:
                            right_count += 1
                    except Exception:
                        # ignore and continue
                        pass

                has_strong_right = (total_check > 0 and (right_count / total_check) >= 0.5)

                # Additional strict check: if any right-border exists on the header row itself,
                # treat that as a definitive column separator and force a split. This prevents
                # collapsing identical header labels when explicit vertical borders separate columns.
                try:
                    hdr_cell = sheet.cell(header_row, col_left)
                    if hdr_cell and hdr_cell.border and hdr_cell.border.right and getattr(hdr_cell.border.right, 'style', None):
                        has_strong_right = True
                except Exception as e:
                    pass  # XML解析エラーは無視

                # Also check merged-cell masters for differences across header rows
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
                    # force split between this and next
                    split_points.append(a + idx + 1)
                elif masters_differ:
                    # also force split when masters differ across header rows
                    split_points.append(a + idx + 1)

            split_points.append(b)
            # build ranges from split_points
            for si in range(len(split_points) - 1):
                final_groups.append((split_points[si], split_points[si+1]))

        # Post-process final_groups: if a group's header fragments are all empty, do not compress that group.
        processed_groups = []
        for (a, b) in final_groups:
            # build sample header fragments across the group's header positions
            try:
                fragments = []
                for idx in range(a, b):
                    if idx < len(headers):
                        fragments.append(str(headers[idx] or '').strip())
                    else:
                        fragments.append('')
                # determine if all fragments are empty (i.e., effectively no header label)
                if all((not f) for f in fragments):
                    # expand to single-column groups to preserve all underlying columns
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
        group_column_ranges = []  # list of (col_start, col_end) inclusive
        for (a, b) in final_groups:
            if a < len(header_positions):
                col_start = header_positions[a]
            else:
                col_start = start_col
            if b < len(header_positions):
                col_end = header_positions[b] - 1
            else:
                col_end = end_col
            # normalize bounds
            if col_start < start_col:
                col_start = start_col
            if col_end > end_col:
                col_end = end_col
            if col_end < col_start:
                col_end = col_start
            group_column_ranges.append((col_start, col_end))
        debug_print(f"[DEBUG] group_column_ranges={group_column_ranges}")

        # Build a helper to get cell value considering merged cells
        def _get_cell_value(r, c):
            key = f"{r}_{c}"
            if key in merged_info and merged_info[key]['is_merged']:
                mi = merged_info[key]
                mc = sheet.cell(mi['master_row'], mi['master_col'])
                return self._format_cell_content(mc) if mc.value is not None else ''
            cell = sheet.cell(r, c)
            return self._format_cell_content(cell) if cell.value is not None else ''

        # For each group, compute column priority based on number of distinct non-empty values
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
                # compute dominance: frequency of most common value
                max_freq = 0
                if vals:
                    from collections import Counter
                    counts = Counter(vals)
                    max_freq = max(counts.values())
                dominance = (max_freq / nonempty) if nonempty > 0 else 0
                # score tuple: prefer more distinct values, then lower dominance (less dominated by single token),
                # then more non-empty cells, finally leftmost column
                col_scores.append((c, distinct, dominance, nonempty))
            # sort by: distinct desc, dominance asc, nonempty desc, col asc
            col_scores.sort(key=lambda x: (-x[1], x[2], -x[3], x[0]))
            ordered_cols = [c for (c, _, _, _) in col_scores]
            group_column_priority.append(ordered_cols)
        debug_print(f"[DEBUG] group_column_priority={group_column_priority}")

        # write compact group/priority info to debug file for offline analysis
        sheet_name = getattr(sheet, 'title', None)

        # データ行を構築（ヘッダー行の次から）。各グループ内では行ごとに優先列順で最初の非空セルを参照して値を取得する
        # header_heightを考慮してヘッダー行をスキップ
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
        
        # 2列最適化チェック（正規化後のヘッダーとgroup_positionsを使用）
        debug_print(f"[DEBUG] 2列最適化チェック開始: headers={compressed_headers}, positions={group_positions}")
        optimized_structure = self._optimize_table_for_two_columns(sheet, region, compressed_headers, group_positions)
        if optimized_structure:
            debug_print(f"[DEBUG] 2列最適化成功、テーブルサイズ: {len(optimized_structure)}行")
            return self._trim_edge_empty_columns(optimized_structure)
        else:
            debug_print(f"[DEBUG] 2列最適化スキップ")
        
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
                # prepare two regex patterns: a stricter primary one and a permissive fallback
                # primary: stricter split pattern but avoid explicit Unicode range checks.
                # accept a middle token that is non-whitespace and does not contain obvious path/XML chars
                primary_re = re.compile(r'^(.*?)\s+([^\\\/<>:\"\s]{1,60})\s+(.+)$')
                # permissive: allow many chars for middle token but exclude path/XML chars later
                # also accept Japanese quotes and fullwidth punctuation after normalizing
                permissive_re = re.compile(r'^(.*?)\s+([^\\\/<>:\\"]{1,60})\s+(.+)$')

                def _normalize_for_split(s: str) -> str:
                    # normalize full-width spaces/quotes/parentheses to improve matching
                    if not s:
                        return ''
                    s = s.replace('\u3000', ' ')
                    s = s.replace('\uFF08', '(').replace('\uFF09', ')')  # fullwidth parens
                    s = s.replace('（', '(').replace('）', ')')
                    s = s.replace('「', ' ').replace('」', ' ')
                    s = s.replace('”', '"').replace('“', '"')
                    # collapse multiple spaces
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
                                # 次に緩いパターンを試みるが、パスやXML等のトークンを含む場合は除外して誤検出を抑制
                                m2 = permissive_re.match(norm)
                                if m2:
                                    mid = m2.group(2)
                                    # exclude obvious path-like or xml-like tokens
                                    if ('\\' not in mid and '/' not in mid and '<' not in mid and '>' not in mid and ':' not in mid):
                                        matches += 1
                    # 非空行が一定数以上かつマッチ率が高ければ分割候補とする
                    ratio = (matches / non_empty) if non_empty > 0 else 0
                    col_details.append((col_idx, non_empty, matches, ratio))
                    # increase threshold to reduce false positives for splitting
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
                                # match against normalized form but preserve original pieces where possible
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
                                    # try to extract corresponding substrings from original cell loosely
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

    
    def _find_table_title_in_region(self, sheet, region: Tuple[int, int, int, int]) -> Optional[str]:
        """テーブル領域内からタイトルを検出（汎用版: 特定キーワードには依存しない）"""
        start_row, end_row, start_col, end_col = region

        # テーブル領域の前後でタイトルを探す（より広い範囲）
        search_start = max(1, start_row - 10)
        search_end = min(start_row + 5, end_row + 1)

        # 最適なタイトル候補を探す
        title_candidates = []

        for row in range(search_start, search_end):
            for col in range(max(1, start_col - 5), min(start_col + 15, end_col + 5)):
                cell = sheet.cell(row, col)
                if cell.value:
                    text = str(cell.value).strip()
                    # ファイルパスや長すぎる文字列はタイトル候補から除外
                    if any(x in text for x in ['\\', '/', 'xml']) or len(text) > 120:
                        continue

                    # 太字やMarkdown強調は優先的にタイトル候補とする
                    if cell.font and cell.font.bold:
                        distance = abs(row - start_row)
                        # mark as bold/high-priority
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
                # row_relation: 0: above table, 1: same row, 2: below table
                return (immediate_priority, kind_priority, row_relation, length_priority, distance)

            best_title = min(title_candidates, key=_title_key)
            # record the detected title row so callers can use it as an anchor
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

        # clear any previous title row if no title found
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

        # Evaluate single-row and multi-row header candidates (up to height 3).
        for row in candidate_rows:
            for height in (1, 2, 3):
                if row + height - 1 > end_row:
                    break

                header_values = []
                # build combined header text per column across `height` rows
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
                            # avoid duplicates when stacking rows
                            if not parts or parts[-1] != text:
                                parts.append(text)

                    combined = '<br>'.join(parts) if parts else ''
                    header_values.append(combined)

                # count groups
                group_count = 0
                prev = None
                nonempty = 0
                for v in header_values:
                    if v:
                        nonempty += 1
                    if v and v != prev:
                        group_count += 1
                    prev = v or prev

                # determine how many columns actually draw header fragments from multiple
                # physical rows. If only a small fraction (<25%) of columns have fragments
                # coming from multiple rows, it's likely the '<br>'s are internal to a single
                # cell rather than a true multi-row header; skip multi-row candidate.
                try:
                    multirow_cols = 0
                    total_columns = min(start_col + 20, end_col + 1) - start_col
                    # contributors_per_col recorded alongside header_values where available
                    # rebuild lightweight contributors detection for this candidate
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
                # multirow_fracが低くても、全体として多くの非空セルがあれば有効なヘッダーとする
                # (行3-4のような「上段と下段で異なる列にテキストがある」構造に対応)
                if height > 1 and multirow_frac < 0.25:
                    # nonemptyが多ければ（全体の50%以上）、有効な複数行ヘッダーとして扱う
                    if nonempty >= total_columns * 0.5:
                        debug_print(f"[DEBUG] 複数行ヘッダ候補を維持（非空セルが多い）: row={row}, height={height}, multirow_frac={multirow_frac:.2f}, nonempty={nonempty}/{total_columns}")
                    else:
                        debug_print(f"[DEBUG] 複数行ヘッダ候補をスキップ（実際には単一セル内改行が多い）: row={row}, height={height}, multirow_frac={multirow_frac:.2f}, nonempty={nonempty}/{total_columns}")
                        continue

                # compute fraction of columns in the top row that are part of multi-column merged masters
                top_row = row
                top_merged_count = 0
                total_columns = min(start_col + 20, end_col + 1) - start_col
                for col in range(start_col, start_col + total_columns):
                    key = f"{top_row}_{col}"
                    if key in merged_info and merged_info[key].get('span_cols', 1) > 1:
                        top_merged_count += 1
                top_merged_fraction = (top_merged_count / total_columns) if total_columns > 0 else 0

                # debug
                debug_print(f"[DEBUG] 行{row}..{row+height-1} combined header_values (first16): {header_values[:16]}")
                debug_print(f"[DEBUG] 行{row} height={height} group_count={group_count}, nonempty_cols={nonempty}")

                # prefer larger group_count; tie-breaker: prefer full-column coverage, then stronger bottom-border alignment,
                # then fewer top-row multi-column merges (favor lower-level splits), then deeper bottom row, then larger height
                bottom_row = row + height - 1
                total_columns = min(start_col + 20, end_col + 1) - start_col
                full_coverage = (nonempty == total_columns)

                # compute header bottom-border alignment: fraction of columns where bottom_row has a bottom border or next row has top border
                border_hits = 0
                border_total = 0
                for c in range(start_col, start_col + total_columns):
                    try:
                        br_cell = sheet.cell(bottom_row, c)
                        border_total += 1
                        if (br_cell.border and br_cell.border.bottom and getattr(br_cell.border.bottom, 'style', None)):
                            border_hits += 1
                        else:
                            # check next row's top border when available
                            if bottom_row + 1 <= end_row:
                                nx = sheet.cell(bottom_row + 1, c)
                                if nx.border and nx.border.top and getattr(nx.border.top, 'style', None):
                                    border_hits += 1
                    except Exception:
                        continue
                header_border_fraction = (border_hits / border_total) if border_total > 0 else 0.0

                # compute merged-master alignment score: fraction of header columns whose merged master is located within the header rows
                master_aligned = 0
                master_total = 0
                for c in range(start_col, start_col + total_columns):
                    key = f"{row}_{c}"
                    mi = merged_info.get(key)
                    if mi:
                        master_total += 1
                        mr = int(mi.get('master_row', row))
                        # aligned if master_row is within the header candidate rows
                        if row <= mr <= bottom_row:
                            master_aligned += 1
                masters_alignment_frac = (master_aligned / master_total) if master_total > 0 else 0.0

                # compute a lightweight "header-likeness" score for the bottom row of this candidate
                # higher score favors rows that look like labels (short tokens, few <br>, not path-like, bold)
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
                        # weights chosen conservatively to avoid overfitting
                        score = short_frac + 0.45 * nobr_frac + 0.35 * bold_frac - 0.9 * path_frac
                        return max(0.0, score)
                    except Exception:
                        return 0.0

                likeness_score = _row_header_likeness(bottom_row)
                # debug print for inspection
                debug_print(f"[DEBUG] header-likeness(bottom_row={bottom_row})={likeness_score:.3f}")
                debug_print(f"[DEBUG] header_border_fraction(bottom_row={bottom_row})={header_border_fraction:.3f}")

                # 拡張列範囲でのグループ数を計算（テーブル範囲外のヘッダーも考慮）
                # height>1の場合は全行を走査して結合した値でグループをカウント
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

                # build metric tuple for this candidate
                # 罫線を最優先、次に拡張グループ数を考慮
                # 同等の場合は上部の行を優先（-rowで小さい行番号が大きい値になる）
                # heightは小さい方を優先（より保守的なヘッダー検出）
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


                # compare with best metrics
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
                        # compare lexicographically
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
                            # fallback to previous tie-breaker
                            if group_count > best_group_count:
                                best_group_count = group_count
                                best_row = row
                                best_height = height

        if best_row:
            # store detected header start and height for downstream heuristics
            self._detected_header_start = best_row
            self._detected_header_height = best_height
            # Guard: prefer single-row header when promoting to multi-row yields little benefit.
            try:
                if best_height and best_height > 1:
                    # recompute group_count for height=1 at the chosen start row
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

                    # require a meaningful gain to keep multi-row header; small gains often indicate
                    # the lower rows are data-like and should not be absorbed. Threshold set to 1.
                    if (best_group_count - group_count_one) <= 1:
                        debug_print(f"[DEBUG] ヘッダー高さの見直し: 複数行によるグループ増分が小さいため単一行を優先します (row={best_row}, before_height={best_height}, groups_before={best_group_count}, groups_one={group_count_one})")
                        best_height = 1
                        self._detected_header_height = best_height
            except (ValueError, TypeError) as e:
                debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")
            # Additional guard: if the bottom row of the selected multi-row header
            # by itself provides equal or better grouping coverage, prefer it as a
            # single-row header (avoids pulling first data row into header).
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
                    # prefer bottom row when its grouping equals the multi-row grouping and covers many columns
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
                                     merged_info: Dict[str, Any]) -> List[List[str]]:
        """結合セルを考慮してテーブルデータを構築（ヘッダー行の検出とテーブル構造改善）"""
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
            header_min_col = start_col
            header_max_col = end_col
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
        # capture initial useful columns for diagnostics
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
                # skip header row (index 0) when counting data-bearing cells
                for r in filtered_table_data[1:]:
                    if ci < len(r) and r[ci] and str(r[ci]).strip():
                        cnt += 1
                col_counts.append(cnt)

            # decision: keep if cnt >= 1 OR fraction >= 0.05 (5%)
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

        # Diagnostics: for every original sheet column in the region, record why it was
        # kept or dropped. This helps trace which branch collapsed columns.
        try:
            per_column_diag = []
            num_all_cols = end_col - start_col + 1
            # map original relative index -> whether in initial_useful, header_keep, guard_keep
            for rel in range(0, num_all_cols):
                abs_col = start_col + rel
                in_initial = rel in initial_useful_columns
                in_final = rel in useful_columns
                header_present = False
                header_texts = []
                # attempt to pull header fragments from detected header rows if available
                try:
                    hdr_row = header_row or actual_start_row
                    detected_h = int(getattr(self, '_detected_header_height') or 1)
                    for hr in range(hdr_row, min(hdr_row + detected_h, end_row + 1)):
                        val = sheet.cell(hr, abs_col).value
                        if val is not None and str(val).strip():
                            header_present = True
                            header_texts.append(str(val).strip())
                except (ValueError, TypeError):
                    # fallback: use filtered_table_data header if exists
                    try:
                        if filtered_table_data and rel < len(filtered_table_data[0]):
                            hv = filtered_table_data[0][rel]
                            if hv and str(hv).strip():
                                header_present = True
                                header_texts.append(str(hv).strip())
                    except (ValueError, TypeError) as e:
                        debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")

                # data non-empty count for this original rel column
                data_count = 0
                for r in filtered_table_data[1:]:
                    if rel < len(r) and r[rel] and str(r[rel]).strip():
                        data_count += 1

                reason = 'kept' if in_final else 'dropped'
                per_column_diag.append((abs_col, rel, in_initial, header_present, header_texts, data_count, reason))

            # print diagnostics
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

        # dump after column compression step
        debug_print(f"[DEBUG-DUMP] after useful_columns compression: useful_columns={useful_columns}, table_rows={len(table_data)} sample (first 6):")
        for i, r in enumerate(table_data[:6]):
            debug_print(f"[DEBUG-DUMP] compressed row {i}: cols={len(r)} -> {r}")
        
        # --- 追加: ヘッダーに同一テキストが連続している場合、それらの列をまとめる ---
        if table_data:
            header = table_data[0]
            # compute per-column non-empty counts to decide whether to preserve columns
            col_nonempty_counts = []
            for ci in range(len(header)):
                cnt = 0
                for r in table_data[1:]:
                    if ci < len(r) and r[ci] and str(r[ci]).strip():
                        cnt += 1
                col_nonempty_counts.append(cnt)
            debug_print(f"[DEBUG-DUMP] per-column nonempty counts (after useful_columns): {col_nonempty_counts}")
            # groups: list of (start_idx, end_idx) for consecutive identical headers
            groups = []
            i = 0
            while i < len(header):
                j = i + 1
                # Only merge consecutive identical headers when the header value is non-empty.
                # If the header is empty, treat each column separately to avoid accidental collapse
                # of multiple data columns into one when header row lacks labels.
                while j < len(header) and header[j] == header[i] and (header[i] and str(header[i]).strip()):
                    j += 1
                groups.append((i, j))
                i = j

            # Post-process groups: if a group (whether header non-empty or not) contains
            # multiple columns that have non-empty data, avoid collapsing that group.
            final_groups = []
            for (a, b) in groups:
                # count how many columns in this group have non-empty data
                nonempty_cols_in_group = sum(1 for k in range(a, b) if k < len(col_nonempty_counts) and col_nonempty_counts[k] > 0)
                # if more than one data-bearing column, split into singletons to preserve columns
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
            # In this path we have only table_data available higher up; keep the table-data based check
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

                # Require a reasonable fraction of rows to match before collapsing
                total_data_rows = max(1, len(table_data) - 1)
                required = max(1, int(total_data_rows * 0.5))  # at least 50% of rows
                # extra diagnostic dump for 2-col optimization decision
                debug_print(f"[DEBUG-DUMP] 2col optimization: total_data_rows={total_data_rows}, matched={matched}, required={required}")
                # show sample of rows used for decision
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

        # Markdownテーブル内で問題となる文字をエスケープ
        # '|' はテーブル区切りになるためエスケープ
        cell_text = cell_text.replace('|', '\\|')

        # '&' は既存のエンティティを壊さないように一律に変換
        cell_text = cell_text.replace('&', '&amp;')
        # 角括弧はユーザ要件によりそのまま保持する (< and > are preserved)

        # 書式設定を適用
        cell_text = self._apply_cell_formatting(cell, cell_text)

        return cell_text

    def _collapse_repeated_sequence(self, parts: List[str]) -> List[str]:
        """Detect if parts is a repeated sequence (like [A,B,A,B,A,B]) and return the minimal repeating pattern once.

        If no perfect repetition is found, return parts unchanged.
        """
        try:
            if not parts:
                return parts
            n = len(parts)
            # try all possible pattern lengths up to n//2
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

            # Preserve programmatically inserted <br> (and common variants) so they are
            # not escaped — Excel-originated '<' '>' should still be escaped.
            # We replace allowed tags with placeholders, perform generic escaping,
            # then restore the placeholders back to the literal tags.
            allowed_tags = []
            # normalize tag variants we want to keep (lowercase)
            for m in re.finditer(r'(?i)<br\s*/?>', t):
                allowed_tags.append(m.group(0))

            placeholders = {}
            for i, tag in enumerate(allowed_tags):
                ph = f'___BR_TAG_PLACEHOLDER_{i}___'
                # replace only the first occurrence each time to keep mapping
                t = t.replace(tag, ph, 1)
                placeholders[ph] = tag

            # Protect existing HTML entities: convert '&' that are not part of an entity
            t = re.sub(r'&(?![A-Za-z]+;|#\d+;)', '&amp;', t)

            # Escape remaining angle brackets (these come from Excel cell content)
            t = t.replace('<', '&lt;').replace('>', '&gt;')

            # Escape Markdown table pipe
            t = t.replace('|', '\\|')

            # Restore allowed tags (placeholders) back to their literal forms
            for ph, tag in placeholders.items():
                # use the normalized '<br>' form
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

        # normalize row lengths
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

        # pick header rows count heuristically
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

        # respect previously detected header height if available
        if hasattr(self, '_detected_header_height'):
            detected = int(getattr(self, '_detected_header_height') or 0)
            if 1 <= detected <= max_header_rows:
                # If the first header row already contains combined pieces (i.e. many '<br>'),
                # avoid re-joining additional rows which can duplicate fragments.
                try:
                    first_row = table_data[0]
                    # count non-empty header cells and those that already contain '<br>'
                    nonempty_hdr = sum(1 for h in first_row if h and str(h).strip())
                    combined_hdr = sum(1 for h in first_row if h and '<br>' in str(h))
                    # if any header cells already contain '<br>', treat as single combined header
                    # (because _build_table_with_header_row already merged multi-row headers with '<br>')
                    if combined_hdr > 0:
                        header_rows_count = 1
                    else:
                        header_rows_count = detected
                except (ValueError, TypeError):
                    header_rows_count = detected

        # previously a fill-down copied the last seen value into later data rows
        # which could populate truly-empty data cells with unrelated values from above.
        # Limit fill-down to header-area completion only: if there are multiple header
        # rows and some header row cells are empty, fill them from nearby header rows
        # rather than propagating into the data area.
        try:
            if header_rows_count > 1:
                for col_idx in range(num_cols):
                    # gather header row values for this column
                    hdr_vals = [table_data[r][col_idx] if col_idx < len(table_data[r]) else '' for r in range(0, header_rows_count)]
                    # forward-fill within header rows only
                    last = None
                    for ri in range(0, header_rows_count):
                        v = hdr_vals[ri]
                        if v and str(v).strip():
                            last = v
                        else:
                            if last is not None:
                                # write back into table_data header cell
                                if col_idx < len(table_data[ri]):
                                    table_data[ri][col_idx] = last
        except (ValueError, TypeError) as e:
            debug_print(f"[DEBUG] 型変換エラー（無視）: {e}")

        # build header cells by joining header rows
        # NOTE: each header row cell may already contain '<br>' sequences (from merged/header assembly).
        # To avoid duplicating subparts when joining multiple header rows, split on '<br>' and dedupe
        # consecutive subparts across rows.
        header_cells = []
        for col in range(num_cols):
            subparts = []
            for ri in range(0, header_rows_count):
                v = table_data[ri][col]
                if v is not None and str(v).strip():
                    # normalize similarly to header builder to avoid near-duplicate parts
                    try:
                        import re as _re
                        vv = str(v).replace('\r\n', '\n').replace('\r', '\n').replace('\n', '<br>')
                        vv = _re.sub(r'(<br>\s*){2,}', '<br>', vv)
                        vv = _re.sub(r'^(?:<br>\s*)+', '', vv)
                        vv = _re.sub(r'(?:\s*<br>)+$', '', vv)
                        vv = vv.strip()
                    except (ValueError, TypeError):
                        vv = str(v).replace('\n', '<br>').strip()
                    # split already-normalized vv into atomic parts and extend
                    for part in [p.strip() for p in vv.split('<br>') if p.strip()]:
                        subparts.append(part)

            # dedupe consecutive subparts to avoid repeated fragments introduced by per-row combined cells
            dedup = []
            for p in subparts:
                if not dedup or dedup[-1] != p:
                    dedup.append(p)
            # collapse perfect repeated sequences like [A,B,A,B,A,B] -> [A,B]
            try:
                collapsed = self._collapse_repeated_sequence(dedup)
            except Exception:
                collapsed = dedup
            header_cells.append('<br>'.join(collapsed) if collapsed else '')

        # collapse consecutive identical headers conservatively
        try:
            groups = []
            i = 0
            while i < len(header_cells):
                j = i + 1
                while j < len(header_cells) and header_cells[j] == header_cells[i]:
                    j += 1
                groups.append((i, j))
                i = j

            # Only perform conservative collapsing of consecutive identical headers
            # when the header value is non-empty. If header_cells are empty strings
            # (common when header row is absent or misaligned), skip collapsing to
            # avoid merging multiple data columns into a single column.
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

    # output header
        safe_header = [self._escape_cell_for_table(h) for h in header_cells]
        self.markdown_lines.append("| " + " | ".join(safe_header) + " |")
        self.markdown_lines.append("| " + " | ".join(["---"] * len(header_cells)) + " |")

        # filter out immediate rows that are actually fragments of the header
        # (some table builders return raw header rows followed by a combined header;
        # detect and skip those to avoid duplicated header fragments in the data)
        try:
            skip_count = 0
            max_check = min(max_header_rows, len(table_data) - header_rows_count)
            # prepare header parts per column
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
                    # Skip overly long cells which are unlikely to be header fragments
                    if len(cell) > 200:
                        # treat as data-like
                        continue

                    # if the cell equals any of the header parts for that column, count as match
                    # but be conservative: require either exact short match or that both
                    # header-part and cell are reasonably short and do not contain '<br>' mismatch.
                    part_match = False
                    for hp in parts:
                        if not hp:
                            continue
                        # if one contains <br> and the other does not, avoid matching
                        if ('<br>' in hp) != ('<br>' in cell):
                            continue
                        # allow exact match
                        if cell == hp:
                            part_match = True
                            break
                        # allow short fuzzy match: both short and similar length
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

        # output data rows (skip detected header-like rows)
        start_idx = header_rows_count + skip_count
        # prepare mapping if source_rows provided
        sheet_map = None
        # Only obtain existing authoritative mapping; do not create it here.
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
            # record mapping from source row to markdown index if available
            if source_rows and idx < len(source_rows) and sheet_map is not None:
                src = source_rows[idx]
                # append line, then map src -> index (guarded)
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

def convert_xls_to_xlsx(xls_file_path: str) -> Optional[str]:
    """XLSファイルをXLSXに変換"""
    print(f"[INFO] XLSファイルをXLSXに変換中: {xls_file_path}")
    
    # 一時ディレクトリを作成
    temp_dir = tempfile.mkdtemp(prefix='xls2md_conversion_')
    
    try:
        # 出力ファイル名を決定
        xls_path = Path(xls_file_path)
        xlsx_filename = xls_path.stem + '.xlsx'
        xlsx_output_path = os.path.join(temp_dir, xlsx_filename)
        
        # LibreOfficeを使用してXLSをXLSXに変換
        cmd = [
            LIBREOFFICE_PATH,
            '--headless',
            '--convert-to', 'xlsx',
            '--outdir', temp_dir,
            xls_file_path
        ]
        
        debug_print(f"[DEBUG] LibreOffice変換コマンド: {' '.join(cmd)}")
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        
        if result.returncode != 0:
            print(f"[ERROR] LibreOffice変換失敗: {result.stderr}")
            shutil.rmtree(temp_dir)
            return None
        
        # 変換されたファイルが存在するか確認
        if not os.path.exists(xlsx_output_path):
            print(f"[ERROR] 変換後のXLSXファイルが見つかりません: {xlsx_output_path}")
            shutil.rmtree(temp_dir)
            return None
        
        print(f"[SUCCESS] XLS→XLSX変換完了: {xlsx_output_path}")
        return xlsx_output_path
        
    except subprocess.TimeoutExpired:
        print("[ERROR] LibreOffice変換がタイムアウトしました")
        shutil.rmtree(temp_dir)
        return None
    except Exception as e:
        print(f"[ERROR] XLS変換エラー: {e}")
        shutil.rmtree(temp_dir)
        return None


def main():
    """メイン関数"""
    import argparse
    
    parser = argparse.ArgumentParser(description='ExcelファイルをMarkdownに変換')
    parser.add_argument('excel_file', help='変換するExcelファイル（.xlsx/.xls）')
    parser.add_argument('-o', '--output-dir', type=str, 
                       help='出力ディレクトリを指定（デフォルト: ./output）')
    parser.add_argument('--debug', action='store_true',
                       help='デバッグモード：debug_workbooks、pdfs、diagnosticsフォルダを出力')
    parser.add_argument('--shape-metadata', action='store_true',
                       help='図形メタデータを画像の後に出力（テキスト形式とJSON形式）')
    parser.add_argument('-v', '--verbose', action='store_true',
                       help='デバッグ情報を出力')
    
    args = parser.parse_args()
    
    set_verbose(args.verbose)
    
    if not os.path.exists(args.excel_file):
        debug_print(f"エラー: ファイル '{args.excel_file}' が見つかりません。")
        sys.exit(1)
    
    if not args.excel_file.endswith(('.xlsx', '.xls')):
        debug_print("エラー: .xlsxまたは.xls形式のファイルを指定してください。")
        sys.exit(1)
    
    # XLSファイルの場合は事前にXLSXに変換
    processing_file = args.excel_file
    converted_file = None
    converted_temp_dir = None
    
    if args.excel_file.endswith('.xls'):
        debug_print("XLSファイルが指定されました。XLSXに変換します...")
        converted_file = convert_xls_to_xlsx(args.excel_file)
        if converted_file is None:
            debug_print("❌ XLS→XLSX変換に失敗しました。")
            sys.exit(1)
        processing_file = converted_file
        converted_temp_dir = Path(converted_file).parent
        debug_print(f"✅ XLS→XLSX変換完了: {converted_file}")
    
    try:
        converter = ExcelToMarkdownConverter(processing_file, output_dir=args.output_dir, debug_mode=args.debug, shape_metadata=args.shape_metadata)
        output_file = converter.convert()
        debug_print("\n✅ 変換完了!")
        debug_print(f"📄 出力ファイル: {output_file}")
        debug_print(f"🖼️  画像フォルダ: {converter.images_dir}")
        
    except Exception as e:
        debug_print(f"❌ 変換エラー: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
    finally:
        # 一時的に作成したXLSXファイルとその親ディレクトリを削除
        if converted_temp_dir:
            try:
                if converted_temp_dir.exists() and converted_temp_dir.name.startswith('xls2md_conversion_'):
                    shutil.rmtree(converted_temp_dir)
                    debug_print(f"🗑️  一時ディレクトリを削除: {converted_temp_dir}")
            except Exception as cleanup_error:
                debug_print(f"⚠️  一時ファイル削除に失敗: {cleanup_error}")


if __name__ == "__main__":
    main()
