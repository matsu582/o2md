#!/usr/bin/env python3
"""
PDF表検出・処理Mixinモジュール

PDFToMarkdownConverterクラスの表検出・処理機能を提供します。
このモジュールはMixinクラスとして設計されており、メインクラスから継承されます。

機能:
- カラム内の表領域検出
- 罫線ベースの表検出（PyMuPDFのfind_tables使用）
- 表構造の検出とMarkdownテーブル生成
- フル幅領域での表検出
"""

import re
from typing import List, Dict, Tuple, Set, Any


def debug_print(*args, **kwargs):
    """デバッグ出力（pdf2mdモジュールに委譲）"""
    try:
        from pdf2md import debug_print as _dp
        _dp(*args, **kwargs)
    except ImportError:
        pass


class _TablesMixin:
    """表検出・処理機能を提供するMixinクラス
    
    このクラスはPDFToMarkdownConverterに継承され、
    表検出、Markdownテーブル生成機能を提供します。
    
    注意: このクラスは単独では使用できません。
    PDFToMarkdownConverterクラスと組み合わせて使用してください。
    """

    def _format_markdown_table(self, rows: List[List[str]]) -> str:
        """表データをMarkdown形式に整形（列幅調整付き）
        
        Args:
            rows: 表データ（2次元リスト、最初の行がヘッダー）
            
        Returns:
            Markdown形式の表文字列
        """
        if not rows or len(rows) < 2:
            return ""
        
        col_count = len(rows[0])
        if col_count == 0:
            return ""
        
        # 各列の最大幅を計算（日本語文字は2文字分としてカウント）
        col_widths = [3] * col_count  # 最小幅は3（---）
        for row in rows:
            for i, cell in enumerate(row):
                if i < col_count:
                    # 日本語文字（マルチバイト）は2文字分
                    width = sum(2 if ord(c) > 127 else 1 for c in cell)
                    col_widths[i] = max(col_widths[i], width)
        
        md_lines = []
        
        # ヘッダー行
        md_lines.append("| " + " | ".join(rows[0]) + " |")
        
        # 区切り行（列幅に応じた長さ）
        separator_cells = ["-" * w for w in col_widths]
        md_lines.append("|" + "|".join(separator_cells) + "|")
        
        # データ行
        for row in rows[1:]:
            # 列数を揃える
            padded_row = row + [""] * (col_count - len(row))
            md_lines.append("| " + " | ".join(padded_row[:col_count]) + " |")
        
        return "\n".join(md_lines)

    def _detect_table_regions(
        self, lines_data: List[Dict], page_center: float
    ) -> List[Dict]:
        """カラム内の表領域を検出
        
        同じY座標に複数のセルがある行が連続する領域を表として検出する。
        
        Args:
            lines_data: 行データのリスト
            page_center: ページの中央X座標
            
        Returns:
            表領域のリスト（各領域は{y_start, y_end, column, rows}を含む）
        """
        import re as re_mod
        # キャプションパターン（図X、表Xで始まる行は表検出から除外）
        caption_pattern = re_mod.compile(r'^[図表]\s*\d+')
        
        # キャプション行を除外してから左右カラムを分離
        filtered_lines = [l for l in lines_data 
                         if not caption_pattern.match(l.get("text", ""))]
        left_lines = [l for l in filtered_lines if l["column"] == "left"]
        right_lines = [l for l in filtered_lines if l["column"] == "right"]
        
        table_regions = []
        
        for column_lines, column_name in [(left_lines, "left"), (right_lines, "right")]:
            if not column_lines:
                continue
            
            # Y座標でグループ化（同じ行にある要素を検出）
            y_tolerance = 5  # Y座標の許容誤差
            y_groups = {}
            
            for line in column_lines:
                y_key = round(line["y"] / y_tolerance) * y_tolerance
                if y_key not in y_groups:
                    y_groups[y_key] = []
                y_groups[y_key].append(line)
            
            # 複数セルがある行を検出（3セル以上で表として認識）
            # ただし、短いテキスト（3文字以下）だけで構成される行は除外（図内ラベルの誤検出防止）
            multi_cell_rows = []
            all_rows = []  # 全ての行（単一セル含む）
            for y_key in sorted(y_groups.keys()):
                cells = y_groups[y_key]
                # X座標でソートして、異なるX位置にあるセルをカウント
                x_positions = sorted(set(round(c["x"] / 20) * 20 for c in cells))
                # 短いテキストだけで構成される行は除外
                texts = [c.get("text", "") for c in cells]
                has_long_text = any(len(t) > 3 for t in texts)
                is_multi = len(x_positions) >= 3 and has_long_text
                row_data = {
                    "y": y_key,
                    "cells": sorted(cells, key=lambda c: c["x"]),
                    "x_positions": x_positions,
                    "is_multi_cell": is_multi
                }
                all_rows.append(row_data)
                if is_multi:
                    multi_cell_rows.append(row_data)
            
            # 連続する複数セル行を表領域としてグループ化
            if not multi_cell_rows:
                continue
            
            current_region = {
                "y_start": multi_cell_rows[0]["y"],
                "y_end": multi_cell_rows[0]["y"] + 20,
                "column": column_name,
                "rows": [multi_cell_rows[0]],
                "all_rows": all_rows  # 単一セル行も含む全行を保持
            }
            
            for i in range(1, len(multi_cell_rows)):
                row = multi_cell_rows[i]
                prev_row = multi_cell_rows[i - 1]
                
                # 連続している場合（Y座標の差が小さい）
                if row["y"] - prev_row["y"] < 50:  # 許容範囲を広げる
                    current_region["rows"].append(row)
                    current_region["y_end"] = row["y"] + 20
                else:
                    # 2行以上の連続した複数セル行があれば表として認識
                    if len(current_region["rows"]) >= 2:
                        table_regions.append(current_region)
                    
                    current_region = {
                        "y_start": row["y"],
                        "y_end": row["y"] + 20,
                        "column": column_name,
                        "rows": [row],
                        "all_rows": all_rows
                    }
            
            # 最後の領域をチェック
            if len(current_region["rows"]) >= 2:
                table_regions.append(current_region)
        
        return table_regions

    def _calc_table_quality_score(self, tables_list) -> float:
        """表検出結果の品質スコアを計算
        
        Args:
            tables_list: 検出された表のリスト
            
        Returns:
            品質スコア（高いほど良い）
        """
        score = 0.0
        for table in tables_list:
            rows = table.extract()
            if not rows:
                continue
            
            # 行数・列数が多いほど高スコア
            row_count = len(rows)
            col_count = len(rows[0]) if rows else 0
            score += row_count * col_count
            
            # 空セル比率が低いほど高スコア
            total_cells = row_count * col_count
            empty_cells = sum(
                1 for row in rows for cell in row
                if cell is None or str(cell).strip() == ""
            )
            if total_cells > 0:
                score *= (1 - empty_cells / total_cells)
        
        return score

    def _detect_line_based_tables(
        self, page, lines_data: List[Dict]
    ) -> List[Dict]:
        """罫線ベースの表検出（PyMuPDFのfind_tables()を使用）
        
        複数の検出戦略を試行し、最適な結果を選択する。
        テキストベースの検出で見逃した表を補完する。
        検出された表領域内のテキストを除外し、Markdownテーブルを生成する。
        
        Args:
            page: PyMuPDFのページオブジェクト
            lines_data: 行データのリスト（表領域内の行を除外するために使用）
            
        Returns:
            検出された表のリスト（各表は{bbox, markdown, y_start}を含む）
        """
        detected_tables = []
        
        try:
            # 複数の戦略を試行して最適な結果を選択
            best_tables = None
            best_score = 0
            
            # 罫線ベースの戦略のみ使用
            # strategy="text"は段組みテキストを表として誤検出するため除外
            strategies = [
                {},  # デフォルト
                {"strategy": "lines_strict"},  # 厳密な罫線検出
            ]
            
            for params in strategies:
                try:
                    tables = page.find_tables(**params)
                    if tables.tables:
                        score = self._calc_table_quality_score(tables.tables)
                        if score > best_score:
                            best_score = score
                            best_tables = tables.tables
                            debug_print(
                                f"[DEBUG] find_tables戦略: {params}, "
                                f"score={score:.1f}, tables={len(tables.tables)}"
                            )
                except Exception as e:
                    debug_print(f"[DEBUG] find_tables戦略エラー: {params}, {e}")
                    continue
            
            if not best_tables:
                return []
            
            for table in best_tables:
                bbox = table.bbox
                rows = table.extract()
                
                if not rows:
                    continue
                
                # 列数を確認（最低2列以上）
                col_count = len(rows[0]) if rows[0] else 0
                if col_count < 2:
                    continue
                
                # 空セル比率と高さによるフィルタリング
                # ヘッダ部分の装飾的な罫線を誤検出しないようにする
                table_height = bbox[3] - bbox[1]
                empty_cell_count = 0
                total_cell_count = 0
                for row in rows:
                    for cell in row:
                        total_cell_count += 1
                        if cell is None or str(cell).strip() == "":
                            empty_cell_count += 1
                
                if total_cell_count > 0:
                    empty_ratio = empty_cell_count / total_cell_count
                    # 空セル比率が50%以上かつ高さが50px以下の場合は除外
                    if empty_ratio >= 0.5 and table_height <= 50:
                        debug_print(f"[DEBUG] 罫線ベース表スキップ（空セル多/高さ小）: bbox={bbox}, empty_ratio={empty_ratio:.2f}, height={table_height:.1f}")
                        continue
                
                # 各行を処理（改行を含むセルはスペースで結合）
                # doclingと同様に、セル内改行を展開せず1セル1行として出力
                processed_rows = []
                for row in rows:
                    processed_row = []
                    for cell in row:
                        if cell is None:
                            cell_text = ""
                        else:
                            # 改行をスペースで結合（論理的な表構造を維持）
                            lines = str(cell).split("\n")
                            cell_text = " ".join(
                                line.strip() for line in lines if line.strip()
                            )
                        # パイプ文字をエスケープ
                        cell_text = cell_text.replace("|", "\\|")
                        processed_row.append(cell_text)
                    processed_rows.append(processed_row)
                
                # 行数を確認（最低2行以上）
                if len(processed_rows) < 2:
                    continue
                
                # 共通メソッドでMarkdownテーブルを生成
                markdown = self._format_markdown_table(processed_rows)
                if not markdown:
                    continue
                
                # 表領域内の行を除外対象としてマーク
                for line in lines_data:
                    line_y = line.get("y", 0)
                    line_x = line.get("x", 0)
                    # 表のbbox内にある行を除外
                    if (bbox[1] - 5 <= line_y <= bbox[3] + 5 and
                        bbox[0] - 5 <= line_x <= bbox[2] + 5):
                        line["in_line_based_table"] = True
                
                detected_tables.append({
                    "bbox": bbox,
                    "markdown": markdown,
                    "y_start": bbox[1],
                    "y_end": bbox[3]
                })
                
                debug_print(f"[DEBUG] 罫線ベース表検出: bbox={bbox}, rows={len(rows)}, cols={col_count}")
                
        except Exception as e:
            debug_print(f"[DEBUG] find_tables()エラー: {e}")
        
        return detected_tables

    def _format_table_region(self, table_region: Dict) -> str:
        """表領域をMarkdownテーブル形式に変換
        
        単一セル行（継続行）も含めて処理し、適切にマージする。
        
        Args:
            table_region: 表領域データ
            
        Returns:
            Markdownテーブル文字列
        """
        rows = table_region.get("rows", [])
        all_rows = table_region.get("all_rows", [])
        if not rows:
            return ""
        
        y_start = table_region.get("y_start", 0)
        y_end = table_region.get("y_end", 0)
        
        # 表領域内の全行を取得（単一セル行も含む）
        table_all_rows = []
        for row in all_rows:
            if y_start - 10 <= row["y"] <= y_end + 10:
                table_all_rows.append(row)
        
        if not table_all_rows:
            table_all_rows = rows
        
        # 複数セル行からヘッダ行を特定（最初の複数セル行）
        header_row = None
        for row in table_all_rows:
            if row.get("is_multi_cell", False):
                header_row = row
                break
        
        if not header_row:
            return ""
        
        # ヘッダ行の列位置を基準にする
        column_positions = sorted(header_row.get("x_positions", []))
        if len(column_positions) < 2:
            return ""
        
        # 各行のセルを列に割り当て
        table_rows = []
        for row in table_all_rows:
            cells = row.get("cells", [])
            row_data = [""] * len(column_positions)
            
            for cell in cells:
                cell_x = round(cell["x"] / 20) * 20
                # 最も近い列に割り当て
                min_dist = float("inf")
                best_idx = 0
                for idx, pos in enumerate(column_positions):
                    dist = abs(cell_x - pos)
                    if dist < min_dist:
                        min_dist = dist
                        best_idx = idx
                
                cell_text = cell["text"].strip()
                if not row_data[best_idx]:
                    row_data[best_idx] = cell_text
                else:
                    row_data[best_idx] += " " + cell_text
            
            table_rows.append({
                "data": row_data,
                "is_multi_cell": row.get("is_multi_cell", False),
                "y": row["y"]
            })
        
        if not table_rows:
            return ""
        
        # 継続行をマージ（単一セル行を前の行にマージ）
        merged_rows = []
        for row in table_rows:
            if row["is_multi_cell"]:
                merged_rows.append(row["data"])
            else:
                # 単一セル行: 前の行にマージ
                if merged_rows:
                    prev_row = merged_rows[-1]
                    for i, cell in enumerate(row["data"]):
                        if cell:
                            if prev_row[i]:
                                prev_row[i] += " " + cell
                            else:
                                prev_row[i] = cell
                else:
                    merged_rows.append(row["data"])
        
        if not merged_rows:
            return ""
        
        # Markdownテーブルを生成
        md_lines = []
        
        # ヘッダ行
        header = merged_rows[0]
        md_lines.append("| " + " | ".join(header) + " |")
        
        # 区切り行
        md_lines.append("| " + " | ".join(["---"] * len(header)) + " |")
        
        # データ行
        for row in merged_rows[1:]:
            md_lines.append("| " + " | ".join(row) + " |")
        
        return "\n".join(md_lines)

    def _detect_table_in_fullwidth(
        self, text_dict: Dict, header_footer_patterns: Set[str]
    ) -> List[List[Dict]]:
        """フル幅領域での表検出
        
        段組みページでも、フル幅領域（ページ幅の60%以上）にある
        表構造を検出する。
        
        Args:
            text_dict: PyMuPDFのテキスト辞書
            header_footer_patterns: ヘッダ・フッタパターン
            
        Returns:
            表の行リスト
        """
        page_width = text_dict.get("width", 612)
        
        # フル幅領域の行を収集
        y_groups: Dict[int, List[Dict]] = {}
        
        for block in text_dict.get("blocks", []):
            if block.get("type") != 0:
                continue
            
            block_bbox = block.get("bbox", (0, 0, 0, 0))
            block_width = block_bbox[2] - block_bbox[0]
            
            # フル幅ブロックのみ対象
            if block_width < page_width * 0.6:
                continue
            
            for line in block.get("lines", []):
                bbox = line.get("bbox", (0, 0, 0, 0))
                y_key = round(bbox[1] / 8) * 8  # 8ピクセル単位でグループ化
                
                line_text = ""
                for span in line.get("spans", []):
                    line_text += span.get("text", "")
                
                if line_text.strip():
                    # ヘッダ・フッタを除外
                    if self._is_header_footer(line_text, header_footer_patterns):
                        continue
                    
                    if y_key not in y_groups:
                        y_groups[y_key] = []
                    y_groups[y_key].append({
                        "text": line_text.strip(),
                        "x": bbox[0],
                        "bbox": bbox
                    })
        
        # 複数のセルがある行を表の行として抽出
        table_rows = []
        for y_key in sorted(y_groups.keys()):
            cells = y_groups[y_key]
            if len(cells) >= 2:
                cells_sorted = sorted(cells, key=lambda c: c["x"])
                avg_text_len = sum(len(c["text"]) for c in cells_sorted) / len(cells_sorted)
                if avg_text_len < 50:
                    table_rows.append(cells_sorted)
        
        # 表として認識する条件
        if len(table_rows) >= 3:
            col_counts = [len(row) for row in table_rows]
            most_common_cols = max(set(col_counts), key=col_counts.count)
            consistent_rows = sum(1 for c in col_counts if c == most_common_cols)
            if consistent_rows / len(table_rows) >= 0.8:
                return table_rows
        
        return []

    def _detect_table_structure(self, text_dict: Dict) -> List[List[Dict]]:
        """表構造を検出
        
        同じY座標に複数のテキストブロックがある場合、表として検出する。
        段組みレイアウトの場合は表検出を無効化する。
        
        Args:
            text_dict: PyMuPDFのテキスト辞書
            
        Returns:
            表の行リスト（各行はセルのリスト）
        """
        # 段組みレイアウトの場合は表検出をスキップ
        column_count = self._detect_column_layout(text_dict)
        if column_count >= 2:
            debug_print("[DEBUG] 段組みレイアウトのため表検出をスキップ")
            return []
        
        page_width = text_dict.get("width", 612)
        
        # 行の高さを収集して動的許容値を計算
        line_heights = []
        for block in text_dict.get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                bbox = line.get("bbox", (0, 0, 0, 0))
                height = bbox[3] - bbox[1]
                if height > 0:
                    line_heights.append(height)
        
        # 動的許容値: 行の高さの中央値の半分（最小3、最大10）
        if line_heights:
            median_height = sorted(line_heights)[len(line_heights) // 2]
            y_tolerance = max(3, min(10, median_height / 2))
        else:
            y_tolerance = 5
        
        # Y座標でグループ化（動的許容値を使用）
        y_groups: Dict[int, List[Dict]] = {}
        
        for block in text_dict.get("blocks", []):
            if block.get("type") != 0:
                continue
            
            for line in block.get("lines", []):
                bbox = line.get("bbox", (0, 0, 0, 0))
                y_key = round(bbox[1] / y_tolerance) * int(y_tolerance)
                
                line_text = ""
                for span in line.get("spans", []):
                    line_text += span.get("text", "")
                
                if line_text.strip():
                    if y_key not in y_groups:
                        y_groups[y_key] = []
                    y_groups[y_key].append({
                        "text": line_text.strip(),
                        "x": bbox[0],
                        "bbox": bbox
                    })
        
        # 複数のセルがある行を表の行として抽出
        table_rows = []
        for y_key in sorted(y_groups.keys()):
            cells = y_groups[y_key]
            if len(cells) >= 2:  # 2つ以上のセルがある行
                # X座標でソート
                cells_sorted = sorted(cells, key=lambda c: c["x"])
                
                # 表の追加条件: セルの平均文字数が短め（長文は段組みの可能性）
                avg_text_len = sum(len(c["text"]) for c in cells_sorted) / len(cells_sorted)
                if avg_text_len < 50:  # 平均50文字未満
                    table_rows.append(cells_sorted)
        
        # 表として認識する条件を強化
        # 1. 連続する行が3行以上
        # 2. 列数が行ごとに大きくブレない
        if len(table_rows) >= 3:
            # 列数の一貫性をチェック
            col_counts = [len(row) for row in table_rows]
            most_common_cols = max(set(col_counts), key=col_counts.count)
            consistent_rows = sum(1 for c in col_counts if c == most_common_cols)
            
            # 80%以上の行が同じ列数なら表として認識
            if consistent_rows / len(table_rows) >= 0.8:
                return table_rows
        
        return []

    def _merge_table_blocks(
        self, blocks: List[Dict], table_rows: List[List[Dict]]
    ) -> List[Dict]:
        """表構造をブロックリストにマージ
        
        Args:
            blocks: 既存のブロックリスト
            table_rows: 検出された表の行
            
        Returns:
            更新されたブロックリスト
        """
        if not table_rows:
            return blocks
        
        # 表データを2次元リストに変換
        rows_data = []
        for row in table_rows:
            row_cells = [cell["text"] for cell in row]
            rows_data.append(row_cells)
        
        # 共通メソッドでMarkdownテーブルを生成
        table_text = self._format_markdown_table(rows_data)
        if not table_text:
            return blocks
        
        # 表ブロックを追加（既存ブロックの最後に）
        # 表に含まれるテキストを持つブロックを除外
        table_texts = set()
        for row in table_rows:
            for cell in row:
                table_texts.add(cell["text"])
        
        filtered_blocks = []
        for block in blocks:
            block_text = block["text"].strip()
            # 表のセルテキストと完全一致するブロックは除外
            if block_text not in table_texts:
                filtered_blocks.append(block)
        
        # 表ブロックを追加
        filtered_blocks.append({
            "type": "table",
            "text": table_text,
            "font_size": 12,
            "bbox": (0, 0, 0, 0)
        })
        
        return filtered_blocks
