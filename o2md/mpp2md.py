#!/usr/bin/env python3
"""
MS Project to Markdown Converter
MS Projectファイル (.mpp, .mpt, .mpx, .xml) をMarkdownに変換するツール

対応形式:
- .mpp (MS Project バイナリ)
- .mpt (MS Project テンプレート)
- .mpx (MS Project Exchange)
- .xml (MSPDI XML形式)

内部構造:
- ネイティブバイナリ (mpp-reader) を subprocess で呼び出し
- mpp-reader は mpxj を GraalVM Native Image でコンパイルしたもの
- JRE不要で動作する
- タスク情報をJSON形式で取得し、PythonでMarkdownに変換

必要な依存関係:
- mpp-reader ネイティブバイナリ (o2md/bin/mpp-reader)
"""

import json
import logging
import os
import subprocess
import sys
import argparse

from o2md.i18n import _
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)

# MS Projectファイルの対応拡張子
MPP_EXTENSIONS = ('.mpp', '.mpt', '.mpx')

# グローバルverboseフラグ
_VERBOSE = False


def set_verbose(verbose: bool):
    """verboseモードを設定"""
    global _VERBOSE
    _VERBOSE = verbose
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.WARNING,
        format='[%(levelname)s] %(message)s',
    )


def is_verbose() -> bool:
    """verboseモードかどうかを返す"""
    return _VERBOSE


def debug_print(*args, **kwargs):
    """verboseモード時のみ出力するデバッグ用print"""
    if _VERBOSE:
        print(*args, **kwargs)


def _find_mpp_reader() -> str:
    """mpp-readerバイナリのパスを検索する

    検索順序:
    1. o2md/bin/mpp-reader（パッケージ同梱）
    2. PATH上のmpp-reader

    Returns:
        mpp-readerの絶対パス

    Raises:
        FileNotFoundError: バイナリが見つからない場合
    """
    # パッケージ同梱のバイナリを検索
    pkg_dir = Path(__file__).parent
    bundled = pkg_dir / 'bin' / 'mpp-reader'
    if bundled.exists() and os.access(str(bundled), os.X_OK):
        return str(bundled)

    # PATH上を検索
    import shutil
    path_bin = shutil.which('mpp-reader')
    if path_bin:
        return path_bin

    raise FileNotFoundError(
        "mpp-readerバイナリが見つかりません。"
        "o2md/bin/mpp-reader に配置するか、PATHに追加してください。"
        "ビルド方法: cd native/mpp-reader && ./gradlew nativeCompile"
    )


def read_project_json(file_path: str) -> dict:
    """MS Projectファイルを読み込み、JSON辞書として返す

    mpp-readerネイティブバイナリをsubprocessで実行し、
    JSON出力をパースして返す。

    Args:
        file_path: MS Projectファイルのパス

    Returns:
        プロジェクト情報のJSON辞書

    Raises:
        RuntimeError: バイナリ実行エラー
        FileNotFoundError: バイナリが見つからない
    """
    binary = _find_mpp_reader()
    abs_path = os.path.abspath(file_path)

    debug_print(f"mpp-reader実行: {binary} {abs_path}")

    result = subprocess.run(
        [binary, abs_path],
        capture_output=True,
        timeout=60,
    )

    if result.returncode != 0:
        err_msg = result.stderr.decode('utf-8', errors='replace').strip()
        raise RuntimeError(
            f"mpp-readerの実行に失敗しました: {err_msg}"
        )

    stdout = result.stdout.decode('utf-8')
    return json.loads(stdout)


def _escape_md_cell(text: str) -> str:
    """Markdownテーブルセル内のパイプ文字をエスケープする"""
    return text.replace("|", "\\|")


def tasks_to_markdown_table(tasks: list[dict]) -> str:
    """タスク情報をMarkdownテーブルに変換する

    アウトラインレベルに応じてタスク名にインデントを付与する。
    サマリータスクは太字で表示する。

    Args:
        tasks: タスク情報の辞書リスト

    Returns:
        Markdownテーブル文字列
    """
    if not tasks:
        return ""

    # テーブルヘッダ
    header = "| タスク名 | 開始日 | 終了日 | 期間 | 進捗 | 担当者 |"
    separator = "| --- | --- | --- | --- | --- | --- |"
    rows = [header, separator]

    # 最小アウトラインレベルを基準にインデント計算
    min_level = min(t['outline_level'] for t in tasks)

    for task in tasks:
        indent_depth = task['outline_level'] - min_level
        # ノーブレークスペースでインデント表現
        indent = "\u00a0\u00a0\u00a0\u00a0" * indent_depth

        name = _escape_md_cell(task['name'])
        if task.get('is_summary', False):
            name = f"**{name}**"

        resources = _escape_md_cell(task.get('resources', '-'))

        row = (
            f"| {indent}{name}"
            f" | {task.get('start', '-')}"
            f" | {task.get('finish', '-')}"
            f" | {task.get('duration', '-')}"
            f" | {task.get('percent_complete', '-')}"
            f" | {resources} |"
        )
        rows.append(row)

    return "\n".join(rows)


def resources_to_markdown_table(resources: list[dict]) -> str:
    """リソース情報をMarkdownテーブルに変換する

    Args:
        resources: リソース情報の辞書リスト

    Returns:
        Markdownテーブル文字列
    """
    if not resources:
        return ""

    header = "| ID | リソース名 |"
    separator = "| --- | --- |"
    rows = [header, separator]

    for res in resources:
        name = _escape_md_cell(res['name'])
        rows.append(f"| {res['id']} | {name} |")

    return "\n".join(rows)


def project_info_to_markdown(info: dict) -> str:
    """プロジェクト情報をMarkdown形式に変換する

    Args:
        info: プロジェクト情報の辞書

    Returns:
        Markdown文字列
    """
    lines = []
    title = info.get('title', '')
    if title:
        lines.append(f"- **プロジェクト名**: {title}")
    author = info.get('author', '')
    if author:
        lines.append(f"- **作成者**: {author}")
    start = info.get('start', '')
    if start and start != '-':
        lines.append(f"- **開始日**: {start}")
    finish = info.get('finish', '')
    if finish and finish != '-':
        lines.append(f"- **終了日**: {finish}")
    return "\n".join(lines)


class MppToMarkdownConverter:
    """MS ProjectファイルをMarkdownに変換するコンバータ

    mpp-readerネイティブバイナリ（mpxj + GraalVM Native Image）を使用して
    MS Projectファイル(.mpp/.mpt/.mpx)を解析し、
    タスク一覧テーブルとリソース一覧テーブルを含むMarkdownを出力する。
    JRE不要で動作する。
    """

    def __init__(
        self,
        file_path: str,
        output_dir: Optional[str] = None,
    ):
        """
        Args:
            file_path: MS Projectファイルのパス
            output_dir: 出力ディレクトリ(省略時はカレントディレクトリ)
        """
        self.file_path = file_path
        self.base_name = Path(file_path).stem
        self.output_image_count = 0

        if output_dir:
            self.output_dir = output_dir
        else:
            self.output_dir = os.path.join(os.getcwd(), "output")

        os.makedirs(self.output_dir, exist_ok=True)

    def get_auto_generated_patterns(self) -> list:
        """このコンバータが自動付与する見出しの正規表現パターンを返す"""
        import re
        return [
            re.compile(r'^' + re.escape(self.base_name) + r'$'),
            re.compile(r'^タスク一覧$'),
            re.compile(r'^リソース一覧$'),
            re.compile(r'^プロジェクト情報$'),
        ]

    def get_auto_generated_html_tags(self) -> list:
        """このコンバータが自動付与するHTMLタグを返す"""
        return []

    def get_auto_generated_line_patterns(self) -> list:
        """このコンバータが自動付与するメタデータ行パターンを返す"""
        import re
        return [
            re.compile(r'^\*\*プロジェクト名\*\*:'),
            re.compile(r'^\*\*作成者\*\*:'),
            re.compile(r'^\*\*開始日\*\*:'),
            re.compile(r'^\*\*終了日\*\*:'),
        ]

    def convert(self) -> str:
        """変換メイン処理

        Returns:
            出力ファイルのパス（.mdまたは.txt）
        """
        from o2md.utils import is_text_only

        print(_("MS Project変換開始: {file}").format(file=self.file_path))

        # mpp-readerでJSON取得
        data = read_project_json(self.file_path)

        proj_info = data.get('project', {})
        tasks = data.get('tasks', [])
        resources = data.get('resources', [])

        debug_print(
            f"抽出完了: タスク{len(tasks)}件, "
            f"リソース{len(resources)}件"
        )

        # Markdown生成
        md_lines = self._build_markdown(proj_info, tasks, resources)
        md_content = "\n".join(md_lines)

        # テキストモード
        if is_text_only():
            return self._write_text(md_content)

        # 通常モード: .md出力
        return self._write_markdown(md_content)

    def _build_markdown(
        self,
        proj_info: dict,
        tasks: list[dict],
        resources: list[dict],
    ) -> list[str]:
        """Markdownコンテンツを構築する"""
        lines = []

        # タイトル
        title = proj_info.get('title', '') or self.base_name
        lines.append(f"# {title}")
        lines.append("")

        # プロジェクト情報
        info_md = project_info_to_markdown(proj_info)
        if info_md:
            lines.append("## プロジェクト情報")
            lines.append("")
            lines.append(info_md)
            lines.append("")

        # タスク一覧テーブル
        if tasks:
            lines.append("## タスク一覧")
            lines.append("")
            lines.append(tasks_to_markdown_table(tasks))
            lines.append("")

        # リソース一覧テーブル
        if resources:
            lines.append("## リソース一覧")
            lines.append("")
            lines.append(resources_to_markdown_table(resources))
            lines.append("")

        return lines

    def _write_text(self, md_content: str) -> str:
        """テキストモードで.txtを出力する"""
        from o2md.o2md import strip_markdown

        auto_patterns = {
            'heading_patterns': self.get_auto_generated_patterns(),
            'html_tags': self.get_auto_generated_html_tags(),
            'line_patterns': self.get_auto_generated_line_patterns(),
        }
        text_content = strip_markdown(
            md_content, auto_patterns=auto_patterns
        )
        output_path = os.path.join(
            self.output_dir, f"{self.base_name}.txt"
        )
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(text_content)
        logger.info(f"変換完了: {output_path}")
        return output_path

    def _write_markdown(self, md_content: str) -> str:
        """通常モードで.mdを出力する"""
        output_path = os.path.join(
            self.output_dir, f"{self.base_name}.md"
        )
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(md_content)
        logger.info(f"変換完了: {output_path}")
        return output_path


def main():
    """コマンドラインエントリポイント"""
    parser = argparse.ArgumentParser(
        description='MS ProjectファイルをMarkdownに変換'
    )
    parser.add_argument(
        'file',
        help='変換するMS Projectファイル (.mpp/.mpt/.mpx/.xml)'
    )
    parser.add_argument(
        '-o', '--output-dir', type=str,
        help='出力ディレクトリを指定（デフォルト: ./output）'
    )
    parser.add_argument(
        '-v', '--verbose', action='store_true',
        help='デバッグ情報を出力'
    )
    parser.add_argument(
        '--text', action='store_true',
        help='テキスト抽出モード（.txtのみ出力）'
    )

    args = parser.parse_args()

    set_verbose(args.verbose)

    if args.text:
        from o2md.utils import set_text_only
        set_text_only(True)

    converter = MppToMarkdownConverter(
        args.file,
        output_dir=args.output_dir,
    )

    output_file = converter.convert()

    print("\n" + "=" * 50)
    print(_("出力ファイル: {output_file}").format(output_file=output_file))
    print("=" * 50)


if __name__ == "__main__":
    main()
