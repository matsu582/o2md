#!/usr/bin/env python3
"""
MS Project to Markdown Converter (JPype版)
MS Projectファイル (.mpp, .mpt, .mpx, .xml) をMarkdownに変換するツール

対応形式:
- .mpp (MS Project バイナリ)
- .mpt (MS Project テンプレート)
- .mpx (MS Project Exchange)
- .xml (MSPDI XML形式)

内部構造:
- JPypeを使用してJVMからmpxjライブラリを直接呼び出す
- JDK/JRE 11以上が必要
- タスク情報をPython辞書として取得し、Markdownに変換
"""

import logging
import os
import re
import sys
from pathlib import Path
from typing import Optional
from datetime import datetime
from o2md.i18n import _, setup_i18n

logger = logging.getLogger(__name__)

# グローバルverboseフラグ
_VERBOSE = False
_JPY_INITIALIZED = False


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


def _init_jpype():
    """JPypeとJVMを初期化する"""
    global _JPY_INITIALIZED
    
    if _JPY_INITIALIZED:
        return
    
    try:
        from jpype import isJVMStarted, startJVM, getDefaultJVMPath
        
        if not isJVMStarted():
            debug_print(_("JVM初期化中..."))
            jvm_path = getDefaultJVMPath()
            if not jvm_path:
                raise RuntimeError(
                    _(
                        "JDK/JREが見つかりません。 "
                        "JDK 11以上をインストールしてください。"
                    )
                )
            
            # CLASSPATH 環境変数があれば優先、なければ JAR を自動取得
            classpath = os.environ.get('CLASSPATH', '')
            if not classpath:
                from o2md.jar_manager import ensure_mpxj_jars
                classpath = ensure_mpxj_jars(verbose=is_verbose())
            
            # JVMを起動
            startJVM(jvm_path, "-Xmx4G", classpath=classpath)
            debug_print(_(f"JVM起動完了: {jvm_path}"))
        
        _JPY_INITIALIZED = True
    except ImportError as e:
        raise ImportError(
            _(
                "JPypeが見つかりません。 "
                "pip install jpype1>=1.5.0 でインストールしてください。"
            )
        ) from e
    except Exception as e:
        raise RuntimeError(_(f"JVM初期化エラー: {e}")) from e


def _get_java_classes():
    """必要なJavaクラスをロードして返す"""
    from jpype import JClass
    
    return {
        'UniversalProjectReader': JClass('net.sf.mpxj.reader.UniversalProjectReader'),
        'MPPReader': JClass('net.sf.mpxj.mpp.MPPReader'),
        'File': JClass('java.io.File'),
        'DateTimeFormatter': JClass('java.time.format.DateTimeFormatter'),
    }


def _java_date_to_string(java_date, formatter=None) -> str:
    """JavaのLocalDateTime をPython文字列に変換"""
    if not java_date:
        return '-'
    
    try:
        if formatter is None:
            # フォーマッタを作成
            from jpype import JClass
            fmt = JClass('java.time.format.DateTimeFormatter')
            formatter = fmt.ofPattern('yyyy/MM/dd')
        
        return str(java_date.format(formatter))
    except Exception as e:
        debug_print(_(f"日時フォーマットエラー: {e}"))
        return '-'


def _java_duration_to_string(java_duration) -> str:
    """JavaのDurationをPython文字列に変換"""
    if java_duration is None:
        return '-'
    
    try:
        value = java_duration.getDuration()
        units = str(java_duration.getUnits())
        
        # ユニット名をマッピング（国際化対応）
        unit_map = {
            'd': _('日'), 'ed': _('日'),
            'h': _('時間'), 'eh': _('時間'),
            'w': _('週'), 'ew': _('週'),
            'mo': _('ヶ月'), 'emo': _('ヶ月'),
            'm': _('分'), 'em': _('分'),
            'y': _('年'), 'ey': _('年'),
        }
        label = unit_map.get(units, units)
        
        if value == int(value):
            return f"{int(value)}{label}"
        return f"{value:.1f}{label}"
    except Exception as e:
        debug_print(_(f"期間フォーマットエラー: {e}"))
        return '-'


def _java_percent_to_string(java_percent) -> str:
    """JavaのNumber (百分比)をPython文字列に変換"""
    if java_percent is None:
        return '-'
    
    try:
        value = float(java_percent)
        if value == int(value):
            return f"{int(value)}%"
        return f"{value:.1f}%"
    except Exception as e:
        debug_print(_(f"パーセント変換エラー: {e}"))
        return '-'


def read_project_mpxj(file_path: str) -> dict:
    """MS Projectファイルを読み込み、Python辞書として返す

    JPypeを使用してmpxjライブラリを直接呼び出し、
    Python辞書フォーマットでプロジェクト情報を返す。

    Args:
        file_path: MS Projectファイルのパス

    Returns:
        プロジェクト情報の辞書

    Raises:
        RuntimeError: ファイル読み込みエラー
        FileNotFoundError: ファイルが見つからない
        ImportError: mpxjが見つからない
    """
    _init_jpype()
    
    import os
    abs_path = os.path.abspath(file_path)
    
    if not os.path.exists(abs_path):
        raise FileNotFoundError(_(f"ファイルが見つかりません: {abs_path}"))
    
    debug_print(_(f"MPP読み込み: {abs_path}"))
    
    try:
        from jpype import JClass
        classes = _get_java_classes()
        
        # ファイルを開く
        file_obj = classes['File'](abs_path)
        reader = classes['UniversalProjectReader']()
        
        # Presentation data を無効化（AWTエラー回避）
        proxy = reader.getProjectReaderProxy(file_obj)
        java_reader = proxy.getProjectReader()
        
        # MPPReader の場合は presentation data 無効化
        if isinstance(java_reader, classes['MPPReader']):
            debug_print(_("MPPReader detected - disabling presentation data"))
            java_reader.setReadPresentationData(False)
            # RTF パーサーキットが無い場合の対応
            try:
                java_reader.setPreserveNullTasks(True)
            except Exception:
                pass  # RTFパーサキットなくても動作
        
        # プロジェクト読み込み
        project = proxy.read()
        
        if not project:
            raise RuntimeError(_(f"ファイルの読み込みに失敗: {abs_path}"))
        
        # Python辞書に変換
        result = {
            'project': _extract_project_info(project),
            'tasks': _extract_tasks(project),
            'resources': _extract_resources(project),
        }
        
        debug_print(f"抽出完了: タスク{len(result['tasks'])}件, "
                   f"リソース{len(result['resources'])}件")
        return result
        
    except Exception as e:
        raise RuntimeError(_(f"MPP読み込みエラー: {e}"))


def _extract_project_info(project) -> dict:
    """ProjectFile → dict変換"""
    try:
        props = project.getProjectProperties()
        return {
            'title': str(props.getProjectTitle() or ''),
            'author': str(props.getAuthor() or ''),
            'start': _java_date_to_string(props.getStartDate()),
            'finish': _java_date_to_string(props.getFinishDate()),
        }
    except Exception as e:
        debug_print(_(f"プロジェクト情報抽出エラー: {e}"))
        return {'title': '', 'author': '', 'start': '-', 'finish': '-'}


def _extract_tasks(project) -> list[dict]:
    """ProjectFile.getChildTasks() → list[dict]変換"""
    tasks = []
    
    def process_tasks(task_list, parent_tasks=None):
        if not task_list:
            return
        for task in task_list:
            try:
                # ネイティブバイナリ版と同じロジック: タスク名がnullの場合はスキップ
                task_name = task.getName()
                if not task_name:
                    continue
                
                task_dict = _task_to_dict(task)
                tasks.append(task_dict)
                # 子タスクを再帰処理
                child_tasks = task.getChildTasks()
                if child_tasks:
                    process_tasks(child_tasks)
            except Exception as e:
                debug_print(f"タスク抽出エラー: {e}")
    
    try:
        process_tasks(project.getChildTasks())
    except Exception as e:
        debug_print(f"タスク抽出エラー: {e}")
    
    return tasks


def _task_to_dict(task) -> dict:
    """Task → dict変換"""
    # 依存関係を抽出
    predecessors = []
    try:
        preds = task.getPredecessors()
        if preds:
            for rel in preds:
                try:
                    pred_task = rel.getTargetTask()
                    if pred_task:
                        pred_id = pred_task.getID()
                        if pred_id:
                            predecessors.append({
                                'task_id': int(pred_id),
                                'type': str(rel.getType() or 'FS'),
                                'lag': _java_duration_to_string(rel.getLag()),
                            })
                except Exception as e:
                    debug_print(_(f"依存関係抽出エラー: {e}"))
    except Exception as e:
        debug_print(_(f"getPredecessorsエラー: {e}"))
    
    # 担当者を抽出
    resources_str = '-'
    try:
        assignments = task.getResourceAssignments()
        if assignments:
            res_names = []
            for assign in assignments:
                res = assign.getResource()
                if res and res.getName():
                    res_names.append(str(res.getName()))
            if res_names:
                resources_str = ', '.join(res_names)
    except Exception as e:
        debug_print(f"リソース抽出エラー: {e}")
    
    # サマリータスク判定
    is_summary = False
    try:
        child_tasks = task.getChildTasks()
        is_summary = bool(child_tasks)
    except Exception as e:
        debug_print(_(f"サマリー判定エラー: {e}"))
    
    return {
        'task_id': int(task.getID() or 0),
        'unique_id': int(task.getUniqueID() or 0),
        'name': str(task.getName() or 'Unnamed'),
        'outline_level': int(task.getOutlineLevel() or 0),
        'start': _java_date_to_string(task.getStart()),
        'finish': _java_date_to_string(task.getFinish()),
        'duration': _java_duration_to_string(task.getDuration()),
        'percent_complete': _java_percent_to_string(task.getPercentageComplete()),
        'resources': resources_str,
        'is_summary': is_summary,
        'predecessors': predecessors,
    }


def _extract_resources(project) -> list[dict]:
    """ProjectFile.getResources() → list[dict]変換"""
    resources = []
    
    try:
        for res in project.getResources():
            try:
                res_id = res.getID()
                res_name = res.getName()
                if res_name:
                    resources.append({
                        'id': int(res_id) if res_id else 0,
                        'name': str(res_name),
                    })
            except Exception as e:
                debug_print(_(f"リソース抽出エラー: {e}"))
    except Exception as e:
        debug_print(_(f"getResourcesエラー: {e}"))
    
    return resources


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
    header = _("| ID | タスク名 | 開始日 | 終了日 | 期間 | 進捗 | 担当者 | 依存タスク |")
    separator = "| --- | --- | --- | --- | --- | --- | --- | --- |"
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
        
        # task_idを取得
        task_id = task.get('task_id', 0)
        
        # predecessorsを整形
        preds = task.get('predecessors', [])
        if preds:
            pred_strs = [f"{p['task_id']}({p['type']})" for p in preds]
            pred_text = _escape_md_cell(", ".join(pred_strs))
        else:
            pred_text = "-"

        row = (
            f"| {task_id}"
            f" | {indent}{name}"
            f" | {task.get('start', '-')}"
            f" | {task.get('finish', '-')}"
            f" | {task.get('duration', '-')}"
            f" | {task.get('percent_complete', '-')}"
            f" | {resources}"
            f" | {pred_text} |"
        )
        rows.append(row)

    return "\n".join(rows)


def resources_to_markdown_table(resources: list[dict], tasks: list[dict] = None) -> str:
    """リソース情報をMarkdownテーブルに変換する

    Args:
        resources: リソース情報の辞書リスト
        tasks: タスク情報の辞書リスト（リソース毎の統計計算用）

    Returns:
        Markdownテーブル文字列
    """
    if not resources:
        return ""

    # リソース毎の統計を計算
    resource_stats = {}
    if tasks:
        for task in tasks:
            task_resources = task.get('resources', '-')
            if task_resources and task_resources != '-':
                # リソース名をカンマで分割（複数リソース対応）
                res_names = [r.strip() for r in task_resources.split(',')]
                
                # タスク期間を数値に変換
                duration_str = task.get('duration', '-')
                duration_days = 0
                if duration_str and duration_str != '-':
                    try:
                        # "124日" から "124" を抽出
                        match = re.match(r'(\d+(?:\.\d+)?)', duration_str)
                        if match:
                            duration_days = float(match.group(1))
                    except Exception as e:
                        debug_print(f"期間パース失敗: {duration_str} - {e}")
                
                # リソース毎に集計
                for res_name in res_names:
                    if res_name not in resource_stats:
                        resource_stats[res_name] = {'count': 0, 'days': 0.0}
                    resource_stats[res_name]['count'] += 1
                    resource_stats[res_name]['days'] += duration_days

    header = _("| ID | リソース名 | 割当タスク数 | 合計日数 |")
    separator = "| --- | --- | --- | --- |"
    rows = [header, separator]

    for res in resources:
        name = _escape_md_cell(res['name'])
        
        # 統計情報を取得
        stats = resource_stats.get(res['name'], {'count': 0, 'days': 0.0})
        task_count = stats['count']
        total_days = stats['days']
        
        # 日数をフォーマット
        if total_days == int(total_days):
            days_str = f"{int(total_days)}{_('日')}"
        else:
            days_str = f"{total_days:.1f}{_('日')}"
        
        rows.append(f"| {res['id']} | {name} | {task_count} | {days_str} |")

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
        lines.append(f"- **{_('プロジェクト名')}**: {title}")
    author = info.get('author', '')
    if author:
        lines.append(f"- **{_('作成者')}**: {author}")
    start = info.get('start', '')
    if start and start != '-':
        lines.append(f"- **{_('開始日')}**: {start}")
    finish = info.get('finish', '')
    if finish and finish != '-':
        lines.append(f"- **{_('終了日')}**: {finish}")
    return "\n".join(lines)


class MppToMarkdownConverter:
    """MS ProjectファイルをMarkdownに変換するコンバータ (JPype版)

    JPypeを使用してmpxjライブラリを直接呼び出し、
    MS Projectファイル(.mpp/.mpt/.mpx)を解析して
    タスク一覧テーブルとリソース一覧テーブルを含むMarkdownを出力する。
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
        """このコンバータが自動付与する見出しの正規表現パターンを返す（国際化対応）"""
        import re
        
        # _build_markdown() と同じメッセージを使用（翻訳ファイルで管理）
        heading_project = _("## プロジェクト情報").replace("## ", "")
        heading_task = _("## タスク一覧").replace("## ", "")
        heading_resource = _("## リソース一覧").replace("## ", "")
        
        return [
            re.compile(r'^' + re.escape(self.base_name) + r'$'),
            re.compile(r'^' + re.escape(heading_task) + r'$'),
            re.compile(r'^' + re.escape(heading_resource) + r'$'),
            re.compile(r'^' + re.escape(heading_project) + r'$'),
        ]

    def get_auto_generated_html_tags(self) -> list:
        """このコンバータが自動付与するHTMLタグを返す"""
        return []

    def get_auto_generated_line_patterns(self) -> list:
        """このコンバータが自動付与するメタデータ行パターンを返す（国際化対応）"""
        import re
        
        # project_info_to_markdown() と同じメッセージを使用（翻訳ファイルで管理）
        label_project = _('プロジェクト名')
        label_author = _('作成者')
        label_start = _('開始日')
        label_finish = _('終了日')
        
        return [
            re.compile(r'^- \*\*' + re.escape(label_project) + r'\*\*:'),
            re.compile(r'^- \*\*' + re.escape(label_author) + r'\*\*:'),
            re.compile(r'^- \*\*' + re.escape(label_start) + r'\*\*:'),
            re.compile(r'^- \*\*' + re.escape(label_finish) + r'\*\*:'),
        ]

    def convert(self) -> str:
        """変換メイン処理

        Returns:
            出力ファイルのパス（.mdまたは.txt）
        """
        from o2md.utils import is_text_only

        print(_("MS Project変換開始: {file}").format(file=self.file_path))

        # mpxjでJSON取得
        data = read_project_mpxj(self.file_path)

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
            lines.append(_("## プロジェクト情報"))
            lines.append("")
            lines.append(info_md)
            lines.append("")

        # タスク一覧
        if tasks:
            lines.append(_("## タスク一覧"))
            lines.append("")
            lines.append(tasks_to_markdown_table(tasks))
            lines.append("")

        # リソース一覧
        if resources:
            lines.append(_("## リソース一覧"))
            lines.append("")
            lines.append(resources_to_markdown_table(resources, tasks))
            lines.append("")

        return lines

    def _write_markdown(self, content: str) -> str:
        """Markdownファイルに書き込む"""
        output_path = os.path.join(self.output_dir, f"{self.base_name}.md")
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(content)
        return output_path

    def _write_text(self, content: str) -> str:
        """Markdown書式を除去してテキストファイルに書き込む"""
        from o2md.o2md import strip_markdown
        auto_patterns = {
            'heading_patterns': self.get_auto_generated_patterns(),
            'html_tags': self.get_auto_generated_html_tags(),
            'line_patterns': self.get_auto_generated_line_patterns(),
        }
        text_content = strip_markdown(content, auto_patterns=auto_patterns)
        output_path = os.path.join(self.output_dir, f"{self.base_name}.txt")
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(text_content)
        return output_path


def main():
    """メイン関数 - コマンドラインから実行される"""
    import argparse
    import sys

    # i18n初期化
    setup_i18n()

    parser = argparse.ArgumentParser(
        description=_('MS Projectファイル(.mpp/.mpt/.mpx)をMarkdownに変換')
    )
    parser.add_argument(
        'mpp_file',
        help=_('変換するMS Projectファイル')
    )
    parser.add_argument(
        '-o', '--output-dir',
        type=str,
        help=_('出力ディレクトリを指定（デフォルト: ./output）')
    )
    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help=_('デバッグ情報を出力')
    )
    parser.add_argument(
        '--text',
        action='store_true',
        help=_('.mdと.txtの両方を出力（プレーンテキスト変換）')
    )

    args = parser.parse_args()

    set_verbose(args.verbose)

    if not os.path.exists(args.mpp_file):
        print(_("エラー: ファイル '{file}' が見つかりません。").format(file=args.mpp_file))
        sys.exit(1)

    # ファイル拡張子チェック
    valid_extensions = ('.mpp', '.mpt', '.mpx', '.xml')
    if not args.mpp_file.lower().endswith(valid_extensions):
        ext_list = ', '.join(valid_extensions)
        print(_("エラー: {ext} 形式のファイルを指定してください。").format(ext=ext_list))
        sys.exit(1)

    try:
        converter = MppToMarkdownConverter(
            args.mpp_file,
            output_dir=args.output_dir
        )

        # テキストモード設定
        if args.text:
            from o2md.utils import set_text_only
            set_text_only(True)

        output_file = converter.convert()

        print("\n" + "=" * 50)
        print(_("出力ファイル: {output_file}").format(output_file=output_file))
        print("=" * 50)

    except FileNotFoundError as e:
        print(_("エラー: {error}").format(error=e))
        sys.exit(1)
    except Exception as e:
        print(_("エラー: {error}").format(error=e))
        if args.verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)
