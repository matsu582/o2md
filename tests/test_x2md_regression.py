"""x2md の変換結果を期待出力と完全一致で比較する回帰テスト

期待値は ``tests/expected/x2md/<stem>/`` 以下に配置する:

- ``<stem>.md`` … 変換結果の Markdown 全文（完全一致）
- ``images.txt`` … 生成画像ファイル名のソート済み一覧（1行1件。
  画像バイナリはレンダリング環境差があるため比較しない）
- ``debug_workbooks.txt`` … debug_workbooks が生成される場合のみ。
  各 xlsx のファイル名と zip 内エントリ名のソート済み一覧
  （``docProps/core.xml`` はタイムスタンプを含むため除外）

期待値の更新方法::

    O2MD_UPDATE_EXPECTED=1 uv run pytest tests/test_x2md_regression.py -q

追加の検証ファイル::

    O2MD_REGRESSION_EXTRA_DIR=<dir> を設定すると、<dir> 内の *.xlsx についても
    同じテストを実行する。期待値は <dir>/<stem>/ 以下に同じ3ファイル構成で置く。
    環境変数が未設定、またはディレクトリが存在しない場合は skip する。
    （このディレクトリはリポジトリには含めない運用を想定）
"""

import os
import zipfile
from pathlib import Path

import pytest

from o2md.x2md import ExcelToMarkdownConverter, set_verbose

REPO_ROOT = Path(__file__).parent.parent
INPUT_DIR = REPO_ROOT / "input_files"
EXPECTED_ROOT = Path(__file__).parent / "expected" / "x2md"

CORE_XML = "docProps/core.xml"


def _input_files():
    return sorted(INPUT_DIR.glob("*.xlsx"))


def _extra_files():
    """環境変数で指定された追加 xlsx の一覧と、その期待値ルートを返す"""
    extra_dir = os.environ.get("O2MD_REGRESSION_EXTRA_DIR")
    if not extra_dir:
        return None, []
    d = Path(extra_dir)
    if not d.is_dir():
        return d, []
    return d, sorted(d.glob("*.xlsx"))


def _update_expected():
    return os.environ.get("O2MD_UPDATE_EXPECTED") == "1"


def _collect_debug_workbooks(output_dir: Path) -> str:
    """debug_workbooks 内の xlsx について、ファイル名と zip エントリ一覧を記録する"""
    dbg = output_dir / "debug_workbooks"
    if not dbg.is_dir():
        return ""
    lines = []
    for xlsx in sorted(dbg.glob("*.xlsx")):
        lines.append(xlsx.name)
        with zipfile.ZipFile(xlsx) as z:
            for name in sorted(z.namelist()):
                if name == CORE_XML:
                    continue
                lines.append(f"  {name}")
    return "\n".join(lines) + "\n"


def _run_conversion(xlsx_path: Path, tmp_path: Path):
    """変換を実行し (markdown 文字列, images 一覧, debug_workbooks 一覧) を返す"""
    output_dir = tmp_path / "out"
    output_dir.mkdir(parents=True, exist_ok=True)
    set_verbose(True)
    try:
        converter = ExcelToMarkdownConverter(str(xlsx_path), output_dir=str(output_dir))
        md_path = Path(converter.convert())
    finally:
        set_verbose(False)
    md_text = md_path.read_text(encoding="utf-8")
    # 出力ディレクトリの絶対パスが埋め込まれている場合に備えて正規化
    md_text = md_text.replace(str(output_dir), "<OUTPUT_DIR>")
    md_text = md_text.replace(str(tmp_path), "<OUTPUT_DIR>")

    images_dir = output_dir / "images"
    image_names = []
    if images_dir.is_dir():
        image_names = sorted(p.name for p in images_dir.iterdir() if p.is_file())
    images_txt = "\n".join(image_names)
    if images_txt:
        images_txt += "\n"

    dbg_txt = _collect_debug_workbooks(output_dir)
    return md_text, images_txt, dbg_txt


def _assert_or_update(expected_dir: Path, stem: str, md_text: str, images_txt: str, dbg_txt: str):
    """期待値と比較。O2MD_UPDATE_EXPECTED=1 なら期待値を書き出す"""
    if _update_expected():
        expected_dir.mkdir(parents=True, exist_ok=True)
        (expected_dir / f"{stem}.md").write_text(md_text, encoding="utf-8")
        (expected_dir / "images.txt").write_text(images_txt, encoding="utf-8")
        dbg_file = expected_dir / "debug_workbooks.txt"
        if dbg_txt:
            dbg_file.write_text(dbg_txt, encoding="utf-8")
        elif dbg_file.exists():
            dbg_file.unlink()
        return

    exp_md = expected_dir / f"{stem}.md"
    assert exp_md.exists(), (
        f"期待値がありません: {exp_md}\n"
        f"O2MD_UPDATE_EXPECTED=1 で生成してください"
    )
    assert md_text == exp_md.read_text(encoding="utf-8"), \
        f"{stem}.md が期待値と一致しません"

    exp_images = expected_dir / "images.txt"
    expected_images = exp_images.read_text(encoding="utf-8") if exp_images.exists() else ""
    assert images_txt == expected_images, \
        f"{stem}: 生成画像ファイル一覧が期待値と一致しません"

    exp_dbg = expected_dir / "debug_workbooks.txt"
    expected_dbg = exp_dbg.read_text(encoding="utf-8") if exp_dbg.exists() else ""
    assert dbg_txt == expected_dbg, \
        f"{stem}: debug_workbooks の構成が期待値と一致しません"


@pytest.mark.parametrize("xlsx_path", _input_files(), ids=lambda p: p.stem)
def test_x2md_output_matches_expected(xlsx_path, tmp_path):
    """input_files 内の各 xlsx の変換結果が期待出力と一致すること"""
    expected_dir = EXPECTED_ROOT / xlsx_path.stem
    md_text, images_txt, dbg_txt = _run_conversion(xlsx_path, tmp_path)
    _assert_or_update(expected_dir, xlsx_path.stem, md_text, images_txt, dbg_txt)


_extra_dir, _extra_xlsx = _extra_files()


@pytest.mark.parametrize(
    "xlsx_path",
    _extra_xlsx if _extra_xlsx else [
        pytest.param(None, marks=pytest.mark.skip(
            reason="O2MD_REGRESSION_EXTRA_DIR が未設定か追加 xlsx がありません"))
    ],
    ids=lambda p: p.stem if p else "no-extra",
)
def test_x2md_extra_output_matches_expected(xlsx_path, tmp_path):
    """O2MD_REGRESSION_EXTRA_DIR 内の追加 xlsx の変換結果が期待出力と一致すること"""
    expected_dir = Path(os.environ["O2MD_REGRESSION_EXTRA_DIR"]) / xlsx_path.stem
    md_text, images_txt, dbg_txt = _run_conversion(xlsx_path, tmp_path)
    _assert_or_update(expected_dir, xlsx_path.stem, md_text, images_txt, dbg_txt)
