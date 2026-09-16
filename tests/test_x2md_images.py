"""x2md の画像出力に関するテスト

分離グループ（iso_group）画像のファイル名がシート名を含み、
同一ブック内で図形インデックス集合が同じ複数シートがあっても
衝突しないことを検証する。
"""
import re
import zipfile
from pathlib import Path

import pytest

from o2md.utils import get_libreoffice_path

INPUT_XLSX = Path(__file__).parent.parent / "input_files" / "excel_shapes_sample.xlsx"


def _make_two_same_sheets_xlsx(src: Path, dst: Path) -> Path:
    """1シートの xlsx を複製して同じ図形を持つ2シート版を tmp に作る

    sheet1.xml / そのrels を sheet2 として複製し、workbook.xml・
    workbook.xml.rels・[Content_Types].xml に参照を追加する。
    """
    zin = zipfile.ZipFile(src)
    wb = zin.read("xl/workbook.xml").decode("utf-8")
    rels = zin.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    ct = zin.read("[Content_Types].xml").decode("utf-8")
    sheet_xml = zin.read("xl/worksheets/sheet1.xml")
    sheet_rels = zin.read("xl/worksheets/_rels/sheet1.xml.rels")

    # workbook.xml.rels 内の既存 rId 最大値を調べて新しい rId を採番する
    used_ids = [int(m) for m in re.findall(r'Id="rId(\d+)"', rels)]
    new_rid = max(used_ids) + 1

    m = re.search(r'name="([^"]+)"[^>]*sheetId="1"', wb)
    sheet_name = m.group(1)
    wb = wb.replace(
        "</sheets>",
        f'<sheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        f'name="{sheet_name}_copy" sheetId="2" state="visible" r:id="rId{new_rid}"/></sheets>',
    )
    rels = rels.replace(
        "</Relationships>",
        f'<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        f'Target="/xl/worksheets/sheet2.xml" Id="rId{new_rid}"/></Relationships>',
    )
    ct = ct.replace(
        "</Types>",
        '<Override PartName="/xl/worksheets/sheet2.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
    )

    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in zin.namelist():
            if name == "xl/workbook.xml":
                zout.writestr(name, wb)
            elif name == "xl/_rels/workbook.xml.rels":
                zout.writestr(name, rels)
            elif name == "[Content_Types].xml":
                zout.writestr(name, ct)
            else:
                zout.writestr(name, zin.read(name))
        zout.writestr("xl/worksheets/sheet2.xml", sheet_xml)
        zout.writestr("xl/worksheets/_rels/sheet2.xml.rels", sheet_rels)
    zin.close()
    return dst


@pytest.mark.skipif(
    get_libreoffice_path() == "soffice",
    reason="LibreOfficeがインストールされていません",
)
def test_iso_group_images_have_unique_names_per_sheet(tmp_path):
    """同一図形を持つ2シートで、シートごとに別の iso_group 画像が出力される"""
    from o2md.x2md import ExcelToMarkdownConverter

    xlsx = _make_two_same_sheets_xlsx(INPUT_XLSX, tmp_path / "two_same_sheets.xlsx")
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    md_path = ExcelToMarkdownConverter(str(xlsx), output_dir=str(output_dir)).convert()

    # (a) images/ に iso_group 画像が2枚ある（同名衝突で1枚に潰れない）
    images = list((output_dir / "images").iterdir())
    iso_images = [p for p in images if "_iso_group" in p.name]
    assert len(iso_images) == 2, f"iso_group画像が2枚ではありません: {[p.name for p in images]}"

    # (b) md に両シートそれぞれの画像参照が1つずつ含まれる
    md_text = Path(md_path).read_text(encoding="utf-8")
    refs = re.findall(r"!\[[^\]]*\]\(images/[^)]+\)", md_text)
    iso_refs = [r for r in refs if "_iso_group" in r]
    assert len(iso_refs) == 2, f"iso_group画像参照が2件ではありません: {refs}"
    assert len(set(iso_refs)) == 2, f"両シートが同じ画像を参照しています: {iso_refs}"
