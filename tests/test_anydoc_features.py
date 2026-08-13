"""MSPDIとWord拡張変換機能の回帰テスト。"""

import tempfile
import zipfile
import io
import sys
from pathlib import Path

import pytest
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from o2md.d2md_fields import convert_paragraph
from o2md.d2md_notes import NoteManager
from o2md.d2md_numbering import NumberingResolver
from o2md.d2md_tables import render_table
from o2md.filter import detect_type_from_bytes
import o2md.filter as filter_cli
from o2md.mspdi import is_mspdi_xml
from o2md.o2md import detect_file_type
from o2md.x2md import ExcelToMarkdownConverter


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _numbering_xml():
    return f"""<w:numbering xmlns:w="{W}">
      <w:abstractNum w:abstractNumId="1">
        <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
        <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1.%2."/></w:lvl>
      </w:abstractNum>
      <w:num w:numId="9"><w:abstractNumId w:val="1"/></w:num>
    </w:numbering>""".encode()


def _paragraph_with_num(level, number_id="9"):
    document = Document()
    paragraph = document.add_paragraph("項目")
    num_pr = OxmlElement("w:numPr")
    ilvl = OxmlElement("w:ilvl")
    ilvl.set(qn("w:val"), str(level))
    num_id = OxmlElement("w:numId")
    num_id.set(qn("w:val"), number_id)
    num_pr.extend([ilvl, num_id])
    paragraph._p.get_or_add_pPr().append(num_pr)
    return paragraph


def test_numbering_resolver_tracks_nested_counters():
    resolver = NumberingResolver(_numbering_xml())
    assert resolver.marker(_paragraph_with_num(0))[1] == "1."
    assert resolver.marker(_paragraph_with_num(1))[1] == "1.a."
    assert resolver.marker(_paragraph_with_num(1))[1] == "1.b."
    assert resolver.marker(_paragraph_with_num(0))[1] == "2."
    assert resolver.marker(_paragraph_with_num(1))[1] == "2.a."


def test_numbering_resolver_honors_start_override_and_lvl_restart():
    blob = f"""<w:numbering xmlns:w="{W}">
      <w:abstractNum w:abstractNumId="2">
        <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
        <w:lvl w:ilvl="1"><w:start w:val="1"/><w:lvlRestart w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2."/></w:lvl>
      </w:abstractNum>
      <w:num w:numId="10"><w:abstractNumId w:val="2"/>
        <w:lvlOverride w:ilvl="0"><w:startOverride w:val="3"/></w:lvlOverride>
      </w:num>
    </w:numbering>""".encode()
    resolver = NumberingResolver(blob)
    assert resolver.marker(_paragraph_with_num(0, "10"))[1] == "3."
    assert resolver.marker(_paragraph_with_num(1, "10"))[1] == "3.1."
    assert resolver.marker(_paragraph_with_num(1, "10"))[1] == "3.2."
    paragraph = _paragraph_with_num(0, "10")
    assert resolver.marker(paragraph)[1] == "4."
    paragraph = _paragraph_with_num(1, "10")
    assert resolver.marker(paragraph)[1] == "4.1."


def test_unknown_field_result_is_kept_and_hyperlink_is_rendered():
    document = Document()
    paragraph = document.add_paragraph()
    begin = OxmlElement("w:r")
    begin_char = OxmlElement("w:fldChar")
    begin_char.set(qn("w:fldCharType"), "begin")
    begin.append(begin_char)
    instruction = OxmlElement("w:r")
    instr_text = OxmlElement("w:instrText")
    instr_text.text = 'HYPERLINK "https://example.test"'
    instruction.append(instr_text)
    separate = OxmlElement("w:r")
    separate_char = OxmlElement("w:fldChar")
    separate_char.set(qn("w:fldCharType"), "separate")
    separate.append(separate_char)
    result = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = "リンク"
    result.append(text)
    end = OxmlElement("w:r")
    end_char = OxmlElement("w:fldChar")
    end_char.set(qn("w:fldCharType"), "end")
    end.append(end_char)
    paragraph._p.extend([begin, instruction, separate, result, end])
    assert convert_paragraph(paragraph) == "[リンク](https://example.test)"


def test_field_paragraph_keeps_bold_run_with_note_reference(tmp_path):
    path = tmp_path / "source.docx"
    document = Document()
    paragraph = document.add_paragraph()
    run = paragraph.add_run("太字")
    run.bold = True
    reference_run = paragraph.add_run()
    reference = OxmlElement("w:footnoteReference")
    reference.set(qn("w:id"), "1")
    reference_run._r.append(reference)
    document.save(path)
    rebuilt = tmp_path / "with-note.docx"
    footnotes = f"""<w:footnotes xmlns:w="{W}">
      <w:footnote w:id="1"><w:p><w:r><w:t>脚注本文</w:t></w:r></w:p></w:footnote>
    </w:footnotes>""".encode()
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(rebuilt, "w") as target:
        for item in source.infolist():
            target.writestr(item, source.read(item.filename))
        target.writestr("word/footnotes.xml", footnotes)
    from o2md.d2md import WordToMarkdownConverter

    result = WordToMarkdownConverter(str(rebuilt), output_dir=str(tmp_path)).convert()
    output = Path(result).read_text()
    assert "**太字**[^fn1]" in output


def test_broken_optional_docx_part_keeps_body(tmp_path):
    path = tmp_path / "source.docx"
    document = Document()
    document.add_paragraph("本文")
    document.save(path)
    rebuilt = tmp_path / "broken.docx"
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(rebuilt, "w") as target:
        for item in source.infolist():
            if item.filename == "word/numbering.xml":
                target.writestr(item, b"<broken")
            else:
                target.writestr(item, source.read(item.filename))
    from o2md.d2md import WordToMarkdownConverter

    result = WordToMarkdownConverter(str(rebuilt), output_dir=str(tmp_path)).convert()
    assert "本文" in Path(result).read_text()


def test_broken_docx_and_degradation_failure_preserves_cause(tmp_path):
    path = tmp_path / "broken.docx"
    path.write_bytes(b"not a zip archive")
    from o2md.d2md import WordToMarkdownConverter

    converter = WordToMarkdownConverter.__new__(WordToMarkdownConverter)
    with pytest.raises(RuntimeError, match="DOCX本文の読み込みに失敗しました") as error:
        converter._load_document_with_optional_degradation(str(path))
    assert error.value.__cause__ is not None
    assert not isinstance(error.value.__cause__, NameError)


def test_table_cell_formatting_preserves_runs_and_fields(tmp_path):
    path = tmp_path / "table-cell.docx"
    document = Document()
    table = document.add_table(rows=1, cols=1)
    cell = table.cell(0, 0)
    paragraph = cell.paragraphs[0]
    bold = paragraph.add_run("太字")
    bold.bold = True
    hidden = paragraph.add_run("隠し文字")
    hidden._r.get_or_add_rPr().append(OxmlElement("w:vanish"))
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), " DATE ")
    result = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = "フィールド結果"
    result.append(text)
    field.append(result)
    paragraph._p.append(field)
    document.save(path)

    from o2md.d2md import WordToMarkdownConverter

    output_path = WordToMarkdownConverter(
        str(path), output_dir=str(tmp_path)
    ).convert()
    output = Path(output_path).read_text()
    assert "**太字**" in output
    assert "隠し文字" not in output
    assert "フィールド結果" in output


def test_table_cell_line_breaks_are_markdown_safe(tmp_path):
    path = tmp_path / "table-cell-breaks.docx"
    document = Document()
    table = document.add_table(rows=1, cols=1)
    paragraph = table.cell(0, 0).paragraphs[0]
    paragraph.add_run("行内")
    paragraph.add_run().add_break()
    paragraph.add_run("改行")
    table.cell(0, 0).add_paragraph("段落")
    document.save(path)

    from o2md.d2md import WordToMarkdownConverter

    output_path = WordToMarkdownConverter(
        str(path), output_dir=str(tmp_path)
    ).convert()
    output = Path(output_path).read_text()
    assert "| 行内<br>改行<br>段落 |" in output
    assert "行内\n改行" not in output


def test_table_keeps_legacy_spacing_after_table(tmp_path):
    path = tmp_path / "table-spacing.docx"
    document = Document()
    table = document.add_table(rows=1, cols=1)
    table.cell(0, 0).text = "表"
    document.add_paragraph("後続段落")
    document.save(path)

    from o2md.d2md import WordToMarkdownConverter

    output_path = WordToMarkdownConverter(
        str(path), output_dir=str(tmp_path)
    ).convert()
    output = Path(output_path).read_text()
    assert "| 表 |" in output
    assert "| 表 |\n| --- |\n\n\n後続段落" in output


def test_mspdi_content_sniffing_is_not_extension_only():
    data = f'<Project xmlns="{W.replace("openxmlformats.org/wordprocessingml/2006/main", "schemas.microsoft.com/project")}"/>'.encode()
    assert is_mspdi_xml(data)
    assert detect_type_from_bytes(data) == "msproject"
    assert detect_type_from_bytes(b"<Project xmlns='http://example.test'/>") == "unknown"


def test_notes_manager_maps_references_and_definitions(tmp_path):
    document = Document()
    path = tmp_path / "notes.docx"
    document.save(path)
    footnotes = f"""<w:footnotes xmlns:w="{W}">
      <w:footnote w:id="-1" w:type="separator"/>
      <w:footnote w:id="1"><w:p><w:r><w:t>脚注本文</w:t></w:r></w:p></w:footnote>
    </w:footnotes>""".encode()
    rebuilt = tmp_path / "notes-with-footnote.docx"
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(rebuilt, "w") as target:
        for item in source.infolist():
            target.writestr(item, source.read(item.filename))
        target.writestr("word/footnotes.xml", footnotes)
    manager = NoteManager(str(rebuilt))
    assert manager.reference("fn", "1") == "[^fn1]"
    assert manager.definitions() == ["[^fn1]: 脚注本文"]


def test_word_table_grid_keeps_column_count():
    document = Document()
    table = document.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "見出し"
    table.cell(0, 1).text = "列2"
    table.cell(1, 0).text = "値"
    lines = render_table(table, lambda cell: cell.text)
    assert lines[0] == "| 見出し | 列2 |"
    assert all(line.count("|") == 3 for line in lines)


def test_word_table_grid_without_tbl_grid_keeps_columns():
    document = Document()
    table = document.add_table(rows=2, cols=3)
    table.cell(0, 0).text = "A"
    table.cell(0, 1).text = "B"
    table.cell(0, 2).text = "C"
    table._tbl.remove(table._tbl.tblGrid)
    lines = render_table(table, lambda cell: cell.text)
    assert all(line.count("|") == 4 for line in lines)


def test_word_table_complex_merges_and_headers_keep_grid():
    document = Document()
    table = document.add_table(rows=5, cols=4)
    for row_index, row in enumerate(table.rows):
        for col_index, cell in enumerate(row.cells):
            cell.text = f"{row_index}-{col_index}"
    rows = table._tbl.tr_lst
    for row in rows[:2]:
        tr_pr = row.get_or_add_trPr()
        header = OxmlElement("w:tblHeader")
        tr_pr.append(header)
    first = rows[0].tc_lst[0]
    first.get_or_add_tcPr().append(OxmlElement("w:gridSpan"))
    first.tcPr.gridSpan.set(qn("w:val"), "2")
    rows[0].remove(rows[0].tc_lst[1])
    row1_pr = rows[1].get_or_add_trPr()
    before = OxmlElement("w:gridBefore")
    before.set(qn("w:val"), "1")
    after = OxmlElement("w:gridAfter")
    after.set(qn("w:val"), "1")
    row1_pr.extend([before, after])
    for row in (rows[2], rows[3], rows[4]):
        merge = OxmlElement("w:vMerge")
        merge.set(qn("w:val"), "restart" if row is rows[2] else "continue")
        row.tc_lst[0].get_or_add_tcPr().append(merge)
    span = OxmlElement("w:gridSpan")
    span.set(qn("w:val"), "2")
    rows[3].tc_lst[1].get_or_add_tcPr().append(span)
    rows[3].remove(rows[3].tc_lst[2])
    lines = render_table(table, lambda cell: cell.text)
    assert all(line.count("|") == 5 for line in lines)
    assert "0-0" in lines[0]
    assert "0-2" in lines[0]


def test_word_table_header_inherits_horizontal_merge_labels():
    document = Document()
    table = document.add_table(rows=2, cols=4)
    values = [
        ("見出しA", "横結合対象", "見出しC", "見出しD"),
        ("補助見出しA", "補助見出しB", "補助見出しC", "補助見出しD"),
    ]
    for row, values_row in zip(table.rows, values):
        for cell, value in zip(row.cells, values_row):
            cell.text = value
    rows = table._tbl.tr_lst
    for row in rows:
        tr_pr = row.get_or_add_trPr()
        tr_pr.append(OxmlElement("w:tblHeader"))
    table._tbl.tblGrid.append(OxmlElement("w:gridCol"))
    first = rows[0].tc_lst[0]
    first.get_or_add_tcPr().append(OxmlElement("w:gridSpan"))
    first.tcPr.gridSpan.set(qn("w:val"), "2")
    rows[0].remove(rows[0].tc_lst[1])
    tr_pr = rows[1].get_or_add_trPr()
    before = OxmlElement("w:gridBefore")
    before.set(qn("w:val"), "1")
    after = OxmlElement("w:gridAfter")
    after.set(qn("w:val"), "0")
    tr_pr.extend([before, after])

    lines = render_table(table, lambda cell: cell.text)

    assert lines[0] == (
        "| 見出しA | 見出しA 補助見出しA | 見出しC 補助見出しB | "
        "見出しD 補助見出しC | 補助見出しD |"
    )


def test_excel_sheet_failure_isolated_and_all_failure_raises(tmp_path):
    from openpyxl import Workbook

    path = tmp_path / "sheets.xlsx"
    workbook = Workbook()
    workbook.active["A1"] = "first"
    workbook.create_sheet("Second")["A1"] = "second"
    workbook.save(path)
    converter = ExcelToMarkdownConverter(str(path), output_dir=str(tmp_path))
    original = converter._convert_sheet
    calls = {"count": 0}

    def fail_first(sheet):
        calls["count"] += 1
        if calls["count"] == 1:
            raise RuntimeError("壊れたシート")
        return original(sheet)

    converter._convert_sheet = fail_first
    output = Path(converter.convert()).read_text()
    assert "second" in output

    converter = ExcelToMarkdownConverter(str(path), output_dir=str(tmp_path))
    converter._convert_sheet = lambda sheet: (_ for _ in ()).throw(RuntimeError("壊れたシート"))
    with pytest.raises(RuntimeError, match="すべてのシート"):
        converter.convert()


def test_mspdi_file_path_detection_and_plain_xml(tmp_path):
    project = tmp_path / "project.xml"
    project.write_text('<Project xmlns="http://schemas.microsoft.com/project"/>')
    plain = tmp_path / "plain.xml"
    plain.write_text("<root/>")
    assert detect_file_type(str(project)) == "msproject"
    assert detect_file_type(str(plain)) == "unknown"


def test_filter_file_and_stdin_routes_sniff_mspdi(tmp_path, monkeypatch, capsys):
    project = tmp_path / "project.xml"
    data = b'<Project xmlns="http://schemas.microsoft.com/project"/>'
    project.write_bytes(data)
    monkeypatch.setattr(filter_cli, "filter_file", lambda path, **kwargs: "変換結果")

    monkeypatch.setattr(sys, "argv", ["o2md-filter", str(project)])
    filter_cli.main()
    assert capsys.readouterr().out == "変換結果"

    monkeypatch.setattr(sys, "argv", ["o2md-filter"])
    monkeypatch.setattr(sys, "stdin", io.TextIOWrapper(io.BytesIO(data)))
    filter_cli.main()
    assert capsys.readouterr().out == "変換結果"
