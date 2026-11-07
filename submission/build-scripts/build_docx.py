#!/usr/bin/env python3
"""Build 02-项目文档.docx from the competition template and Markdown content.

Usage: /tmp/hapdoc-env/bin/python submission/build-scripts/build_docx.py
"""
import os
import re
from copy import deepcopy

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import (WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT,
                            WD_TAB_LEADER)
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.normpath(os.path.join(HERE, ".."))
TEMPLATE = os.path.join(ROOT, "templates", "xxx团队-赛道一-开发设计文档模板.docx")
CONTENT = os.path.join(ROOT, "02-项目文档.md")
OUT = os.path.join(ROOT, "02-项目文档.docx")

BODY_FONT_EA = "宋体"
BODY_FONT_ASCII = "Times New Roman"
HEAD_FONT_EA = "黑体"
HEAD_FONT_ASCII = "Arial"
CODE_FONT = "Consolas"
ACCENT = RGBColor(0x1F, 0x3A, 0x5F)
GRAY = RGBColor(0x59, 0x59, 0x59)


def set_run_font(run, ascii_font=BODY_FONT_ASCII, ea_font=BODY_FONT_EA,
                 size=10.5, bold=None, color=None, italic=None):
    run.font.name = ascii_font
    run._element.rPr.rFonts.set(qn("w:eastAsia"), ea_font)
    run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color


def shade(element, fill):
    pPr = element.get_or_add_pPr() if hasattr(element, "get_or_add_pPr") else element
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    pPr.append(shd)


def cell_shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def set_cell_margins(table, top=40, bottom=40, left=80, right=80):
    tblPr = table._tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for name, value in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        node = OxmlElement(f"w:{name}")
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")
        mar.append(node)
    tblPr.append(mar)


def set_repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    header = OxmlElement("w:tblHeader")
    header.set(qn("w:val"), "true")
    trPr.append(header)


def add_field(paragraph, instr, placeholder="1"):
    run = paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    run._r.append(fld_begin)
    run2 = paragraph.add_run()
    instr_el = OxmlElement("w:instrText")
    instr_el.set(qn("xml:space"), "preserve")
    instr_el.text = instr
    run2._r.append(instr_el)
    run3 = paragraph.add_run()
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    run3._r.append(fld_sep)
    run4 = paragraph.add_run(placeholder)
    run5 = paragraph.add_run()
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run5._r.append(fld_end)
    return [run, run2, run3, run4, run5]


def add_inline(paragraph, text, size=10.5, bold=False, color=None, ea_font=BODY_FONT_EA):
    """Parse **bold** and `code` inline spans."""
    pattern = re.compile(r"(\*\*.+?\*\*|`[^`]+`)")
    pos = 0
    for match in pattern.finditer(text):
        if match.start() > pos:
            run = paragraph.add_run(text[pos:match.start()])
            set_run_font(run, size=size, bold=bold, color=color, ea_font=ea_font)
        token = match.group(0)
        if token.startswith("**") and token.endswith("**"):
            run = paragraph.add_run(token[2:-2])
            set_run_font(run, size=size, bold=True, color=color, ea_font=ea_font)
        else:
            run = paragraph.add_run(token[1:-1])
            set_run_font(run, ascii_font=CODE_FONT, ea_font=CODE_FONT, size=size - 1,
                         bold=bold, color=RGBColor(0x9C, 0x27, 0x1B))
        pos = match.end()
    if pos < len(text):
        run = paragraph.add_run(text[pos:])
        set_run_font(run, size=size, bold=bold, color=color, ea_font=ea_font)


def style_heading(paragraph, level):
    sizes = {1: 16, 2: 14, 3: 12, 4: 10.5}
    paragraph.paragraph_format.space_before = Pt(12 if level == 1 else 8)
    paragraph.paragraph_format.space_after = Pt(6 if level == 1 else 4)
    paragraph.paragraph_format.keep_with_next = True
    for run in paragraph.runs:
        set_run_font(run, ascii_font=HEAD_FONT_ASCII, ea_font=HEAD_FONT_EA,
                     size=sizes.get(level, 10.5), bold=True, color=ACCENT)


def add_body_paragraph(doc, text, indent=True):
    p = doc.add_paragraph(style="Normal")
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(3)
    if indent:
        p.paragraph_format.first_line_indent = Pt(21)  # 2 chars at 10.5pt
    add_inline(p, text, size=10.5)
    return p


def add_bullet(doc, text):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.line_spacing = 1.35
    p.paragraph_format.space_after = Pt(2)
    add_inline(p, text, size=10.5)
    return p


def add_code_block(doc, lines):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    set_cell_margins(table, top=60, bottom=60, left=100, right=100)
    cell = table.cell(0, 0)
    cell_shade(cell, "F6F8FA")
    cell.text = ""
    first = True
    for line in lines:
        p = cell.paragraphs[0] if first else cell.add_paragraph()
        first = False
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        run = p.add_run(line if line else " ")
        set_run_font(run, ascii_font=CODE_FONT, ea_font=CODE_FONT, size=8.5)
    # small spacer after table
    doc.add_paragraph().paragraph_format.space_after = Pt(0)


def add_table(doc, rows):
    if not rows:
        return
    header = rows[0]
    body = rows[1:]
    table = doc.add_table(rows=1 + len(body), cols=len(header))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    table.autofit = True
    set_cell_margins(table)
    # widths
    if len(header) == 2:
        widths = [Inches(1.55), Inches(4.75)]
    elif len(header) == 3:
        widths = [Inches(1.45), Inches(2.2), Inches(2.65)]
    else:
        widths = [Inches(6.3 / len(header))] * len(header)
    for idx, text in enumerate(header):
        cell = table.cell(0, idx)
        cell.width = widths[idx]
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        add_inline(p, text, size=9, bold=True)
        for run in p.runs:
            run.font.bold = True
        cell_shade(cell, "DCE6F5")
    set_repeat_header(table.rows[0])
    for row in table.rows:
        trPr = row._tr.get_or_add_trPr()
        cant = OxmlElement("w:cantSplit")
        trPr.append(cant)
    for r_idx, row in enumerate(body, start=1):
        for c_idx, text in enumerate(row):
            if c_idx >= len(header):
                continue
            cell = table.cell(r_idx, c_idx)
            cell.width = widths[c_idx]
            cell.text = ""
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.15
            add_inline(p, text, size=9)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)


def add_figure(doc, path, caption):
    if not os.path.exists(path):
        p = doc.add_paragraph(style="Normal")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_inline(p, f"[缺图：{caption}]", size=9, color=GRAY)
        doc.add_paragraph()
        return
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(3)
    run = p.add_run()
    run.add_picture(path, width=Inches(6.15))
    cap = doc.add_paragraph(style="Normal")
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.space_after = Pt(8)
    add_inline(cap, caption, size=9, color=GRAY)


def parse_markdown(path):
    blocks = []
    lines = open(path, encoding="utf-8").read().splitlines()
    i = 0
    while i < len(lines):
        line = lines[i].rstrip()
        if not line.strip():
            i += 1
            continue
        if line.startswith("```"):
            i += 1
            code = []
            while i < len(lines) and not lines[i].startswith("```"):
                code.append(lines[i].rstrip("\n"))
                i += 1
            i += 1
            blocks.append(("code", code))
            continue
        if line.startswith("#### "):
            blocks.append(("h4", line[5:].strip()))
            i += 1
            continue
        if line.startswith("### "):
            blocks.append(("h3", line[4:].strip()))
            i += 1
            continue
        if line.startswith("## "):
            blocks.append(("h2", line[3:].strip()))
            i += 1
            continue
        if line.startswith("# "):
            blocks.append(("h1", line[2:].strip()))
            i += 1
            continue
        if line.startswith("![") and "](" in line and line.endswith(")"):
            m = re.match(r"!\[(.*?)\]\((.*?)\)", line)
            if m:
                blocks.append(("image", (m.group(2), m.group(1))))
                i += 1
                continue
        if line.startswith("|"):
            rows = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                raw = lines[i].strip()
                if re.match(r"^\|[\s:\-|]+\|$", raw):
                    i += 1
                    continue
                cells = [c.strip() for c in raw.strip("|").split("|")]
                rows.append(cells)
                i += 1
            if rows:
                blocks.append(("table", rows))
            continue
        if re.match(r"^\d+\.\s", line):
            blocks.append(("p", line.strip()))
            i += 1
            continue
        if line.startswith("- "):
            blocks.append(("bullet", line[2:].strip()))
            i += 1
            continue
        if line.startswith("> "):
            blocks.append(("quote", line[2:].strip()))
            i += 1
            continue
        # paragraph: consume until blank or special
        para = [line.strip()]
        i += 1
        while i < len(lines):
            nxt = lines[i].rstrip()
            if (not nxt.strip() or nxt.startswith(("#", "```", "|", "- ", "> ", "!["))):
                break
            para.append(nxt.strip())
            i += 1
        blocks.append(("p", " ".join(para)))
    return blocks


def fill_cover(doc):
    # paragraph 5 is the 赛题 line in the template
    for p in doc.paragraphs:
        if p.text.strip().startswith("赛题："):
            for r in p.runs:
                r.text = ""
            run = p.runs[0] if p.runs else p.add_run()
            run.text = "赛题：HapTest: 面向 OpenHarmony 应用的多模态智能测试系统"
            set_run_font(run, ascii_font=HEAD_FONT_ASCII, ea_font=HEAD_FONT_EA,
                         size=12, bold=True, color=ACCENT)
            br = p.add_run()
            br.add_break()
            r2 = p.add_run("定位：OpenHarmony 应用的智能测试与性能观测基础设施")
            set_run_font(r2, ascii_font=HEAD_FONT_ASCII, ea_font=HEAD_FONT_EA,
                         size=10.5, bold=False, color=GRAY)
            break
    table = doc.tables[0]
    values = {
        0: ("学校名称：", "北京航空航天大学"),
        1: ("团队名称：", "SMAT 团队"),
        2: ("日期：", "2026 年 11 月"),
        3: ("队长：", "刘发荣"),
        4: ("队员1：", "无"),
        5: ("队员2：", "—"),
    }
    for idx, (left, right) in values.items():
        row = table.rows[idx]
        for c_idx, text in ((0, left), (1, right)):
            cell = row.cells[c_idx]
            cell.text = ""
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 else WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(text)
            set_run_font(run, size=12, bold=(c_idx == 0), color=ACCENT if c_idx == 0 else None)


def remove_heading_numbering(doc):
    for name in ("Heading 1", "Heading 2", "Heading 3", "Heading 4"):
        style = doc.styles[name]
        pPr = style._element.pPr
        if pPr is not None and pPr.numPr is not None:
            pPr.remove(pPr.numPr)


def collect_toc_entries(blocks):
    entries = []
    for kind, value in blocks:
        if kind in ("h1", "h2", "h3"):
            entries.append((kind, value.replace("`", "").replace("**", "")))
    return entries


def build_static_toc(doc, blocks, page_map=None):
    page_map = page_map or {}
    # locate the TOC title and remove the template's static entries / field / blanks
    toc_title = None
    for p in doc.paragraphs:
        if p.text.strip().replace(" ", "") == "目录":
            toc_title = p
            break
    if toc_title is None:
        return
    for child in list(toc_title._element):
        if child.tag != qn("w:pPr"):
            toc_title._element.remove(child)
    run = toc_title.add_run("目  录")
    set_run_font(run, ascii_font="SimSun", ea_font="SimSun", size=18, bold=True)

    # delete everything between the TOC title and the first body heading
    el = toc_title._element.getnext()
    while el is not None:
        nxt = el.getnext()
        if el.tag == qn("w:p"):
            from docx.text.paragraph import Paragraph
            para = Paragraph(el, doc._body)
            if para.style.name == "Heading 1":
                break
            el.getparent().remove(el)
        el = nxt

    anchor = toc_title._element
    entries = collect_toc_entries(blocks)
    for kind, value in entries:
        p_el = OxmlElement("w:p")
        anchor.addnext(p_el)
        anchor = p_el
        from docx.text.paragraph import Paragraph
        para = Paragraph(p_el, doc._body)
        para.paragraph_format.tab_stops.add_tab_stop(
            Inches(6.25), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        if kind == "h1":
            para.paragraph_format.space_before = Pt(3)
            para.paragraph_format.space_after = Pt(0)
            indent, size, bold = 0.0, 10.5, True
        elif kind == "h2":
            para.paragraph_format.space_before = Pt(0)
            para.paragraph_format.space_after = Pt(0)
            indent, size, bold = 0.28, 9.5, False
        else:
            para.paragraph_format.space_before = Pt(0)
            para.paragraph_format.space_after = Pt(0)
            indent, size, bold = 0.56, 9.0, False
        para.paragraph_format.left_indent = Inches(indent)
        para.paragraph_format.line_spacing = 0.98
        page = page_map.get(value, "00")
        r = para.add_run(value)
        set_run_font(r, size=size, bold=bold)
        r2 = para.add_run("\t" + str(page))
        set_run_font(r2, size=size, bold=bold)

    # page break after the TOC
    p_el = OxmlElement("w:p")
    anchor.addnext(p_el)
    from docx.text.paragraph import Paragraph
    para = Paragraph(p_el, doc._body)
    run = para.add_run()
    run.add_break(WD_BREAK.PAGE)


def update_footer(doc):
    for section in doc.sections:
        footer = section.footer
        for p in list(footer.paragraphs):
            p._element.getparent().remove(p._element)
        p = footer.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run("第 ")
        set_run_font(r, size=9, color=GRAY)
        add_field(p, " PAGE ")
        r = p.add_run(" 页，共 ")
        set_run_font(r, size=9, color=GRAY)
        add_field(p, " NUMPAGES ")
        r = p.add_run(" 页")
        set_run_font(r, size=9, color=GRAY)
        for run in p.runs:
            set_run_font(run, size=9, color=GRAY)


def enable_update_fields(doc):
    settings = doc.settings.element
    update = OxmlElement("w:updateFields")
    update.set(qn("w:val"), "true")
    settings.append(update)


def main():
    doc = Document(TEMPLATE)

    # cover + TOC + footer
    fill_cover(doc)
    remove_heading_numbering(doc)
    blocks = parse_markdown(CONTENT)
    page_map = {}
    page_map_path = os.path.join(ROOT, "build", "page_map.json")
    if os.path.exists(page_map_path):
        import json
        page_map = json.load(open(page_map_path, encoding="utf-8"))
    build_static_toc(doc, blocks, page_map)
    update_footer(doc)

    # delete the template body from the first Heading 1 to the end
    body_started = False
    for p in list(doc.paragraphs):
        if p.style.name == "Heading 1" and p.text.strip() == "简介":
            body_started = True
        if body_started:
            p._element.getparent().remove(p._element)

    # metadata
    doc.core_properties.title = "HapTest: 面向 OpenHarmony 应用的多模态智能测试系统"
    doc.core_properties.author = "刘发荣"
    doc.core_properties.subject = "2026 开源鸿蒙大学生创新大赛-赛道一 项目设计文档"
    doc.core_properties.keywords = "OpenHarmony, HapTest, 自动化测试, PTG, HarmonyOS, 性能观测"

    first_h1 = True
    for kind, value in blocks:
        if kind == "h1":
            value = value.replace("`", "").replace("**", "")
            if not first_h1:
                doc.add_page_break()
            first_h1 = False
            p = doc.add_paragraph(style="Heading 1")
            p.add_run(value)
            style_heading(p, 1)
        elif kind == "h2":
            value = value.replace("`", "").replace("**", "")
            p = doc.add_paragraph(style="Heading 2")
            p.add_run(value)
            style_heading(p, 2)
        elif kind == "h3":
            value = value.replace("`", "").replace("**", "")
            p = doc.add_paragraph(style="Heading 3")
            p.add_run(value)
            style_heading(p, 3)
        elif kind == "h4":
            value = value.replace("`", "").replace("**", "")
            p = doc.add_paragraph(style="Heading 4")
            p.add_run(value)
            style_heading(p, 4)
        elif kind == "p":
            add_body_paragraph(doc, value)
        elif kind == "bullet":
            add_bullet(doc, value)
        elif kind == "quote":
            p = doc.add_paragraph(style="Normal")
            p.paragraph_format.left_indent = Inches(0.3)
            add_inline(p, value, size=9.5, color=GRAY)
        elif kind == "code":
            add_code_block(doc, value)
        elif kind == "table":
            add_table(doc, value)
        elif kind == "image":
            rel, caption = value
            add_figure(doc, os.path.join(ROOT, rel), caption)

    doc.save(OUT)
    print("wrote", OUT)


if __name__ == "__main__":
    main()
