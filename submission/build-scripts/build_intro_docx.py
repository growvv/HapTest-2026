#!/usr/bin/env python3
"""Build 01-参赛作品简介.docx from the Markdown summary."""
import os
import sys

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import build_docx as bd  # noqa: E402

ROOT = os.path.normpath(os.path.join(HERE, ".."))
CONTENT = os.path.join(ROOT, "01-参赛作品简介.md")
OUT = os.path.join(ROOT, "01-参赛作品简介.docx")


def main():
    doc = Document()
    section = doc.sections[0]
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.left_margin = section.right_margin = Inches(0.9)
    section.top_margin = section.bottom_margin = Inches(0.9)

    # base style
    style = doc.styles["Normal"]
    style.font.name = bd.BODY_FONT_ASCII
    style.font.size = Pt(10.5)
    style._element.rPr.rFonts.set(bd.qn("w:eastAsia"), bd.BODY_FONT_EA)

    # cover block
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("2026 开源鸿蒙大学生创新大赛 · 赛道一")
    bd.set_run_font(r, size=11, bold=True, color=bd.ACCENT)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(18)
    r = p.add_run("HapTest")
    bd.set_run_font(r, ascii_font=bd.HEAD_FONT_ASCII, ea_font=bd.HEAD_FONT_EA,
                    size=30, bold=True, color=bd.ACCENT)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("面向 OpenHarmony 应用的多模态智能测试系统")
    bd.set_run_font(r, ascii_font=bd.HEAD_FONT_ASCII, ea_font=bd.HEAD_FONT_EA,
                    size=15, bold=True, color=RGBColor(0x2F, 0x5C, 0xA8))
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    r = p.add_run("OpenHarmony 应用的智能测试与性能观测基础设施")
    bd.set_run_font(r, size=10, bold=False, color=bd.GRAY)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    r = p.add_run("参赛作品简介")
    bd.set_run_font(r, size=12, bold=False, color=bd.GRAY)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(30)
    r = p.add_run("北京航空航天大学 · SMAT 团队 · 刘发荣")
    bd.set_run_font(r, size=11, bold=True)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("2026 年 11 月")
    bd.set_run_font(r, size=10, color=bd.GRAY)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(18)
    r = p.add_run("作品仓库：github.com/growvv/HapTest-2026")
    bd.set_run_font(r, size=9, color=bd.GRAY)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("GitCode：gitcode.com/rogn/HapTest-2026")
    bd.set_run_font(r, size=9, color=bd.GRAY)
    doc.add_page_break()

    # body from Markdown, skipping the H1 and the info table (cover already states it)
    blocks = bd.parse_markdown(CONTENT)
    skip = True
    first_body = True
    for kind, value in blocks:
        if kind == "h1":
            if skip:
                skip = False
                continue
            p = doc.add_paragraph(style="Heading 1")
            p.add_run(value)
            bd.style_heading(p, 1)
        elif kind == "h2":
            p = doc.add_paragraph(style="Heading 2")
            p.add_run(value)
            bd.style_heading(p, 2)
        elif kind == "h3":
            p = doc.add_paragraph(style="Heading 3")
            p.add_run(value)
            bd.style_heading(p, 3)
        elif kind == "p":
            if first_body and "2026 开源鸿蒙大学生创新大赛" in value:
                first_body = False
                continue
            first_body = False
            bd.add_body_paragraph(doc, value)
        elif kind == "bullet":
            bd.add_bullet(doc, value)
        elif kind == "table":
            # The first table on the cover page is the metadata table; skip it.
            if value and value[0] and value[0][0] == "项目":
                continue
            bd.add_table(doc, value)
        elif kind == "code":
            bd.add_code_block(doc, value)
        elif kind == "image":
            rel, caption = value
            bd.add_figure(doc, os.path.join(ROOT, rel), caption)

    bd.update_footer(doc)
    doc.core_properties.title = "HapTest: 面向 OpenHarmony 应用的多模态智能测试系统 · 参赛作品简介"
    doc.core_properties.author = "刘发荣"
    doc.core_properties.subject = "2026 开源鸿蒙大学生创新大赛-赛道一"
    doc.save(OUT)
    print("wrote", OUT)


if __name__ == "__main__":
    main()
