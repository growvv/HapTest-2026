#!/usr/bin/env python3
"""Build 03-答辩PPT.pptx from the competition template and the submission content.

Usage: /tmp/hapdoc-env/bin/python submission/build-scripts/build_pptx.py
"""
import os
from copy import deepcopy

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.util import Inches, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.normpath(os.path.join(HERE, ".."))
TEMPLATE = os.path.join(ROOT, "templates", "xxx团队-赛道一-答辩模板.pptx")
OUT = os.path.join(ROOT, "03-答辩PPT.pptx")
ASSETS = os.path.join(ROOT, "assets")

FONT = "Arial Unicode MS"
MONO = "Courier New"
INK = RGBColor(0x0F, 0x14, 0x23)
GRAY = RGBColor(0x5B, 0x65, 0x77)
MUTED = RGBColor(0x8A, 0x94, 0xA6)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT = RGBColor(0xF4, 0xF7, 0xFC)
BORDER = RGBColor(0xE1, 0xE7, 0xF1)
BLUE = RGBColor(0x48, 0x74, 0xCB)
BLUE_D = RGBColor(0x2F, 0x5C, 0xA8)
BLUE_L = RGBColor(0xEA, 0xF0, 0xFA)
ORANGE = RGBColor(0xEE, 0x82, 0x2F)
ORANGE_L = RGBColor(0xFD, 0xF0, 0xE6)
TEAL = RGBColor(0x30, 0xC0, 0xB4)
TEAL_L = RGBColor(0xE6, 0xF7, 0xF5)
GREEN = RGBColor(0x75, 0xBD, 0x42)
GREEN_L = RGBColor(0xEE, 0xF7, 0xE8)
RED = RGBColor(0xE5, 0x4C, 0x5E)
RED_L = RGBColor(0xFC, 0xE9, 0xEC)
COVER_BLUE = RGBColor(0x08, 0xA8, 0xF6)

TOTAL = 16


def has_cjk(value):
    return any("\u4e00" <= ch <= "\u9fff" for ch in value)


def set_font(run, size=14, bold=False, color=INK, name=FONT, italic=False):
    latin_name = "Courier New" if name == MONO else "Arial Unicode MS"
    ea_name = "Arial Unicode MS"
    run.font.name = latin_name
    run.font.size = Pt(size)
    run.font.bold = bool(bold) and not has_cjk(run.text)
    run.font.italic = italic
    run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn("a:ea"))
    if ea is None:
        ea = parse_xml(f'<a:ea {nsdecls("a")} typeface="{ea_name}"/>')
        latin = rPr.find(qn("a:latin"))
        if latin is not None:
            latin.addnext(ea)
        else:
            rPr.append(ea)
    else:
        ea.set("typeface", ea_name)


def text(slide, x, y, w, h, value, size=14, bold=False, color=INK,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, line=1.08, name=FONT):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    p.line_spacing = line
    run = p.add_run()
    run.text = value
    set_font(run, size, bold, color, name)
    return box


def lines(slide, x, y, w, h, items, anchor=MSO_ANCHOR.TOP):
    """items: list of dicts {text, size, bold, color, space_after, align, font, line}"""
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    first = True
    for item in items:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = item.get("align", PP_ALIGN.LEFT)
        p.line_spacing = item.get("line", 1.08)
        p.space_after = Pt(item.get("space_after", 0))
        p.space_before = Pt(item.get("space_before", 0))
        run = p.add_run()
        run.text = item["text"]
        set_font(run, item.get("size", 12), item.get("bold", False),
                 item.get("color", INK), item.get("font", FONT))
    return box


def rect(slide, x, y, w, h, fill=WHITE, line_color=BORDER, radius=0.08,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, line_w=0.75):
    shp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
    if line_color is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line_color
        shp.line.width = Pt(line_w)
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        try:
            shp.adjustments[0] = radius
        except Exception:
            pass
    shp.shadow.inherit = False
    return shp


def hline(slide, x1, y1, x2, y2, color=BORDER, width=0.75):
    conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    conn.line.color.rgb = color
    conn.line.width = Pt(width)
    return conn


def image_fit(slide, path, x, y, w, h):
    if not os.path.exists(path):
        rect(slide, x, y, w, h, fill=LIGHT, line_color=BORDER)
        text(slide, x, y + h / 2 - 0.1, w, 0.2, f"缺图：{os.path.basename(path)}", 9, False, MUTED, PP_ALIGN.CENTER)
        return None
    with Image.open(path) as im:
        iw, ih = im.size
    scale = min(w / iw, h / ih)
    nw, nh = iw * scale, ih * scale
    left = x + (w - nw) / 2
    top = y + (h - nh) / 2
    return slide.shapes.add_picture(path, Inches(left), Inches(top), Inches(nw), Inches(nh))


def header(slide, section, title, subtitle=None):
    rect(slide, 0, 0, 2.15, 0.075, fill=BLUE, line_color=None, shape=MSO_SHAPE.RECTANGLE)
    rect(slide, 2.15, 0, 0.55, 0.075, fill=ORANGE, line_color=None, shape=MSO_SHAPE.RECTANGLE)
    text(slide, 0.65, 0.30, 8.0, 0.22, section, 9.5, True, BLUE)
    text(slide, 0.65, 0.54, 12.0, 0.55, title, 25, True, INK)
    if subtitle:
        text(slide, 0.66, 1.12, 12.0, 0.28, subtitle, 11, False, GRAY)
    hline(slide, 0.65, 1.54, 12.68, 1.54, BORDER, 0.7)


def footer(slide, idx):
    hline(slide, 0.65, 7.06, 12.68, 7.06, BORDER, 0.6)
    text(slide, 0.65, 7.13, 8.0, 0.18, "HapTest · 面向 OpenHarmony 应用的多模态智能测试系统", 8, False, MUTED)
    text(slide, 11.9, 7.13, 0.8, 0.18, f"{idx:02d} / {TOTAL:02d}", 8, False, MUTED, PP_ALIGN.RIGHT)


def notes(slide, value):
    slide.notes_slide.notes_text_frame.text = value


def metric(slide, x, y, w, h, value, label, accent=BLUE, note=None):
    rect(slide, x, y, w, h, fill=WHITE, line_color=BORDER, radius=0.07)
    rect(slide, x, y, 0.07, h, fill=accent, line_color=None, shape=MSO_SHAPE.RECTANGLE)
    text(slide, x + 0.18, y + 0.10, w - 0.28, 0.45, value, 25, True, accent)
    text(slide, x + 0.18, y + 0.53, w - 0.28, 0.22, label, 9.5, False, GRAY)
    if note:
        text(slide, x + 0.18, y + 0.73, w - 0.28, 0.18, note, 7.5, False, MUTED)


def card(slide, x, y, w, h, title, body, accent=BLUE, fill=WHITE, title_size=14, body_size=9.8):
    rect(slide, x, y, w, h, fill=fill, line_color=BORDER, radius=0.06)
    rect(slide, x, y, w, 0.06, fill=accent, line_color=None, shape=MSO_SHAPE.RECTANGLE)
    text(slide, x + 0.16, y + 0.16, w - 0.3, 0.3, title, title_size, True, INK)
    text(slide, x + 0.16, y + 0.52, w - 0.3, h - 0.6, body, body_size, False, GRAY, line=1.28)


def terminal(slide, x, y, w, h, rows, title="haptest harmony"):
    rect(slide, x, y, w, h, fill=RGBColor(0x0F, 0x14, 0x23), line_color=None, radius=0.035)
    rect(slide, x, y, w, 0.34, fill=RGBColor(0x1A, 0x22, 0x35), line_color=None, radius=0.035)
    for i, c in enumerate((RED, ORANGE, GREEN)):
        rect(slide, x + 0.16 + i * 0.16, y + 0.12, 0.08, 0.08, fill=c, line_color=None, shape=MSO_SHAPE.OVAL)
    text(slide, x + 0.68, y + 0.075, w - 0.8, 0.18, title, 8.5, False, RGBColor(0xC9, 0xD3, 0xE3), name=MONO)
    top = y + 0.46
    for row in rows:
        value, color = row
        box = slide.shapes.add_textbox(Inches(x + 0.18), Inches(top), Inches(w - 0.32), Inches(0.24))
        tf = box.text_frame
        tf.clear()
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.line_spacing = 1.0
        run = p.add_run()
        run.text = value
        set_font(run, 9.2, False, color, MONO)
        top += 0.235
    return top


def add_table(slide, x, y, w, h, rows, widths=None, font_size=9):
    shape = slide.shapes.add_table(len(rows), len(rows[0]), Inches(x), Inches(y), Inches(w), Inches(h))
    table = shape.table
    table.first_row = True
    table.horz_banding = False
    if widths:
        for i, width in enumerate(widths):
            table.columns[i].width = Inches(width)
    for r_idx, row in enumerate(rows):
        for c_idx, value in enumerate(row):
            cell = table.cell(r_idx, c_idx)
            cell.margin_left = Inches(0.08)
            cell.margin_right = Inches(0.08)
            cell.margin_top = Inches(0.04)
            cell.margin_bottom = Inches(0.04)
            cell.fill.solid()
            cell.fill.fore_color.rgb = BLUE if r_idx == 0 else (WHITE if r_idx % 2 else LIGHT)
            tf = cell.text_frame
            tf.clear()
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.line_spacing = 1.0
            run = p.add_run()
            run.text = value
            set_font(run, font_size, r_idx == 0, WHITE if r_idx == 0 else INK)
    return table


def bar_chart(slide, x, y, w, rows, max_value, accent_map):
    """rows: list of (label, value). Draw horizontal bars."""
    label_w = 1.55
    value_w = 0.7
    bar_max = w - label_w - value_w - 0.2
    row_h = 0.62
    for i, (label, value) in enumerate(rows):
        top = y + i * row_h
        text(slide, x, top + 0.12, label_w, 0.24, label, 10, True, INK)
        bg = rect(slide, x + label_w, top + 0.14, bar_max, 0.22, fill=RGBColor(0xEE, 0xF2, 0xF8), line_color=None, radius=0.11)
        width = max(0.12, bar_max * value / max_value)
        rect(slide, x + label_w, top + 0.14, width, 0.22, fill=accent_map[label], line_color=None, radius=0.11)
        text(slide, x + label_w + width + 0.08, top + 0.10, value_w, 0.24, str(value), 10, True, accent_map[label])


def build():
    prs = Presentation(TEMPLATE)

    # The competition template uses 微软雅黑, which is not available to the
    # LibreOffice renderer used for the PDF export. Re-point layout/master text
    # at a CJK font that is installed on macOS and available to LibreOffice.
    for layout in list(prs.slide_layouts) + [prs.slide_master]:
        for shape in layout.shapes:
            if not shape.has_text_frame:
                continue
            for p in shape.text_frame.paragraphs:
                for r in p.runs:
                    r.font.name = "Arial Unicode MS"
                    if has_cjk(r.text):
                        r.font.bold = False
                    rPr = r._r.get_or_add_rPr()
                    ea = rPr.find(qn("a:ea"))
                    if ea is None:
                        ea = parse_xml(f'<a:ea {nsdecls("a")} typeface="{FONT}"/>')
                        latin = rPr.find(qn("a:latin"))
                        if latin is not None:
                            latin.addnext(ea)
                        else:
                            rPr.append(ea)
                    else:
                        ea.set("typeface", FONT)

    # ------------------------------------------------------------------ cover
    cover = prs.slides[0]
    for shape in cover.shapes:
        if not shape.has_text_frame:
            continue
        name = shape.name
        if name == "文本框 6":
            tf = shape.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = "2026开源鸿蒙大学生创新大赛-赛道一"
            set_font(run, 30, True, COVER_BLUE)
        elif name == "文本框 10":
            shape.top = Inches(2.55)
            shape.height = Inches(1.65)
            tf = shape.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            r = p.add_run()
            r.text = "HapTest"
            set_font(r, 44, True, COVER_BLUE)
            p2 = tf.add_paragraph()
            p2.alignment = PP_ALIGN.CENTER
            p2.space_before = Pt(3)
            r2 = p2.add_run()
            r2.text = "面向 OpenHarmony 应用的多模态智能测试系统"
            set_font(r2, 15, False, BLUE_D)
            p3 = tf.add_paragraph()
            p3.alignment = PP_ALIGN.CENTER
            p3.space_before = Pt(2)
            r3 = p3.add_run()
            r3.text = "OpenHarmony 应用的智能测试与性能观测基础设施"
            set_font(r3, 10, False, GRAY)
        elif name == "文本框 4":
            shape.top = Inches(5.42)
            shape.height = Inches(0.85)
            tf = shape.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            r = p.add_run()
            r.text = "北京航空航天大学 · SMAT 团队"
            set_font(r, 14.5, True, INK)
            p2 = tf.add_paragraph()
            p2.alignment = PP_ALIGN.CENTER
            p2.space_before = Pt(3)
            r2 = p2.add_run()
            r2.text = "刘发荣 · 2026 年 11 月"
            set_font(r2, 11.5, False, GRAY)
    text(cover, 3.15, 4.38, 7.05, 0.30, "多模态观察 · 智能测试 · 性能观测", 11.5, True, BLUE_D, PP_ALIGN.CENTER)
    hline(cover, 4.7, 4.76, 8.65, 4.76, COVER_BLUE, 1.0)
    text(cover, 2.9, 6.55, 7.55, 0.24, "基于 FSE'25 / ASE'25 研究积累与 2026 系统级实现", 9, False, MUTED, PP_ALIGN.CENTER)
    notes(cover, "各位评委老师好，我是北京航空航天大学 SMAT 团队的刘发荣。今天汇报的作品是 HapTest：面向 OpenHarmony 应用的多模态智能测试系统，也是应用质量保障与性能观测的基础设施。作品希望把测试意图变成可执行、可复现、可观测的证据，而不是再写一个点击脚本。")

    # ------------------------------------------------------------------ TOC
    toc = prs.slides[1]
    for shape in list(toc.shapes):
        shape._element.getparent().remove(shape._element)
    text(toc, 0.55, 2.05, 1.8, 0.22, "五个部分", 9, False, MUTED)
    entries = [
        ("01", "背景与问题", "生态在长大，测试仍在用单点脚本追赶", BLUE),
        ("02", "方案与创新", "五模块闭环、PTG、设备 API、CLI、性能观测", ORANGE),
        ("03", "系统实现", "四层架构、关键模块、工程规范", TEAL),
        ("04", "验证与成果", "FSE'25、ASE'25、当前仓库可复现证据", GREEN),
        ("05", "价值与展望", "应用场景、路线图、团队与研究积累", RED),
    ]
    y0 = 1.62
    for i, (num, title, desc, accent) in enumerate(entries):
        y = y0 + i * 0.98
        rect(toc, 2.55, y, 0.56, 0.56, fill=accent, line_color=None, radius=0.12)
        text(toc, 2.55, y + 0.13, 0.56, 0.26, num, 13, True, WHITE, PP_ALIGN.CENTER)
        rect(toc, 3.28, y, 8.95, 0.56, fill=LIGHT, line_color=None, radius=0.08)
        rect(toc, 3.28, y, 0.06, 0.56, fill=accent, line_color=None, shape=MSO_SHAPE.RECTANGLE)
        text(toc, 3.48, y + 0.12, 8.4, 0.3, title, 17, True, INK)
        text(toc, 3.48, y + 0.66, 8.6, 0.24, desc, 10, False, MUTED)
    footer(toc, 2)
    notes(toc, "汇报分五个部分：先说明 OpenHarmony 应用质量保障的问题边界；再介绍 HapTest 的五模块闭环与四个创新点；然后说明工程实现；接着用论文与当前仓库的证据做验证；最后讨论应用价值与后续路线。")

    blank = prs.slide_layouts[8]

    def new_slide():
        s = prs.slides.add_slide(blank)
        for ph in list(s.placeholders):
            ph._element.getparent().remove(ph._element)
        s.background.fill.solid()
        s.background.fill.fore_color.rgb = WHITE
        return s

    # ------------------------------------------------------------------ 03 background
    s = new_slide()
    header(s, "01 · 背景与问题", "OpenHarmony 应用的多模态智能测试，缺的不是工具，而是闭环",
           "系统在长大，测试仍在用单点脚本追赶；多模态观察、状态、执行、证据彼此割裂")
    metric(s, 0.65, 1.72, 3.75, 1.05, "15,000+", "OpenHarmony 应用（公开研究数据）", BLUE)
    metric(s, 4.78, 1.72, 3.75, 1.05, "9 亿", "覆盖设备（公开研究数据）", TEAL)
    metric(s, 8.91, 1.72, 3.75, 1.05, "26", "20 款商业应用中发现的问题（FSE'25）", ORANGE)
    card(s, 0.65, 3.02, 3.75, 2.28, "人工测试", "依赖业务经验，能判断页面含义；\n但重复操作耗时，操作路径难以完整保留。", BLUE, body_size=10)
    card(s, 4.78, 3.02, 3.75, 2.28, "脚本测试", "稳定、可回归；\n但定位器、等待条件与初始化逻辑需要持续维护，页面一改就失效。", ORANGE, body_size=10)
    card(s, 8.91, 3.02, 3.75, 2.28, "随机 / 规则探索", "能自动产生操作；\n但难以表达业务目标，也难以把一次探索沉淀为可复现资产。", TEAL, body_size=10)
    hline(s, 0.81, 4.55, 4.24, 4.55, BORDER, 0.5)
    hline(s, 4.94, 4.55, 8.37, 4.55, BORDER, 0.5)
    hline(s, 9.07, 4.55, 12.50, 4.55, BORDER, 0.5)
    text(s, 0.81, 4.66, 3.3, 0.45, "例：登录态、弹窗与返回路径靠人记", 8.2, False, MUTED, line=1.2)
    text(s, 4.94, 4.66, 3.3, 0.45, "例：页面改版后定位器批量失效", 8.2, False, MUTED, line=1.2)
    text(s, 9.07, 4.66, 3.3, 0.45, "例：能点遍页面，说不清业务是否完成", 8.2, False, MUTED, line=1.2)
    rect(s, 0.65, 5.62, 12.01, 1.08, fill=BLUE_L, line_color=None, radius=0.07)
    text(s, 0.92, 5.76, 11.5, 0.3, "缺口：没有一个层把“观察 → 状态 → 动作 → 证据”串起来", 14, True, BLUE_D)
    text(s, 0.92, 6.14, 11.4, 0.3, "多模态观察是入口：控件树 + 截图 + 日志 + 任务语义；HapTest 补的正是这层编排。", 10.5, False, GRAY)
    footer(s, 3)
    notes(s, "OpenHarmony 生态已经有很大的应用与设备规模。但应用质量保障仍有一个结构性缺口：人工、脚本、随机探索各自只解决一段问题，没有一个层把观察、状态、动作与证据串起来。HapTest 就是补这层编排。")

    # ------------------------------------------------------------------ 04 five modules
    s = new_slide()
    header(s, "02 · 方案与创新", "多模态智能测试：把测试意图变成可执行、可复现、可观测的证据",
           "五模块闭环，共享同一份多模态观察、同一张页面图、同一套执行器")
    modules = [
        ("01", "多模态观察", "截图、控件树、日志与任务语义，生成带 observationId 的多模态观察记录。", BLUE, BLUE_L),
        ("02", "状态理解", "解析页面与控件，构建 PTG 页面迁移图；保留业务字段差异。", TEAL, TEAL_L),
        ("03", "路径规划", "候选生成、过滤、评分与选择；随机、DFS、静态、LLM 可替换。", ORANGE, ORANGE_L),
        ("04", "受约束执行", "观察先行、过期即拒、歧义不猜、设备独占；执行后立即失效 observation。", GREEN, GREEN_L),
        ("05", "证据沉淀", "日志、截图、PTG、检查点与失败现场；回放策略支持回归。", RED, RED_L),
    ]
    outputs = ["产出：observationId + 树 + 截图", "产出：PTG 节点与迁移边", "产出：下一步动作 a(t)", "产出：新观察 O(t+1)", "产出：用例 / 报告 / 证据"]
    for i, (num, title, body, accent, light) in enumerate(modules):
        x = 0.55 + i * 2.52
        rect(s, x, 1.78, 2.3, 3.42, fill=WHITE, line_color=BORDER, radius=0.06)
        rect(s, x, 1.78, 2.3, 0.06, fill=accent, line_color=None, shape=MSO_SHAPE.RECTANGLE)
        rect(s, x + 0.16, 1.96, 0.46, 0.46, fill=light, line_color=None, radius=0.12)
        text(s, x + 0.16, 2.06, 0.46, 0.24, num, 12, True, accent, PP_ALIGN.CENTER)
        text(s, x + 0.16, 2.56, 1.95, 0.3, title, 14, True, INK)
        text(s, x + 0.16, 2.98, 1.95, 1.9, body, 9.2, False, GRAY, line=1.32)
        hline(s, x + 0.16, 4.55, x + 2.14, 4.55, BORDER, 0.5)
        text(s, x + 0.16, 4.66, 1.95, 0.4, outputs[i], 8.2, False, MUTED, line=1.2)
        if i < 4:
            text(s, x + 2.31, 3.28, 0.2, 0.3, "›", 22, True, MUTED, PP_ALIGN.CENTER)
    bands = [
        ("研究底座", "FSE'25 模块化动态分析框架", BLUE, 0.55, 3.95),
        ("系统级新增", "2026：设备 API / CLI / 会话 / E2E", ORANGE, 4.65, 3.95),
        ("研究延伸", "ASE'25 遮挡诊断 / 性能观测", TEAL, 8.75, 3.8),
    ]
    for label, body, accent, x, w in bands:
        rect(s, x, 5.42, w, 1.22, fill=LIGHT, line_color=None, radius=0.07)
        rect(s, x, 5.42, 0.06, 1.22, fill=accent, line_color=None, shape=MSO_SHAPE.RECTANGLE)
        text(s, x + 0.18, 5.56, w - 0.3, 0.22, label, 9, True, accent)
        text(s, x + 0.18, 5.84, w - 0.3, 0.55, body, 10, True, INK)
    footer(s, 4)
    notes(s, "HapTest 用五个模块构成闭环：观察采集、状态理解、路径规划、受约束执行、证据沉淀。它们共享同一份观察、同一张 PTG 和同一套执行器。下面按三层展开：FSE'25 的研究底座、2026 年新增的系统级能力、以及 ASE'25 与性能观测的研究延伸。")

    # ------------------------------------------------------------------ 05 architecture
    s = new_slide()
    header(s, "02 · 方案与创新", "分层架构：设备差异下沉，策略差异上浮",
           "四层 + 一个贯穿始终的 observation / 证据模型；同一套 API 服务 CLI、Fuzz、Web 与未来 Agent")
    image_fit(s, os.path.join(ASSETS, "architecture.png"), 0.72, 1.68, 11.9, 5.28)
    text(s, 0.72, 6.92, 11.9, 0.16, "图 1  HapTest 系统总体架构（实现与设计分层标注）", 7.5, False, MUTED, PP_ALIGN.CENTER)
    footer(s, 5)
    notes(s, "架构分四层：驱动入口、核心服务、分析与观测、设备适配。设备差异尽量下沉到 HDC、uitest、Hypium；策略差异尽量上浮到 Policy 管线；observation 与证据贯穿所有层。这样同一套能力可以同时服务 CLI、Fuzz、Web 查看器和未来的 Agent 宿主。")

    # ------------------------------------------------------------------ 06 PTG
    s = new_slide()
    header(s, "02 · 方案与创新", "创新 1：PTG 把智能测试过程变成可回溯的状态图",
           "每一条边都来自真实执行后的观察，而不是模型猜测；状态图同时约束重复与预算")
    image_fit(s, os.path.join(ASSETS, "ptg-loop.png"), 0.55, 1.7, 12.2, 3.9)
    card(s, 0.65, 5.78, 3.8, 1.05, "状态图", "页面节点 + 迁移边 + 执行证据；支持可达性、路径回溯与 SVG 导出。", BLUE, body_size=8.6, title_size=12)
    card(s, 4.62, 5.78, 3.8, 1.05, "策略管线", "候选 → 过滤 → 评分 → 选择；随机 / DFS / 静态 / LLM 可替换。", ORANGE, body_size=8.6, title_size=12)
    card(s, 8.59, 5.78, 3.8, 1.05, "复用", "同一轨迹进入回放与回归；失败证据与页面变化一起保存。", GREEN, body_size=8.6, title_size=12)
    footer(s, 6)
    notes(s, "PTG 是 HapTest 的核心数据结构。它不是模型预测出来的图，而是每次真实执行后观察到的迁移。策略只负责在图上选择下一步，执行器负责安全执行并回写观察。这样探索过程既能减少重复，也能保留触发缺陷所需的关键条件。")

    # ------------------------------------------------------------------ 07 device API
    s = new_slide()
    header(s, "02 · 方案与创新", "创新 2：让“不猜”成为设备层的系统行为",
           "observationId、过期检测、歧义拒绝、跨进程租约、锁屏与敏感字段保护，全部由 HarmonyDevice 强制")
    image_fit(s, os.path.join(ASSETS, "device-api.png"), 0.55, 1.7, 12.2, 3.55)
    card(s, 0.65, 5.45, 3.8, 1.35, "观察先行", "动作前必须有 observation；mutate 后立即失效。没有 observation 的动作直接拒绝。", BLUE, body_size=8.8, title_size=12)
    card(s, 4.62, 5.45, 3.8, 1.35, "失败关闭", "歧义返回候选、过期返回 STALE、异常标记 ACTION_OUTCOME_UNCERTAIN；不猜、不重放。", RED, body_size=8.8, title_size=12)
    card(s, 8.59, 5.45, 3.8, 1.35, "设备独占", "同一 serial 跨进程文件租约；锁屏不返回树与截图；认证字段拒绝代输。", ORANGE, body_size=8.8, title_size=12)
    footer(s, 7)
    notes(s, "设备 API 的核心不是方法数量，而是失败时是否可控。HarmonyDevice 强制七条安全不变量：观察先行、坐标绑定 observationId、过期即拒、歧义不猜、异常不重放、设备独占、认证字段不代输。这些不是文档承诺，而是有单元测试覆盖的回归约束。")

    # ------------------------------------------------------------------ 08 CLI / serve
    s = new_slide()
    header(s, "02 · 方案与创新", "创新 3：CLI + JSON-lines，让 Agent 与 CI 用同一套语义",
           "stdout 永远是 JSON；退出码稳定；失败带 errorCode 与 nextStep")
    terminal(s, 0.6, 1.72, 7.55, 4.75, [
        ("$ haptest harmony observe --screenshot before.png", RGBColor(0x7F, 0xD9, 0xCF)),
        ('{"ok":true,"observation":{"observationId":"o-7f3a",...}}', RGBColor(0xC9, 0xD3, 0xE3)),
        ("", RGBColor(0xC9, 0xD3, 0xE3)),
        ('$ haptest harmony tap --text "搜索" --index 0', RGBColor(0x7F, 0xD9, 0xCF)),
        ('{"ok":true,"result":{"outcome":"completed"}}', RGBColor(0xC9, 0xD3, 0xE3)),
        ("", RGBColor(0xC9, 0xD3, 0xE3)),
        ('$ haptest harmony swipe --from 500,1800 --to 500,600', RGBColor(0x7F, 0xD9, 0xCF)),
        ('{"ok":true,"result":{"outcome":"completed"}}', RGBColor(0xC9, 0xD3, 0xE3)),
        ("", RGBColor(0xC9, 0xD3, 0xE3)),
        ("$ haptest harmony doctor", RGBColor(0x7F, 0xD9, 0xCF)),
        ('{"ok":false,"errorCode":"DEVICE_NOT_READY",...}', RGBColor(0xFF, 0xB4, 0x8A)),
        ('nextStep: "Connect and authorize one device..."', RGBColor(0x9F, 0xB0, 0xC8)),
        ("", RGBColor(0xC9, 0xD3, 0xE3)),
        ("$ haptest harmony serve", RGBColor(0x7F, 0xD9, 0xCF)),
        ('{"id":1,"method":"observe","params":{}}', RGBColor(0xC9, 0xD3, 0xE3)),
        ('{"id":1,"ok":true,"result":{"observationId":"o1"}}', RGBColor(0xC9, 0xD3, 0xE3)),
    ], title="haptest harmony · 真实命令与输出格式")
    card(s, 8.35, 1.72, 4.35, 1.42, "一次性命令", "doctor / devices / observe / tap / type / key / swipe / launch / stop；先 observe 再动作。", BLUE, body_size=8.8, title_size=12)
    card(s, 8.35, 3.28, 4.35, 1.42, "有状态会话", "serve 每行一个请求、每行一个响应；observationId 可跨多步复用。", ORANGE, body_size=8.8, title_size=12)
    card(s, 8.35, 4.84, 4.35, 1.42, "退出码", "0 成功 / 1 动作失败 / 2 用法错误 / 3 前置条件不满足；错误面可被脚本判断。", TEAL, body_size=8.8, title_size=12)
    rect(s, 0.65, 6.45, 12.01, 0.42, fill=BLUE_L, line_color=None, radius=0.06)
    text(s, 0.82, 6.53, 11.7, 0.24, "同一套 API 语义 → CLI / serve / 未来 Agent 宿主；不需要为每个宿主重写一层设备适配。", 9.5, True, BLUE_D)
    footer(s, 8)
    notes(s, "CLI 与 serve 是这次新增的系统级交付形态。一次性命令适合 shell、Python、CI；serve 适合多步、有状态的 Agent 会话。两者共用同一套 API 语义：stdout 永远是 JSON，退出码稳定，失败带 errorCode 与 nextStep。这样同一套能力不需要为每个宿主重写。")

    # ------------------------------------------------------------------ 09 perf
    s = new_slide()
    header(s, "02 · 方案与创新", "创新 4：性能观测，从“fps 采样”转向“逐帧切片”",
           "hitrace → .htrace → trace_streamer → SQLite frame_slice → FrameSample；缺数据就明确报错")
    image_fit(s, os.path.join(ASSETS, "perf-pipeline.png"), 0.55, 1.7, 12.2, 3.75)
    card(s, 0.65, 5.62, 3.8, 1.2, "为什么不用 fps", "fps 采样回答不了“哪一帧、哪个动作、哪段时间”；逐帧切片才能把性能与步骤对齐。", ORANGE, body_size=8.6, title_size=12)
    card(s, 4.62, 5.62, 3.8, 1.2, "probe / trace", "probeMetrics 探测来源；trace(action, run) 用同进程包住动作窗口；FrameSample 输出帧数与 p95。", BLUE, body_size=8.6, title_size=12)
    card(s, 8.59, 5.62, 3.8, 1.2, "失败关闭", "来源不可用 → UNSUPPORTED_METRIC；不返回 fps:0，不编造帧数；报告显式列出缺失。", RED, body_size=8.6, title_size=12)
    footer(s, 9)
    notes(s, "性能观测体现了我们的方法：不用 fps 采样冒充逐帧掉帧，而是解析 DevEco 的 trace，用 vsync 对齐的帧切片和 jankFlag 定义掉帧。缺数据时明确报 UNSUPPORTED_METRIC，不编造数字。")

    # ------------------------------------------------------------------ 10 engineering
    s = new_slide()
    header(s, "03 · 系统实现", "工程实现：可构建、可测试、可复现",
           "TypeScript · vitest · 真机 E2E · GitHub Actions；提交前 build 与 test 必须全绿")
    metric(s, 0.55, 1.72, 2.95, 0.95, "74", "src/ TypeScript 文件", BLUE)
    metric(s, 3.63, 1.72, 2.95, 0.95, "10,843", "src/ 代码行数", TEAL)
    metric(s, 6.71, 1.72, 2.95, 0.95, "60", "单元测试（20 个文件）", ORANGE)
    metric(s, 9.79, 1.72, 2.87, 0.95, "37", "本次分支变更文件（+6,794）", GREEN)
    terminal(s, 0.55, 2.9, 7.3, 3.55, [
        ("$ npm run build", RGBColor(0x7F, 0xD9, 0xCF)),
        ("> tsc  （exit 0）", RGBColor(0x9F, 0xB0, 0xC8)),
        ("", RGBColor(0xC9, 0xD3, 0xE3)),
        ("$ npm test -- --run", RGBColor(0x7F, 0xD9, 0xCF)),
        ("Test Files  20 passed (20)", RGBColor(0x9F, 0xE8, 0xB0)),
        ("Tests       60 passed (60)", RGBColor(0x9F, 0xE8, 0xB0)),
        ("Duration    3.89s", RGBColor(0x9F, 0xB0, 0xC8)),
        ("", RGBColor(0xC9, 0xD3, 0xE3)),
        ("$ npm run e2e", RGBColor(0x7F, 0xD9, 0xCF)),
        ("SKIPPED: set HAPTEST_E2E_SERIAL ...", RGBColor(0xFF, 0xD1, 0x8A)),
        ("（无设备时不伪造结果，退出 0）", RGBColor(0x9F, 0xB0, 0xC8)),
    ], title="构建与测试记录 · 2026-11-07")
    card(s, 8.15, 2.9, 4.55, 0.78, "代码结构", "src/cli · device · model · policy · runner · ui；lib/ 为构建产物。", BLUE, body_size=8.2, title_size=10.5)
    card(s, 8.15, 3.78, 4.55, 0.78, "测试标准", "只写会失败的测试；真实输出、真实 IO、状态机韧性、安全不变量。", ORANGE, body_size=8.2, title_size=10.5)
    card(s, 8.15, 4.66, 4.55, 0.78, "CI", "GitHub Actions：Linux / Windows，Node 18 / 20 / 21。", TEAL, body_size=8.2, title_size=10.5)
    card(s, 8.15, 5.54, 4.55, 0.78, "文档", "设计规格、设备 API 手册、性能逆向记录、AGENTS.md 提交规范。", GREEN, body_size=8.2, title_size=10.5)
    footer(s, 10)
    notes(s, "工程上，仓库有 74 个 TypeScript 文件、约 1.08 万行源码、20 个测试文件、60 个测试。构建与单测全绿；E2E 在没有设备时明确 SKIPPED 并退出 0。我们坚持“只写会失败的测试”，不把 mock 回放当真实。")

    # ------------------------------------------------------------------ 11 FSE
    s = new_slide()
    header(s, "04 · 验证与成果", "真实应用验证：20 款商业应用，26 个未报告问题",
           "FSE 2025 Industry Track · HUAWEI Mate60 / OpenHarmony 5.0.0 · 贪心 DFS + 故障日志")
    metric(s, 0.65, 1.72, 3.6, 0.95, "20", "参与实验的商业应用", BLUE)
    metric(s, 4.45, 1.72, 3.6, 0.95, "11", "发现问题的应用", ORANGE)
    metric(s, 8.25, 1.72, 4.41, 0.95, "26", "此前未报告的运行问题", RED)
    text(s, 0.72, 2.95, 7.0, 0.3, "问题类型分布（项）", 11, True, INK)
    bar_chart(s, 0.72, 3.28, 7.1, [("ArkTS 崩溃", 13), ("C++ 崩溃", 10), ("卡死", 3)], 13,
              {"ArkTS 崩溃": RED, "C++ 崩溃": ORANGE, "卡死": TEAL})
    text(s, 0.72, 5.25, 7.0, 0.5, "13 ArkTS 崩溃  ·  10 C++ 崩溃  ·  3 卡死；问题分布在 11 / 20 款应用。", 9.5, False, GRAY)
    rect(s, 8.15, 2.92, 4.5, 3.35, fill=WHITE, line_color=BORDER, radius=0.05)
    image_fit(s, os.path.join(ROOT, "05-辅助材料", "FSE2025-HapTest-首页.png"), 8.32, 3.08, 1.75, 2.45)
    text(s, 10.18, 3.12, 2.3, 0.28, "FSE 2025", 11, True, BLUE)
    text(s, 10.18, 3.42, 2.3, 0.5, "Industry Track", 9, False, GRAY)
    text(s, 10.18, 3.95, 2.3, 1.2, "HapTest: The Dynamic Analysis Framework for OpenHarmony", 8.2, True, INK, line=1.2)
    text(s, 10.18, 5.18, 2.3, 0.55, "第一作者：刘发荣\n北京航空航天大学", 8, False, GRAY, line=1.2)
    text(s, 0.72, 6.25, 12.0, 0.5, "来源标注：以上数据来自 FSE'25 论文表 4；本仓库为方法实现与工程化，不将论文实验重复声称为本次新增。", 8.5, False, MUTED)
    footer(s, 11)
    notes(s, "研究验证来自 FSE 2025 Industry Track 论文。我们在 20 款热门商业应用上发现 26 个此前未报告的运行问题，分布在 11 款应用：13 个 ArkTS 崩溃、10 个 C++ 崩溃、3 个卡死。这里明确标注来源，不把论文实验说成本次新增。")

    # ------------------------------------------------------------------ 12 ASE
    s = new_slide()
    header(s, "04 · 验证与成果", "多模态遮挡诊断：100 款应用、3326 万个重叠实例",
           "ASE 2025 Industry Track · HCO-Eye / HCOBench · 精确率 97.14% / 召回率 89.47% / F1 93.12%")
    metric(s, 0.55, 1.72, 2.95, 0.95, "100", "分析的热门应用", BLUE)
    metric(s, 3.63, 1.72, 2.95, 0.95, "33.26M", "UI 重叠实例", TEAL)
    metric(s, 6.71, 1.72, 2.95, 0.95, "34", "高开销遮挡案例", ORANGE)
    metric(s, 9.79, 1.72, 2.87, 0.95, "93.12%", "HCOBench F1", RED)
    text(s, 0.62, 2.95, 7.0, 0.3, "HCOBench 指标（%）", 11, True, INK)
    bar_chart(s, 0.62, 3.28, 7.2, [("精确率", 97.14), ("召回率", 89.47), ("F1", 93.12)], 100,
              {"精确率": BLUE, "召回率": TEAL, "F1": RED})
    text(s, 0.62, 5.28, 7.2, 0.55, "HCOBench：152 个挑战样本、38 个人工确认正例；正确检出 34 个高开销遮挡案例。", 9.2, False, GRAY)
    rect(s, 8.15, 2.92, 4.5, 3.35, fill=WHITE, line_color=BORDER, radius=0.05)
    image_fit(s, os.path.join(ROOT, "05-辅助材料", "ASE2025-UI-Overlap-首页.png"), 8.32, 3.08, 1.75, 2.45)
    text(s, 10.18, 3.12, 2.3, 0.28, "ASE 2025", 11, True, BLUE)
    text(s, 10.18, 3.42, 2.3, 0.5, "Industry Track", 9, False, GRAY)
    text(s, 10.18, 3.95, 2.3, 1.25, "An Empirical Study on UI Overlap in OpenHarmony Applications", 8.2, True, INK, line=1.2)
    text(s, 10.18, 5.2, 2.3, 0.55, "第一作者：刘发荣\n北京航空航天大学", 8, False, GRAY, line=1.2)
    text(s, 0.62, 6.25, 12.0, 0.55, "来源标注：数据来自 ASE'25 论文；VLM 多模态诊断为研究延伸，未并入当前仓库主链路。", 8.5, False, MUTED)
    footer(s, 12)
    notes(s, "第二项研究验证是 ASE 2025 的多模态 UI 遮挡诊断 HCO-Eye。我们分析 100 款应用、3300 多万个重叠实例，在 HCOBench 上取得 93.12% 的 F1，并发现 34 个高开销遮挡案例。同样标注来源：VLM 诊断是研究延伸，当前仓库主链路没有把它写成已实现。")

    # ------------------------------------------------------------------ 13 evidence
    s = new_slide()
    header(s, "04 · 验证与成果", "当前仓库的可复现证据",
           "构建、单测、E2E、CLI 输出都可重跑；无设备时不伪造结果")
    tbl = add_table(s, 0.62, 1.72, 8.35, 4.15, [
        ["证据", "命令 / 文件", "结果"],
        ["构建", "npm run build", "tsc 退出 0"],
        ["单元测试", "npm test -- --run", "20 files / 60 tests passed"],
        ["真机 E2E", "npm run e2e", "无 serial 时 SKIPPED，退出 0"],
        ["CLI 诊断", "haptest harmony doctor", "DEVICE_NOT_READY + nextStep"],
        ["代码规模", "src/", "74 files / 10,843 lines"],
        ["本次变更", "git diff main...HEAD", "37 files / +6,794 / -23"],
    ], widths=[1.35, 3.05, 3.95], font_size=9)
    tbl.rows[0].height = Inches(0.42)
    for row in list(tbl.rows)[1:]:
        row.height = Inches(0.58)
    rect(s, 9.2, 1.72, 3.48, 4.15, fill=BLUE_L, line_color=None, radius=0.06)
    text(s, 9.4, 1.9, 3.1, 0.28, "证据链", 13, True, BLUE_D)
    steps = [("1", "观察", "observationId + 树 + 截图"), ("2", "动作", "selector / 坐标 + 约束"), ("3", "PTG", "页面节点 + 迁移边"), ("4", "报告", "日志 + 截图 + 失败现场")]
    for i, (num, title, body) in enumerate(steps):
        y = 2.35 + i * 0.83
        rect(s, 9.4, y, 0.34, 0.34, fill=BLUE, line_color=None, radius=0.1)
        text(s, 9.4, y + 0.06, 0.34, 0.2, num, 9, True, WHITE, PP_ALIGN.CENTER)
        text(s, 9.86, y - 0.02, 2.5, 0.22, title, 10.5, True, INK)
        text(s, 9.86, y + 0.2, 2.5, 0.2, body, 7.8, False, GRAY)
    text(s, 9.4, 5.5, 3.0, 0.22, "可复现 · 可追溯 · 不造假", 9, True, GREEN)
    rect(s, 0.65, 6.1, 12.01, 0.72, fill=ORANGE_L, line_color=None, radius=0.06)
    text(s, 0.85, 6.22, 11.6, 0.24, "验证原则：不把 mock 当真实；无设备时明确 SKIPPED；每个数字有来源；论文与专利按来源标注。", 10, True, ORANGE)
    footer(s, 13)
    notes(s, "这是当前仓库的可复现证据：构建、60 个单测、E2E 的 SKIPPED 行为、CLI 诊断输出、代码规模与本次变更量。我们明确区分“当前仓库可重跑”与“论文实验数据”，不混在一起说。")

    # ------------------------------------------------------------------ 14 value
    s = new_slide()
    header(s, "05 · 价值与展望", "从一次多模态智能测试到可复用的质量资产",
           "开发者、测试团队与生态伙伴复用同一套观察、执行与证据")
    card(s, 0.55, 1.78, 3.9, 2.55, "应用开发者", "新增/修改页面后，用自然语言提出验证目标；\n获得可执行步骤、检查点与失败证据；\n有效轨迹沉淀为回归用例。", BLUE, body_size=9, title_size=13)
    card(s, 4.72, 1.78, 3.9, 2.55, "测试团队", "缺陷复现、稳定性压测与探索性测试；\n统一定位器、等待条件与证据格式；\n减少重复操作与用例维护。", ORANGE, body_size=9, title_size=13)
    card(s, 8.89, 1.78, 3.78, 2.55, "生态伙伴", "面向 OpenHarmony 开发者的测试基础设施；\n可接入 CI、教学与工具链；\n方法、代码与文档可复用。", TEAL, body_size=9, title_size=13)
    text(s, 0.8, 3.55, 3.4, 0.5, "典型：新增页面后的保存/返回验收", 8.2, False, MUTED, line=1.2)
    text(s, 4.97, 3.55, 3.4, 0.5, "典型：崩溃复现与版本回归", 8.2, False, MUTED, line=1.2)
    text(s, 9.14, 3.55, 3.3, 0.5, "典型：CI 冒烟与证据归档", 8.2, False, MUTED, line=1.2)
    text(s, 0.65, 4.65, 5.0, 0.3, "路线图", 12, True, INK)
    hline(s, 0.9, 5.32, 12.35, 5.32, BORDER, 1.0)
    phases = [
        ("现在", "设备 API / CLI / 会话 / E2E", "已实现并有测试覆盖", BLUE, 0.9),
        ("下一步", "帧解析 / 性能报告 / 多设备", "真机日补齐枚举值，扩展采集", ORANGE, 4.75),
        ("后续", "CI 集成 / Agent 宿主 / 应用试点", "把证据链接入研发流程", TEAL, 8.6),
    ]
    for label, title, body, accent, x in phases:
        rect(s, x, 5.18, 0.26, 0.26, fill=accent, line_color=None, shape=MSO_SHAPE.OVAL)
        text(s, x, 5.52, 3.4, 0.22, label, 8.5, True, accent)
        text(s, x, 5.78, 3.4, 0.28, title, 10.5, True, INK)
        text(s, x, 6.1, 3.4, 0.35, body, 8.2, False, GRAY)
    footer(s, 14)
    notes(s, "价值上，开发者可以用它做页面验收与回归；测试团队可以用它做缺陷复现、稳定性压测与探索性测试；生态伙伴可以把这套能力接进 CI、教学与工具链。路线图分三步：现在完成设备 API/CLI/E2E；下一步补齐帧解析与多设备；后续接入 CI 与 Agent 宿主。")

    # ------------------------------------------------------------------ 15 team
    s = new_slide()
    header(s, "05 · 价值与展望", "团队与研究积累",
           "北京航空航天大学 SMAT 团队 · 刘发荣 · 全栈交付")
    rect(s, 0.55, 1.75, 4.0, 3.7, fill=WHITE, line_color=BORDER, radius=0.06)
    rect(s, 0.55, 1.75, 4.0, 0.06, fill=BLUE, line_color=None, shape=MSO_SHAPE.RECTANGLE)
    rect(s, 0.85, 2.0, 0.9, 0.9, fill=BLUE, line_color=None, shape=MSO_SHAPE.OVAL)
    text(s, 0.85, 2.24, 0.9, 0.42, "刘", 22, True, WHITE, PP_ALIGN.CENTER)
    text(s, 1.95, 2.04, 2.3, 0.3, "刘发荣", 18, True, INK)
    text(s, 1.95, 2.4, 2.4, 0.45, "队长 · 架构 / 设备层 / CLI / 测试 / 文档", 8.5, False, GRAY, line=1.2)
    text(s, 0.85, 3.1, 3.4, 0.3, "北京航空航天大学 · SMAT 团队", 10, True, BLUE_D)
    text(s, 0.85, 3.48, 3.4, 1.6, "覆盖需求、架构、实现、测试、文档与演示；\n论文与专利为研究积累，按来源标注。", 9, False, GRAY, line=1.4)
    text(s, 0.85, 4.18, 3.4, 0.65, "研究：FSE'25 / ASE'25 / 发明专利\n开源：SMAT-Lab / HapTest（Apache-2.0）", 8.2, False, MUTED, line=1.3)
    lines(s, 0.85, 4.92, 3.4, 0.5, [
        {"text": "作品仓库：github.com/growvv/HapTest-2026", "size": 7.5, "color": MUTED},
        {"text": "GitCode：gitcode.com/rogn/HapTest-2026", "size": 7.5, "color": MUTED},
    ])
    achievements = [
        ("FSE 2025", "HapTest 动态分析框架\n第一作者", os.path.join(ROOT, "05-辅助材料", "FSE2025-HapTest-首页.png")),
        ("ASE 2025", "UI 重叠实证研究\n第一作者", os.path.join(ROOT, "05-辅助材料", "ASE2025-UI-Overlap-首页.png")),
        ("发明专利", "OpenHarmony 动态分析框架\n申请号 202510375959.1", os.path.join(ASSETS, "patent-notice.png")),
        ("开源项目", "SMAT-Lab / HapTest\nApache-2.0", os.path.join(ASSETS, "icon-code.png")),
    ]
    for i, (title, body, img) in enumerate(achievements):
        col = i % 2
        row = i // 2
        x = 4.85 + col * 4.0
        y = 1.75 + row * 1.9
        rect(s, x, y, 3.8, 1.7, fill=WHITE, line_color=BORDER, radius=0.05)
        rect(s, x, y, 0.06, 1.7, fill=[BLUE, ORANGE, TEAL, GREEN][i], line_color=None, shape=MSO_SHAPE.RECTANGLE)
        image_fit(s, img, x + 0.16, y + 0.18, 0.95, 1.3)
        text(s, x + 1.2, y + 0.18, 2.45, 0.25, title, 10.5, True, INK)
        text(s, x + 1.2, y + 0.5, 2.45, 0.9, body, 7.8, False, GRAY, line=1.25)
    rect(s, 0.65, 5.75, 11.95, 0.8, fill=LIGHT, line_color=None, radius=0.06)
    text(s, 0.85, 5.9, 11.6, 0.22, "工作流", 8.5, True, BLUE)
    text(s, 0.85, 6.12, 11.6, 0.28, "架构设计 · 设备适配 · CLI/会话 · 策略与模型 · 测试/E2E · 文档 · PPT · 视频", 9.5, True, INK)
    footer(s, 15)
    notes(s, "我承担了架构、设备层、CLI、测试、文档和演示的全部工作。研究积累包括 FSE'25、ASE'25 两篇第一作者论文和一项北京航空航天大学的发明专利；这些在文档里都按来源标注，不与本次代码混为一谈。")

    # ------------------------------------------------------------------ 16 closing
    closing = prs.slides[3]  # the template's "谢谢" slide (section slide is still present)
    for ph in list(closing.placeholders):
        ph._element.getparent().remove(ph._element)
    # edit the existing 谢 谢 text box
    for shape in closing.shapes:
        if shape.has_text_frame and "谢" in shape.text_frame.text:
            shape.left = Inches(3.4)
            shape.top = Inches(1.16)
            shape.width = Inches(6.5)
            shape.height = Inches(0.85)
            tf = shape.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = "谢  谢"
            set_font(run, 38, True, WHITE)
    text(closing, 1.2, 0.52, 10.9, 0.4, "让测试意图变成可执行、可复现、可观测的证据。", 18, True, WHITE, PP_ALIGN.CENTER)
    hline(closing, 3.0, 1.0, 10.3, 1.0, RGBColor(0xBF, 0xE3, 0xFF), 1.0)
    cards = [
        ("closing-perf.png", "性能观察", "看清每一次卡顿"),
        ("closing-devices.png", "多设备支持", "覆盖更多终端"),
        ("closing-ci.png", "日常自动检查", "融入团队流程"),
    ]
    cw, ch, gap = 3.55, 2.65, 0.35
    x0 = (13.333 - (cw * 3 + gap * 2)) / 2
    cy = 2.22
    for i, (img, title_txt, body_txt) in enumerate(cards):
        cx = x0 + i * (cw + gap)
        rect(closing, cx, cy, cw, ch, fill=WHITE, line_color=None, radius=0.1)
        image_fit(closing, os.path.join(ASSETS, img), cx + 0.55, cy + 0.10, 2.45, 1.5)
        text(closing, cx, cy + 1.7, cw, 0.3, title_txt, 12, True, INK, PP_ALIGN.CENTER)
        text(closing, cx, cy + 2.02, cw, 0.3, body_txt, 8.8, False, GRAY, PP_ALIGN.CENTER)
    text(closing, 1.2, 5.14, 10.9, 0.3, "HapTest · 面向 OpenHarmony 应用的多模态智能测试系统", 12, True, RGBColor(0xE6, 0xF4, 0xFF), PP_ALIGN.CENTER)
    text(closing, 1.2, 5.51, 10.9, 0.28, "北京航空航天大学 SMAT 团队 · 刘发荣 · 2026 年 11 月", 10.5, False, RGBColor(0xD8, 0xEB, 0xFF), PP_ALIGN.CENTER)
    text(closing, 1.2, 5.92, 10.9, 0.24, "作品仓库：github.com/growvv/HapTest-2026 · GitCode：gitcode.com/rogn/HapTest-2026", 8.2, False, RGBColor(0xB9, 0xD4, 0xF2), PP_ALIGN.CENTER)
    notes(closing, "我的汇报到这里。HapTest 想做的事情可以概括成一句话：让测试意图变成可执行、可复现、可观测的证据。谢谢各位评委老师。")

    # remove the template's placeholder section slide, then move the closing slide to the end
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    rId = slides[2].get(qn("r:id"))
    prs.part.drop_rel(rId)
    xml_slides.remove(slides[2])
    slides = list(xml_slides)
    xml_slides.remove(slides[2])
    xml_slides.append(slides[2])

    prs.core_properties.title = "HapTest: 面向 OpenHarmony 应用的多模态智能测试系统"
    prs.core_properties.author = "刘发荣"
    prs.core_properties.subject = "2026 开源鸿蒙大学生创新大赛-赛道一 答辩 PPT"
    prs.save(OUT)
    print("wrote", OUT, "slides:", len(prs.slides))


if __name__ == "__main__":
    build()
