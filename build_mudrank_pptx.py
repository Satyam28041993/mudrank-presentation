# -*- coding: utf-8 -*-
"""Build Secure Legal Paper presentation (16:9) — light theme, simple English."""

from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

ROOT = Path(r"D:\Users\user\Desktop\Mudrank Presentation")
ASSETS = ROOT / "_ppt_assets"
OUT = ROOT / "Propix_Mudrank_Secure_Legal_Paper_v2.pptx"

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

NAVY = RGBColor(0x12, 0x2B, 0x3C)
INK = RGBColor(0x1A, 0x2B, 0x36)
MUTED = RGBColor(0x5A, 0x6B, 0x78)
LINE = RGBColor(0xDE, 0xE5, 0xEB)
ACCENT = RGBColor(0xC4, 0x5E, 0x24)
ACCENT_SOFT = RGBColor(0xF7, 0xE8, 0xDC)
SOFT = RGBColor(0xF6, 0xF8, 0xFA)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
TEAL = RGBColor(0x1F, 0x6F, 0x7A)
RED = RGBColor(0xB5, 0x3A, 0x2E)
GREEN = RGBColor(0x2F, 0x7D, 0x4F)
SKY = RGBColor(0xE8, 0xF0, 0xF4)

TOTAL = 24


def set_run(run, size=14, bold=False, color=INK, font="Candara"):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font


def add_text(shape, text, size=14, bold=False, color=INK, align=PP_ALIGN.LEFT, font="Candara"):
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    set_run(run, size=size, bold=bold, color=color, font=font)
    return tf


def add_para(tf, text, size=13, bold=False, color=INK, align=PP_ALIGN.LEFT, space_before=6, font="Candara"):
    p = tf.add_paragraph()
    p.alignment = align
    p.space_before = Pt(space_before)
    run = p.add_run()
    run.text = text
    set_run(run, size=size, bold=bold, color=color, font=font)
    return p


def rect(slide, x, y, w, h, fill=None, line=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shape.line.fill.background()
    if fill is not None:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    else:
        shape.fill.background()
    if line is not None:
        shape.line.color.rgb = line
        shape.line.width = Pt(1)
    return shape


def round_rect(slide, x, y, w, h, fill=WHITE, line=LINE):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
        shape.line.width = Pt(1)
    try:
        shape.adjustments[0] = 0.08
    except Exception:
        pass
    return shape


def oval(slide, x, y, w, h, fill=ACCENT):
    shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.fill.background()
    return shape


def picture(slide, path: Path, x, y, w=None, h=None):
    if not path.exists():
        return None
    if w and h:
        return slide.shapes.add_picture(str(path), x, y, width=w, height=h)
    if w:
        return slide.shapes.add_picture(str(path), x, y, width=w)
    if h:
        return slide.shapes.add_picture(str(path), x, y, height=h)
    return slide.shapes.add_picture(str(path), x, y)


def top_bar(slide, page_no: int, total: int, section: str = ""):
    rect(slide, 0, 0, SLIDE_W, Inches(0.58), fill=WHITE)
    rect(slide, 0, Inches(0.58), SLIDE_W, Inches(0.035), fill=ACCENT)
    box = slide.shapes.add_textbox(Inches(0.45), Inches(0.14), Inches(5.2), Inches(0.35))
    add_text(box, "PRAKRUTI GRAPHIC PVT. LTD.", size=12, bold=True, color=NAVY)
    if section:
        mid = slide.shapes.add_textbox(Inches(5.0), Inches(0.14), Inches(3.8), Inches(0.35))
        add_text(mid, section, size=12, color=MUTED, align=PP_ALIGN.CENTER)
    right = slide.shapes.add_textbox(Inches(10.5), Inches(0.14), Inches(2.4), Inches(0.35))
    add_text(right, f"{page_no}  /  {total}", size=12, color=MUTED, align=PP_ALIGN.RIGHT)


def footer(slide, text="Secure Legal Paper  |  Paper Security + QR Check + Tracking"):
    rect(slide, Inches(0.45), Inches(7.05), Inches(12.4), Inches(0.015), fill=LINE)
    left = slide.shapes.add_textbox(Inches(0.45), Inches(7.12), Inches(8), Inches(0.28))
    add_text(left, text, size=9, color=MUTED)
    right = slide.shapes.add_textbox(Inches(8.5), Inches(7.12), Inches(4.35), Inches(0.28))
    add_text(right, "Confidential  ·  www.pgpltechprint.com", size=9, color=MUTED, align=PP_ALIGN.RIGHT)


def blank_content(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def content_bg(slide):
    rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill=SOFT)


def title_block(slide, title, subtitle=""):
    t = slide.shapes.add_textbox(Inches(0.5), Inches(0.78), Inches(12.3), Inches(0.5))
    add_text(t, title, size=28, bold=True, color=INK, font="Georgia")
    if subtitle:
        s = slide.shapes.add_textbox(Inches(0.5), Inches(1.28), Inches(12.3), Inches(0.35))
        add_text(s, subtitle, size=14, color=MUTED)


def bullet_card(slide, x, y, w, h, title, lines, accent=ACCENT):
    round_rect(slide, x, y, w, h, fill=WHITE, line=LINE)
    rect(slide, x, y, Inches(0.1), h, fill=accent)
    tb = slide.shapes.add_textbox(x + Inches(0.28), y + Inches(0.18), w - Inches(0.4), Inches(0.35))
    add_text(tb, title, size=15, bold=True, color=INK)
    body = slide.shapes.add_textbox(x + Inches(0.28), y + Inches(0.55), w - Inches(0.4), h - Inches(0.7))
    tf = add_text(body, "•  " + lines[0], size=12, color=MUTED)
    for line in lines[1:]:
        add_para(tf, "•  " + line, size=12, color=MUTED, space_before=4)


def metric_card(slide, x, y, w, h, value, label):
    round_rect(slide, x, y, w, h, fill=WHITE, line=LINE)
    v = slide.shapes.add_textbox(x + Inches(0.15), y + Inches(0.25), w - Inches(0.3), Inches(0.55))
    add_text(v, value, size=26, bold=True, color=ACCENT, align=PP_ALIGN.CENTER, font="Georgia")
    l = slide.shapes.add_textbox(x + Inches(0.15), y + Inches(0.85), w - Inches(0.3), Inches(0.55))
    add_text(l, label, size=11, color=MUTED, align=PP_ALIGN.CENTER)


def simple_bar_chart(slide, x, y, w, h, values, labels, title=""):
    round_rect(slide, x, y, w, h, fill=WHITE, line=LINE)
    if title:
        t = slide.shapes.add_textbox(x + Inches(0.25), y + Inches(0.15), w - Inches(0.5), Inches(0.3))
        add_text(t, title, size=13, bold=True, color=INK)

    chart_top = y + Inches(0.55)
    chart_bottom = y + h - Inches(0.55)
    chart_left = x + Inches(0.55)
    chart_right = x + w - Inches(0.35)
    chart_h = chart_bottom - chart_top
    chart_w = chart_right - chart_left
    max_v = max(values) * 1.08
    n = len(values)
    gap = chart_w / n
    bar_w = gap * 0.55

    for i in range(4):
        gy = chart_top + chart_h * i / 3
        rect(slide, chart_left, gy, chart_w, Pt(1), fill=LINE)

    for i, (val, lab) in enumerate(zip(values, labels)):
        bh = int(chart_h * (val / max_v))
        bx = chart_left + gap * i + (gap - bar_w) / 2
        by = chart_bottom - bh
        rect(slide, bx, by, bar_w, bh, fill=ACCENT if i == n - 1 else TEAL)
        vt = slide.shapes.add_textbox(bx - Inches(0.15), by - Inches(0.28), bar_w + Inches(0.3), Inches(0.25))
        add_text(vt, f"{val:,}", size=10, bold=True, color=INK, align=PP_ALIGN.CENTER)
        lt = slide.shapes.add_textbox(bx - Inches(0.2), chart_bottom + Inches(0.08), bar_w + Inches(0.4), Inches(0.35))
        add_text(lt, lab, size=10, color=MUTED, align=PP_ALIGN.CENTER)


def feature_slide(prs, no, section, title, subtitle, bullets, image_name, process_tag=""):
    slide = blank_content(prs)
    content_bg(slide)
    top_bar(slide, no, TOTAL, section)
    if process_tag:
        round_rect(slide, Inches(0.5), Inches(0.78), Inches(1.8), Inches(0.32), fill=ACCENT_SOFT, line=None)
        tg = slide.shapes.add_textbox(Inches(0.5), Inches(0.82), Inches(1.8), Inches(0.28))
        add_text(tg, process_tag, size=11, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
        t = slide.shapes.add_textbox(Inches(2.5), Inches(0.75), Inches(10), Inches(0.4))
        add_text(t, title, size=26, bold=True, color=INK, font="Georgia")
        s = slide.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(6.2), Inches(0.35))
        add_text(s, subtitle, size=13, color=MUTED)
    else:
        title_block(slide, title, subtitle)

    body = slide.shapes.add_textbox(Inches(0.5), Inches(1.7), Inches(6.2), Inches(4.6))
    tf = add_text(body, "•  " + bullets[0], size=15, color=INK)
    for b in bullets[1:]:
        add_para(tf, "•  " + b, size=15, color=INK, space_before=10)

    img = ASSETS / image_name
    round_rect(slide, Inches(7.0), Inches(1.55), Inches(5.8), Inches(5.1), fill=WHITE, line=LINE)
    if img.exists():
        picture(slide, img, Inches(7.25), Inches(1.75), h=Inches(4.7))
    footer(slide)


# ---------------------------------------------------------------------------
# SLIDES
# ---------------------------------------------------------------------------

def slide_cover(prs):
    slide = blank_content(prs)
    # light full background
    rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill=SOFT)
    # left soft panel
    rect(slide, 0, 0, Inches(7.1), SLIDE_H, fill=WHITE)
    rect(slide, 0, 0, Inches(0.12), SLIDE_H, fill=ACCENT)
    # soft right wash
    rect(slide, Inches(7.1), 0, Inches(6.233), SLIDE_H, fill=SKY)

    badge = round_rect(slide, Inches(0.55), Inches(0.55), Inches(4.3), Inches(0.38), fill=ACCENT_SOFT, line=None)
    bt = slide.shapes.add_textbox(Inches(0.55), Inches(0.61), Inches(4.3), Inches(0.3))
    add_text(bt, "MAHARASHTRA  ·  REGISTRATION & STAMPS", size=10, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)

    brand = slide.shapes.add_textbox(Inches(0.55), Inches(1.25), Inches(6.2), Inches(0.35))
    add_text(brand, "PRAKRUTI GRAPHIC PVT. LTD.", size=13, bold=True, color=MUTED)

    title = slide.shapes.add_textbox(Inches(0.55), Inches(1.75), Inches(6.2), Inches(1.5))
    tf = add_text(title, "Secure Legal Paper", size=40, bold=True, color=NAVY, font="Georgia")
    add_para(tf, "for Property Registration", size=28, bold=True, color=NAVY, space_before=4, font="Georgia")

    sub = slide.shapes.add_textbox(Inches(0.55), Inches(3.4), Inches(6.0), Inches(0.4))
    add_text(sub, "Safe paper  +  QR check  +  easy tracking", size=16, bold=True, color=ACCENT)

    points = [
        "Strong paper that does not tear easily",
        "Hidden marks, hologram, foil and raised print",
        "Scan QR to check if paper is real",
        "Made for property registration documents",
    ]
    y = Inches(4.05)
    for p in points:
        oval(slide, Inches(0.65), y + Inches(0.08), Inches(0.14), Inches(0.14), fill=ACCENT)
        tb = slide.shapes.add_textbox(Inches(0.95), y, Inches(5.6), Inches(0.32))
        add_text(tb, p, size=14, color=INK)
        y += Inches(0.4)

    # framed paper image
    img = ASSETS / "cover_paper.jpg"
    if not img.exists():
        img = ASSETS / "full_paper.jpg"
    round_rect(slide, Inches(7.55), Inches(0.45), Inches(5.25), Inches(6.55), fill=WHITE, line=LINE)
    if img.exists():
        picture(slide, img, Inches(7.75), Inches(0.65), h=Inches(6.15))

    c = slide.shapes.add_textbox(Inches(0.55), Inches(6.85), Inches(6.2), Inches(0.35))
    add_text(c, "prakruti@prakrutigraphic.com  ·  +91 98212 32349", size=11, color=MUTED)


def slide_agenda(prs, no):
    slide = blank_content(prs)
    content_bg(slide)
    top_bar(slide, no, TOTAL, "Overview")
    title_block(slide, "What's Inside", "Problem → Facts → Product → Features → Digital Check → Result")

    items = [
        ("01", "The Problem", "Fraud, fake papers and weak checks"),
        ("02", "Maharashtra Facts", "Big money + recent fraud cases"),
        ("03", "Our Secure Paper", "Strong paper with many safety marks"),
        ("04", "Paper Features", "Step-by-step safety layers on paper"),
        ("05", "QR & Tracking", "Scan, check and follow each paper"),
        ("06", "Final Benefit", "More control, more trust, less loss"),
    ]
    for i, (num, title, desc) in enumerate(items):
        col = i % 3
        row = i // 3
        x = Inches(0.5) + Inches(4.15) * col
        y = Inches(1.9) + Inches(2.2) * row
        round_rect(slide, x, y, Inches(3.95), Inches(1.9), fill=WHITE, line=LINE)
        n = slide.shapes.add_textbox(x + Inches(0.25), y + Inches(0.3), Inches(3.4), Inches(0.4))
        add_text(n, num, size=22, bold=True, color=ACCENT, font="Georgia")
        t = slide.shapes.add_textbox(x + Inches(0.25), y + Inches(0.8), Inches(3.4), Inches(0.35))
        add_text(t, title, size=16, bold=True, color=INK)
        d = slide.shapes.add_textbox(x + Inches(0.25), y + Inches(1.2), Inches(3.4), Inches(0.5))
        add_text(d, desc, size=13, color=MUTED)
    footer(slide)


def slide_why_now(prs, no):
    slide = blank_content(prs)
    content_bg(slide)
    top_bar(slide, no, TOTAL, "Context")
    title_block(slide, "Why Do We Need Secure Legal Paper Now?", "Property papers handle big money — fake papers cause big loss")

    metric_card(slide, Inches(0.5), Inches(1.85), Inches(3.9), Inches(1.6), "₹59,150 Cr", "Stamp & registration money\ncollected in FY 2024–25")
    metric_card(slide, Inches(4.7), Inches(1.85), Inches(3.9), Inches(1.6), "13.4 Lakh+", "Property documents\nregistered in one year")
    metric_card(slide, Inches(8.9), Inches(1.85), Inches(3.9), Inches(1.6), "₹10–20k Cr", "Possible loss under\nstate-level fraud probe*")

    bullet_card(
        slide, Inches(0.5), Inches(3.75), Inches(6.0), Inches(2.85),
        "What is at risk",
        [
            "Property papers prove who owns the land or house",
            "Banks also depend on these papers for loans",
            "Fake papers can cheat buyers and the government",
            "State needs stronger paper to protect public money",
        ],
    )
    bullet_card(
        slide, Inches(6.8), Inches(3.75), Inches(6.0), Inches(2.85),
        "Problem with today's paper",
        [
            "Mostly simple green / plain legal paper",
            "Easy to copy, tear, or make a fake look real",
            "No easy way to scan and confirm it is original",
            "Weak safety checks in the full process",
        ],
        accent=RED,
    )
    footer(slide)
    note = slide.shapes.add_textbox(Inches(0.5), Inches(6.7), Inches(12), Inches(0.25))
    add_text(note, "*Sources: CAG report, IGR Maharashtra, news reports on stamp duty probe.", size=9, color=MUTED)


def slide_problem_5points(prs, no):
    slide = blank_content(prs)
    content_bg(slide)
    top_bar(slide, no, TOTAL, "Problem")
    title_block(slide, "Five Big Problems in Documents", "These gaps make fraud and fake papers easier")

    points = [
        ("1", "Fraud in the past", "Fake registration and wrong process cases found", RED),
        ("2", "Same paper copied again", "Fake sale deeds, fake seals, double selling", RED),
        ("3", "No strong checks", "No QR / safety marks at every step", ACCENT),
        ("4", "Paper and digital not joined", "Online systems and paper safety stay separate", TEAL),
        ("5", "Need easy tracking", "Anyone should quickly check if paper is real", GREEN),
    ]
    for i, (num, title, desc, color) in enumerate(points):
        y = Inches(1.75) + Inches(0.95) * i
        round_rect(slide, Inches(0.5), y, Inches(12.3), Inches(0.85), fill=WHITE, line=LINE)
        oval(slide, Inches(0.7), y + Inches(0.2), Inches(0.45), Inches(0.45), fill=color)
        nt = slide.shapes.add_textbox(Inches(0.7), y + Inches(0.28), Inches(0.45), Inches(0.35))
        add_text(nt, num, size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tt = slide.shapes.add_textbox(Inches(1.4), y + Inches(0.15), Inches(10.8), Inches(0.3))
        add_text(tt, title, size=16, bold=True, color=INK)
        dd = slide.shapes.add_textbox(Inches(1.4), y + Inches(0.45), Inches(10.8), Inches(0.3))
        add_text(dd, desc, size=13, color=MUTED)
    footer(slide)


def slide_fraud_facts(prs, no):
    slide = blank_content(prs)
    content_bg(slide)
    top_bar(slide, no, TOTAL, "Problem 01 · Fraud")
    title_block(slide, "Fraud Found in Past", "Recent Maharashtra cases show how weak checks can cause big loss")

    cases = [
        ("Navi Mumbai", "₹13.91 Cr", "Many property papers of illegal buildings were registered with less stamp duty. Case filed against officers and agents."),
        ("Full State Check", "₹10–20k Cr*", "Government ordered a big check on stamp duty cases from 2021 to 2026 after fraud reports."),
        ("Marathwada Cases", "16+ cases", "Fake ownership papers, fake inheritance papers, and double selling. Even AI-made fake papers reported."),
        ("Satara Land Case", "Fake seals", "Fake sale deed and fake office seals used. Registration number was not found in real records."),
    ]
    for i, (place, value, text) in enumerate(cases):
        col = i % 2
        row = i // 2
        x = Inches(0.5) + Inches(6.35) * col
        y = Inches(1.8) + Inches(2.35) * row
        round_rect(slide, x, y, Inches(6.1), Inches(2.15), fill=WHITE, line=LINE)
        rect(slide, x, y, Inches(0.12), Inches(2.15), fill=RED)
        p = slide.shapes.add_textbox(x + Inches(0.35), y + Inches(0.2), Inches(5.5), Inches(0.3))
        add_text(p, place, size=12, bold=True, color=MUTED)
        v = slide.shapes.add_textbox(x + Inches(0.35), y + Inches(0.5), Inches(5.5), Inches(0.4))
        add_text(v, value, size=24, bold=True, color=RED, font="Georgia")
        d = slide.shapes.add_textbox(x + Inches(0.35), y + Inches(1.05), Inches(5.5), Inches(0.9))
        add_text(d, text, size=12, color=INK)
    footer(slide)


def slide_duplication(prs, no):
    slide = blank_content(prs)
    content_bg(slide)
    top_bar(slide, no, TOTAL, "Problem 02 · Copying")
    title_block(slide, "Document Copying Found in Past", "When paper is easy to copy, fake papers look almost real")

    left = [
        "Fake sale deeds used to transfer property",
        "Fake office seals and fake signatures",
        "Same property sold more than once",
        "Fake ownership / inheritance papers",
        "Computer-made fake papers also found",
    ]
    right = [
        "Simple paper copies can look good enough",
        "No clear 'FAKE COPY' mark on photocopy",
        "Hard for officers to feel or see the difference",
        "Buyers and banks cannot check quickly",
        "Each paper has no unique digital ID",
    ]
    bullet_card(slide, Inches(0.5), Inches(1.85), Inches(6.0), Inches(4.4), "What police / officers keep finding", left, accent=RED)
    bullet_card(slide, Inches(6.8), Inches(1.85), Inches(6.0), Inches(4.4), "Why plain green paper fails", right, accent=ACCENT)
    footer(slide)


def slide_no_control(prs, no):
    slide = blank_content(prs)
    content_bg(slide)
    top_bar(slide, no, TOTAL, "Problem 03 · No Checks")
    title_block(slide, "No Strong Control on the Process", "Because safety checks are missing from start to end")

    stages = [
        ("Draft", "Paper written\non plain sheet"),
        ("Stamp", "Stamp duty\npaid / paper given"),
        ("Register", "Office signs\nthe document"),
        ("Record", "Land record\ngets updated"),
        ("Bank", "Bank / buyer\nchecks paper"),
    ]
    for i, (title, desc) in enumerate(stages):
        x = Inches(0.45) + Inches(2.55) * i
        round_rect(slide, x, Inches(1.95), Inches(2.35), Inches(1.7), fill=WHITE, line=LINE)
        t = slide.shapes.add_textbox(x + Inches(0.1), Inches(2.15), Inches(2.15), Inches(0.35))
        add_text(t, f"{i+1}. {title}", size=14, bold=True, color=INK, align=PP_ALIGN.CENTER)
        d = slide.shapes.add_textbox(x + Inches(0.1), Inches(2.6), Inches(2.15), Inches(0.8))
        add_text(d, desc, size=12, color=MUTED, align=PP_ALIGN.CENTER)
        if i < 4:
            ar = slide.shapes.add_textbox(x + Inches(2.2), Inches(2.55), Inches(0.4), Inches(0.4))
            add_text(ar, "→", size=20, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)

    round_rect(slide, Inches(0.5), Inches(4.0), Inches(12.3), Inches(2.5), fill=WHITE, line=LINE)
    rect(slide, Inches(0.5), Inches(4.0), Inches(12.3), Inches(0.08), fill=RED)
    h = slide.shapes.add_textbox(Inches(0.8), Inches(4.25), Inches(11.7), Inches(0.35))
    add_text(h, "Gap today: process is open — safety locks are missing", size=16, bold=True, color=RED)
    gaps = [
        "No strong secure paper on every important sheet",
        "No unique QR linked to each paper",
        "No quick Real / Fake check for officers and banks",
        "No live record of who scanned the paper and where",
    ]
    body = slide.shapes.add_textbox(Inches(0.8), Inches(4.75), Inches(11.7), Inches(1.5))
    tf = add_text(body, "•  " + gaps[0], size=13, color=INK)
    for g in gaps[1:]:
        add_para(tf, "•  " + g, size=13, color=INK, space_before=4)
    footer(slide)


def slide_solution_intro(prs, no):
    slide = blank_content(prs)
    content_bg(slide)
    top_bar(slide, no, TOTAL, "Solution")
    title_block(slide, "Our Secure Legal Paper", "Strong paper made in-house — with many safety marks + QR check")

    img = ASSETS / "process_01.jpg"
    picture(slide, img, Inches(0.5), Inches(1.8), h=Inches(4.9))

    specs = [
        ("Use", "Property registration / stamp department papers"),
        ("Size", "Legal size — 216 × 356 mm"),
        ("Paper", "Thick strong paper (about 140 micron) — hard to tear"),
        ("For", "Maharashtra Registration & Stamps Department"),
        ("On paper", "Watermark, copy-stop mark, UV mark, hologram, foil, raised print"),
        ("Digital", "QR code on paper for Real / Fake check and tracking"),
    ]
    y = Inches(1.8)
    for title, desc in specs:
        round_rect(slide, Inches(5.4), y, Inches(7.4), Inches(0.72), fill=WHITE, line=LINE)
        t = slide.shapes.add_textbox(Inches(5.65), y + Inches(0.08), Inches(7.0), Inches(0.28))
        add_text(t, title, size=11, bold=True, color=ACCENT)
        d = slide.shapes.add_textbox(Inches(5.65), y + Inches(0.35), Inches(7.0), Inches(0.3))
        add_text(d, desc, size=13, color=INK)
        y += Inches(0.8)
    footer(slide)


def slide_combo(prs, no):
    slide = blank_content(prs)
    content_bg(slide)
    top_bar(slide, no, TOTAL, "Problem 04 · Together")
    title_block(slide, "Paper Safety + Digital Check — Together", "Only paper or only online is not enough. Both must work as one.")

    round_rect(slide, Inches(0.5), Inches(1.85), Inches(5.9), Inches(4.55), fill=WHITE, line=LINE)
    rect(slide, Inches(0.5), Inches(1.85), Inches(5.9), Inches(0.55), fill=NAVY)
    ht = slide.shapes.add_textbox(Inches(0.5), Inches(1.95), Inches(5.9), Inches(0.4))
    add_text(ht, "ON THE PAPER", size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    phys = [
        "Strong paper that does not tear easily",
        "Watermark and fine line patterns",
        "Copy shows warning mark",
        "Rainbow border design",
        "Hidden mark seen under special light",
        "Hologram and shiny foil logo",
        "Raised MH print you can feel",
    ]
    body = slide.shapes.add_textbox(Inches(0.85), Inches(2.6), Inches(5.3), Inches(3.5))
    tf = add_text(body, "•  " + phys[0], size=14, color=INK)
    for p in phys[1:]:
        add_para(tf, "•  " + p, size=14, color=INK, space_before=6)

    round_rect(slide, Inches(6.9), Inches(1.85), Inches(5.9), Inches(4.55), fill=WHITE, line=LINE)
    rect(slide, Inches(6.9), Inches(1.85), Inches(5.9), Inches(0.55), fill=TEAL)
    ht2 = slide.shapes.add_textbox(Inches(6.9), Inches(1.95), Inches(5.9), Inches(0.4))
    add_text(ht2, "ON PHONE / COMPUTER", size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    dig = [
        "Unique QR on every sheet",
        "Check against secure online record",
        "Instant Real / Fake result",
        "See basic paper information",
        "Scan place and time gets saved",
        "Admin can track all scans",
        "Safety check at every step",
    ]
    body2 = slide.shapes.add_textbox(Inches(7.25), Inches(2.6), Inches(5.3), Inches(3.5))
    tf2 = add_text(body2, "•  " + dig[0], size=14, color=INK)
    for p in dig[1:]:
        add_para(tf2, "•  " + p, size=14, color=INK, space_before=6)
    footer(slide)


def slide_process_overview(prs, no):
    slide = blank_content(prs)
    content_bg(slide)
    top_bar(slide, no, TOTAL, "How Paper is Made")
    title_block(slide, "8 Safety Steps on One Paper", "Each step adds one more lock against fake papers")

    processes = [
        ("1", "Strong Base", "Thick green secure paper"),
        ("2", "Watermark", "Hidden mark in paper"),
        ("3", "Copy Stop", "Copy shows warning"),
        ("4", "Colour Print", "Official government look"),
        ("5", "Hidden Ink", "Seen under special light"),
        ("6", "Hologram", "Shiny security stamp"),
        ("7", "Metal Logo", "Hot foil government seal"),
        ("8", "Raised Print", "MH mark you can feel"),
    ]
    for i, (num, title, desc) in enumerate(processes):
        col = i % 4
        row = i // 4
        x = Inches(0.5) + Inches(3.2) * col
        y = Inches(1.9) + Inches(2.35) * row
        round_rect(slide, x, y, Inches(3.0), Inches(2.1), fill=WHITE, line=LINE)
        oval(slide, x + Inches(1.15), y + Inches(0.25), Inches(0.7), Inches(0.7), fill=ACCENT)
        n = slide.shapes.add_textbox(x + Inches(1.15), y + Inches(0.4), Inches(0.7), Inches(0.4))
        add_text(n, num, size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        t = slide.shapes.add_textbox(x + Inches(0.15), y + Inches(1.1), Inches(2.7), Inches(0.35))
        add_text(t, title, size=14, bold=True, color=INK, align=PP_ALIGN.CENTER)
        d = slide.shapes.add_textbox(x + Inches(0.15), y + Inches(1.45), Inches(2.7), Inches(0.45))
        add_text(d, desc, size=12, color=MUTED, align=PP_ALIGN.CENTER)
    footer(slide)


def slide_feature_gallery(prs, no):
    slide = blank_content(prs)
    content_bg(slide)
    top_bar(slide, no, TOTAL, "Real Photos")
    title_block(slide, "Real Close-up Photos of Safety Marks", "Taken from our sample secure legal paper")

    gallery = [
        ("uv_guilloche.jpg", "Rainbow border + special light"),
        ("uv_invisible.jpg", "Hidden mark under light"),
        ("mh_emboss_b.jpg", "Raised MH print"),
        ("gold_foil.jpg", "Shiny foil seal"),
        ("hologram_sq.jpg", "Security hologram"),
        ("full_paper.jpg", "Full paper sample"),
    ]
    for i, (name, label) in enumerate(gallery):
        col = i % 3
        row = i // 3
        x = Inches(0.45) + Inches(4.25) * col
        y = Inches(1.75) + Inches(2.5) * row
        round_rect(slide, x, y, Inches(4.05), Inches(2.3), fill=WHITE, line=LINE)
        img = ASSETS / name
        if img.exists():
            picture(slide, img, x + Inches(0.15), y + Inches(0.12), h=Inches(1.7))
        lb = slide.shapes.add_textbox(x + Inches(0.1), y + Inches(1.85), Inches(3.85), Inches(0.35))
        add_text(lb, label, size=12, bold=True, color=INK, align=PP_ALIGN.CENTER)
    footer(slide)


def slide_features_matrix(prs, no):
    slide = blank_content(prs)
    content_bg(slide)
    top_bar(slide, no, TOTAL, "Paper Features")
    title_block(slide, "Safety Marks — Easy View", "See · Feel · Scan — many ways to confirm the paper is real")

    rows = [
        ("Feature", "Type", "How it helps"),
        ("Strong thick paper", "See / Feel", "Hard to tear or replace with normal paper"),
        ("Watermark", "See in light", "Copy cannot carry the same mark"),
        ("Copy-stop mark", "On photocopy", "Copy shows a warning word"),
        ("Rainbow border", "See", "Fine design is hard to print at home"),
        ("Hidden ink", "Special light", "Secret Maharashtra marks appear"),
        ("Hologram + foil", "See", "Shiny marks are hard to fake"),
        ("Raised MH print", "Feel", "Touch and confirm it is original"),
        ("Unique QR", "Phone scan", "Online check + reuse warning"),
    ]
    y0 = Inches(1.7)
    for i, (a, b, c) in enumerate(rows):
        y = y0 + Inches(0.52) * i
        fill = NAVY if i == 0 else (SOFT if i % 2 == 0 else WHITE)
        fc = WHITE if i == 0 else INK
        rect(slide, Inches(0.5), y, Inches(3.2), Inches(0.5), fill=fill)
        rect(slide, Inches(3.7), y, Inches(2.2), Inches(0.5), fill=fill)
        rect(slide, Inches(5.9), y, Inches(6.9), Inches(0.5), fill=fill)
        ta = slide.shapes.add_textbox(Inches(0.6), y + Inches(0.1), Inches(3.0), Inches(0.35))
        tb = slide.shapes.add_textbox(Inches(3.8), y + Inches(0.1), Inches(2.0), Inches(0.35))
        tc = slide.shapes.add_textbox(Inches(6.0), y + Inches(0.1), Inches(6.7), Inches(0.35))
        add_text(ta, a, size=12, bold=(i == 0), color=fc)
        add_text(tb, b, size=12, bold=(i == 0), color=fc)
        add_text(tc, c, size=12, bold=(i == 0), color=fc)
    footer(slide)


def slide_digital_support(prs, no):
    slide = blank_content(prs)
    content_bg(slide)
    top_bar(slide, no, TOTAL, "Problem 05 · Tracking")
    title_block(slide, "Strong Digital Support for Checking & Tracking", "Scan once → know if paper is real → see where it was checked")

    steps = [
        ("1", "Scan QR", "Officer / bank / person scans the paper"),
        ("2", "Online Check", "System matches paper with secure record"),
        ("3", "See Info", "Basic paper status comes on screen"),
        ("4", "Result", "Real · Fake · Already Used"),
        ("5", "Save Log", "Time and place saved for tracking"),
    ]
    for i, (num, title, desc) in enumerate(steps):
        x = Inches(0.4) + Inches(2.55) * i
        round_rect(slide, x, Inches(1.9), Inches(2.4), Inches(2.4), fill=WHITE, line=LINE)
        oval(slide, x + Inches(0.85), Inches(2.15), Inches(0.7), Inches(0.7), fill=TEAL)
        n = slide.shapes.add_textbox(x + Inches(0.85), Inches(2.3), Inches(0.7), Inches(0.4))
        add_text(n, num, size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        t = slide.shapes.add_textbox(x + Inches(0.15), Inches(3.05), Inches(2.1), Inches(0.35))
        add_text(t, title, size=13, bold=True, color=INK, align=PP_ALIGN.CENTER)
        d = slide.shapes.add_textbox(x + Inches(0.15), Inches(3.4), Inches(2.1), Inches(0.7))
        add_text(d, desc, size=11, color=MUTED, align=PP_ALIGN.CENTER)

    for i, name in enumerate(["qr_scan.jpg", "verified.jpg", "dashboard_v3.jpg"]):
        x = Inches(0.5) + Inches(4.2) * i
        round_rect(slide, x, Inches(4.55), Inches(4.0), Inches(2.0), fill=WHITE, line=LINE)
        img = ASSETS / name
        if img.exists():
            picture(slide, img, x + Inches(0.15), Inches(4.65), h=Inches(1.8))
    footer(slide)


def slide_qr_detail(prs, no):
    slide = blank_content(prs)
    content_bg(slide)
    top_bar(slide, no, TOTAL, "QR Check")
    title_block(slide, "Unique QR on Every Sheet", "Each paper gets its own code linked to a secure online record")

    bullets = [
        "Every sheet gets its own number and QR",
        "QR is saved in a secure online system",
        "Every scan checks the original record",
        "Useful for citizens, banks and officers",
        "Admin can make batches and watch scans",
        "System can warn if same QR is used again",
    ]
    body = slide.shapes.add_textbox(Inches(0.5), Inches(1.85), Inches(6.3), Inches(4.5))
    tf = add_text(body, "•  " + bullets[0], size=15, color=INK)
    for b in bullets[1:]:
        add_para(tf, "•  " + b, size=15, color=INK, space_before=10)

    img = ASSETS / "unique_qr_v3.jpg"
    round_rect(slide, Inches(7.0), Inches(1.75), Inches(5.8), Inches(4.7), fill=WHITE, line=LINE)
    if img.exists():
        picture(slide, img, Inches(7.2), Inches(2.1), w=Inches(5.4))
    footer(slide)


def slide_track_trace(prs, no):
    slide = blank_content(prs)
    content_bg(slide)
    top_bar(slide, no, TOTAL, "Tracking")
    title_block(slide, "Tracking + Live Scan Record", "From issue to registration to later check — full visibility")

    left = [
        "Create unique number + QR in secure system",
        "Print / issue on our secure legal paper",
        "Scan at stamp, registration or bank points",
        "Citizen or officer can verify on phone / portal",
        "Admin sees Real / Fake / Already-used with place & time",
    ]
    bullet_card(slide, Inches(0.5), Inches(1.85), Inches(6.2), Inches(4.55), "Full control loop", left, accent=TEAL)

    img = ASSETS / "dashboard_v3.jpg"
    round_rect(slide, Inches(7.0), Inches(1.85), Inches(5.8), Inches(4.55), fill=WHITE, line=LINE)
    if img.exists():
        picture(slide, img, Inches(7.35), Inches(2.2), w=Inches(5.1))
    footer(slide)


def slide_revenue_chart(prs, no):
    slide = blank_content(prs)
    content_bg(slide)
    top_bar(slide, no, TOTAL, "Facts & Numbers")
    title_block(slide, "Maharashtra Stamp & Registration Money", "A very big system — needs very strong paper")

    values = [25288, 35029, 40099, 57670]
    labels = ["FY21-22", "FY22-23", "FY23-24", "FY24-25*"]
    simple_bar_chart(
        slide, Inches(0.5), Inches(1.75), Inches(7.4), Inches(4.8),
        values, labels, title="Money collected from stamp & registration (₹ Crore)",
    )

    round_rect(slide, Inches(8.2), Inches(1.75), Inches(4.6), Inches(4.8), fill=WHITE, line=LINE)
    h = slide.shapes.add_textbox(Inches(8.45), Inches(2.0), Inches(4.1), Inches(0.4))
    add_text(h, "Why this matters", size=16, bold=True, color=INK)
    notes = [
        "Many documents every year",
        "Even small fraud causes big loss",
        "Recent cases show pressure",
        "Secure paper reduces fake risk",
        "Quick check builds trust",
    ]
    body = slide.shapes.add_textbox(Inches(8.45), Inches(2.55), Inches(4.1), Inches(3.5))
    tf = add_text(body, "•  " + notes[0], size=13, color=MUTED)
    for n in notes[1:]:
        add_para(tf, "•  " + n, size=13, color=MUTED, space_before=8)

    note = slide.shapes.add_textbox(Inches(0.5), Inches(6.65), Inches(12.3), Inches(0.3))
    add_text(note, "*FY24–25 about ₹57,670 Cr (news). Older years from IGR data. Shown for simple understanding.", size=9, color=MUTED)
    footer(slide)


def slide_before_after(prs, no):
    slide = blank_content(prs)
    content_bg(slide)
    top_bar(slide, no, TOTAL, "Result")
    title_block(slide, "Before vs After", "From weak plain paper to strong secure legal paper")

    headers = ["Point", "Today (Plain / Green)", "Our Secure Paper"]
    rows = [
        ["Paper strength", "Tears / easy to replace", "Strong, hard to tear"],
        ["Copy safety", "Copy looks almost same", "Copy shows warning mark"],
        ["Visual check", "Few clear signs", "Hologram, foil, border, raised print"],
        ["Digital ID", "Usually missing", "Unique QR on every sheet"],
        ["Tracking", "Hard / manual", "Live scan tracking"],
        ["Process control", "Weak checks", "Paper + QR checks together"],
    ]
    y = Inches(1.7)
    widths = [Inches(3.0), Inches(4.5), Inches(4.8)]
    xs = [Inches(0.5), Inches(3.5), Inches(8.0)]
    for x, w, htxt in zip(xs, widths, headers):
        rect(slide, x, y, w, Inches(0.45), fill=NAVY)
        t = slide.shapes.add_textbox(x + Inches(0.1), y + Inches(0.08), w - Inches(0.2), Inches(0.3))
        add_text(t, htxt, size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    for i, row in enumerate(rows):
        y = Inches(2.15) + Inches(0.7) * i
        fills = [SOFT if i % 2 == 0 else WHITE, RGBColor(0xFB, 0xEE, 0xEB), RGBColor(0xE8, 0xF3, 0xEC)]
        colors = [INK, RED, GREEN]
        for j, (x, w, txt) in enumerate(zip(xs, widths, row)):
            rect(slide, x, y, w, Inches(0.7), fill=fills[j])
            t = slide.shapes.add_textbox(x + Inches(0.12), y + Inches(0.18), w - Inches(0.24), Inches(0.4))
            add_text(t, txt, size=12, bold=(j == 0), color=colors[j], align=PP_ALIGN.CENTER if j else PP_ALIGN.LEFT)
    footer(slide)


def slide_benefits(prs, no):
    slide = blank_content(prs)
    content_bg(slide)
    top_bar(slide, no, TOTAL, "Who Gains")
    title_block(slide, "Who Gains — and How", "One secure paper helps government, people and banks")

    cards = [
        ("State / Department", ["Less fake paper risk", "Better process control", "Protects public money"]),
        ("Registration Office", ["Easy Real / Fake check", "Feel, see and scan marks", "Fewer disputes"]),
        ("Citizens", ["Harder to fake deeds", "Scan and feel safe", "Safer property deals"]),
        ("Banks", ["Quick paper check", "Lower document risk", "Faster trust on papers"]),
    ]
    for i, (title, lines) in enumerate(cards):
        x = Inches(0.45) + Inches(3.2) * i
        round_rect(slide, x, Inches(1.9), Inches(3.05), Inches(4.4), fill=WHITE, line=LINE)
        rect(slide, x, Inches(1.9), Inches(3.05), Inches(0.7), fill=NAVY)
        t = slide.shapes.add_textbox(x, Inches(2.05), Inches(3.05), Inches(0.4))
        add_text(t, title, size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        body = slide.shapes.add_textbox(x + Inches(0.2), Inches(2.9), Inches(2.65), Inches(3.0))
        tf = add_text(body, "•  " + lines[0], size=13, color=INK)
        for line in lines[1:]:
            add_para(tf, "•  " + line, size=13, color=INK, space_before=12)
    footer(slide)


def slide_closing(prs, no):
    slide = blank_content(prs)
    rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill=SOFT)
    rect(slide, 0, 0, Inches(0.12), SLIDE_H, fill=ACCENT)
    round_rect(slide, Inches(0.8), Inches(1.2), Inches(11.7), Inches(5.1), fill=WHITE, line=LINE)

    t = slide.shapes.add_textbox(Inches(1.1), Inches(1.7), Inches(11.1), Inches(0.7))
    add_text(t, "Make the Paper Safe. Control the Process.", size=28, bold=True, color=NAVY, align=PP_ALIGN.CENTER, font="Georgia")

    s = slide.shapes.add_textbox(Inches(1.6), Inches(2.5), Inches(10.1), Inches(0.7))
    add_text(
        s,
        "Our secure legal paper joins strong paper marks with QR checking and tracking\nfor Maharashtra Registration & Stamps documents.",
        size=15, color=MUTED, align=PP_ALIGN.CENTER,
    )

    points = ["Stop easy copying", "Add real checks", "Join paper + digital", "Track and verify"]
    for i, p in enumerate(points):
        x = Inches(1.3) + Inches(2.7) * i
        round_rect(slide, x, Inches(3.5), Inches(2.5), Inches(0.9), fill=ACCENT_SOFT, line=None)
        tb = slide.shapes.add_textbox(x + Inches(0.1), Inches(3.75), Inches(2.3), Inches(0.45))
        add_text(tb, p, size=13, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)

    c = slide.shapes.add_textbox(Inches(1.1), Inches(4.7), Inches(11.1), Inches(1.2))
    tf = add_text(c, "Next step: Sample check / small pilot with the department", size=15, bold=True, color=INK, align=PP_ALIGN.CENTER)
    add_para(tf, "prakruti@prakrutigraphic.com  ·  +91 98212 32349  ·  www.pgpltechprint.com", size=13, color=MUTED, align=PP_ALIGN.CENTER, space_before=10)
    add_para(tf, "Thank You", size=20, bold=True, color=NAVY, align=PP_ALIGN.CENTER, space_before=14, font="Georgia")


def build():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    slide_cover(prs)
    n = 2
    slide_agenda(prs, n); n += 1
    slide_why_now(prs, n); n += 1
    slide_problem_5points(prs, n); n += 1
    slide_fraud_facts(prs, n); n += 1
    slide_duplication(prs, n); n += 1
    slide_no_control(prs, n); n += 1
    slide_revenue_chart(prs, n); n += 1
    slide_solution_intro(prs, n); n += 1
    slide_combo(prs, n); n += 1
    slide_process_overview(prs, n); n += 1

    feature_slide(
        prs, n, "Paper Features",
        "Strong Paper That Does Not Tear Easily",
        "Made in-house — thicker and stronger than normal legal paper",
        [
            "Special strong paper made by us",
            "About 140 micron thick — hard to tear",
            "Harder to replace with normal office paper",
            "Feels better and safer than plain green paper",
            "Base for all other safety marks",
        ],
        "process_02.jpg", "STEP 01",
    ); n += 1

    feature_slide(
        prs, n, "Paper Features",
        "Watermark Safety",
        "Hidden mark inside the paper — copy cannot keep it",
        [
            "Watermark added during paper making",
            "Can be checked in light",
            "Normal photocopy cannot copy it properly",
            "Helps confirm the paper is original",
        ],
        "process_03.jpg", "STEP 02",
    ); n += 1

    feature_slide(
        prs, n, "Paper Features",
        "Copy-Stop Printing",
        "If someone photocopies, a warning mark can appear",
        [
            "Special print that stops easy copying",
            "Original looks clean",
            "Copy can show a warning word",
            "Direct help against fake duplicate papers",
        ],
        "process_04.jpg", "STEP 03",
    ); n += 1

    feature_slide(
        prs, n, "Paper Features",
        "Rainbow Border · Official Print · Hidden Ink",
        "Fine border + government look + secret marks under special light",
        [
            "Fine rainbow / wave border — hard to print at home",
            "Official Maharashtra department printing",
            "Hidden 'महाराष्ट्र' marks on the sides",
            "Seen under special light for quick check",
            "Real close-up photos available for demo",
        ],
        "uv_guilloche.jpg", "STEP 04–05",
    ); n += 1

    feature_slide(
        prs, n, "Paper Features",
        "Hologram · Shiny Foil · Raised Print",
        "See and feel these marks to confirm the paper is real",
        [
            "Hologram stamp for shiny optical safety",
            "Metal-look Maharashtra logo with hot foil",
            "Raised MH print that you can feel by hand",
            "Very hard to make on normal printers",
            "Strong visual warning against fake paper",
        ],
        "hologram_sq.jpg", "STEP 06–08",
    ); n += 1

    slide_feature_gallery(prs, n); n += 1
    slide_features_matrix(prs, n); n += 1
    slide_digital_support(prs, n); n += 1
    slide_qr_detail(prs, n); n += 1
    slide_track_trace(prs, n); n += 1
    slide_before_after(prs, n); n += 1
    slide_benefits(prs, n); n += 1
    slide_closing(prs, n)

    print(f"Slides built: {len(prs.slides)}")
    prs.save(str(OUT))
    print(f"Saved: {OUT} ({OUT.stat().st_size/1024/1024:.1f} MB)")


if __name__ == "__main__":
    build()
