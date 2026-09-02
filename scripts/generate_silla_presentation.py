#!/usr/bin/env python3
"""Generate a branded SILA Dawai PowerPoint from the 12-slide section in the strategy Markdown.

Usage:
    pip install -r scripts/requirements-silla-presentation.txt
    python scripts/generate_silla_presentation.py

The script intentionally reads the slide copy from the Markdown file instead of duplicating
it, so edits to the 12-slide section flow into the deck on the next run.
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import List, Dict

try:
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.util import Inches, Pt
except ImportError:
    raise SystemExit(
        "المكتبة المطلوبة غير مثبتة. شغّل: "
        "pip install -r scripts/requirements-silla-presentation.txt"
    )

# SILA-inspired palette: deep teal, aqua, warm sand and white.
NAVY = RGBColor(7, 48, 58)
TEAL = RGBColor(0, 137, 145)
AQUA = RGBColor(55, 194, 190)
SAND = RGBColor(245, 239, 226)
CREAM = RGBColor(252, 250, 245)
WHITE = RGBColor(255, 255, 255)
INK = RGBColor(27, 48, 53)
MUTED = RGBColor(91, 112, 116)
CORAL = RGBColor(224, 116, 93)
FONT = "Arial"  # Replace with a locally installed Arabic font if preferred.


def parse_slides(markdown: str) -> List[Dict[str, object]]:
    """Extract ### 1 ... ### 12 blocks from the presentation section."""
    marker = re.search(r"^## 10\) هيكل العرض التقديمي", markdown, re.MULTILINE)
    if not marker:
        raise ValueError("لم يتم العثور على قسم «هيكل العرض التقديمي» في الملف.")
    section = markdown[marker.start():]
    matches = list(re.finditer(r"^### (\d+)[.:] (.+)$", section, re.MULTILINE))
    slides: List[Dict[str, object]] = []
    for index, match in enumerate(matches):
        number = int(match.group(1))
        if not 1 <= number <= 12:
            continue
        end = matches[index + 1].start() if index + 1 < len(matches) else len(section)
        raw = section[match.end():end].strip()
        title = match.group(2).strip()
        title = re.sub(r"^([^—]+) — ", "", title) if " — " in title else title
        # Exclude the descriptive label and the قوتها line from the main bullets.
        raw = re.sub(r"^\s*### العنوان.*?(?=\n|$)", "", raw, flags=re.MULTILINE)
        raw = re.sub(r"^\s*\*\*قوتها الإقناعية:\*\*.*?(?:\n|$)", "", raw, flags=re.MULTILINE)
        raw = re.sub(r"^\s*### التصميم المقترح.*?(?:\n|$)", "", raw, flags=re.MULTILINE)
        bullets: List[str] = []
        for line in raw.splitlines():
            line = line.strip()
            if not line or line.startswith(">") or line.startswith("###"):
                continue
            if line.startswith("-"):
                bullets.append(line[1:].strip())
            elif re.match(r"^\d+\.\s", line):
                bullets.append(re.sub(r"^\d+\.\s", "", line))
            elif line.startswith("**") and line.endswith("**"):
                bullets.append(line.strip("*"))
            elif line.startswith("CTA:"):
                bullets.append(line)
        slides.append({"number": number, "title": title, "bullets": bullets})
    slides.sort(key=lambda item: int(item["number"]))
    if len(slides) != 12:
        raise ValueError(f"تم استخراج {len(slides)} شريحة بدلاً من 12.")
    return slides


def set_rtl(paragraph) -> None:
    """Set right alignment and RTL on a paragraph using python-pptx's XML layer."""
    paragraph.alignment = PP_ALIGN.RIGHT
    p_pr = paragraph._p.get_or_add_pPr()
    rtl = p_pr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}rtl")
    if rtl is None:
        from pptx.oxml.xmlchemy import OxmlElement
        rtl = OxmlElement("a:rtl")
        p_pr.append(rtl)
    rtl.set("val", "1")


def add_text(slide, text: str, x, y, w, h, size=20, color=INK,
             bold=False, align=PP_ALIGN.RIGHT, font=FONT):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = Inches(0.08)
    tf.margin_right = Inches(0.08)
    tf.margin_top = Inches(0.04)
    tf.margin_bottom = Inches(0.04)
    p = tf.paragraphs[0]
    p.text = text
    p.alignment = align
    p.font.name = font
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.color.rgb = color
    if align == PP_ALIGN.RIGHT:
        set_rtl(p)
    return box


def add_background(slide, number: int, title: str) -> None:
    bg = slide.background.fill
    bg.solid()
    bg.fore_color.rgb = CREAM
    # Header band.
    band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(1.0))
    band.fill.solid(); band.fill.fore_color.rgb = NAVY; band.line.fill.background()
    # Accent blocks on the left.
    for i, color in enumerate((TEAL, AQUA, CORAL)):
        block = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(i * 0.16), 0, Inches(0.12), Inches(1.0))
        block.fill.solid(); block.fill.fore_color.rgb = color; block.line.fill.background()
    add_text(slide, f"{number:02d}", Inches(0.55), Inches(0.23), Inches(0.7), Inches(0.4), 18, AQUA, True, PP_ALIGN.LEFT)
    add_text(slide, title, Inches(2.0), Inches(0.15), Inches(10.75), Inches(0.58), 25, WHITE, True)
    add_text(slide, "صِلة دوائي  |  SILA IQ", Inches(0.55), Inches(7.05), Inches(3.0), Inches(0.2), 9, MUTED, False, PP_ALIGN.LEFT)
    add_text(slide, "دواؤك… يصلك", Inches(10.8), Inches(7.03), Inches(1.95), Inches(0.2), 9, TEAL, True, PP_ALIGN.RIGHT)


def add_bullets(slide, bullets: List[str], x=Inches(1.0), y=Inches(1.45), w=Inches(11.25), h=Inches(4.8), size=22):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.clear(); tf.word_wrap = True
    tf.margin_left = Inches(0.22); tf.margin_right = Inches(0.18)
    tf.margin_top = Inches(0.08); tf.margin_bottom = Inches(0.08)
    for i, item in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = f"• {item}"
        p.level = 0
        p.font.name = FONT; p.font.size = Pt(size); p.font.color.rgb = INK
        p.space_after = Pt(12)
        p.alignment = PP_ALIGN.RIGHT
        set_rtl(p)
        # Use a visible aqua bullet independent of the RTL bullet direction.
        p._p.get_or_add_pPr().set("marL", "420")
    return box


def add_callout(slide, text: str, color=TEAL):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(5.95), Inches(11.25), Inches(0.65))
    shape.fill.solid(); shape.fill.fore_color.rgb = color; shape.line.fill.background()
    add_text(slide, text, Inches(1.25), Inches(6.05), Inches(10.75), Inches(0.35), 16, WHITE, True)


def build_deck(slides: List[Dict[str, object]], output: Path) -> None:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    for data in slides:
        number = int(data["number"])
        title = str(data["title"])
        bullets = [str(value) for value in data["bullets"]]
        slide = prs.slides.add_slide(blank)
        if number == 1:
            slide.background.fill.solid(); slide.background.fill.fore_color.rgb = NAVY
            # Decorative circles.
            for x, y, diameter, color in [(9.8, 0.5, 3.7, TEAL), (10.6, 1.4, 2.7, AQUA), (8.8, 4.7, 2.0, CORAL)]:
                c = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(diameter), Inches(diameter))
                c.fill.solid(); c.fill.fore_color.rgb = color; c.fill.transparency = 18; c.line.fill.background()
            add_text(slide, "صِلة دوائي", Inches(0.85), Inches(1.25), Inches(7.7), Inches(0.9), 44, WHITE, True)
            add_text(slide, title, Inches(0.85), Inches(2.25), Inches(8.4), Inches(1.35), 31, AQUA, True)
            add_text(slide, "الخطة التسويقية والاستراتيجية للنمو في العراق", Inches(0.85), Inches(4.1), Inches(7.4), Inches(0.5), 20, WHITE)
            add_text(slide, "بغداد → البصرة / النجف / كربلاء", Inches(0.85), Inches(5.0), Inches(6.5), Inches(0.4), 15, SAND)
            add_text(slide, "2 أيلول 2026", Inches(0.85), Inches(6.65), Inches(2.0), Inches(0.25), 10, AQUA, False, PP_ALIGN.LEFT)
            add_text(slide, "دواؤك… يصلك", Inches(10.3), Inches(6.65), Inches(2.0), Inches(0.25), 12, WHITE, True)
            continue

        add_background(slide, number, title)
        # The parser gives us concise bullet content. Long slides get a smaller font.
        size = 19 if len(bullets) > 8 else 22
        add_bullets(slide, bullets[:12], size=size)
        if number == 12:
            add_callout(slide, "الخطوة التالية: إطلاق Pilot مدفوع لمدة 14 يوماً وقياس الطلبات الآمنة والمربحة.", CORAL)
        elif number == 4:
            add_callout(slide, "التموضع: شبكة صيدليات موثّقة، لا متجر واحد ولا دليفري عام.", TEAL)
        elif number == 7:
            add_callout(slide, "الثقة ليست حملة إعلانية؛ إنها نظام تشغيل.", NAVY)
        elif number == 11:
            add_callout(slide, "North Star: طلبات دوائية مكتملة وآمنة لكل صيدلية نشطة.", TEAL)

    output.parent.mkdir(parents=True, exist_ok=True)
    prs.save(output)


def main() -> int:
    parser = argparse.ArgumentParser(description="إنشاء عرض صِلة دوائي من ملف Markdown")
    parser.add_argument("--input", type=Path, default=Path("docs/silla-iq-marketing-strategy-ar.md"))
    parser.add_argument("--output", type=Path, default=Path("docs/silla-iq-marketing-strategy-ar.pptx"))
    args = parser.parse_args()
    if not args.input.exists():
        print(f"الملف غير موجود: {args.input}", file=sys.stderr)
        return 1
    try:
        slides = parse_slides(args.input.read_text(encoding="utf-8"))
        build_deck(slides, args.output)
    except (OSError, ValueError) as exc:
        print(f"تعذر إنشاء العرض: {exc}", file=sys.stderr)
        return 1
    print(f"تم إنشاء العرض: {args.output} ({len(slides)} شريحة)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
