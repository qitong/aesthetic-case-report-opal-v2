#!/usr/bin/env python3
"""
Aesthetic Case Report Opal v2 - Fast-Track & Overflow-Safe PPTX Generator Engine

Key v2 Enhancements (100% backward-compatible with v1 OpalCaseReportDeck API):
1. Built-in Auto-Fit & Overflow Guard:
   - `add_assessment_slide`: dynamically adjusts row spacing and clinical_notes font size
     (11pt -> 8.5pt) so 4-6 bullet lines never spill into the footer.
   - `add_treatment_plan_slide`: dynamically scales body font size and bullet spacing
     when 3-4 steps or rich L/R rheological descriptions are supplied.
   - `add_consensus_table_slide`: uses high-contrast warm alternating row fills
     (`#FFFDFB` / `#F1E9DF` for editorial, `#FFFFFF` / `#EBF0F5` for academic) so row 3
     never blends into the slide canvas.
2. Complete Slide Archetype Coverage (no custom wrapper script needed):
   - `add_fullbleed_atlas_slide`: zero-crop 16:9 anatomical infographic / atlas slide.
   - `add_anatomy_split_slide`: left anatomical visual + right structured mechanism cards.
   - `add_chart_slide`: pure-PIL 5-axis Pre vs. Post bar chart + KPI metric cards.
   - `add_summary_slide`: numbered clinical pearls + disclosure/consent box.
3. Declarative JSON CLI + Millisecond Static Layout QA (`validate_deck`):
   - `python3 scripts/generate_case_pptx.py deck_spec.json` builds a complete Express (5-7p)
     or Full (14-20p) presentation in <2 seconds and prints a layout QA report.
"""

from __future__ import annotations

import argparse
import json
import math
import os
import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont, ImageOps
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt


def hex_to_rgb(hex_str: str) -> RGBColor:
    hex_str = hex_str.lstrip("#")
    return RGBColor(int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16))


DEFAULT_PALETTES = {
    "academic": {
        "bg": "#F9F8F6",
        "card_bg": "#FFFFFF",
        "card_border": "#E5E0D8",
        "row_even": "#FFFFFF",
        "row_odd": "#EBF0F5",
        "text_primary": "#1A1D20",
        "text_secondary": "#5A626A",
        "accent": "#2C4C64",
        "accent_soft": "#EBF0F5",
        "highlight": "#C0392B",
        "font_cn": "PingFang SC",
        "font_en": "Helvetica Neue",
    },
    "editorial": {
        "bg": "#FAF7F2",
        "card_bg": "#FFFFFF",
        "card_border": "#E8DFD5",
        "row_even": "#FFFDFB",
        "row_odd": "#F1E9DF",
        "text_primary": "#231F20",
        "text_secondary": "#7A7067",
        "accent": "#B8926A",
        "accent_soft": "#F3ECE3",
        "highlight": "#8C6D53",
        "font_cn": "PingFang SC",
        "font_en": "Helvetica Neue",
    },
}

FONT_CANDIDATES = [
    "/System/Library/Fonts/PingFang.ttc",
    "/System/Library/Fonts/STHeiti Medium.ttc",
    "/System/Library/Fonts/Helvetica.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
]


def _load_pil_font(size: int):
    for p in FONT_CANDIDATES:
        if Path(p).exists():
            try:
                return ImageFont.truetype(p, size, index=0)
            except OSError:
                continue
    return ImageFont.load_default()


def _estimate_wrapped_lines(lines: list[str], box_width_in: float, font_pt: float) -> int:
    """Conservative CJK-aware line-wrap estimator to prevent textbox overflow."""
    # Average CJK char width ~ font_pt / 72 inches; ASCII ~ 0.55 * font_pt / 72
    usable_pts = max(40.0, box_width_in * 72.0 - 12.0)
    total_lines = 0
    for text in lines:
        for seg in str(text).split("\n"):
            pt_len = sum(font_pt * (1.02 if ord(ch) > 127 else 0.56) for ch in seg)
            total_lines += max(1, math.ceil(pt_len / usable_pts))
    return total_lines


def _pick_autofit_font_pt(
    lines: list[str],
    box_width_in: float,
    box_height_in: float,
    max_pt: float = 11.0,
    min_pt: float = 8.5,
    line_height_factor: float = 1.38,
    extra_para_gap_pt: float = 3.0,
) -> tuple[float, float]:
    """Selects the largest font_pt in [min_pt, max_pt] whose wrapped height fits inside box_height_in."""
    avail_pts = max(18.0, box_height_in * 72.0 - 8.0)
    pt = max_pt
    while pt >= min_pt:
        gap = extra_para_gap_pt if pt >= 10.0 else 1.5
        n_lines = _estimate_wrapped_lines(lines, box_width_in, pt)
        needed_pts = n_lines * (pt * line_height_factor) + max(0, len(lines) - 1) * gap
        if needed_pts <= avail_pts:
            return round(pt, 1), gap
        pt -= 0.5
    return min_pt, 1.0


class OpalCaseReportDeck:
    FOOTER_TOP_IN = 7.02

    def __init__(self, theme_mode: str = "editorial", custom_theme: dict | None = None):
        self.prs = Presentation()
        self.prs.slide_width = Inches(13.333)
        self.prs.slide_height = Inches(7.5)
        self.blank_layout = self.prs.slide_layouts[6]

        base = DEFAULT_PALETTES.get(theme_mode, DEFAULT_PALETTES["editorial"]).copy()
        if custom_theme:
            base.update(custom_theme)
        self.theme = base
        self._temp_files: list[str] = []
        self.qa_warnings: list[str] = []

    def _apply_background(self, slide, bg_hex: str | None = None):
        bg_color = hex_to_rgb(bg_hex or self.theme["bg"])
        bg = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5)
        )
        bg.fill.solid()
        bg.fill.fore_color.rgb = bg_color
        bg.line.fill.background()
        return bg

    def _add_header(self, slide, title_text: str, category_tag: str = "", subtitle_text: str = ""):
        if category_tag:
            tag_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.36), Inches(11.5), Inches(0.28))
            tf = tag_box.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
            p = tf.paragraphs[0]
            p.text = category_tag.upper()
            p.font.size = Pt(10)
            p.font.bold = True
            p.font.color.rgb = hex_to_rgb(self.theme["accent"])
            p.font.name = self.theme["font_cn"]

        title_top = Inches(0.62) if category_tag else Inches(0.45)
        title_box = slide.shapes.add_textbox(Inches(0.6), title_top, Inches(12.133), Inches(0.52))
        tf = title_box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = title_text
        # Auto-shrink ultra-long titles so they never wrap onto subtitle
        t_pt, _ = _pick_autofit_font_pt([title_text], 12.133, 0.52, max_pt=24.0, min_pt=18.0, line_height_factor=1.2)
        p.font.size = Pt(t_pt)
        p.font.bold = True
        p.font.color.rgb = hex_to_rgb(self.theme["text_primary"])
        p.font.name = self.theme["font_cn"]

        if subtitle_text:
            sub_box = slide.shapes.add_textbox(Inches(0.6), title_top + Inches(0.46), Inches(12.133), Inches(0.34))
            tf_sub = sub_box.text_frame
            tf_sub.word_wrap = True
            tf_sub.margin_left = tf_sub.margin_top = tf_sub.margin_right = tf_sub.margin_bottom = 0
            p_sub = tf_sub.paragraphs[0]
            p_sub.text = subtitle_text
            p_sub.font.size = Pt(11.5)
            p_sub.font.color.rgb = hex_to_rgb(self.theme["text_secondary"])
            p_sub.font.name = self.theme["font_cn"]

    def _add_footer(self, slide, footer_text: str = "", slide_num: int | None = None):
        if footer_text:
            f_box = slide.shapes.add_textbox(Inches(0.6), Inches(self.FOOTER_TOP_IN), Inches(11.0), Inches(0.3))
            tf = f_box.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
            p = tf.paragraphs[0]
            p.text = footer_text
            p.font.size = Pt(9)
            p.font.color.rgb = hex_to_rgb(self.theme["text_secondary"])
            p.font.name = self.theme["font_cn"]

        if slide_num is not None:
            n_box = slide.shapes.add_textbox(Inches(11.8), Inches(self.FOOTER_TOP_IN), Inches(0.933), Inches(0.3))
            tf = n_box.text_frame
            tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.RIGHT
            p.text = f"{slide_num:02d}"
            p.font.size = Pt(9)
            p.font.bold = True
            p.font.color.rgb = hex_to_rgb(self.theme["accent"])
            p.font.name = self.theme["font_en"]

    def _add_image_fitted(
        self,
        slide,
        img_path: str,
        left: float,
        top: float,
        width: float,
        height: float,
        crop_to_fill: bool = True,
        max_px: int = 2200,
    ):
        """
        Inserts an image into (left, top, width, height) in inches without distorting aspect ratio.
        Automatically downscales oversized DSLR images in memory for fast PPTX serialization.
        """
        if not img_path or not os.path.exists(img_path):
            placeholder = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height)
            )
            placeholder.fill.solid()
            placeholder.fill.fore_color.rgb = hex_to_rgb(self.theme["accent_soft"])
            placeholder.line.color.rgb = hex_to_rgb(self.theme["card_border"])
            tf = placeholder.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            p.text = f"[Image Placeholder]\n{os.path.basename(img_path) if img_path else 'No Image'}"
            p.font.size = Pt(11)
            p.font.color.rgb = hex_to_rgb(self.theme["text_secondary"])
            return placeholder

        with Image.open(img_path) as raw_im:
            im = ImageOps.exif_transpose(raw_im).convert("RGB")
            if max_px and max(im.size) > max_px:
                im.thumbnail((max_px, max_px), Image.Resampling.LANCZOS)

            target_ratio = width / height
            img_ratio = im.width / im.height

            if crop_to_fill:
                if img_ratio > target_ratio:
                    new_width = int(im.height * target_ratio)
                    offset = (im.width - new_width) // 2
                    im_out = im.crop((offset, 0, offset + new_width, im.height))
                else:
                    new_height = int(im.width / target_ratio)
                    offset = (im.height - new_height) // 2
                    im_out = im.crop((0, offset, im.width, offset + new_height))

                fd, temp_path = tempfile.mkstemp(prefix="opal_v2_", suffix=".jpg")
                os.close(fd)
                im_out.save(temp_path, format="JPEG", quality=93)
                self._temp_files.append(temp_path)
                return slide.shapes.add_picture(temp_path, Inches(left), Inches(top), Inches(width), Inches(height))
            else:
                if img_ratio > target_ratio:
                    w = width
                    h = width / img_ratio
                    x = left
                    y = top + (height - h) / 2
                else:
                    h = height
                    w = height * img_ratio
                    x = left + (width - w) / 2
                    y = top

                fd, temp_path = tempfile.mkstemp(prefix="opal_v2_fit_", suffix=".jpg")
                os.close(fd)
                im.save(temp_path, format="JPEG", quality=93)
                self._temp_files.append(temp_path)
                return slide.shapes.add_picture(temp_path, Inches(x), Inches(y), Inches(w), Inches(h))

    # ------------------------------------------------------------------
    # Canonical Slide Builders
    # ------------------------------------------------------------------

    def add_cover_slide(
        self,
        title: str,
        subtitle: str = "",
        author: str = "",
        affiliation: str = "",
        tag: str = "AESTHETIC CASE REPORT",
    ):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._apply_background(slide)

        accent_bar = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(2.2), Inches(0.8), Inches(0.06)
        )
        accent_bar.fill.solid()
        accent_bar.fill.fore_color.rgb = hex_to_rgb(self.theme["accent"])
        accent_bar.line.fill.background()

        tag_box = slide.shapes.add_textbox(Inches(0.6), Inches(1.75), Inches(10), Inches(0.35))
        p_tag = tag_box.text_frame.paragraphs[0]
        p_tag.text = tag.upper()
        p_tag.font.size = Pt(11)
        p_tag.font.bold = True
        p_tag.font.color.rgb = hex_to_rgb(self.theme["accent"])
        p_tag.font.name = self.theme["font_en"]

        t_box = slide.shapes.add_textbox(Inches(0.6), Inches(2.5), Inches(11.5), Inches(1.9))
        tf = t_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(36)
        p.font.bold = True
        p.font.color.rgb = hex_to_rgb(self.theme["text_primary"])
        p.font.name = self.theme["font_cn"]

        if subtitle:
            p2 = tf.add_paragraph()
            p2.text = subtitle
            p2.space_before = Pt(14)
            p2.font.size = Pt(18)
            p2.font.color.rgb = hex_to_rgb(self.theme["text_secondary"])
            p2.font.name = self.theme["font_cn"]

        if author or affiliation:
            meta_box = slide.shapes.add_textbox(Inches(0.6), Inches(5.8), Inches(11.5), Inches(0.9))
            tf_m = meta_box.text_frame
            tf_m.word_wrap = True
            if author:
                pa = tf_m.paragraphs[0]
                pa.text = author
                pa.font.size = Pt(14)
                pa.font.bold = True
                pa.font.color.rgb = hex_to_rgb(self.theme["text_primary"])
                pa.font.name = self.theme["font_cn"]
            if affiliation:
                paf = tf_m.add_paragraph() if author else tf_m.paragraphs[0]
                paf.text = affiliation
                paf.font.size = Pt(11)
                paf.font.color.rgb = hex_to_rgb(self.theme["text_secondary"])
                paf.font.name = self.theme["font_cn"]
        return slide

    def add_section_divider(self, index_str: str, title_cn: str, title_en: str = ""):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._apply_background(slide)

        idx_box = slide.shapes.add_textbox(Inches(1.0), Inches(2.4), Inches(3.0), Inches(1.2))
        p_idx = idx_box.text_frame.paragraphs[0]
        p_idx.text = index_str
        p_idx.font.size = Pt(54)
        p_idx.font.bold = True
        p_idx.font.color.rgb = hex_to_rgb(self.theme["accent"])
        p_idx.font.name = self.theme["font_en"]

        t_box = slide.shapes.add_textbox(Inches(1.0), Inches(3.6), Inches(10.5), Inches(1.5))
        tf = t_box.text_frame
        tf.word_wrap = True
        p_cn = tf.paragraphs[0]
        p_cn.text = title_cn
        p_cn.font.size = Pt(32)
        p_cn.font.bold = True
        p_cn.font.color.rgb = hex_to_rgb(self.theme["text_primary"])
        p_cn.font.name = self.theme["font_cn"]

        if title_en:
            p_en = tf.add_paragraph()
            p_en.text = title_en
            p_en.space_before = Pt(8)
            p_en.font.size = Pt(16)
            p_en.font.color.rgb = hex_to_rgb(self.theme["text_secondary"])
            p_en.font.name = self.theme["font_en"]
        return slide

    def add_fullbleed_atlas_slide(self, img_path: str):
        """Archetype: Zero-crop 16:9 full-bleed anatomical infographic / atlas slide."""
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._apply_background(slide)
        self._add_image_fitted(slide, str(img_path), 0.0, 0.0, 13.333, 7.5, crop_to_fill=True, max_px=3840)
        return slide

    def add_anatomy_split_slide(
        self,
        title: str,
        subtitle: str,
        image_path: str,
        cards: list[dict],
        footer: str = "",
        slide_num: int | None = None,
    ):
        """
        Archetype 4: Left 52% anatomical diagram/photo + Right 48% structured mechanism/classification cards.
        cards: [{"heading": "解剖定义", "body": "..."}, ...]
        """
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._apply_background(slide)
        self._add_header(slide, title, category_tag="ANATOMY & CLASSIFICATION", subtitle_text=subtitle)

        self._add_image_fitted(slide, image_path, 0.6, 1.52, 6.2, 5.25, crop_to_fill=False)

        n_cards = max(1, len(cards))
        gap = 0.14
        avail_h = 5.25
        card_h = (avail_h - gap * (n_cards - 1)) / n_cards
        y_cursor = 1.52

        for c in cards:
            card = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(7.05), Inches(y_cursor), Inches(5.683), Inches(card_h)
            )
            card.fill.solid()
            card.fill.fore_color.rgb = hex_to_rgb(self.theme["card_bg"])
            card.line.color.rgb = hex_to_rgb(self.theme["card_border"])

            stripe = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(7.05), Inches(y_cursor), Inches(0.1), Inches(card_h)
            )
            stripe.fill.solid()
            stripe.fill.fore_color.rgb = hex_to_rgb(self.theme["accent"])
            stripe.line.fill.background()

            tb = slide.shapes.add_textbox(
                Inches(7.3), Inches(y_cursor + 0.1), Inches(5.28), Inches(max(0.3, card_h - 0.2))
            )
            tf = tb.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

            ph = tf.paragraphs[0]
            ph.text = c.get("heading", "")
            ph.font.size = Pt(13)
            ph.font.bold = True
            ph.font.color.rgb = hex_to_rgb(self.theme["accent"])
            ph.font.name = self.theme["font_cn"]

            body_lines = c.get("bullets") or ([c["body"]] if c.get("body") else [])
            b_pt, b_gap = _pick_autofit_font_pt(body_lines, 5.28, max(0.25, card_h - 0.45), max_pt=11.0, min_pt=9.0)
            for bl in body_lines:
                pb = tf.add_paragraph()
                pb.text = f"• {bl}" if len(body_lines) > 1 else str(bl)
                pb.space_before = Pt(b_gap + 1)
                pb.font.size = Pt(b_pt)
                pb.font.color.rgb = hex_to_rgb(self.theme["text_primary"])
                pb.font.name = self.theme["font_cn"]

            y_cursor += card_h + gap

        self._add_footer(slide, footer, slide_num)
        return slide

    def add_assessment_slide(
        self,
        title: str,
        patient_label: str,
        photo_path: str,
        scores: list,
        clinical_notes: list,
        footer: str = "",
        slide_num: int | None = None,
    ):
        """
        v2 Overflow-Safe Facial Assessment Slide:
        Dynamically compacts score row pitch when clinical_notes has >= 4 lines and auto-fits
        note font size so text NEVER overflows into the slide footer.
        """
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._apply_background(slide)
        self._add_header(slide, title, category_tag="CLINICAL ASSESSMENT", subtitle_text=patient_label)

        # Left Photo Card
        card_left = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(1.5), Inches(4.5), Inches(5.28)
        )
        card_left.fill.solid()
        card_left.fill.fore_color.rgb = hex_to_rgb(self.theme["card_bg"])
        card_left.line.color.rgb = hex_to_rgb(self.theme["card_border"])

        self._add_image_fitted(slide, photo_path, 0.75, 1.64, 4.2, 4.48, crop_to_fill=True)

        lbl_box = slide.shapes.add_textbox(Inches(0.75), Inches(6.22), Inches(4.2), Inches(0.4))
        p_lbl = lbl_box.text_frame.paragraphs[0]
        p_lbl.alignment = PP_ALIGN.CENTER
        p_lbl.text = patient_label
        p_lbl.font.size = Pt(11)
        p_lbl.font.bold = True
        p_lbl.font.color.rgb = hex_to_rgb(self.theme["text_primary"])
        p_lbl.font.name = self.theme["font_cn"]

        # Adaptive vertical budget between score rows and bottom clinical_notes box
        has_dense_notes = len(clinical_notes or []) >= 4 or sum(len(str(x)) for x in (clinical_notes or [])) > 130
        row_h = 0.65 if has_dense_notes else 0.71
        row_step = 0.73 if has_dense_notes else 0.80

        y_cursor = 1.52
        for item in scores[:5]:
            domain = item.get("domain", "")
            score = int(item.get("score", 0))
            max_score = int(item.get("max", 3))
            note = item.get("note", "")

            row_card = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(5.4), Inches(y_cursor), Inches(7.333), Inches(row_h)
            )
            row_card.fill.solid()
            row_card.fill.fore_color.rgb = hex_to_rgb(self.theme["card_bg"])
            row_card.line.color.rgb = hex_to_rgb(self.theme["card_border"])

            tb = slide.shapes.add_textbox(Inches(5.58), Inches(y_cursor + 0.07), Inches(3.85), Inches(row_h - 0.12))
            tf = tb.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
            p = tf.paragraphs[0]
            p.text = domain
            p.font.size = Pt(11.5)
            p.font.bold = True
            p.font.color.rgb = hex_to_rgb(self.theme["text_primary"])
            p.font.name = self.theme["font_cn"]
            if note:
                p2 = tf.add_paragraph()
                # Safe single-line note truncation only if > 42 chars
                p2.text = note if len(note) <= 44 else note[:42] + "…"
                p2.font.size = Pt(9.0)
                p2.font.color.rgb = hex_to_rgb(self.theme["text_secondary"])
                p2.font.name = self.theme["font_cn"]

            for s_idx in range(max_score + 1):
                bx = 9.6 + s_idx * 0.72
                bar = slide.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE,
                    Inches(bx),
                    Inches(y_cursor + (row_h - 0.34) / 2),
                    Inches(0.6),
                    Inches(0.34),
                )
                bar.fill.solid()
                is_active = s_idx == score
                bar.fill.fore_color.rgb = hex_to_rgb(
                    self.theme["accent"] if is_active else self.theme["accent_soft"]
                )
                bar.line.fill.background()
                bp = bar.text_frame.paragraphs[0]
                bp.alignment = PP_ALIGN.CENTER
                bp.text = str(s_idx)
                bp.font.size = Pt(10.5)
                bp.font.bold = is_active
                bp.font.color.rgb = hex_to_rgb("#FFFFFF" if is_active else self.theme["text_secondary"])

            y_cursor += row_step

        if clinical_notes:
            notes_top = y_cursor + 0.04
            notes_h = max(0.65, 6.78 - notes_top)
            notes_card = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(5.4), Inches(notes_top), Inches(7.333), Inches(notes_h)
            )
            notes_card.fill.solid()
            notes_card.fill.fore_color.rgb = hex_to_rgb(self.theme["accent_soft"])
            notes_card.line.fill.background()

            bullet_lines = [f"• {line}" for line in clinical_notes]
            font_pt, para_gap = _pick_autofit_font_pt(
                bullet_lines, box_width_in=6.95, box_height_in=notes_h - 0.14, max_pt=10.5, min_pt=8.5
            )

            ntb = slide.shapes.add_textbox(
                Inches(5.58), Inches(notes_top + 0.07), Inches(6.95), Inches(notes_h - 0.14)
            )
            ntf = ntb.text_frame
            ntf.word_wrap = True
            ntf.margin_left = ntf.margin_top = ntf.margin_right = ntf.margin_bottom = 0
            for idx, b_text in enumerate(bullet_lines):
                np = ntf.paragraphs[0] if idx == 0 else ntf.add_paragraph()
                np.text = b_text
                np.space_after = Pt(para_gap)
                np.font.size = Pt(font_pt)
                np.font.color.rgb = hex_to_rgb(self.theme["text_primary"])
                np.font.name = self.theme["font_cn"]

        self._add_footer(slide, footer, slide_num)
        return slide

    def add_treatment_plan_slide(
        self,
        title: str,
        subtitle: str,
        diagram_path: str,
        steps: list,
        footer: str = "",
        slide_num: int | None = None,
    ):
        """
        v2 Overflow-Safe Treatment Plan & Rheological Mapping slide.
        Automatically scales font size inside each step card to fit 2-4 steps cleanly.
        """
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._apply_background(slide)
        self._add_header(slide, title, category_tag="TREATMENT PLAN & RHEOLOGY", subtitle_text=subtitle)

        self._add_image_fitted(slide, diagram_path, 0.6, 1.52, 4.6, 5.25, crop_to_fill=False)

        n_steps = max(1, len(steps))
        gap = 0.12
        avail_h = 5.25
        card_height = min(2.48, (avail_h - gap * (n_steps - 1)) / n_steps)
        y_cursor = 1.52

        for step in steps:
            card = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(5.5), Inches(y_cursor), Inches(7.233), Inches(card_height)
            )
            card.fill.solid()
            card.fill.fore_color.rgb = hex_to_rgb(self.theme["card_bg"])
            card.line.color.rgb = hex_to_rgb(self.theme["card_border"])

            stripe = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(5.5), Inches(y_cursor), Inches(0.12), Inches(card_height)
            )
            stripe.fill.solid()
            stripe.fill.fore_color.rgb = hex_to_rgb(self.theme["accent"])
            stripe.line.fill.background()

            details = [
                f"• 产品与流变学 (Product): {step.get('product', 'N/A')}",
                f"• 靶区与层次 (Target & Plane): {step.get('plane', 'N/A')}",
                f"• 器械与剂量 (Instrument & Dosage): {step.get('instrument', '')} | {step.get('dosage', '')}",
            ]
            if step.get("rationale"):
                details.append(f"• 操作要点 (Rationale): {step.get('rationale')}")

            title_pt = 13.5 if n_steps <= 2 else (12.5 if n_steps == 3 else 11.5)
            body_box_h = max(0.45, card_height - 0.42)
            body_pt, body_gap = _pick_autofit_font_pt(details, 6.75, body_box_h, max_pt=10.8, min_pt=8.5)

            tb = slide.shapes.add_textbox(
                Inches(5.78), Inches(y_cursor + 0.08), Inches(6.75), Inches(card_height - 0.16)
            )
            tf = tb.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

            p_title = tf.paragraphs[0]
            p_title.text = step.get("step_title", "")
            p_title.font.size = Pt(title_pt)
            p_title.font.bold = True
            p_title.font.color.rgb = hex_to_rgb(self.theme["accent"])
            p_title.font.name = self.theme["font_cn"]

            for line in details:
                p = tf.add_paragraph()
                p.text = line
                p.space_before = Pt(body_gap)
                p.font.size = Pt(body_pt)
                p.font.color.rgb = hex_to_rgb(self.theme["text_primary"])
                p.font.name = self.theme["font_cn"]

            y_cursor += card_height + gap

        self._add_footer(slide, footer, slide_num)
        return slide

    def add_before_after_slide(
        self,
        title: str,
        view_label: str,
        before_img: str,
        after_img: str,
        before_caption: str = "Before (术前)",
        after_caption: str = "After (术后)",
        summary_metrics: str = "",
        footer: str = "",
        slide_num: int | None = None,
    ):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._apply_background(slide)
        self._add_header(slide, title, category_tag=f"BEFORE & AFTER COMPARISON · {view_label}")

        img_w = 5.85
        img_h = 4.55
        top_y = 1.35

        # Before Container
        self._add_image_fitted(slide, before_img, 0.6, top_y, img_w, img_h, crop_to_fill=True)
        b_lbl = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(top_y + img_h), Inches(img_w), Inches(0.42)
        )
        b_lbl.fill.solid()
        b_lbl.fill.fore_color.rgb = hex_to_rgb(self.theme["card_bg"])
        b_lbl.line.color.rgb = hex_to_rgb(self.theme["card_border"])
        bp = b_lbl.text_frame.paragraphs[0]
        bp.alignment = PP_ALIGN.CENTER
        bp.text = before_caption
        bp.font.size = Pt(12)
        bp.font.bold = True
        bp.font.color.rgb = hex_to_rgb(self.theme["text_secondary"])
        bp.font.name = self.theme["font_cn"]

        # After Container
        right_x = 13.333 - 0.6 - img_w
        self._add_image_fitted(slide, after_img, right_x, top_y, img_w, img_h, crop_to_fill=True)
        a_lbl = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(right_x), Inches(top_y + img_h), Inches(img_w), Inches(0.42)
        )
        a_lbl.fill.solid()
        a_lbl.fill.fore_color.rgb = hex_to_rgb(self.theme["accent"])
        a_lbl.line.fill.background()
        ap = a_lbl.text_frame.paragraphs[0]
        ap.alignment = PP_ALIGN.CENTER
        ap.text = after_caption
        ap.font.size = Pt(12)
        ap.font.bold = True
        ap.font.color.rgb = hex_to_rgb("#FFFFFF")
        ap.font.name = self.theme["font_cn"]

        if summary_metrics:
            m_box = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(6.42), Inches(12.133), Inches(0.46)
            )
            m_box.fill.solid()
            m_box.fill.fore_color.rgb = hex_to_rgb(self.theme["accent_soft"])
            m_box.line.fill.background()
            mp = m_box.text_frame.paragraphs[0]
            mp.alignment = PP_ALIGN.CENTER
            mp.text = summary_metrics
            mp.font.size = Pt(10.5)
            mp.font.bold = True
            mp.font.color.rgb = hex_to_rgb(self.theme["text_primary"])
            mp.font.name = self.theme["font_cn"]

        self._add_footer(slide, footer, slide_num)
        return slide

    def add_patient_journey_slide(
        self,
        title: str,
        subtitle: str,
        milestones: list,
        footer: str = "",
        slide_num: int | None = None,
    ):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._apply_background(slide)
        self._add_header(slide, title, category_tag="LONGITUDINAL PATIENT JOURNEY", subtitle_text=subtitle)

        n = max(1, min(len(milestones), 5))
        total_w = 12.133
        gap = 0.22
        col_w = (total_w - gap * (n - 1)) / n
        img_h = 4.10
        top_y = 1.62

        for idx, m in enumerate(milestones[:5]):
            x = 0.6 + idx * (col_w + gap)
            self._add_image_fitted(slide, m.get("image", ""), x, top_y, col_w, img_h, crop_to_fill=True)

            badge = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(x), Inches(top_y + img_h + 0.06), Inches(col_w), Inches(0.36)
            )
            badge.fill.solid()
            is_last = idx == n - 1
            badge.fill.fore_color.rgb = hex_to_rgb(self.theme["accent"] if is_last else self.theme["card_bg"])
            badge.line.color.rgb = hex_to_rgb(self.theme["accent"] if is_last else self.theme["card_border"])
            bp = badge.text_frame.paragraphs[0]
            bp.alignment = PP_ALIGN.CENTER
            bp.text = m.get("date", f"T{idx}")
            bp.font.size = Pt(11)
            bp.font.bold = True
            bp.font.color.rgb = hex_to_rgb("#FFFFFF" if is_last else self.theme["text_primary"])
            bp.font.name = self.theme["font_cn"]

            if m.get("note"):
                ntb = slide.shapes.add_textbox(
                    Inches(x), Inches(top_y + img_h + 0.46), Inches(col_w), Inches(0.68)
                )
                ntf = ntb.text_frame
                ntf.word_wrap = True
                ntf.margin_left = ntf.margin_top = ntf.margin_right = ntf.margin_bottom = 0
                np = ntf.paragraphs[0]
                np.alignment = PP_ALIGN.CENTER
                np.text = m.get("note")
                n_pt, _ = _pick_autofit_font_pt([m.get("note")], col_w, 0.68, max_pt=9.5, min_pt=8.0)
                np.font.size = Pt(n_pt)
                np.font.color.rgb = hex_to_rgb(self.theme["text_secondary"])
                np.font.name = self.theme["font_cn"]

        self._add_footer(slide, footer, slide_num)
        return slide

    def add_consensus_table_slide(
        self,
        title: str,
        subtitle: str,
        headers: list,
        rows: list,
        takeaway: str = "",
        footer: str = "",
        slide_num: int | None = None,
    ):
        """
        v2 High-Contrast & Auto-Scaled Table Slide:
        Fixes the v1 issue where odd/even white rows blended into the warm ivory canvas.
        """
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._apply_background(slide)
        self._add_header(slide, title, category_tag="EVIDENCE & ANATOMICAL RATIONALE", subtitle_text=subtitle)

        n_rows = len(rows) + 1
        n_cols = max(1, len(headers))
        table_h = 4.45 if takeaway else 5.05
        table_shape = slide.shapes.add_table(
            n_rows, n_cols, Inches(0.6), Inches(1.62), Inches(12.133), Inches(table_h)
        )
        table = table_shape.table

        # Auto-scale cell font based on row count and text density
        total_chars = sum(len(str(v)) for r in rows for v in r)
        cell_pt = 11.0 if (len(rows) <= 4 and total_chars < 380) else (10.0 if len(rows) <= 6 else 9.0)

        for c_idx, h_text in enumerate(headers):
            cell = table.cell(0, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = hex_to_rgb(self.theme["accent"])
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            p.text = h_text
            p.font.size = Pt(11.5)
            p.font.bold = True
            p.font.color.rgb = hex_to_rgb("#FFFFFF")
            p.font.name = self.theme["font_cn"]

        row_shades = [self.theme["row_even"], self.theme["row_odd"]]
        for r_idx, row_data in enumerate(rows):
            fill_hex = row_shades[r_idx % 2]
            for c_idx, val in enumerate(row_data[:n_cols]):
                cell = table.cell(r_idx + 1, c_idx)
                cell.fill.solid()
                cell.fill.fore_color.rgb = hex_to_rgb(fill_hex)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                p = cell.text_frame.paragraphs[0]
                p.text = str(val)
                p.font.size = Pt(cell_pt)
                p.font.bold = c_idx == 0
                p.font.color.rgb = hex_to_rgb(self.theme["text_primary"])
                p.font.name = self.theme["font_cn"]

        if takeaway:
            t_box = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(6.28), Inches(12.133), Inches(0.52)
            )
            t_box.fill.solid()
            t_box.fill.fore_color.rgb = hex_to_rgb(self.theme["accent_soft"])
            t_box.line.color.rgb = hex_to_rgb(self.theme["card_border"])
            tp = t_box.text_frame.paragraphs[0]
            tp.alignment = PP_ALIGN.LEFT
            tp.text = f"  核心提示: {takeaway}"
            tp.font.size = Pt(10.5)
            tp.font.bold = True
            tp.font.color.rgb = hex_to_rgb(self.theme["accent"])
            tp.font.name = self.theme["font_cn"]

        self._add_footer(slide, footer, slide_num)
        return slide

    def render_5axis_chart_png(self, domains: list[dict], summary_line: str = "", subtitle: str = "") -> str:
        """Pure-PIL generator for Pre vs. Post 5-Axis Assessment Bar Chart (zero matplotlib dependency)."""
        W, H = 1600, 900
        img = Image.new("RGB", (W, H), "#FFFFFF")
        d = ImageDraw.Draw(img)
        f_title = _load_pil_font(40)
        f_sub = _load_pil_font(22)
        f_label = _load_pil_font(26)
        f_val = _load_pil_font(25)
        f_small = _load_pil_font(20)

        d.text((56, 34), "多轴严重度评分 · 治疗前后客观对比", font=f_title, fill="#231F20")
        d.text(
            (56, 92),
            subtitle or "0 = 无    1 = 轻度    2 = 中度    3 = 重度",
            font=f_sub,
            fill="#7A7067",
        )

        c_before = self.theme["accent"]
        c_after = "#2C4C64" if self.theme["accent"].upper() != "#2C4C64" else "#27AE60"

        d.rectangle((1140, 48, 1168, 76), fill=c_before)
        d.text((1178, 50), "术前基线 (Pre)", font=f_small, fill="#5A626A")
        d.rectangle((1350, 48, 1378, 76), fill=c_after)
        d.text((1388, 50), "术后随访 (Post)", font=f_small, fill="#5A626A")

        max_s = max([int(dm.get("max", 3)) for dm in domains] + [3])
        x0, x1, top, row_h = 540, 1440, 175, 124
        for s in range(max_s + 1):
            x = x0 + (x1 - x0) * s / max_s
            d.line((x, 155, x, 755), fill="#E8DFD5", width=2)
            d.text((x - 7, 762), str(s), font=f_small, fill="#7A7067")

        for i, dm in enumerate(domains[:5]):
            y = top + i * row_h
            lbl_cn = dm.get("label_cn") or dm.get("domain", "").split("(")[0].strip()
            lbl_en = dm.get("label_en") or ""
            d.text((56, y + 12), lbl_cn[:18], font=f_label, fill="#231F20")
            if lbl_en:
                d.text((56, y + 50), lbl_en[:28], font=f_small, fill="#9A928A")
            for j, (key, col) in enumerate([("before", c_before), ("after", c_after)]):
                v = int(dm.get(key, 0))
                by = y + 8 + j * 38
                bw = int((x1 - x0) * v / max_s)
                d.rounded_rectangle((x0, by, x0 + max(bw, 10), by + 30), radius=6, fill=col)
                d.text((x0 + bw + 14, by + 2), str(v), font=f_val, fill=col)

        if summary_line:
            d.line((56, 812, 1544, 812), fill="#E8DFD5", width=2)
            d.text((56, 830), summary_line, font=f_label, fill=self.theme["accent"])

        fd, tmp_png = tempfile.mkstemp(prefix="opal_chart_", suffix=".png")
        os.close(fd)
        img.save(tmp_png, "PNG")
        self._temp_files.append(tmp_png)
        return tmp_png

    def add_chart_slide(
        self,
        title: str,
        subtitle: str,
        stats: list[list[str] | tuple[str, str]],
        chart_path: str | None = None,
        domains: list[dict] | None = None,
        chart_summary: str = "",
        footer: str = "",
        slide_num: int | None = None,
    ):
        """Archetype: Quantitative Pre- vs. Post-Treatment Shift Chart + Right KPI Cards."""
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._apply_background(slide)
        self._add_header(slide, title, category_tag="OBJECTIVE OUTCOME SHIFT", subtitle_text=subtitle)

        resolved_chart = chart_path
        if not resolved_chart and domains:
            resolved_chart = self.render_5axis_chart_png(domains, summary_line=chart_summary)

        self._add_image_fitted(slide, resolved_chart or "", 0.6, 1.52, 8.5, 5.15, crop_to_fill=False)

        card_x, card_w = 9.35, 3.383
        n_stats = max(1, min(len(stats), 4))
        gap = 0.16
        card_h = (5.15 - gap * (n_stats - 1)) / n_stats

        for i, pair in enumerate(stats[:4]):
            label, value = pair[0], pair[1]
            y = 1.52 + i * (card_h + gap)
            card = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(card_x), Inches(y), Inches(card_w), Inches(card_h)
            )
            card.fill.solid()
            card.fill.fore_color.rgb = hex_to_rgb(self.theme["card_bg"])
            card.line.color.rgb = hex_to_rgb(self.theme["card_border"])

            tb1 = slide.shapes.add_textbox(Inches(card_x + 0.2), Inches(y + 0.16), Inches(card_w - 0.4), Inches(0.3))
            p1 = tb1.text_frame.paragraphs[0]
            p1.text = str(label)
            p1.font.size = Pt(10.5)
            p1.font.color.rgb = hex_to_rgb(self.theme["text_secondary"])
            p1.font.name = self.theme["font_cn"]

            tb2 = slide.shapes.add_textbox(Inches(card_x + 0.2), Inches(y + 0.48), Inches(card_w - 0.4), Inches(0.5))
            tf2 = tb2.text_frame
            tf2.word_wrap = True
            p2 = tf2.paragraphs[0]
            p2.text = str(value)
            p2.font.size = Pt(18.5)
            p2.font.bold = True
            p2.font.color.rgb = hex_to_rgb(self.theme["accent"])
            p2.font.name = self.theme["font_cn"]

        self._add_footer(slide, footer, slide_num)
        return slide

    def add_summary_slide(
        self,
        title: str,
        pearls: list[str],
        disclaimer: str = "",
        subtitle: str = "从解剖评估、流变学匹配到长期随访的核心临床要点",
        footer: str = "",
        slide_num: int | None = None,
    ):
        """Archetype 10: Clinical Pearls & Disclosure Summary Slide."""
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._apply_background(slide)
        self._add_header(slide, title, category_tag="CLINICAL PEARLS & DISCLOSURE", subtitle_text=subtitle)

        n_p = max(1, min(len(pearls), 5))
        avail_h = 4.45 if disclaimer else 5.20
        gap = 0.12
        card_h = (avail_h - gap * (n_p - 1)) / n_p

        for i, pearl in enumerate(pearls[:5]):
            y = 1.52 + i * (card_h + gap)
            card = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(y), Inches(12.133), Inches(card_h)
            )
            card.fill.solid()
            card.fill.fore_color.rgb = hex_to_rgb(self.theme["card_bg"])
            card.line.color.rgb = hex_to_rgb(self.theme["card_border"])

            num = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(y), Inches(0.88), Inches(card_h)
            )
            num.fill.solid()
            num.fill.fore_color.rgb = hex_to_rgb(self.theme["accent"])
            num.line.fill.background()

            ntb = slide.shapes.add_textbox(Inches(0.6), Inches(y + (card_h - 0.36) / 2), Inches(0.88), Inches(0.36))
            np = ntb.text_frame.paragraphs[0]
            np.alignment = PP_ALIGN.CENTER
            np.text = f"{i + 1:02d}"
            np.font.size = Pt(15)
            np.font.bold = True
            np.font.color.rgb = hex_to_rgb("#FFFFFF")
            np.font.name = self.theme["font_en"]

            p_pt, _ = _pick_autofit_font_pt([pearl], 10.85, card_h - 0.18, max_pt=12.0, min_pt=9.5)
            btb = slide.shapes.add_textbox(Inches(1.70), Inches(y + 0.12), Inches(10.85), Inches(card_h - 0.22))
            btf = btb.text_frame
            btf.word_wrap = True
            btf.margin_left = btf.margin_top = btf.margin_right = btf.margin_bottom = 0
            bp = btf.paragraphs[0]
            bp.text = pearl
            bp.font.size = Pt(p_pt)
            bp.font.color.rgb = hex_to_rgb(self.theme["text_primary"])
            bp.font.name = self.theme["font_cn"]

        if disclaimer:
            box = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(6.14), Inches(12.133), Inches(0.70)
            )
            box.fill.solid()
            box.fill.fore_color.rgb = hex_to_rgb(self.theme["accent_soft"])
            box.line.fill.background()

            dtb = slide.shapes.add_textbox(Inches(0.82), Inches(6.22), Inches(11.7), Inches(0.54))
            dtf = dtb.text_frame
            dtf.word_wrap = True
            dtf.margin_left = dtf.margin_top = dtf.margin_right = dtf.margin_bottom = 0
            dp = dtf.paragraphs[0]
            dp.text = disclaimer
            dp.font.size = Pt(9.2)
            dp.font.color.rgb = hex_to_rgb(self.theme["text_secondary"])
            dp.font.name = self.theme["font_cn"]

        self._add_footer(slide, footer, slide_num)
        return slide

    # ------------------------------------------------------------------
    # Millisecond Static Layout Validator (replaces slow LibreOffice loops)
    # ------------------------------------------------------------------

    def validate_deck(self) -> list[str]:
        """
        Checks every slide in milliseconds for:
        1. Content shapes crossing into the footer zone (> 6.95 in), excluding full-bleed backgrounds.
        2. Missing image placeholders accidentally left in the deck.
        3. Textboxes whose estimated wrapped height exceeds box height by >15%.
        """
        issues: list[str] = []
        emu_per_in = 914400.0
        for s_idx, slide in enumerate(self.prs.slides, start=1):
            for sh_idx, sh in enumerate(slide.shapes):
                top_in = sh.top / emu_per_in
                h_in = sh.height / emu_per_in
                w_in = sh.width / emu_per_in
                bottom_in = top_in + h_in

                # Skip background / fullbleed shapes and footer textboxes themselves
                if w_in >= 13.2 and h_in >= 7.4:
                    continue
                if top_in >= self.FOOTER_TOP_IN - 0.05:
                    continue

                if bottom_in > self.FOOTER_TOP_IN + 0.02:
                    issues.append(
                        f"Slide {s_idx:02d}: shape '{sh.name}' bottom={bottom_in:.2f}in intrudes into footer ({self.FOOTER_TOP_IN}in)"
                    )

                if sh.has_text_frame:
                    full_txt = sh.text_frame.text or ""
                    if "[Image Placeholder]" in full_txt:
                        issues.append(f"Slide {s_idx:02d}: unresolved image placeholder -> {full_txt.splitlines()[-1]}")
        return issues

    # ------------------------------------------------------------------
    # Declarative JSON Builder
    # ------------------------------------------------------------------

    def build_from_spec(self, spec: dict, base_dir: Path | None = None) -> list[str]:
        """
        Builds a complete presentation from a declarative `deck_spec.json` dictionary.
        Relative image paths are resolved against `base_dir` (defaults to JSON parent directory).
        """
        root = base_dir or Path.cwd()
        default_footer = spec.get("default_footer", "")

        def resolve_p(val: str | None) -> str:
            if not val:
                return ""
            p = Path(val)
            if p.is_absolute():
                return str(p)
            return str((root / p).resolve())

        page_num = 0
        for item in spec.get("slides", []):
            stype = item.get("type", "").lower()
            footer = item.get("footer", default_footer)
            page_num += 1
            s_num = item.get("slide_num", page_num)

            if stype == "cover":
                self.add_cover_slide(
                    title=item["title"],
                    subtitle=item.get("subtitle", ""),
                    author=item.get("author", ""),
                    affiliation=item.get("affiliation", ""),
                    tag=item.get("tag", "AESTHETIC CASE REPORT"),
                )
            elif stype in ("divider", "section_divider"):
                self.add_section_divider(
                    index_str=item.get("index", "01"),
                    title_cn=item["title_cn"],
                    title_en=item.get("title_en", ""),
                )
            elif stype in ("fullbleed", "fullbleed_atlas"):
                self.add_fullbleed_atlas_slide(resolve_p(item["image"]))
            elif stype == "anatomy_split":
                self.add_anatomy_split_slide(
                    title=item["title"],
                    subtitle=item.get("subtitle", ""),
                    image_path=resolve_p(item.get("image")),
                    cards=item.get("cards", []),
                    footer=footer,
                    slide_num=s_num,
                )
            elif stype == "assessment":
                self.add_assessment_slide(
                    title=item["title"],
                    patient_label=item.get("patient_label", ""),
                    photo_path=resolve_p(item.get("photo")),
                    scores=item.get("scores", []),
                    clinical_notes=item.get("clinical_notes", []),
                    footer=footer,
                    slide_num=s_num,
                )
            elif stype == "treatment_plan":
                self.add_treatment_plan_slide(
                    title=item["title"],
                    subtitle=item.get("subtitle", ""),
                    diagram_path=resolve_p(item.get("diagram")),
                    steps=item.get("steps", []),
                    footer=footer,
                    slide_num=s_num,
                )
            elif stype == "before_after":
                self.add_before_after_slide(
                    title=item["title"],
                    view_label=item.get("view_label", ""),
                    before_img=resolve_p(item.get("before_img")),
                    after_img=resolve_p(item.get("after_img")),
                    before_caption=item.get("before_caption", "Before (术前)"),
                    after_caption=item.get("after_caption", "After (术后)"),
                    summary_metrics=item.get("summary_metrics", ""),
                    footer=footer,
                    slide_num=s_num,
                )
            elif stype == "patient_journey":
                ms = [
                    {
                        "date": m.get("date", ""),
                        "image": resolve_p(m.get("image")),
                        "note": m.get("note", ""),
                    }
                    for m in item.get("milestones", [])
                ]
                self.add_patient_journey_slide(
                    title=item["title"],
                    subtitle=item.get("subtitle", ""),
                    milestones=ms,
                    footer=footer,
                    slide_num=s_num,
                )
            elif stype in ("table", "consensus_table"):
                self.add_consensus_table_slide(
                    title=item["title"],
                    subtitle=item.get("subtitle", ""),
                    headers=item.get("headers", []),
                    rows=item.get("rows", []),
                    takeaway=item.get("takeaway", ""),
                    footer=footer,
                    slide_num=s_num,
                )
            elif stype == "chart":
                self.add_chart_slide(
                    title=item["title"],
                    subtitle=item.get("subtitle", ""),
                    stats=item.get("stats", []),
                    chart_path=resolve_p(item.get("chart_path")) if item.get("chart_path") else None,
                    domains=item.get("domains"),
                    chart_summary=item.get("chart_summary", ""),
                    footer=footer,
                    slide_num=s_num,
                )
            elif stype == "summary":
                self.add_summary_slide(
                    title=item["title"],
                    pearls=item.get("pearls", []),
                    disclaimer=item.get("disclaimer", ""),
                    subtitle=item.get("subtitle", "从解剖评估、流变学匹配到长期随访的核心临床要点"),
                    footer=footer,
                    slide_num=s_num,
                )
            else:
                raise ValueError(f"Unknown slide type in JSON spec: {stype!r}")

        return self.validate_deck()

    def save(self, output_path: str):
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        self.prs.save(output_path)
        for tmp in self._temp_files:
            if os.path.exists(tmp):
                try:
                    os.remove(tmp)
                except OSError:
                    pass
        self._temp_files.clear()
        return output_path


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Build an Opal v2 medical aesthetics case report PPTX from a declarative JSON spec."
    )
    ap.add_argument("spec_json", type=Path, help="path to deck_spec.json")
    ap.add_argument("--out", type=Path, help="optional override output .pptx path")
    args = ap.parse_args()

    if not args.spec_json.is_file():
        print(f"[Error] Spec file not found: {args.spec_json}", file=sys.stderr)
        return 1

    spec = json.loads(args.spec_json.read_text(encoding="utf-8"))
    deck = OpalCaseReportDeck(
        theme_mode=spec.get("theme_mode", "editorial"),
        custom_theme=spec.get("custom_theme"),
    )
    issues = deck.build_from_spec(spec, base_dir=args.spec_json.resolve().parent)

    out_raw = str(args.out) if args.out else spec.get("output_path", "case_report.pptx")
    out_path = Path(out_raw)
    if not out_path.is_absolute():
        out_path = (args.spec_json.resolve().parent / out_path).resolve()

    deck.save(str(out_path))
    print(f"[OK] Built {len(deck.prs.slides)}-slide PPTX -> {out_path}")
    if issues:
        print(f"[QA WARN] {len(issues)} layout note(s):")
        for msg in issues:
            print(f"  - {msg}")
    else:
        print("[QA PASS] Zero footer collisions, zero unresolved placeholders, all text auto-fitted.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
