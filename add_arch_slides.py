"""
Add architecture diagram slides to Cadence PII Shield stakeholder deck.
Inserts 4 new slides after slide 10 (or after the last slide).
"""

import os
import copy
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from lxml import etree
import shutil

# Paths
PPTX_IN  = "Cadence_PII_Shield_Stakeholder_Deck.pptx"
PPTX_OUT = "Cadence_PII_Shield_Stakeholder_Deck_with_Arch.pptx"
ARTIFACTS_DIR = r"C:\Users\adakr\.gemini\antigravity-ide\brain\ff07a976-0822-4596-9a45-f601c84bf7b3"

# Architecture images generated
IMAGES = [
    {
        "file": "pii_shield_pipeline_architecture_1791898829701.jpg",
        "title": "System Architecture: End-to-End PII Processing Pipeline",
        "subtitle": "100% Local • Fail-Closed • Multi-Layer Detection",
        "portrait": True,   # 9:16 image → portrait, we'll show it tall
    },
    {
        "file": "pii_shield_detection_layers_1791898955915.jpg",
        "title": "Detection Engine: 6-Layer PII Identification Stack",
        "subtitle": "Regex + NER + Context Boost + OCR Repair + Alias Propagation + Fail-Closed",
        "portrait": False,
    },
    {
        "file": "pii_shield_security_threat_model_1791898935226.jpg",
        "title": "Security Architecture: Threat Model & Trust Boundaries",
        "subtitle": "AES-128-CBC Encryption • HMAC-SHA256 • Zero Raw PII Exposure",
        "portrait": False,
    },
    {
        "file": "pii_shield_browser_extension_arch_1791898873300.jpg",
        "title": "Browser Extension Architecture: Real-Time Prompt Shielding",
        "subtitle": "Chrome / Edge / Brave • Local Bridge 127.0.0.1:8765 • Auto-Restore",
        "portrait": False,
    },
]

# Slide dimensions (widescreen 13.33" x 7.5")
SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)

# Dark background color
BG_COLOR   = RGBColor(0x0D, 0x11, 0x17)
# Title color
TTL_COLOR  = RGBColor(0xFF, 0xFF, 0xFF)
# Subtitle/accent color
SUB_COLOR  = RGBColor(0x5B, 0xC4, 0xD4)
# Thin accent bar color
BAR_COLOR  = RGBColor(0x1E, 0x88, 0xE5)

def set_slide_bg(slide, rgb: RGBColor):
    """Fill slide background with a solid color."""
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = rgb

def add_textbox(slide, text, left, top, width, height,
                font_size=24, bold=False, color=RGBColor(0xFF,0xFF,0xFF),
                align=PP_ALIGN.LEFT, wrap=True):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = "Segoe UI"
    return txBox

def add_rect(slide, left, top, width, height, fill_rgb, line_rgb=None, line_width_pt=0):
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        left, top, width, height
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_rgb
    if line_rgb:
        shape.line.color.rgb = line_rgb
        shape.line.width = Pt(line_width_pt)
    else:
        shape.line.fill.background()
    return shape

def add_arch_slide(prs, image_path, title_text, subtitle_text, portrait=False):
    """Add a new dark-themed slide with architecture image and title/subtitle."""
    # Use blank layout (index 6 is usually blank)
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # Dark background
    set_slide_bg(slide, BG_COLOR)

    # Thin accent bar at top
    add_rect(slide,
             left=Inches(0), top=Inches(0),
             width=SLIDE_W, height=Inches(0.06),
             fill_rgb=BAR_COLOR)

    # Header area with title
    HEADER_H = Inches(0.95)

    # Title textbox
    add_textbox(slide,
                text=title_text,
                left=Inches(0.35), top=Inches(0.08),
                width=Inches(11.0), height=Inches(0.55),
                font_size=22, bold=True,
                color=TTL_COLOR, align=PP_ALIGN.LEFT)

    # Subtitle textbox
    add_textbox(slide,
                text=subtitle_text,
                left=Inches(0.35), top=Inches(0.60),
                width=Inches(11.0), height=Inches(0.35),
                font_size=11, bold=False,
                color=SUB_COLOR, align=PP_ALIGN.LEFT)

    # Decorative shield emoji label (top-right)
    add_textbox(slide,
                text="🛡️ PII SHIELD",
                left=Inches(11.0), top=Inches(0.08),
                width=Inches(2.0), height=Inches(0.5),
                font_size=13, bold=True,
                color=SUB_COLOR, align=PP_ALIGN.RIGHT)

    # Thin horizontal rule under header
    add_rect(slide,
             left=Inches(0.35), top=Inches(0.95),
             width=Inches(12.63), height=Inches(0.015),
             fill_rgb=RGBColor(0x1E, 0x88, 0xE5))

    # Image area
    IMAGE_TOP = Inches(1.02)
    IMAGE_H   = Inches(6.35)  # full remaining height with small margin

    if portrait:
        # 9:16 image → preserve ratio, center horizontally
        img_ratio = 9/16   # width/height
        img_h_emu = IMAGE_H
        img_w_emu = int(img_h_emu * img_ratio)
        img_left  = int((SLIDE_W - img_w_emu) / 2)
        slide.shapes.add_picture(image_path, img_left, IMAGE_TOP, img_w_emu, img_h_emu)
    else:
        # 16:9 image → fill width with margin
        IMG_LEFT  = Inches(0.35)
        IMG_WIDTH = Inches(12.63)
        img_ratio = 16/9
        img_h_emu = int(IMG_WIDTH / img_ratio)
        img_top   = IMAGE_TOP + int((IMAGE_H - img_h_emu) / 2)
        slide.shapes.add_picture(image_path, IMG_LEFT, img_top, IMG_WIDTH, img_h_emu)

    # Bottom accent bar
    add_rect(slide,
             left=Inches(0), top=Inches(7.44),
             width=SLIDE_W, height=Inches(0.06),
             fill_rgb=BAR_COLOR)

    print(f"  Added slide: {title_text[:60]}")
    return slide

# ── Main ──────────────────────────────────────────────────────────────────────
prs = Presentation(PPTX_IN)
print(f"Loaded '{PPTX_IN}' — {len(prs.slides)} existing slides.")

# Ensure slide size is 13.33 x 7.5 inches (widescreen)
prs.slide_width  = SLIDE_W
prs.slide_height = SLIDE_H

for img_info in IMAGES:
    img_path = os.path.join(ARTIFACTS_DIR, img_info["file"])
    if not os.path.exists(img_path):
        print(f"  WARNING: Image not found: {img_path}")
        continue
    add_arch_slide(
        prs,
        image_path  = img_path,
        title_text  = img_info["title"],
        subtitle_text = img_info["subtitle"],
        portrait    = img_info["portrait"],
    )

prs.save(PPTX_OUT)
print(f"\nSaved updated deck -> '{PPTX_OUT}'")
print(f"Total slides now: {len(prs.slides)}")
