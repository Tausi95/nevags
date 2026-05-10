"""Generate NEVAGS DEC Presentation PowerPoint"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pptx.enum.dml import MSO_THEME_COLOR
import copy

# ── Colours ───────────────────────────────────────────────
FOREST   = RGBColor(0x1B, 0x43, 0x32)
FOREST2  = RGBColor(0x2D, 0x6A, 0x4F)
FMID     = RGBColor(0x40, 0x91, 0x6C)
ORANGE   = RGBColor(0xE8, 0x69, 0x0A)
CHARCOAL = RGBColor(0x1C, 0x2B, 0x28)
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
CREAM    = RGBColor(0xF8, 0xF4, 0xEE)
LGRAY    = RGBColor(0xF1, 0xF5, 0xF9)
DGRAY    = RGBColor(0x64, 0x74, 0x8B)
GREEN3   = RGBColor(0xD8, 0xF3, 0xDC)
ORANGE2  = RGBColor(0xFD, 0xBA, 0x74)

# ── Slide dimensions (widescreen 16:9) ────────────────────
W = Inches(13.33)
H = Inches(7.5)

prs = Presentation()
prs.slide_width  = W
prs.slide_height = H

BLANK = prs.slide_layouts[6]  # completely blank

# ── Helpers ───────────────────────────────────────────────
def add_rect(slide, x, y, w, h, fill, alpha=None):
    s = slide.shapes.add_shape(1, x, y, w, h)
    s.line.fill.background()
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    return s

def add_text(slide, text, x, y, w, h, size=18, bold=False, color=WHITE,
             align=PP_ALIGN.LEFT, wrap=True, italic=False):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tb.word_wrap = wrap
    p = tb.text_frame.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = "Calibri"
    return tb

def add_para(tf, text, size=14, bold=False, color=CHARCOAL, align=PP_ALIGN.LEFT,
             space_before=0, italic=False):
    p = tf.add_paragraph()
    p.alignment = align
    if space_before:
        p.space_before = Pt(space_before)
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = "Calibri"
    return p

def section_header(slide, label, title, dark=True):
    """Standard section header band at top of slide."""
    bg = CHARCOAL if dark else FOREST
    add_rect(slide, 0, 0, W, Inches(1.3), bg)
    add_text(slide, label.upper(), Inches(0.4), Inches(0.08), Inches(12), Inches(0.3),
             size=9, bold=True, color=ORANGE, align=PP_ALIGN.LEFT)
    add_text(slide, title, Inches(0.4), Inches(0.32), Inches(12.5), Inches(0.85),
             size=28, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

def accent_bar(slide, x, y, w=Inches(0.5), h=Inches(0.05)):
    add_rect(slide, x, y, w, h, ORANGE)

def stat_box(slide, x, y, w, h, number, label, sub="", bg=FOREST, num_color=WHITE):
    add_rect(slide, x, y, w, h, bg)
    add_text(slide, number, x+Inches(0.15), y+Inches(0.12), w-Inches(0.3), Inches(0.55),
             size=26, bold=True, color=num_color, align=PP_ALIGN.CENTER)
    add_text(slide, label, x+Inches(0.1), y+Inches(0.62), w-Inches(0.2), Inches(0.28),
             size=11, bold=True, color=ORANGE2, align=PP_ALIGN.CENTER)
    if sub:
        add_text(slide, sub, x+Inches(0.1), y+Inches(0.88), w-Inches(0.2), Inches(0.28),
                 size=8, color=RGBColor(0xCC,0xCC,0xCC), align=PP_ALIGN.CENTER)

def bullet_box(slide, x, y, w, h, title, bullets, bg=WHITE, title_color=FOREST,
               bullet_color=CHARCOAL, title_size=14, bullet_size=11):
    add_rect(slide, x, y, w, h, bg)
    add_text(slide, title, x+Inches(0.18), y+Inches(0.14), w-Inches(0.3), Inches(0.3),
             size=title_size, bold=True, color=title_color)
    bx = slide.shapes.add_textbox(x+Inches(0.18), y+Inches(0.48), w-Inches(0.3), h-Inches(0.55))
    bx.word_wrap = True
    tf = bx.text_frame
    tf.word_wrap = True
    first = True
    for b in bullets:
        if first:
            p = tf.paragraphs[0]; first = False
        else:
            p = tf.add_paragraph()
        p.space_before = Pt(3)
        run = p.add_run()
        run.text = ("• " if not b.startswith("•") else "") + b
        run.font.size = Pt(bullet_size)
        run.font.color.rgb = bullet_color
        run.font.name = "Calibri"

def add_table(slide, x, y, w, rows, cols, data, col_widths=None,
              header_bg=FOREST, row_bg=WHITE, alt_bg=LGRAY):
    row_h = Inches(0.42)
    tbl = slide.shapes.add_table(rows, cols, x, y, w, row_h * rows).table
    if col_widths:
        for i, cw in enumerate(col_widths):
            tbl.columns[i].width = cw
    for r in range(rows):
        for c in range(cols):
            cell = tbl.cell(r, c)
            cell.margin_left = Inches(0.08)
            cell.margin_right = Inches(0.08)
            cell.margin_top = Inches(0.04)
            cell.margin_bottom = Inches(0.04)
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            run = p.add_run()
            val = data[r][c] if r < len(data) and c < len(data[r]) else ""
            run.text = str(val)
            run.font.name = "Calibri"
            if r == 0:
                run.font.size = Pt(10)
                run.font.bold = True
                run.font.color.rgb = WHITE
                fill = cell.fill; fill.solid(); fill.fore_color.rgb = header_bg
            else:
                run.font.size = Pt(10)
                run.font.color.rgb = CHARCOAL
                bg = alt_bg if r % 2 == 0 else row_bg
                fill = cell.fill; fill.solid(); fill.fore_color.rgb = bg
    return tbl


# ═══════════════════════════════════════════════════════════
# SLIDE 1 — COVER
# ═══════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
add_rect(slide, 0, 0, W, H, CHARCOAL)
# Green left band
add_rect(slide, 0, 0, Inches(0.45), H, FOREST)
# Orange top accent strip
add_rect(slide, Inches(0.45), Inches(1.6), Inches(12.88), Inches(0.06), ORANGE)

add_text(slide, "DISTRICT EXECUTIVE COMMITTEE PRESENTATION",
         Inches(0.65), Inches(0.25), Inches(12), Inches(0.4),
         size=10, bold=True, color=ORANGE)
add_text(slide, "Mulanje District, Malawi  ·  May 2026",
         Inches(0.65), Inches(0.62), Inches(10), Inches(0.35),
         size=11, color=RGBColor(0xAA,0xBB,0xAA))

add_text(slide, "NEVAGS",
         Inches(0.65), Inches(1.72), Inches(12), Inches(1.4),
         size=72, bold=True, color=WHITE)
add_text(slide, "ECO BRICK & CONSTRUCTION",
         Inches(0.65), Inches(3.05), Inches(10), Inches(0.55),
         size=26, bold=True, color=FMID)
add_text(slide, '"Building Tomorrow Sustainably"',
         Inches(0.65), Inches(3.65), Inches(9), Inches(0.45),
         size=17, italic=True, color=RGBColor(0xCC,0xDD,0xCC))

# Key stats row
bw = Inches(2.9)
bh = Inches(1.15)
by = Inches(4.95)
stat_box(slide, Inches(0.65), by, bw, bh, "K450M+", "Capital Invested", "Founder + Atmosfair EUR 175K", FOREST2)
stat_box(slide, Inches(3.65), by, bw, bh, "51 Staff", "Current Workforce", "15 Female · 36 Male", CHARCOAL, ORANGE2)
stat_box(slide, Inches(6.65), by, bw, bh, "K2.07B", "Revenue Potential", "Annual MWK at full capacity", FOREST)
stat_box(slide, Inches(9.65), by, bw, bh, "8,000+", "Jobs in 3 Years", "Direct + Community ecosystem", ORANGE, WHITE)

add_text(slide, "One of the first black-owned industrial companies in Mulanje District",
         Inches(0.65), Inches(4.72), Inches(9), Inches(0.25),
         size=9, bold=True, color=ORANGE2)


# ═══════════════════════════════════════════════════════════
# SLIDE 2 — ABOUT / WHY WE EXIST
# ═══════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
add_rect(slide, 0, 0, W, H, CREAM)
section_header(slide, "Company Background", "Why NEVAGS Exists", dark=True)

# Left column: problem
add_rect(slide, Inches(0.3), Inches(1.5), Inches(5.9), Inches(5.6), WHITE)
add_text(slide, "THE PROBLEM", Inches(0.5), Inches(1.65), Inches(5.5), Inches(0.3),
         size=10, bold=True, color=ORANGE)
accent_bar(slide, Inches(0.5), Inches(1.95))
tx = slide.shapes.add_textbox(Inches(0.5), Inches(2.05), Inches(5.5), Inches(4.8))
tx.word_wrap = True
tf = tx.text_frame; tf.word_wrap = True
add_para(tf, "Malawi's Housing & Deforestation Crisis", 13, True, CHARCOAL)
add_para(tf, "• Over 80–90% of houses in Malawi are built with fired clay bricks",
         10, False, DGRAY, space_before=5)
add_para(tf, "• For decades, bricks were fired using massive quantities of firewood",
         10, False, DGRAY, space_before=3)
add_para(tf, "• Malawi has one of the fastest deforestation rates in Africa", 10, False, DGRAY, space_before=3)
add_para(tf, "• Deforestation causes soil erosion, flooding, biodiversity loss", 10, False, DGRAY, space_before=3)
add_para(tf, "", 6)
add_para(tf, "The Government's Response", 13, True, CHARCOAL, space_before=4)
add_para(tf, "• Malawi Government BANNED firewood-fired bricks", 10, True, RGBColor(0xCC,0x00,0x00), space_before=5)
add_para(tf, "• This creates a critical supply gap across the construction sector", 10, False, DGRAY, space_before=3)
add_para(tf, "• Developers, NGOs, and government projects need a legal alternative", 10, False, DGRAY, space_before=3)
add_para(tf, "• Demand is policy-driven, long-term, and nationwide", 10, True, FOREST, space_before=3)

# Right column: solution
add_rect(slide, Inches(6.5), Inches(1.5), Inches(6.5), Inches(5.6), FOREST)
add_text(slide, "THE SOLUTION", Inches(6.7), Inches(1.65), Inches(6.1), Inches(0.3),
         size=10, bold=True, color=ORANGE)
accent_bar(slide, Inches(6.7), Inches(1.95))
tx2 = slide.shapes.add_textbox(Inches(6.7), Inches(2.05), Inches(6.1), Inches(4.8))
tx2.word_wrap = True
tf2 = tx2.text_frame; tf2.word_wrap = True
add_para(tf2, "NEVAGS Eco Brick & Construction", 14, True, WHITE)
add_para(tf2, "Operating from Musewu, Mulanje District, we deploy VSK (Vertical Shaft Kiln) technology fuelled entirely by biomass briquettes — producing high-quality bricks with zero firewood.",
         10, False, RGBColor(0xCC,0xDD,0xCC), space_before=6)
add_para(tf2, "", 5)
add_para(tf2, "✓  Zero firewood — 100% policy compliant", 11, True, RGBColor(0x90,0xEE,0x90), space_before=4)
add_para(tf2, "✓  Scalable industrial production", 11, False, RGBColor(0xCC,0xDD,0xCC), space_before=3)
add_para(tf2, "✓  Affordable pricing for developers", 11, False, RGBColor(0xCC,0xDD,0xCC), space_before=3)
add_para(tf2, "✓  Located in the heart of Mulanje", 11, False, RGBColor(0xCC,0xDD,0xCC), space_before=3)
add_para(tf2, "✓  Already employing 51 community members", 11, False, RGBColor(0xCC,0xDD,0xCC), space_before=3)
add_para(tf2, "✓  Backed by GIZ, Atmosfair & research partners", 11, False, RGBColor(0xCC,0xDD,0xCC), space_before=3)
add_para(tf2, "", 5)
add_para(tf2, '"When we build, let us think that we build forever."', 10, True,
         RGBColor(0xF5,0x9E,0x0B), space_before=4, italic=True)
add_para(tf2, "— John Ruskin", 9, False, RGBColor(0xAA,0xBB,0xAA), space_before=2)


# ═══════════════════════════════════════════════════════════
# SLIDE 3 — LEADERSHIP & OWNERSHIP
# ═══════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
add_rect(slide, 0, 0, W, H, CREAM)
section_header(slide, "Company Ownership & Location", "Leadership & Registration", dark=True)

# Pioneer badge
add_rect(slide, Inches(0.3), Inches(1.5), Inches(12.73), Inches(0.55), RGBColor(0xFF,0xF3,0xE0))
add_text(slide, "🏆  One of the first black-owned industrial companies in Mulanje District — a pioneer enterprise building a legacy of community ownership",
         Inches(0.5), Inches(1.56), Inches(12.4), Inches(0.42),
         size=10, bold=True, color=ORANGE)

# Charles Nasala card
add_rect(slide, Inches(0.3), Inches(2.2), Inches(4.0), Inches(4.5), FOREST)
accent_bar(slide, Inches(0.3), Inches(2.2), Inches(4.0), Inches(0.08))
add_text(slide, "FOUNDER, OWNER &\nMANAGING DIRECTOR",
         Inches(0.5), Inches(2.3), Inches(3.6), Inches(0.5),
         size=9, bold=True, color=ORANGE)
add_text(slide, "Charles Billy Nasala",
         Inches(0.5), Inches(2.78), Inches(3.6), Inches(0.55),
         size=18, bold=True, color=WHITE)
tx = slide.shapes.add_textbox(Inches(0.5), Inches(3.35), Inches(3.5), Inches(3.1))
tx.word_wrap = True; tf = tx.text_frame; tf.word_wrap = True
add_para(tf, "Entrepreneur, industrial innovator, and community development champion. Founder of one of Malawi's first VSK brick enterprises.",
         9, False, RGBColor(0xCC,0xDD,0xCC))
add_para(tf, "", 4)
add_para(tf, "📞  +265 888 34 75 75", 9, False, RGBColor(0xAA,0xFF,0xAA), space_before=3)
add_para(tf, "📞  +265 99 751 0160", 9, False, RGBColor(0xAA,0xFF,0xAA), space_before=2)
add_para(tf, "✉   nasalacharles.b@gmail.com", 9, False, ORANGE2, space_before=3)

# Chancy Tsonga card
add_rect(slide, Inches(4.5), Inches(2.2), Inches(4.0), Inches(4.5), CHARCOAL)
accent_bar(slide, Inches(4.5), Inches(2.2), Inches(4.0), Inches(0.08))
add_text(slide, "FOUNDING ENGINEER &\nBUSINESS DEVELOPMENT MANAGER",
         Inches(4.7), Inches(2.3), Inches(3.6), Inches(0.5),
         size=9, bold=True, color=ORANGE)
add_text(slide, "Chancy Tausi Tsonga",
         Inches(4.7), Inches(2.78), Inches(3.6), Inches(0.55),
         size=18, bold=True, color=WHITE)
tx2 = slide.shapes.add_textbox(Inches(4.7), Inches(3.35), Inches(3.5), Inches(3.1))
tx2.word_wrap = True; tf2 = tx2.text_frame; tf2.word_wrap = True
add_para(tf2, "VSK technology specialist, construction innovator, and strategic partner. Leads business development and market expansion.",
         9, False, RGBColor(0xCC,0xDD,0xCC))
add_para(tf2, "", 4)
add_para(tf2, "📞  +265 984 000 366", 9, False, RGBColor(0xAA,0xFF,0xAA), space_before=3)
add_para(tf2, "💬  WhatsApp: +27 764 998 4601", 9, False, RGBColor(0xAA,0xFF,0xAA), space_before=2)
add_para(tf2, "✉   chancy.tsonga@yahoo.com", 9, False, ORANGE2, space_before=3)
add_para(tf2, "🌐  chancytsonga.com", 9, False, RGBColor(0x90,0xD0,0xFF), space_before=2)

# Company details card
add_rect(slide, Inches(8.7), Inches(2.2), Inches(4.33), Inches(4.5), WHITE)
accent_bar(slide, Inches(8.7), Inches(2.2), Inches(4.33), Inches(0.08))
add_text(slide, "COMPANY DETAILS",
         Inches(8.9), Inches(2.3), Inches(4.0), Inches(0.3),
         size=9, bold=True, color=ORANGE)
tx3 = slide.shapes.add_textbox(Inches(8.9), Inches(2.65), Inches(3.9), Inches(3.8))
tx3.word_wrap = True; tf3 = tx3.text_frame; tf3.word_wrap = True
add_para(tf3, "Registered Entity", 10, True, CHARCOAL)
add_para(tf3, "New Vision Anenenji Construction", 10, False, DGRAY, space_before=2)
add_para(tf3, "Registration No. 46289", 10, False, DGRAY, space_before=1)
add_para(tf3, "", 4)
add_para(tf3, "Location", 10, True, CHARCOAL, space_before=4)
add_para(tf3, "Musewu, Mulanje District\nSouthern Malawi", 10, False, DGRAY, space_before=2)
add_para(tf3, "P.O. Box 90, Mulanje", 10, False, DGRAY, space_before=1)
add_para(tf3, "", 4)
add_para(tf3, "Operational Reach", 10, True, CHARCOAL, space_before=4)
add_para(tf3, "Blantyre · Lilongwe · Mzuzu\nSouthern Africa expansion planned", 10, False, DGRAY, space_before=2)
add_para(tf3, "", 4)
add_para(tf3, "Careers / HR", 10, True, CHARCOAL, space_before=4)
add_para(tf3, "careers.nevags@gmail.com", 10, False, ORANGE, space_before=2)


# ═══════════════════════════════════════════════════════════
# SLIDE 4 — MISSION & VISION
# ═══════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
add_rect(slide, 0, 0, W, H, RGBColor(0x0F, 0x1F, 0x15))
section_header(slide, "Mission, Vision & Objectives", "Our Strategic Foundation", dark=False)

# Mission
add_rect(slide, Inches(0.3), Inches(1.5), Inches(4.1), Inches(5.6), FOREST2)
add_text(slide, "MISSION", Inches(0.5), Inches(1.65), Inches(3.8), Inches(0.3),
         size=10, bold=True, color=ORANGE)
add_text(slide, "What We Do",
         Inches(0.5), Inches(1.97), Inches(3.8), Inches(0.4),
         size=16, bold=True, color=WHITE)
add_text(slide, "To produce high-quality, environmentally compliant bricks using VSK technology and biomass briquettes — providing affordable building materials that support Malawi's housing sector while protecting forests, creating jobs, and empowering communities in Mulanje District and beyond.",
         Inches(0.5), Inches(2.4), Inches(3.7), Inches(3.2),
         size=11, color=RGBColor(0xCC,0xEE,0xDD))

# Vision
add_rect(slide, Inches(4.6), Inches(1.5), Inches(4.1), Inches(5.6), CHARCOAL)
add_text(slide, "VISION", Inches(4.8), Inches(1.65), Inches(3.8), Inches(0.3),
         size=10, bold=True, color=ORANGE)
add_text(slide, "Where We're Going",
         Inches(4.8), Inches(1.97), Inches(3.8), Inches(0.4),
         size=16, bold=True, color=WHITE)
add_text(slide, "To build a scalable ecosystem for sustainable, affordable, and climate-resilient housing across Malawi and Southern Africa — positioning NEVAGS as the benchmark for eco-industrial construction innovation on the African continent.",
         Inches(4.8), Inches(2.4), Inches(3.7), Inches(3.0),
         size=11, color=RGBColor(0xCC,0xCC,0xCC))

# Objectives
add_rect(slide, Inches(8.9), Inches(1.5), Inches(4.13), Inches(5.6), RGBColor(0x1C,0x3A,0x2C))
add_text(slide, "STRATEGIC OBJECTIVES", Inches(9.1), Inches(1.65), Inches(3.8), Inches(0.3),
         size=10, bold=True, color=ORANGE)
add_text(slide, "Key Goals",
         Inches(9.1), Inches(1.97), Inches(3.8), Inches(0.4),
         size=16, bold=True, color=WHITE)
tx = slide.shapes.add_textbox(Inches(9.1), Inches(2.4), Inches(3.75), Inches(4.4))
tx.word_wrap = True; tf = tx.text_frame; tf.word_wrap = True
for obj in [
    "Produce 4.8M+ bricks/year — zero firewood",
    "Scale to 200+ internal employees in 3 years",
    "Create 8,000+ jobs in the district ecosystem",
    "Achieve 30–45% annual ROI for investors",
    "Commission biogas 2nd kiln shaft in 14 months",
    "Expand to Blantyre, Lilongwe & Mzuzu markets",
    "Maintain ≥30% female employment — SDG 5",
    "Full ESG compliance — SDGs 3, 5, 13, 15",
]:
    add_para(tf, "✓  " + obj, 10, False, RGBColor(0x90,0xEE,0x90), space_before=5)

# Quote at bottom
add_rect(slide, Inches(0.3), Inches(7.05), Inches(12.73), Inches(0.38), RGBColor(0x08,0x14,0x0E))
add_text(slide, '"Civilizations are remembered by what they build. The responsibility of our generation is to build differently."  — Chancy Tausi Tsonga',
         Inches(0.5), Inches(7.08), Inches(12.4), Inches(0.32),
         size=9, italic=True, color=ORANGE2, align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════
# SLIDE 5 — PRODUCTS
# ═══════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
add_rect(slide, 0, 0, W, H, CREAM)
section_header(slide, "Products & Services", "What We Produce", dark=True)

pw = Inches(4.0)
ph = Inches(5.5)
py = Inches(1.55)

for i, (bg, tag, name, cost, price, margin, cap, notes) in enumerate([
    (FOREST, "CORE PRODUCT",
     "Ordinary VSK Bricks",
     "K195/brick", "K300/brick (QS Rate)", "35% · K105/brick",
     "3,000,000 / year",
     ["100% firewood-free — fully legal",
      "Passes Malawi govt compliance",
      "Suitable for all construction",
      "Competitive vs cement blocks"]),
    (CHARCOAL, "PREMIUM PRODUCT",
     "Face Bricks (Double-Faced)",
     "K335/brick", "K510–K650/brick (QS)", "34% · K175/brick",
     "1,800,000 / year",
     ["Premium surface finish",
      "Higher absolute margin",
      "Commercial & residential",
      "Supports early-phase revenue"]),
    (FOREST2, "ECO INNOVATION",
     "Biomass Briquettes",
     "Agri waste feedstock", "Kiln fuel (self-use)", "Cost elimination",
     "Self-sufficient on-site",
     ["Zero firewood replacement",
      "Made from rice husks & biomass",
      "Circular economy model",
      "Community supply chain"]),
]):
    x = Inches(0.3 + i * 4.3)
    add_rect(slide, x, py, pw, ph, bg)
    add_text(slide, tag, x+Inches(0.18), py+Inches(0.12), pw-Inches(0.3), Inches(0.25),
             size=8, bold=True, color=ORANGE)
    add_text(slide, name, x+Inches(0.18), py+Inches(0.35), pw-Inches(0.3), Inches(0.55),
             size=15, bold=True, color=WHITE)
    add_rect(slide, x, py+Inches(0.88), pw, Inches(0.04), ORANGE)

    tx = slide.shapes.add_textbox(x+Inches(0.18), py+Inches(0.97), pw-Inches(0.3), Inches(2.1))
    tx.word_wrap = True; tf = tx.text_frame; tf.word_wrap = True
    for lbl, val in [("Production Cost", cost), ("Selling Price", price), ("Margin", margin), ("Capacity", cap)]:
        p = tf.add_paragraph() if tf.paragraphs[0].runs else tf.paragraphs[0]
        p.space_before = Pt(4)
        r = p.add_run(); r.text = lbl + ": "; r.font.size = Pt(9); r.font.bold = True
        r.font.color.rgb = ORANGE2; r.font.name = "Calibri"
        r2 = p.add_run(); r2.text = val; r2.font.size = Pt(9)
        r2.font.color.rgb = WHITE; r2.font.name = "Calibri"
        tf.add_paragraph()  # spacing

    add_rect(slide, x+Inches(0.18), py+Inches(3.1), pw-Inches(0.36), Inches(0.03),
             RGBColor(0x40,0x70,0x50))
    tx2 = slide.shapes.add_textbox(x+Inches(0.18), py+Inches(3.18), pw-Inches(0.3), Inches(2.1))
    tx2.word_wrap = True; tf2 = tx2.text_frame; tf2.word_wrap = True
    for b in notes:
        p = tf2.add_paragraph()
        p.space_before = Pt(5)
        run = p.add_run(); run.text = "✓  " + b
        run.font.size = Pt(9); run.font.color.rgb = RGBColor(0xCC,0xEE,0xCC)
        run.font.name = "Calibri"

# Annual totals bar
add_rect(slide, Inches(0.3), Inches(7.1), Inches(12.73), Inches(0.35), CHARCOAL)
add_text(slide,
         "Annual Revenue: K900M (Ordinary)  +  K1,170M (Face)  =  K2.07 BILLION   ·   Gross Profit: K774M (~$440,000 USD)",
         Inches(0.5), Inches(7.12), Inches(12.4), Inches(0.3),
         size=10, bold=True, color=ORANGE2, align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════
# SLIDE 6 — INVESTMENT OVERVIEW
# ═══════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
add_rect(slide, 0, 0, W, H, CREAM)
section_header(slide, "Investment Overview", "Capital, Returns & Financial Snapshot", dark=True)

# Capital base row
for i, (num, lbl, sub, bg) in enumerate([
    ("K175M+",  "Founder's Equity",    "Charles Billy Nasala\npersonal investment",          FOREST),
    ("EUR 175K","Atmosfair Loan",       "2nd Amendment Apr 2026\n5 disbursements",            CHARCOAL),
    ("K450M+",  "Total Capital Base",  "Founder equity +\ninternational finance",             FOREST2),
    ("K2.07B",  "Annual Revenue",      "At full production\n(~$1.2M USD)",                   ORANGE),
]):
    x = Inches(0.3 + i * 3.2)
    stat_box(slide, x, Inches(1.5), Inches(3.0), Inches(1.4),
             num, lbl, sub, bg if i<3 else ORANGE)

# ROI boxes
for i, (num, lbl, body) in enumerate([
    ("30–45%", "Annual ROI Potential",
     "Based on full capacity, QS pricing & existing infrastructure. Policy-driven demand = long-term revenue."),
    ("3–5 yrs", "Payback Period",
     "Conservative payback on total investment. Expansion financing accelerates this significantly."),
    ("$2.5M+", "Revenue in 3 Years",
     "With expansion financing, revenue exceeds $2.5M USD within 3 years of scaled operations."),
    ("K774M", "Annual Gross Profit",
     "Annual gross profit at full production. (~$440,000 USD). Grows with biogas kiln addition."),
]):
    x = Inches(0.3 + i * 3.2)
    add_rect(slide, x, Inches(3.1), Inches(3.0), Inches(1.6), WHITE)
    add_text(slide, num, x+Inches(0.15), Inches(3.18), Inches(2.7), Inches(0.55),
             size=24, bold=True, color=FOREST, align=PP_ALIGN.CENTER)
    add_text(slide, lbl, x+Inches(0.15), Inches(3.7), Inches(2.7), Inches(0.25),
             size=9, bold=True, color=ORANGE, align=PP_ALIGN.CENTER)
    add_text(slide, body, x+Inches(0.12), Inches(3.96), Inches(2.76), Inches(0.68),
             size=8, color=DGRAY, align=PP_ALIGN.CENTER)

# Revenue table
add_text(slide, "Annual Revenue & Profit Breakdown",
         Inches(0.3), Inches(4.85), Inches(8), Inches(0.3),
         size=12, bold=True, color=CHARCOAL)
add_table(slide, Inches(0.3), Inches(5.15), Inches(12.73), 4, 5,
    [
        ["Product",             "Capacity",         "Production Cost", "Selling Price (QS)", "Annual Revenue"],
        ["Ordinary VSK Bricks", "3,000,000 / year", "K195 / brick",   "K300 / brick",       "K900,000,000"],
        ["Face Bricks (Premium)","1,800,000 / year","K335 / brick",   "K510–K650 / brick",  "K1,170,000,000"],
        ["TOTAL",               "4,800,000 / year", "—",              "—",                  "K2,070,000,000"],
    ],
    col_widths=[Inches(2.6), Inches(2.2), Inches(2.0), Inches(2.5), Inches(3.43)],
    header_bg=FOREST)


# ═══════════════════════════════════════════════════════════
# SLIDE 7 — MARKET COMPARISON
# ═══════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
add_rect(slide, 0, 0, W, H, CREAM)
section_header(slide, "Pricing & Market Position", "Cost per m² of Wall — NEVAGS vs Alternatives", dark=True)

add_text(slide, "A standard half-brick wall (1m²) requires ~59 standard bricks or ~12 cement blocks. This is how NEVAGS compares:",
         Inches(0.3), Inches(1.42), Inches(12.73), Inches(0.28),
         size=10, italic=True, color=DGRAY)

# Bar chart simulation using rectangles
bar_y = Inches(2.0)
bar_max_w = Inches(7.5)
for i, (label, cost, pct, bg, note) in enumerate([
    ("Cement Blocks\n(12 blocks × K3,000)", "K36,000/m²", 1.0,   RGBColor(0xDC,0x26,0x26),
     "Expensive. Requires extra mortar, rebar & skilled labour. High transport cost to Mulanje."),
    ("NEVAGS VSK Bricks\n(59 bricks × K300)",  "K17,700/m²", 0.492, FOREST,
     "✓ LEGAL  ✓ LOCAL  ✓ AFFORDABLE — 51% cheaper than cement. Policy compliant. Eligible for all govt projects."),
    ("Traditional Firewood Bricks\n(NOW ILLEGAL)", "K8,850/m²", 0.246, RGBColor(0xCA,0x8A,0x04),
     "⚠ BANNED BY MALAWI GOVERNMENT. Using these bricks = illegal. Not eligible for any formal construction project."),
]):
    y = bar_y + Inches(i * 1.55)
    add_text(slide, label, Inches(0.3), y, Inches(3.2), Inches(0.6),
             size=10, bold=True, color=CHARCOAL)
    bw = bar_max_w * pct
    add_rect(slide, Inches(3.6), y + Inches(0.05), bw, Inches(0.45), bg)
    add_text(slide, cost, Inches(3.6) + bw + Inches(0.1), y + Inches(0.08),
             Inches(1.5), Inches(0.35), size=12, bold=True, color=CHARCOAL)
    add_text(slide, note, Inches(0.3), y + Inches(0.65), Inches(12.73), Inches(0.6),
             size=9, color=DGRAY)

# Conclusion
add_rect(slide, Inches(0.3), Inches(6.85), Inches(12.73), Inches(0.52), CHARCOAL)
add_text(slide,
         "NEVAGS IS 51% CHEAPER THAN CEMENT BLOCKS — The only legal, affordable, locally-produced alternative in Mulanje District",
         Inches(0.5), Inches(6.9), Inches(12.4), Inches(0.42),
         size=12, bold=True, color=ORANGE2, align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════
# SLIDE 8 — ENVIRONMENTAL IMPACT & SDGs
# ═══════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
add_rect(slide, 0, 0, W, H, RGBColor(0x0A, 0x18, 0x0F))
section_header(slide, "Environmental & Social Impact", "Built for a Sustainable Malawi", dark=False)

# SDG cards
sdg_data = [
    (RGBColor(0x4C,0x9F,0x38), "SDG 3", "Good Health\n& Well-Being",
     ["Safety compliance ≥95%", "PPE for all 51 workers", "Reduced air pollution", "Community health protection"]),
    (RGBColor(0xBF,0x0D,0x0D), "SDG 5", "Gender\nEquality",
     ["15 female employees (29%)", "Equal pay policy", "Female leadership roles", "Target: ≥30% women"]),
    (RGBColor(0x3F,0x7E,0x44), "SDG 13", "Climate\nAction",
     ["Zero firewood in production", "Biomass briquette fuel", "Carbon emission reduction", "Forest preservation"]),
    (RGBColor(0x56,0xC0,0x2B), "SDG 15", "Life\non Land",
     ["Zero deforestation impact", "Protects Mulanje Massif", "Biodiversity preservation", "Water catchment protection"]),
]
sw = Inches(2.95)
for i, (color, sdg, title, points) in enumerate(sdg_data):
    x = Inches(0.3 + i * 3.25)
    add_rect(slide, x, Inches(1.55), sw, Inches(0.7), color)
    add_text(slide, sdg, x+Inches(0.15), Inches(1.6), sw-Inches(0.3), Inches(0.28),
             size=18, bold=True, color=WHITE)
    add_text(slide, title, x+Inches(0.15), Inches(1.88), sw-Inches(0.3), Inches(0.38),
             size=10, color=WHITE)
    add_rect(slide, x, Inches(2.25), sw, Inches(4.0), RGBColor(0x15,0x30,0x1E))
    tx = slide.shapes.add_textbox(x+Inches(0.15), Inches(2.32), sw-Inches(0.25), Inches(3.85))
    tx.word_wrap = True; tf = tx.text_frame; tf.word_wrap = True
    for p in points:
        add_para(tf, "✓  " + p, 10, False, RGBColor(0x90,0xEE,0x90), space_before=8)

# Zero waste strip
add_rect(slide, Inches(0.3), Inches(6.35), Inches(12.73), Inches(1.05), RGBColor(0x1B,0x43,0x32))
add_text(slide, "ZERO WASTE MODEL",
         Inches(0.5), Inches(6.4), Inches(12.4), Inches(0.25),
         size=9, bold=True, color=ORANGE)
add_text(slide,
         "Production dust & clay → recovered  ·  Agricultural biomass → briquette fuel  ·  Organic waste → Biogas (2nd kiln, 14 months)  ·  Kiln ash → soil amendment",
         Inches(0.5), Inches(6.65), Inches(12.4), Inches(0.68),
         size=11, color=RGBColor(0xCC,0xEE,0xCC), align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════
# SLIDE 9 — EMPLOYMENT & JOB CREATION
# ═══════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
add_rect(slide, 0, 0, W, H, CREAM)
section_header(slide, "Employment & Job Creation", "51 Jobs Today. 8,000+ in 3 Years.", dark=True)

# Big numbers
for i, (num, lbl, sub, bg) in enumerate([
    ("51",     "Current Employees",    "15 Female (29%) · 36 Male (71%)", FOREST),
    ("200+",   "Internal Jobs (3 yrs)","All departments incl.\nMech. Engineering & Maintenance", CHARCOAL),
    ("5,000+", "Community Moulders",   "Trained & buying from them;\nNEVAGS fires in VSK kiln",  FOREST2),
    ("2,500+", "Biomass Suppliers",    "Rice husks & agri waste\nbuyers from community",         ORANGE),
]):
    x = Inches(0.3 + i * 3.2)
    stat_box(slide, x, Inches(1.5), Inches(3.0), Inches(1.45), num, lbl, sub, bg)

# Community model
add_text(slide, "The NEVAGS Job Creation Ecosystem",
         Inches(0.3), Inches(3.1), Inches(12), Inches(0.3),
         size=13, bold=True, color=CHARCOAL)

for i, (title, body, bg) in enumerate([
    ("Community Brick Moulding Programme",
     "NEVAGS trains community members in standard VSK-quality brick moulding. They produce green bricks locally — NEVAGS buys directly from them and fires in our kiln. Self-employment at community scale.",
     WHITE),
    ("Biomass & Rice Husk Supply Chain",
     "We purchase rice husks, maize cobs, and agricultural biomass from community farmers. Farm waste becomes clean energy income — turning zero-value materials into a new rural revenue stream.",
     RGBColor(0xFE,0xF9,0xEE)),
    ("Internal Growth — All Departments",
     "Beyond operations, NEVAGS scales across Sales, HR, Admin, Finance, and a dedicated Mechanical Engineering & Maintenance department. 200+ formal, permanent jobs within 3 years.",
     WHITE),
]):
    x = Inches(0.3 + i * 4.3)
    add_rect(slide, x, Inches(3.5), Inches(4.1), Inches(3.42), bg)
    add_rect(slide, x, Inches(3.5), Inches(4.1), Inches(0.06), ORANGE)
    add_text(slide, title, x+Inches(0.15), Inches(3.58), Inches(3.85), Inches(0.4),
             size=11, bold=True, color=FOREST)
    add_text(slide, body, x+Inches(0.15), Inches(4.02), Inches(3.85), Inches(2.7),
             size=10, color=DGRAY)

# Gender bar at bottom
add_rect(slide, Inches(0.3), Inches(7.0), Inches(12.73), Inches(0.42), CHARCOAL)
# Visual gender bar
add_rect(slide, Inches(0.5), Inches(7.08), Inches(8.5), Inches(0.25), FOREST)  # Male 71%
add_rect(slide, Inches(9.0), Inches(7.08), Inches(3.73), Inches(0.25), ORANGE)  # Female 29%
add_text(slide, "Male 71% (36)",  Inches(0.5), Inches(7.0), Inches(2.5), Inches(0.22), size=8, bold=True, color=WHITE)
add_text(slide, "Female 29% (15) — SDG 5 · Target ≥30%", Inches(9.05), Inches(7.0), Inches(3.9), Inches(0.22), size=8, bold=True, color=ORANGE2)


# ═══════════════════════════════════════════════════════════
# SLIDE 10 — WORKFORCE PLAN
# ═══════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
add_rect(slide, 0, 0, W, H, CREAM)
section_header(slide, "Workforce Plan", "KPI-Driven Phased Scaling — Apr to Jul 2026", dark=True)

add_table(slide, Inches(0.3), Inches(1.5), Inches(12.73), 6, 5,
    [
        ["Phase",         "Period",           "Staff on Site", "Daily Output Target",   "Key KPIs"],
        ["Startup",       "Wk 1–2 (Apr 6–19)","18 staff",     "Site commissioning",    "Safety ≥95% · Downtime <10%"],
        ["Phase 1",       "Wk 3–4 (Apr–May)", "33 staff",     "≥3,000 bricks/day",     "Output ≥3,000/day · Cost ≤K195/brick"],
        ["Phase 2",       "Wk 5–8 (May)",     "39–40 staff",  "≥5,000 bricks/day",     "Kiln Util >80% · Orders >50,000/wk"],
        ["Phase 3 / Full","Wk 9–13 (Jun–Jul)","45 staff",     "≥7,000 bricks/day",     "Active Clients ≥5 · Full shift ops"],
        ["Expansion",     "2026–2029",         "200+ internal","Doubled capacity",       "8,000+ ecosystem jobs · Biogas kiln"],
    ],
    col_widths=[Inches(1.6), Inches(2.4), Inches(1.9), Inches(2.9), Inches(3.93)],
    header_bg=FOREST)

add_text(slide, "Compensation Philosophy",
         Inches(0.3), Inches(4.85), Inches(12), Inches(0.3),
         size=12, bold=True, color=CHARCOAL)

for i, (title, body, bg) in enumerate([
    ("Senior & Management Staff",
     "Competitive market-rate remuneration — above-market base salaries reflecting expertise and business contribution",
     FOREST),
    ("General Workers & Supervisors",
     "Above statutory minimum wage — no exceptions. Stable, year-round income with career progression tied to KPI performance",
     CHARCOAL),
    ("Community Moulders & Biomass Suppliers",
     "Direct purchase payments. Guaranteed offtake agreements create predictable income for community entrepreneurs",
     FOREST2),
]):
    x = Inches(0.3 + i * 4.25)
    add_rect(slide, x, Inches(5.2), Inches(4.1), Inches(2.1), bg)
    add_text(slide, title, x+Inches(0.15), Inches(5.3), Inches(3.85), Inches(0.35),
             size=10, bold=True, color=ORANGE2)
    add_text(slide, body, x+Inches(0.15), Inches(5.66), Inches(3.85), Inches(1.5),
             size=10, color=RGBColor(0xCC,0xDD,0xCC))


# ═══════════════════════════════════════════════════════════
# SLIDE 11 — PARTNERS & STAKEHOLDERS
# ═══════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
add_rect(slide, 0, 0, W, H, CREAM)
section_header(slide, "Partners & Stakeholders", "Backed by Leading Institutions", dark=True)

partners = [
    ("GIZ",       "Deutsche Gesellschaft für\nInternationale Zusammenarbeit",
     "Technical & Financial Partner\nGerman Development Cooperation"),
    ("Atmosfair",  "German Climate Protection NGO\nBerlin, Germany",
     "Carbon Finance & Climate Compliance\nLoan Provider (EUR 175,000)"),
    ("CCODE",      "Centre for Community\nOrganisation & Development",
     "Community Engagement\n& Local Development Partner"),
    ("TERA",       "Technical & Environmental\nResearch Associate",
     "Environmental Compliance\n& Research Partner"),
    ("MUBAS",      "Malawi University of Business\n& Applied Sciences",
     "Academic Research\n& Skills Development"),
    ("CIRA",       "Construction Industry\nRegulatory Authority",
     "Industry Compliance\n& Market Development"),
]
for i, (name, org, role) in enumerate(partners):
    col = i % 3
    row = i // 3
    x = Inches(0.3 + col * 4.25)
    y = Inches(1.55 + row * 2.65)
    add_rect(slide, x, y, Inches(4.1), Inches(2.4), WHITE)
    add_rect(slide, x, y, Inches(4.1), Inches(0.06), ORANGE if row == 0 else FOREST)
    add_text(slide, name, x+Inches(0.15), y+Inches(0.14), Inches(3.8), Inches(0.48),
             size=20, bold=True, color=FOREST)
    add_text(slide, org, x+Inches(0.15), y+Inches(0.62), Inches(3.8), Inches(0.5),
             size=9, color=DGRAY)
    add_text(slide, role, x+Inches(0.15), y+Inches(1.12), Inches(3.8), Inches(0.5),
             size=9, bold=True, color=CHARCOAL)

add_rect(slide, Inches(0.3), Inches(7.0), Inches(12.73), Inches(0.42), CHARCOAL)
add_text(slide,
         "Government  ·  NGOs & Dev Finance  ·  Academic Institutions  ·  Industry Regulators  ·  Community Organisations",
         Inches(0.5), Inches(7.05), Inches(12.4), Inches(0.32),
         size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════
# SLIDE 12 — IMPLEMENTATION ROADMAP
# ═══════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
add_rect(slide, 0, 0, W, H, RGBColor(0x0F,0x1F,0x15))
section_header(slide, "Implementation Roadmap", "Phased Growth Plan — 2026 to 2029", dark=False)

phases = [
    (FOREST2, "PHASE 1", "Foundation &\nCommissioning",
     "Apr–May 2026",
     ["VSK kiln commissioning complete", "18→33 staff on-boarded", "Moulding operations begin",
      "≥3,000 bricks/day output", "First client engagements"]),
    (ORANGE, "PHASE 2", "Scale\nProduction",
     "May–Jun 2026",
     ["Output scaled to ≥5,000/day", "39–40 staff deployed", "Kiln utilisation >80%",
      "50,000+ bricks/week orders", "Market penetration active"]),
    (FMID, "PHASE 3", "Full-Scale\nOperations",
     "Jun–Jul 2026",
     ["45 staff at full capacity", "Output ≥7,000 bricks/day", "≥5 active bulk clients",
      "Double-shift introduced", "Revenue targets met"]),
    (RGBColor(0x3B,0x82,0xF6), "PHASE 4", "Expansion &\nInnovation",
     "2026–2029",
     ["200+ internal staff", "8,000+ ecosystem jobs", "BIOGAS 2nd kiln (14 months)",
      "Revenue >$2.5M USD", "Southern Africa expansion"]),
]
pw = Inches(2.95)
for i, (bg, ph, title, period, pts) in enumerate(phases):
    x = Inches(0.3 + i * 3.25)
    add_rect(slide, x, Inches(1.55), pw, Inches(0.6), bg)
    add_text(slide, ph, x+Inches(0.15), Inches(1.6), pw-Inches(0.3), Inches(0.25),
             size=10, bold=True, color=WHITE)
    add_text(slide, period, x+Inches(0.15), Inches(1.82), pw-Inches(0.3), Inches(0.28),
             size=8, color=RGBColor(0xCC,0xCC,0xCC))
    add_rect(slide, x, Inches(2.15), pw, Inches(4.5), RGBColor(0x15,0x30,0x1E))
    add_text(slide, title, x+Inches(0.15), Inches(2.2), pw-Inches(0.3), Inches(0.5),
             size=14, bold=True, color=WHITE)
    tx = slide.shapes.add_textbox(x+Inches(0.15), Inches(2.72), pw-Inches(0.25), Inches(3.8))
    tx.word_wrap = True; tf = tx.text_frame; tf.word_wrap = True
    for pt in pts:
        add_para(tf, "• " + pt, 10, False, RGBColor(0xCC,0xEE,0xCC), space_before=7)

# Biogas spotlight
add_rect(slide, Inches(0.3), Inches(6.75), Inches(12.73), Inches(0.68), RGBColor(0x92,0x40,0x0E))
add_text(slide, "🔥  IN 14 MONTHS: Biogas-Powered 2nd VSK Kiln Shaft — First Biogas Brick Producer in Malawi · Doubles Capacity",
         Inches(0.5), Inches(6.82), Inches(12.4), Inches(0.55),
         size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════
# SLIDE 13 — WHY THE DISTRICT / WHY NEVAGS
# ═══════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
add_rect(slide, 0, 0, W, H, CREAM)
section_header(slide, "A Relationship of Mutual Benefit",
               "Why Mulanje Needs NEVAGS · Why NEVAGS Needs Mulanje", dark=True)

# Left: Why district needs NEVAGS
add_rect(slide, Inches(0.3), Inches(1.48), Inches(6.0), Inches(5.85), FOREST)
add_text(slide, "WHY MULANJE DISTRICT NEEDS NEVAGS",
         Inches(0.5), Inches(1.58), Inches(5.6), Inches(0.3),
         size=10, bold=True, color=ORANGE)
tx = slide.shapes.add_textbox(Inches(0.5), Inches(1.92), Inches(5.65), Inches(5.2))
tx.word_wrap = True; tf = tx.text_frame; tf.word_wrap = True
for title, body in [
    ("Policy Compliance",
     "NEVAGS is the only VSK producer in the district — the ONLY legal solution to the firewood ban. Every govt & NGO project needs us."),
    ("8,000+ Jobs",
     "The largest potential job creator in Mulanje history — from factory to community supply chain, all within 3 years."),
    ("Forest Protection",
     "Every NEVAGS brick protects Mulanje Massif's forests, water catchments, and biodiversity."),
    ("Black-Owned Pioneer",
     "One of the first black-owned industrial enterprises in Mulanje. Supporting NEVAGS sends a historic signal."),
    ("Tax & Local Economy",
     "Revenue stays in Mulanje — paying wages, buying local supplies, and contributing to district tax revenue."),
]:
    add_para(tf, title, 11, True, ORANGE2, space_before=8)
    add_para(tf, body, 9, False, RGBColor(0xCC,0xEE,0xCC), space_before=2)

# Right: Why NEVAGS needs district
add_rect(slide, Inches(6.5), Inches(1.48), Inches(6.53), Inches(5.85), CHARCOAL)
add_text(slide, "WHY NEVAGS NEEDS MULANJE DISTRICT",
         Inches(6.7), Inches(1.58), Inches(6.1), Inches(0.3),
         size=10, bold=True, color=ORANGE)
tx2 = slide.shapes.add_textbox(Inches(6.7), Inches(1.92), Inches(6.15), Inches(5.2))
tx2.word_wrap = True; tf2 = tx2.text_frame; tf2.word_wrap = True
for title, body in [
    ("DEC Endorsement",
     "Formal endorsement opens government procurement — millions of bricks in guaranteed long-term demand."),
    ("Workforce",
     "Mulanje's community is our talent pool. District cooperation enables the community training programme."),
    ("Biomass Supply",
     "Rice husks and farm biomass from Mulanje's farmers are our fuel. We need the agricultural sector as a partner."),
    ("Infrastructure",
     "Stable ESCOM power, road access, and land tenure security require district-level facilitation and support."),
    ("Policy Champion",
     "As a pioneering black enterprise, NEVAGS needs the DEC to champion, protect, and amplify our growth."),
]:
    add_para(tf2, title, 11, True, ORANGE2, space_before=8)
    add_para(tf2, body, 9, False, RGBColor(0xCC,0xCC,0xCC), space_before=2)


# ═══════════════════════════════════════════════════════════
# SLIDE 14 — THE ASK & CONTACT
# ═══════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
add_rect(slide, 0, 0, W, H, RGBColor(0x0A,0x18,0x0E))
add_rect(slide, 0, 0, Inches(0.45), H, ORANGE)

add_text(slide, "PARTNERSHIP INVITATION",
         Inches(0.65), Inches(0.2), Inches(12), Inches(0.35),
         size=10, bold=True, color=ORANGE)
add_text(slide, "Let's Build Malawi's\nSustainable Future Together",
         Inches(0.65), Inches(0.52), Inches(10), Inches(1.3),
         size=34, bold=True, color=WHITE)
add_text(slide, "NEVAGS asks the District Executive Committee for three things:",
         Inches(0.65), Inches(1.85), Inches(12), Inches(0.3),
         size=11, color=RGBColor(0xCC,0xDD,0xCC))

# 3 Asks
for i, (num, title, body) in enumerate([
    ("1", "Recognition & Endorsement",
     "Formally recognise NEVAGS as a strategic district partner — opening government procurement, NGO supply contracts, and public sector projects."),
    ("2", "Community Programme Support",
     "Facilitate community outreach, training endorsement, and biomass supply chain development through district channels."),
    ("3", "Infrastructure Facilitation",
     "Support stable power supply, road infrastructure, and land tenure security — foundational needs for NEVAGS to scale."),
]):
    x = Inches(0.65 + i * 4.2)
    add_rect(slide, x, Inches(2.2), Inches(3.9), Inches(2.4), RGBColor(0x15,0x35,0x20))
    add_text(slide, num, x+Inches(0.15), Inches(2.3), Inches(0.45), Inches(0.45),
             size=22, bold=True, color=ORANGE)
    add_text(slide, title, x+Inches(0.62), Inches(2.3), Inches(3.1), Inches(0.45),
             size=12, bold=True, color=WHITE)
    add_text(slide, body, x+Inches(0.15), Inches(2.78), Inches(3.65), Inches(1.72),
             size=9, color=RGBColor(0xBB,0xCC,0xBB))

# Contact cards
add_text(slide, "CONTACTS", Inches(0.65), Inches(4.75), Inches(12), Inches(0.28),
         size=9, bold=True, color=ORANGE)

for i, (role, name, lines) in enumerate([
    ("Managing Director",
     "Charles Billy Nasala",
     ["+265 888 34 75 75  |  +265 99 751 0160",
      "nasalacharles.b@gmail.com"]),
    ("Founding Engineer & BD Manager",
     "Chancy Tausi Tsonga",
     ["+265 984 000 366  |  WA: +27 764 998 4601",
      "chancy.tsonga@yahoo.com  |  chancytsonga.com"]),
    ("Careers & Location",
     "Join NEVAGS",
     ["careers.nevags@gmail.com",
      "Musewu, Mulanje District · P.O. Box 90"]),
]):
    x = Inches(0.65 + i * 4.2)
    add_rect(slide, x, Inches(5.05), Inches(3.9), Inches(2.3), RGBColor(0x1C,0x3A,0x25))
    add_text(slide, role, x+Inches(0.15), Inches(5.12), Inches(3.65), Inches(0.25),
             size=8, bold=True, color=ORANGE)
    add_text(slide, name, x+Inches(0.15), Inches(5.35), Inches(3.65), Inches(0.38),
             size=13, bold=True, color=WHITE)
    for j, line in enumerate(lines):
        add_text(slide, line, x+Inches(0.15), Inches(5.77 + j*0.34), Inches(3.65), Inches(0.32),
                 size=8.5, color=RGBColor(0xBB,0xDD,0xBB))

# Final tagline
add_rect(slide, Inches(0.65), Inches(7.12), Inches(12.08), Inches(0.35), ORANGE)
add_text(slide,
         "NEVAGS ECO BRICK & CONSTRUCTION  ·  Reg. 46289  ·  Musewu, Mulanje, Malawi  ·  Building Tomorrow Sustainably",
         Inches(0.8), Inches(7.16), Inches(11.8), Inches(0.28),
         size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════
# SAVE
# ═══════════════════════════════════════════════════════════
out = "NEVAGS_DEC_Presentation_2026.pptx"
prs.save(out)
print(f"Saved: {out}  ({prs.slides.__len__()} slides)")
