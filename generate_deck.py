from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches, Pt
import math
from datetime import date

COMPANY   = "Janatics India Private Limited"
CIN       = "U31103TZ1991PTC003409"
DATE_STR  = date.today().strftime("%d %B %Y")
COMPOSITE = 75.7

DARK_BLUE  = RGBColor(0x1F, 0x38, 0x64)
MID_BLUE   = RGBColor(0x2E, 0x75, 0xB6)
LIGHT_BLUE = RGBColor(0xD6, 0xE4, 0xF0)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
GREEN      = RGBColor(0x00, 0xB0, 0x50)
AMBER      = RGBColor(0xFF, 0xC0, 0x00)
RED_C      = RGBColor(0xFF, 0x00, 0x00)
DARK_GRAY  = RGBColor(0x40, 0x40, 0x40)

def score_color(s):
    if s >= 75: return GREEN
    if s >= 50: return AMBER
    return RED_C

def add_text_box(slide, text, left, top, width, height,
                 font_size=12, bold=False, color=None,
                 bg_color=None, align=PP_ALIGN.LEFT, italic=False):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf    = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size  = Pt(font_size)
    run.font.bold  = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = color
    if bg_color:
        txBox.fill.solid()
        txBox.fill.fore_color.rgb = bg_color
    return txBox

def add_rect(slide, left, top, width, height, fill_color, line_color=None):
    shape = slide.shapes.add_shape(1, left, top, width, height)  # MSO_SHAPE_TYPE.RECTANGLE = 1
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
    else:
        shape.line.fill.background()
    return shape

# ── Presentation setup ─────────────────────────────────────────────────────────

prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)

W = prs.slide_width
H = prs.slide_height

blank_layout = prs.slide_layouts[6]  # blank

# ── Slide 1: Title ─────────────────────────────────────────────────────────────

slide = prs.slides.add_slide(blank_layout)

# Background
add_rect(slide, 0, 0, W, H, DARK_BLUE)
add_rect(slide, 0, Inches(4.8), W, Inches(2.7), MID_BLUE)

add_text_box(slide, "DIGITAL DIAGNOSTIC AND MATURITY REPORT",
             Inches(0.8), Inches(1.0), Inches(11.7), Inches(0.8),
             font_size=28, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

add_text_box(slide, COMPANY,
             Inches(0.8), Inches(1.9), Inches(11.7), Inches(0.9),
             font_size=22, bold=True, color=LIGHT_BLUE, align=PP_ALIGN.CENTER)

add_text_box(slide, f"CIN: {CIN}",
             Inches(0.8), Inches(2.9), Inches(11.7), Inches(0.5),
             font_size=12, color=LIGHT_BLUE, align=PP_ALIGN.CENTER)

# Score callout
add_rect(slide, Inches(4.9), Inches(3.8), Inches(3.5), Inches(1.2), GREEN)
add_text_box(slide, f"{COMPOSITE}/100",
             Inches(4.9), Inches(3.8), Inches(3.5), Inches(0.7),
             font_size=32, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_text_box(slide, "COMPOSITE DDMR SCORE",
             Inches(4.9), Inches(4.5), Inches(3.5), Inches(0.4),
             font_size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

add_text_box(slide, f"Report Date: {DATE_STR}   |   Confidential",
             Inches(0.8), Inches(5.2), Inches(11.7), Inches(0.4),
             font_size=10, color=WHITE, align=PP_ALIGN.CENTER, italic=True)

# ── Slide 2: Company Snapshot ──────────────────────────────────────────────────

slide = prs.slides.add_slide(blank_layout)
add_rect(slide, 0, 0, W, Inches(1.1), DARK_BLUE)
add_text_box(slide, "Company Snapshot",
             Inches(0.3), Inches(0.15), Inches(12.7), Inches(0.8),
             font_size=20, bold=True, color=WHITE)

facts = [
    ("Founded",            "1977 (brand) / 1991 (incorporated)"),
    ("CIN",                "U31103TZ1991PTC003409"),
    ("Headquarters",       "Coimbatore, Tamil Nadu"),
    ("Additional Plants",  "Ahmedabad, Noida"),
    ("Status",             "Active | ROC Coimbatore"),
    ("Revenue (FY2025)",   "₹531 Crore"),
    ("Employees",          "662 (Aug 2025)"),
    ("Products",           "3,500+ SKUs"),
    ("Customers",          "17,000+"),
    ("Countries",          "42+"),
    ("Subsidiaries",       "USA, Germany"),
    ("Credit Rating",      "ICRA — Stable Outlook"),
    ("R&D",                "DSIR-recognized"),
    ("Industry",           "Pneumatics & Industrial Automation"),
]

cols = [facts[:7], facts[7:]]
for ci, col in enumerate(cols):
    for ri, (label, val) in enumerate(col):
        lx = Inches(0.4 + ci * 6.5)
        ty = Inches(1.3 + ri * 0.82)
        add_rect(slide, lx, ty, Inches(2.3), Inches(0.55), MID_BLUE)
        add_text_box(slide, label, lx + Inches(0.05), ty + Inches(0.05),
                     Inches(2.2), Inches(0.45), font_size=9, bold=True, color=WHITE)
        add_rect(slide, lx + Inches(2.35), ty, Inches(3.8), Inches(0.55), LIGHT_BLUE)
        add_text_box(slide, val, lx + Inches(2.4), ty + Inches(0.05),
                     Inches(3.7), Inches(0.45), font_size=10, color=DARK_BLUE)

# ── Slide 3: Scorecard Radar ───────────────────────────────────────────────────

slide = prs.slides.add_slide(blank_layout)
add_rect(slide, 0, 0, W, Inches(1.1), DARK_BLUE)
add_text_box(slide, "DDMR Dimension Scorecard",
             Inches(0.3), Inches(0.15), Inches(12.7), Inches(0.8),
             font_size=20, bold=True, color=WHITE)

# Radar chart
chart_data = ChartData()
chart_data.categories = [
    "Strategy &\nLeadership",
    "Sales &\nMarketing",
    "Operations &\nSupply Chain",
    "Finance",
    "Industry\nContext"
]
chart_data.add_series("Score", (81, 78, 78, 67, 75))
chart_data.add_series("Benchmark (75)", (75, 75, 75, 75, 75))

chart = slide.shapes.add_chart(
    XL_CHART_TYPE.RADAR_FILLED,
    Inches(0.5), Inches(1.2), Inches(7.5), Inches(5.8),
    chart_data
).chart
chart.has_legend = True

# Score boxes on right
dims = [
    ("SL", "Strategy & Leadership",     81),
    ("SM", "Sales & Marketing",         78),
    ("OSC","Operations & Supply Chain", 78),
    ("FIN","Finance",                   67),
    ("IC", "Industry Context",          75),
]
for i, (code, name, score) in enumerate(dims):
    ty = Inches(1.4 + i * 1.1)
    add_rect(slide, Inches(8.3), ty, Inches(0.6), Inches(0.7), score_color(score))
    add_text_box(slide, str(score), Inches(8.3), ty, Inches(0.6), Inches(0.7),
                 font_size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_rect(slide, Inches(8.95), ty, Inches(4.0), Inches(0.7), LIGHT_BLUE)
    add_text_box(slide, f"{code}: {name}", Inches(9.0), ty + Inches(0.08),
                 Inches(3.9), Inches(0.55), font_size=10, color=DARK_BLUE)

add_rect(slide, Inches(8.3), Inches(7.0), Inches(4.65), Inches(0.45), MID_BLUE)
add_text_box(slide, f"Composite Score: {COMPOSITE}/100  —  Strong",
             Inches(8.3), Inches(7.0), Inches(4.65), Inches(0.45),
             font_size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

# ── Slides 4–8: One per dimension ─────────────────────────────────────────────

dimension_slides = [
    {
        "code": "SL", "name": "Strategy & Leadership", "score": 81, "weight": "25%",
        "bullets": [
            "Founded 1977 — 49-year operational track record, incorporated 1991",
            "3 active directors | AGM Sept 2024 | ROC Coimbatore filings current",
            "Global footprint: 42+ countries | Subsidiaries: USA and Germany",
            "DSIR-recognized R&D division | 200 global distribution partners",
            "Strategic pivot: electric actuators, robotics, Industry 4.0 products",
            "ICRA-rated debt profile | Stable outlook | Healthy debt protection",
        ],
        "sub_scores": [("SL01 Tenure",90),("SL02 Directors",75),("SL03 Governance",80),
                       ("SL04 Debt",78),("SL05 Strategy",82),("SL06 Global",88)],
        "highlight": "49-year brand | Global in 42 countries",
        "risk": "No public ESG / independent director disclosures",
    },
    {
        "code": "SM", "name": "Sales & Marketing", "score": 78, "weight": "20%",
        "bullets": [
            "FY2025 Revenue: ₹531 Crore (₹507.9 Cr FY2024 → 4% YoY growth)",
            "Industry CAGR: 6.9% — Janatics growing slightly below sector pace",
            "17,000+ customers across 7 industries — low concentration risk",
            "200 global distribution partners | 3 manufacturing plants nationally",
            "Own-brand (Janatics) | AIA India member | 49-year brand equity",
            "Digital commerce channel not yet activated — potential growth lever",
        ],
        "sub_scores": [("SM01 Rev Growth",55),("SM02 Scale",80),("SM03 Customers",90),
                       ("SM04 Distribution",85),("SM05 Brand",80)],
        "highlight": "17,000+ customers | 7 industries | Global distribution",
        "risk": "Revenue growth 4% vs industry CAGR 6.9%",
    },
    {
        "code": "OSC", "name": "Operations & Supply Chain", "score": 78, "weight": "20%",
        "bullets": [
            "70,000 sqm manufacturing facility in Coimbatore | Ahmedabad + Noida plants",
            "3,500+ distinct SKUs — single-source supplier capability for OEMs",
            "Advanced CNC machining, casting, surface treatment infrastructure",
            "DSIR-recognized R&D — proprietary product development capability",
            "Expanding into electric actuators, sensors, robotic systems",
            "Internal digitization (ERP, MES) not publicly disclosed",
        ],
        "sub_scores": [("OSC01 Scale",85),("OSC02 Products",90),("OSC03 R&D",82),
                       ("OSC04 GST",72),("OSC05 Digital",65)],
        "highlight": "3,500+ SKUs | 70,000 sqm plant | DSIR R&D",
        "risk": "Supply chain digitization level undisclosed",
    },
    {
        "code": "FIN", "name": "Finance", "score": 67, "weight": "25%",
        "bullets": [
            "FY2025 Revenue: ₹531 Cr | Employee productivity: ₹80L per employee",
            "ICRA-rated | Stable outlook | Multiple reaffirmations (2021–2025)",
            "Paid-up capital ₹26.5 Cr | Authorized ₹30 Cr | Self-funded growth",
            "Revenue growth 4% YoY — steady but below inflation-adjusted peer level",
            "EBITDA CAGR -22% (1-yr, Tracxn) — margin compression flagged",
            "Full P&L not public (unlisted Pvt Ltd) — ICRA rationale is key proxy",
        ],
        "sub_scores": [("FIN01 Revenue",80),("FIN02 Growth",55),("FIN03 EBITDA",42),
                       ("FIN04 Credit",82),("FIN05 Capital",75),("FIN06 Productivity",70)],
        "highlight": "ICRA Stable | ₹531 Cr revenue | Self-funded",
        "risk": "EBITDA CAGR -22% (FY25) — needs full P&L for confirmation",
    },
    {
        "code": "IC", "name": "Industry Context", "score": 75, "weight": "10%",
        "bullets": [
            "India pneumatics market: $1.06B (2024) → $1.94B (2033) | CAGR 6.9%",
            "Make in India + PLI schemes driving domestic manufacturing investment",
            "China+1 sourcing strategy creating incremental export demand",
            "Janatics competes with SMC, Festo, Parker in mid-market segment",
            "End-market spread: pharma, auto, packaging, food, medical, textile, printing",
            "Industry 4.0 / smart pneumatics trend aligns with Janatics product roadmap",
        ],
        "sub_scores": [("IC01 Mkt Size",70),("IC02 Growth",78),("IC03 Position",70),
                       ("IC04 Macro",82),("IC05 Diversity",76)],
        "highlight": "6.9% market CAGR | Make in India tailwind | 7 end-markets",
        "risk": "Global competition intensifying — SMC, Festo localizing",
    },
]

for d in dimension_slides:
    slide = prs.slides.add_slide(blank_layout)

    # Header bar
    add_rect(slide, 0, 0, W, Inches(1.1), score_color(d["score"]))
    add_text_box(slide, f"{d['code']}: {d['name']}",
                 Inches(0.3), Inches(0.1), Inches(9.0), Inches(0.55),
                 font_size=20, bold=True, color=WHITE)
    add_text_box(slide, f"Score: {d['score']}/100   |   Weight: {d['weight']}",
                 Inches(0.3), Inches(0.6), Inches(9.0), Inches(0.4),
                 font_size=12, color=WHITE)

    # Score badge
    add_rect(slide, Inches(10.8), Inches(0.1), Inches(2.2), Inches(0.9), DARK_BLUE)
    add_text_box(slide, f"{d['score']}/100",
                 Inches(10.8), Inches(0.1), Inches(2.2), Inches(0.9),
                 font_size=24, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # Bullets
    for bi, bullet in enumerate(d["bullets"]):
        by = Inches(1.25 + bi * 0.87)
        add_rect(slide, Inches(0.25), by + Inches(0.15), Inches(0.08), Inches(0.22),
                 score_color(d["score"]))
        add_text_box(slide, bullet,
                     Inches(0.45), by, Inches(8.1), Inches(0.75),
                     font_size=11, color=DARK_GRAY)

    # Sub-scores panel
    panel_x = Inches(8.8)
    add_rect(slide, panel_x, Inches(1.1), Inches(4.3), Inches(0.4), MID_BLUE)
    add_text_box(slide, "Sub-Metric Scores",
                 panel_x, Inches(1.1), Inches(4.3), Inches(0.4),
                 font_size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    bar_h = Inches(0.52)
    for si, (sm_name, sm_score) in enumerate(d["sub_scores"]):
        sy = Inches(1.55 + si * 0.65)
        # Label
        add_text_box(slide, sm_name, panel_x, sy, Inches(2.0), bar_h,
                     font_size=9, color=DARK_BLUE)
        # Background bar
        add_rect(slide, panel_x + Inches(2.05), sy + Inches(0.12),
                 Inches(1.8), Inches(0.28), LIGHT_BLUE)
        # Filled bar
        bar_w = Inches(1.8 * sm_score / 100)
        add_rect(slide, panel_x + Inches(2.05), sy + Inches(0.12),
                 bar_w, Inches(0.28), score_color(sm_score))
        # Score label
        add_text_box(slide, str(sm_score),
                     panel_x + Inches(3.9), sy, Inches(0.4), bar_h,
                     font_size=9, bold=True, color=score_color(sm_score))

    # Highlight + Risk footer
    add_rect(slide, 0, Inches(6.85), W * 0.6, Inches(0.55), GREEN)
    add_text_box(slide, f"Key Strength: {d['highlight']}",
                 Inches(0.15), Inches(6.87), Inches(7.8), Inches(0.45),
                 font_size=9, bold=True, color=WHITE)
    add_rect(slide, W * 0.6, Inches(6.85), W * 0.4, Inches(0.55), RED_C)
    add_text_box(slide, f"Risk: {d['risk']}",
                 W * 0.6 + Inches(0.1), Inches(6.87), Inches(5.1), Inches(0.45),
                 font_size=9, bold=True, color=WHITE)

# ── Slide 9: Recommendation ────────────────────────────────────────────────────

slide = prs.slides.add_slide(blank_layout)
add_rect(slide, 0, 0, W, Inches(1.1), DARK_BLUE)
add_text_box(slide, "Executive Recommendation",
             Inches(0.3), Inches(0.15), Inches(12.7), Inches(0.8),
             font_size=20, bold=True, color=WHITE)

add_rect(slide, Inches(0.4), Inches(1.2), Inches(12.5), Inches(0.7), GREEN)
add_text_box(slide, f"DDMR Composite Score: {COMPOSITE}/100 — STRONG  |  Recommendation: Proceed with engagement",
             Inches(0.4), Inches(1.2), Inches(12.5), Inches(0.7),
             font_size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

rec_points = [
    ("Strengths to leverage",
     "49-year brand with ICRA-validated credit health. Global distribution in 42 countries "
     "and 17,000+ customers provides revenue resilience. DSIR R&D and 3,500+ SKU portfolio "
     "enable single-source supplier status with OEMs."),
    ("Watch points",
     "EBITDA CAGR of -22% (FY25) requires explanation — request full P&L and management "
     "commentary on margin drivers. Revenue growth at 4% trails the 6.9% sector CAGR."),
    ("Immediate due diligence asks",
     "1. Full audited P&L + balance sheet for FY2023–FY2025. "
     "2. Customer concentration report (top 10 customers as % of revenue). "
     "3. Director disclosure and related-party transactions. "
     "4. Status of expansion into electric actuators / robotics — capex plan."),
    ("Opportunity",
     "Make in India + China+1 export tailwind creates a 3–5 year window for Janatics to "
     "gain global OEM approvals. A partner with procurement or market access in Europe/Americas "
     "could accelerate this significantly."),
]

for ri, (heading, body) in enumerate(rec_points):
    ty = Inches(2.1 + ri * 1.3)
    add_rect(slide, Inches(0.4), ty, Inches(2.8), Inches(1.1), MID_BLUE)
    add_text_box(slide, heading, Inches(0.45), ty + Inches(0.1),
                 Inches(2.7), Inches(0.9), font_size=10, bold=True, color=WHITE)
    add_rect(slide, Inches(3.3), ty, Inches(9.6), Inches(1.1), LIGHT_BLUE)
    add_text_box(slide, body, Inches(3.4), ty + Inches(0.05),
                 Inches(9.4), Inches(1.0), font_size=9, color=DARK_BLUE)

add_text_box(slide, f"Report generated: {DATE_STR}  |  Janatics India Private Limited  |  DDMR v1.0",
             Inches(0.4), Inches(7.1), Inches(12.5), Inches(0.3),
             font_size=8, color=DARK_GRAY, align=PP_ALIGN.CENTER, italic=True)

# ── Save ───────────────────────────────────────────────────────────────────────

out = "/Users/laaksandy/Documents/Claude-testing/DDMR/output/Janatics_DDMR_Deck.pptx"
prs.save(out)
print(f"Saved: {out}")
