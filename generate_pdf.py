from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, HRFlowable, KeepTogether,
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm, mm
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus.flowables import Flowable
from datetime import date
import os

# ── Data ──────────────────────────────────────────────────────────────────────

COMPANY   = "Janatics India Private Limited"
CIN       = "U31103TZ1991PTC003409"
DATE_STR  = date.today().strftime("%d %B %Y")
COMPOSITE = 64

DARK_BLUE  = colors.HexColor("#1F3864")
MID_BLUE   = colors.HexColor("#2E75B6")
LIGHT_BLUE = colors.HexColor("#D6E4F0")
GREEN      = colors.HexColor("#00B050")
AMBER      = colors.HexColor("#FFC000")
RED_       = colors.HexColor("#FF0000")
BLUE_FR    = colors.HexColor("#2563EB")
LIGHT_GREEN= colors.HexColor("#D5E8D4")
LIGHT_AMBER= colors.HexColor("#FFF2CC")
LIGHT_RED  = colors.HexColor("#F8CECC")
LIGHT_BLUE2= colors.HexColor("#DAE8FC")
WHITE      = colors.white
BLACK      = colors.black
GRAY       = colors.HexColor("#555555")
LIGHT_GRAY = colors.HexColor("#F8F9FA")

def maturity_label(s):
    if s >= 84: return "Future Ready"
    if s >= 67: return "Strategic"
    if s >= 34: return "Siloed"
    return "Legacy"

def maturity_color(s):
    if s >= 84: return BLUE_FR
    if s >= 67: return GREEN
    if s >= 34: return AMBER
    return RED_

def maturity_light(s):
    if s >= 84: return colors.HexColor("#DBEAFE")
    if s >= 67: return LIGHT_GREEN
    if s >= 34: return LIGHT_AMBER
    return LIGHT_RED

dimensions = [
    {
        "code": "ST", "name": "Strategy", "score": 70,
        "sub_metrics": [
            ("ST01", "Digital Strategy Articulation", 65, "No formal DX roadmap publicly disclosed; I4.0 product lines signal intent"),
            ("ST02", "Leadership Technology Vision", 72, "Product pivot to electric actuators, robotics, sensors signals digital awareness at leadership level"),
            ("ST03", "Digital Investment Signals", 75, "New IoT-enabled product lines indicate active technology capex commitment"),
            ("ST04", "Technology Partnerships", 68, "DSIR R&D recognition confirmed; no public digital alliance or ecosystem partnerships"),
            ("ST05", "Innovation Pipeline", 70, "Incremental product digitization active; no IP trail, open innovation, or startup signals"),
        ],
        "narrative": (
            "Janatics India's digital strategy is visible primarily through product portfolio decisions. The pivot "
            "toward electric actuators, smart sensors, robotic systems, and Industry 4.0-compatible equipment "
            "demonstrates leadership awareness of the digital megatrend. The DSIR-recognized R&D division "
            "provides structural support for technology investment continuity. However, no formal digital "
            "transformation roadmap, CDO appointment, or publicly articulated DX strategy has been evidenced. "
            "Technology partnerships beyond DSIR are not disclosed. Score: 70/100 — Siloed."
        ),
    },
    {
        "code": "OSC", "name": "Operations & Supply Chain", "score": 65,
        "sub_metrics": [
            ("OSC01", "Digital Manufacturing Readiness", 60, "70,000 sqm Coimbatore plant with CNC; no I4.0 cert or smart factory disclosure"),
            ("OSC02", "Supply Chain Visibility Tools", 50, "No ERP, SCM platform, or supply chain digitization initiative disclosed publicly"),
            ("OSC03", "Process Automation Adoption", 70, "CNC machining, casting automation, surface treatment lines — physical automation established"),
            ("OSC04", "Quality Management (Digital)", 72, "ICRA notes consistent compliance; ISO quality system assumed but not publicly confirmed"),
            ("OSC05", "Industry 4.0 Integration", 73, "I4.0 product portfolio exists; internal adoption for own manufacturing not disclosed"),
        ],
        "narrative": (
            "Janatics operates a world-class physical manufacturing base — a 70,000 sqm Coimbatore plant with "
            "CNC machining and casting, plus Ahmedabad and Noida plants. Physical process automation is well "
            "established. However, the digital operations layer is a significant gap: no ERP, MES, or supply "
            "chain visibility platform disclosures have been found. The company's I4.0 product expertise is not "
            "evidenced internally. Score: 65/100 — Siloed."
        ),
    },
    {
        "code": "SM", "name": "Sales & Marketing", "score": 55,
        "sub_metrics": [
            ("SM01", "Digital Presence & SEO", 60, "Two product websites with online catalog; SEO signals moderate"),
            ("SM02", "Social Media Engagement", 40, "Limited LinkedIn activity; no active social media marketing evidenced"),
            ("SM03", "E-commerce / Digital Ordering", 45, "No B2B e-commerce, digital ordering portal, or online pricing catalog identified"),
            ("SM04", "Digital Marketing Activity", 55, "Websites established; no paid digital campaigns or content marketing evidenced"),
            ("SM05", "CRM & Customer Data Utilization", 75, "17,000+ customers across 7 industries — implies structured data management; CRM tool undisclosed"),
        ],
        "narrative": (
            "Janatics' digital sales and marketing footprint is the lowest-scoring dimension, reflecting a "
            "traditional go-to-market model via distributors and OEM relationships. No B2B e-commerce, digital "
            "ordering portal, or real-time pricing has been evidenced. Social media engagement is minimal. "
            "The positive: 17,000+ customers managed across 7 industries implies structured data, though "
            "the CRM tool is undisclosed. Score: 55/100 — Siloed."
        ),
    },
    {
        "code": "TA", "name": "Technology Adoption", "score": 70,
        "sub_metrics": [
            ("TA01", "Website & Digital Infrastructure", 65, "Functional product websites; no modern tech stack, API integrations, or customer portal"),
            ("TA02", "IoT / Smart Product Portfolio", 80, "Electric actuators, smart sensors, robotics, I4.0 didactics — strong product-layer digitization"),
            ("TA03", "Cloud & SaaS Adoption", 60, "No public cloud migration, SaaS tools, or digital collaboration platform disclosures"),
            ("TA04", "R&D in Digital / Smart Tech", 78, "DSIR-recognized R&D developing robotic systems, sensors, and smart automation equipment"),
            ("TA05", "Cybersecurity & Data Practices", 67, "ICRA-compliant; no public cybersecurity framework, ISO 27001, or data policy disclosed"),
        ],
        "narrative": (
            "Technology Adoption is the strongest dimension alongside Strategy, driven by Janatics' product-layer "
            "digitization. The DSIR R&D division actively develops robotic systems, smart sensors, and I4.0 "
            "equipment. Internal technology infrastructure (cloud, SaaS, cybersecurity) presents a mixed picture "
            "— functional websites without modern architecture; no cloud migration evidenced. Score: 70/100 — Siloed."
        ),
    },
    {
        "code": "SC", "name": "Skills & Capabilities", "score": 60,
        "sub_metrics": [
            ("SC01", "Digital Talent Hiring Signals", 55, "662 employees primarily engineering/manufacturing; limited software or data role hiring signals"),
            ("SC02", "Technical Workforce Depth", 72, "Strong mechanical engineering base; DSIR R&D team; CNC and automation specialists"),
            ("SC03", "Digital Training Programs", 50, "No public L&D investment, digital upskilling programs, or online learning adoption disclosed"),
            ("SC04", "Leadership Digital Literacy", 65, "Management I4.0 awareness via product decisions; no digital leadership hire signals"),
            ("SC05", "Innovation Culture Signals", 58, "DSIR R&D present; no hackathon, innovation lab, startup engagement, or employee innovation program"),
        ],
        "narrative": (
            "Janatics' workforce of 662 employees is engineering-deep with strong mechanical and manufacturing "
            "expertise. The DSIR R&D team enables product innovation. However, digital talent acquisition signals "
            "are limited — no software engineers, data scientists, or digital roles are publicly evidenced. "
            "No L&D, upskilling, or innovation culture programs are disclosed. Score: 60/100 — Siloed."
        ),
    },
]

next_actions = [
    {
        "rank": "01 — HIGH PRIORITY",
        "title": "Activate a B2B Digital Sales Channel",
        "dimension": "Sales & Marketing · SM03: 45/100",
        "initiative": "Launch a B2B digital ordering portal with searchable product catalog, real-time pricing, and quote-request workflows. Pair with SEO content and LinkedIn demand generation.",
        "benefit": "Estimated 15–20% incremental revenue from direct digital channel within 18 months. Reduced cost-per-lead vs. distributor-led acquisition. Improved customer data visibility.",
        "loss": "SMC Corporation, Festo, and Parker Hannifin all operate mature digital ordering platforms. Every quarter without a digital channel cedes online discovery to global competitors.",
        "why_now": "India's B2B e-commerce in industrial components is growing 25%+ annually. Janatics has a unique window — 49-year brand + 17,000 customers + 3,500+ SKUs — to win digital-first industrial buyers.",
    },
    {
        "rank": "02 — HIGH PRIORITY",
        "title": "Deploy ERP and Supply Chain Visibility Platform",
        "dimension": "Operations & Supply Chain · OSC02: 50/100",
        "initiative": "Implement an integrated ERP (SAP/Oracle/Odoo) connecting production planning, procurement, inventory, and dispatch across all three plants. Add supply chain visibility for key customers.",
        "benefit": "10–15% reduction in production inefficiencies; faster order-to-delivery; eligibility for global OEM vendor qualification audits requiring digital supply chain traceability.",
        "loss": "Global OEM customers increasingly mandate supply chain transparency as a vendor qualification criterion. Without an ERP-backed system, Janatics risks disqualification from high-value export RFQs.",
        "why_now": "Make in India + China+1 sourcing shifts are generating an influx of global OEM inquiries. Janatics has a 12–18 month window to establish digital operational credibility before qualification pipelines close.",
    },
    {
        "rank": "03 — MEDIUM PRIORITY",
        "title": "Build a Digital Workforce Capability Program",
        "dimension": "Skills & Capabilities · SC03: 50/100 · SC01: 55/100",
        "initiative": "Launch a 12-month structured digital upskilling program. Hire 2–3 digital roles (data analyst, digital marketing manager, ERP lead). Partner with an L&D platform.",
        "benefit": "Accelerates adoption velocity of all other digital investments. Reduces transformation failure risk. Builds internal champions who sustain change without consultant dependency.",
        "loss": "Technology without people is stranded investment. Every rupee spent on ERP, e-commerce, or analytics will underperform if the workforce lacks confidence and capability to use it.",
        "why_now": "India's digital skills gap in manufacturing is widening. Early movers gain a 2–3 year head start on talent acquisition and internal capability building.",
    },
    {
        "rank": "04 — MEDIUM PRIORITY",
        "title": "Convert Coimbatore Plant into a Smart Factory Showcase",
        "dimension": "Technology Adoption · TA01: 65/100 · Strategy · ST03: 75/100",
        "initiative": "Deploy Janatics' own I4.0 sensors, electric actuators, and robotic systems on the Coimbatore production floor as a live demonstration environment. Create a customer visit programme and virtual tour capability.",
        "benefit": "Validates product performance with proof-by-example. Accelerates B2B sales cycles. Generates authentic case study content. Qualifies Janatics for smart manufacturing certification programmes.",
        "loss": "Selling Industry 4.0 products while operating a non-digitized factory is a credibility contradiction that sophisticated buyers detect during site visits and RFQ evaluations.",
        "why_now": "Janatics already owns the technology — internal deployment requires no new vendor. This is the fastest credibility-building action available and aligns with PLI scheme reporting.",
    },
]

# ── Styles ────────────────────────────────────────────────────────────────────

def make_styles():
    base = getSampleStyleSheet()
    return {
        "cover_title": ParagraphStyle("cover_title", parent=base["Normal"],
            fontSize=22, fontName="Helvetica-Bold", textColor=DARK_BLUE,
            alignment=TA_CENTER, spaceAfter=8),
        "cover_company": ParagraphStyle("cover_company", parent=base["Normal"],
            fontSize=16, fontName="Helvetica-Bold", textColor=MID_BLUE,
            alignment=TA_CENTER, spaceAfter=6),
        "cover_meta": ParagraphStyle("cover_meta", parent=base["Normal"],
            fontSize=10, textColor=GRAY, alignment=TA_CENTER, spaceAfter=4),
        "cover_score": ParagraphStyle("cover_score", parent=base["Normal"],
            fontSize=14, fontName="Helvetica-Bold", textColor=maturity_color(COMPOSITE),
            alignment=TA_CENTER, spaceAfter=4),
        "h1": ParagraphStyle("h1", parent=base["Normal"],
            fontSize=14, fontName="Helvetica-Bold", textColor=DARK_BLUE,
            spaceBefore=14, spaceAfter=6),
        "h2": ParagraphStyle("h2", parent=base["Normal"],
            fontSize=12, fontName="Helvetica-Bold", textColor=MID_BLUE,
            spaceBefore=10, spaceAfter=4),
        "h3": ParagraphStyle("h3", parent=base["Normal"],
            fontSize=10, fontName="Helvetica-Bold", textColor=DARK_BLUE,
            spaceBefore=6, spaceAfter=3),
        "body": ParagraphStyle("body", parent=base["Normal"],
            fontSize=9, textColor=GRAY, leading=14, spaceAfter=6),
        "label": ParagraphStyle("label", parent=base["Normal"],
            fontSize=8, fontName="Helvetica-Bold", textColor=GRAY, spaceAfter=2),
        "small": ParagraphStyle("small", parent=base["Normal"],
            fontSize=8, textColor=GRAY, leading=12),
        "bullet": ParagraphStyle("bullet", parent=base["Normal"],
            fontSize=9, textColor=GRAY, leading=14, leftIndent=12,
            bulletIndent=4, spaceAfter=3),
    }

# ── Cover page ─────────────────────────────────────────────────────────────────

class ColorRect(Flowable):
    def __init__(self, width, height, color, radius=4):
        super().__init__()
        self.width  = width
        self.height = height
        self.color  = color
        self.radius = radius
    def draw(self):
        self.canv.setFillColor(self.color)
        self.canv.roundRect(0, 0, self.width, self.height, self.radius, fill=1, stroke=0)

# ── Build PDF ─────────────────────────────────────────────────────────────────

def build():
    out  = "/Users/laaksandy/Documents/Claude-testing/DDMR/output/Janatics_DDMR_Report.pdf"
    doc  = SimpleDocTemplate(
        out,
        pagesize=A4,
        leftMargin=2*cm, rightMargin=2*cm,
        topMargin=2*cm,  bottomMargin=2*cm,
        title=f"DDMR — {COMPANY}",
        author="DDMR Tool",
    )
    W = doc.width  # usable width
    S = make_styles()
    story = []

    # ── Cover ──────────────────────────────────────────────────────────────────
    story.append(Spacer(1, 1.5*cm))
    story.append(Paragraph("DIGITAL DIAGNOSTIC AND MATURITY REPORT", S["cover_title"]))
    story.append(Paragraph("(DDMR)", S["cover_meta"]))
    story.append(Spacer(1, 0.3*cm))
    story.append(HRFlowable(width=W, color=DARK_BLUE, thickness=2))
    story.append(Spacer(1, 0.4*cm))
    story.append(Paragraph(COMPANY, S["cover_company"]))
    story.append(Paragraph(f"CIN: {CIN}   |   Report Date: {DATE_STR}", S["cover_meta"]))
    story.append(Spacer(1, 0.5*cm))

    # Score badge table
    score_table = Table(
        [[Paragraph(f"{COMPOSITE}/100", ParagraphStyle("sc", fontSize=32,
            fontName="Helvetica-Bold", textColor=WHITE, alignment=TA_CENTER)),
          Paragraph(f"Composite DDMR Score\n{maturity_label(COMPOSITE)} Maturity Band",
            ParagraphStyle("sl", fontSize=12, fontName="Helvetica-Bold",
            textColor=WHITE, alignment=TA_LEFT, leading=18))]],
        colWidths=[5*cm, W-5*cm],
    )
    score_table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), maturity_color(COMPOSITE)),
        ("VALIGN",     (0,0), (-1,-1), "MIDDLE"),
        ("LEFTPADDING",(0,0), (-1,-1), 12),
        ("RIGHTPADDING",(0,0),(-1,-1),12),
        ("TOPPADDING", (0,0), (-1,-1), 14),
        ("BOTTOMPADDING",(0,0),(-1,-1),14),
        ("ROUNDEDCORNERS", (0,0), (-1,-1), [6, 6, 6, 6]),
    ]))
    story.append(score_table)
    story.append(Spacer(1, 0.5*cm))

    # Maturity scale bar
    scale_data = [["Legacy\n0–33", "Siloed\n34–66", "Strategic\n67–83", "Future Ready\n84–100"]]
    scale_table = Table(scale_data, colWidths=[W/4]*4)
    scale_table.setStyle(TableStyle([
        ("BACKGROUND", (0,0),(0,0), RED_),
        ("BACKGROUND", (1,0),(1,0), AMBER),
        ("BACKGROUND", (2,0),(2,0), GREEN),
        ("BACKGROUND", (3,0),(3,0), BLUE_FR),
        ("FONTNAME", (0,0),(-1,-1), "Helvetica-Bold"),
        ("FONTSIZE", (0,0),(-1,-1), 8),
        ("TEXTCOLOR",(0,0),(-1,-1), WHITE),
        ("ALIGN", (0,0),(-1,-1), "CENTER"),
        ("VALIGN", (0,0),(-1,-1), "MIDDLE"),
        ("TOPPADDING",(0,0),(-1,-1), 6),
        ("BOTTOMPADDING",(0,0),(-1,-1), 6),
    ]))
    story.append(scale_table)
    story.append(Spacer(1, 0.6*cm))

    # Company meta grid
    meta_rows = [
        [Paragraph("<b>Founded</b>", S["small"]), Paragraph("1977 (brand) / 1991 (incorp.)", S["small"]),
         Paragraph("<b>HQ</b>", S["small"]), Paragraph("Coimbatore, Tamil Nadu", S["small"])],
        [Paragraph("<b>Revenue</b>", S["small"]), Paragraph("₹531 Cr (FY2025)", S["small"]),
         Paragraph("<b>Employees</b>", S["small"]), Paragraph("662", S["small"])],
        [Paragraph("<b>Countries</b>", S["small"]), Paragraph("42+", S["small"]),
         Paragraph("<b>Credit Rating</b>", S["small"]), Paragraph("ICRA — Stable", S["small"])],
        [Paragraph("<b>Products</b>", S["small"]), Paragraph("3,500+ SKUs", S["small"]),
         Paragraph("<b>Customers</b>", S["small"]), Paragraph("17,000+", S["small"])],
    ]
    meta_table = Table(meta_rows, colWidths=[2.5*cm, 6*cm, 2.5*cm, W-11*cm])
    meta_table.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,-1), LIGHT_GRAY),
        ("GRID",      (0,0),(-1,-1), 0.5, colors.HexColor("#E5E7EB")),
        ("TOPPADDING",(0,0),(-1,-1), 5),
        ("BOTTOMPADDING",(0,0),(-1,-1),5),
        ("LEFTPADDING",(0,0),(-1,-1), 6),
    ]))
    story.append(meta_table)
    story.append(PageBreak())

    # ── Section 1: Executive Summary ───────────────────────────────────────────
    story.append(Paragraph("1. Executive Summary", S["h1"]))
    story.append(Paragraph(
        f"{COMPANY} (CIN: {CIN}) is assessed for digital transformation maturity using five "
        "dimensions derived exclusively from public signals: Strategy, Operations &amp; Supply Chain, "
        "Sales &amp; Marketing, Technology Adoption, and Skills &amp; Capabilities. The company is a "
        "Coimbatore-based manufacturer of pneumatic and industrial automation products, incorporated "
        "in 1991 and operating since 1977, with ₹531 Crore FY2025 revenue, ICRA Stable credit, "
        "and export presence in 42+ countries.", S["body"]))
    story.append(Paragraph(
        f"<b>DDMR Composite Score: <font color='#{hex(int(maturity_color(COMPOSITE).hexval(), 16))[2:]}' >"
        f"{COMPOSITE}/100 — {maturity_label(COMPOSITE)}</font></b>. "
        "Digital intent is visible in Janatics' product strategy — the pivot to electric actuators, "
        "robotics, smart sensors, and Industry 4.0 equipment signals leadership awareness of digital "
        "megatrends. However, systematic digitization of internal operations, sales channels, and "
        "workforce development remains nascent based on public evidence.", S["body"]))

    story.append(Spacer(1, 0.2*cm))

    # Dimension summary table
    hdr  = ["Dimension", "Score /100", "Weight", "Maturity Level"]
    rows = [hdr] + [
        [d["name"], str(d["score"]), "20%", maturity_label(d["score"])]
        for d in dimensions
    ] + [[f"COMPOSITE", str(COMPOSITE), "100%", maturity_label(COMPOSITE)]]

    col_w = [W*0.38, W*0.18, W*0.16, W*0.28]
    dim_table = Table(rows, colWidths=col_w)
    ts = [
        ("BACKGROUND", (0,0),(-1,0), DARK_BLUE),
        ("TEXTCOLOR",  (0,0),(-1,0), WHITE),
        ("FONTNAME",   (0,0),(-1,0), "Helvetica-Bold"),
        ("FONTNAME",   (0,-1),(-1,-1), "Helvetica-Bold"),
        ("BACKGROUND", (0,-1),(-1,-1), MID_BLUE),
        ("TEXTCOLOR",  (0,-1),(-1,-1), WHITE),
        ("GRID",       (0,0),(-1,-1), 0.5, colors.HexColor("#CCCCCC")),
        ("FONTSIZE",   (0,0),(-1,-1), 9),
        ("ALIGN",      (1,0),(-1,-1), "CENTER"),
        ("VALIGN",     (0,0),(-1,-1), "MIDDLE"),
        ("TOPPADDING", (0,0),(-1,-1), 5),
        ("BOTTOMPADDING",(0,0),(-1,-1), 5),
        ("LEFTPADDING",(0,0),(-1,-1), 6),
    ]
    for i, d in enumerate(dimensions, start=1):
        bg = maturity_light(d["score"])
        ts += [("BACKGROUND", (0,i),(-1,i), bg)]
    dim_table.setStyle(TableStyle(ts))
    story.append(dim_table)
    story.append(PageBreak())

    # ── Maturity scale ─────────────────────────────────────────────────────────
    story.append(Paragraph("Maturity Scale Reference", S["h2"]))
    story.append(Paragraph(
        "<b>Legacy (0–33)</b> — Minimal digital presence; paper-based or siloed analogue processes. "
        "<b>Siloed (34–66)</b> — Digital tools adopted in isolation; limited cross-functional integration. "
        "<b>Strategic (67–83)</b> — Deliberate digital strategy with integration across functions. "
        "<b>Future Ready (84–100)</b> — Continuous digital innovation; data-driven; platform-enabled.", S["body"]))

    # ── Sections 2–6: Dimensions ───────────────────────────────────────────────
    section_nums = {0:2, 1:3, 2:4, 3:5, 4:6}
    for idx, d in enumerate(dimensions):
        story.append(Paragraph(
            f"{section_nums[idx]}. {d['name']} — Score: {d['score']}/100 — {maturity_label(d['score'])}",
            S["h1"]))

        # Score + narrative side by side
        score_cell = Table([[
            Paragraph(str(d["score"]), ParagraphStyle("ds", fontSize=28,
                fontName="Helvetica-Bold", textColor=WHITE, alignment=TA_CENTER)),
            Paragraph(f"/100<br/><b>{maturity_label(d['score'])}</b>",
                ParagraphStyle("dl", fontSize=9, textColor=WHITE,
                alignment=TA_LEFT, leading=13))
        ]], colWidths=[2*cm, 2*cm])
        score_cell.setStyle(TableStyle([
            ("BACKGROUND", (0,0),(-1,-1), maturity_color(d["score"])),
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
            ("TOPPADDING",(0,0),(-1,-1),10),
            ("BOTTOMPADDING",(0,0),(-1,-1),10),
            ("LEFTPADDING",(0,0),(0,0),12),
        ]))
        narr_cell = Paragraph(d["narrative"], S["body"])
        layout = Table([[score_cell, narr_cell]], colWidths=[4.2*cm, W-4.2*cm])
        layout.setStyle(TableStyle([
            ("VALIGN",(0,0),(-1,-1),"TOP"),
            ("LEFTPADDING",(1,0),(1,0),10),
        ]))
        story.append(layout)
        story.append(Spacer(1, 0.3*cm))

        # Sub-metrics table
        sm_hdr = ["Code", "Sub-Metric", "Score", "Maturity", "Public Signal / Evidence"]
        sm_rows = [sm_hdr] + [
            [code, name, str(score), maturity_label(score), evidence]
            for code, name, score, evidence in d["sub_metrics"]
        ]
        sm_cw = [1.2*cm, 4.0*cm, 1.3*cm, 1.8*cm, W-8.3*cm]
        sm_table = Table(sm_rows, colWidths=sm_cw, repeatRows=1)
        sm_ts = [
            ("BACKGROUND", (0,0),(-1,0), MID_BLUE),
            ("TEXTCOLOR",  (0,0),(-1,0), WHITE),
            ("FONTNAME",   (0,0),(-1,0), "Helvetica-Bold"),
            ("FONTSIZE",   (0,0),(-1,-1), 8),
            ("GRID",       (0,0),(-1,-1), 0.5, colors.HexColor("#CCCCCC")),
            ("VALIGN",     (0,0),(-1,-1), "TOP"),
            ("TOPPADDING", (0,0),(-1,-1), 4),
            ("BOTTOMPADDING",(0,0),(-1,-1),4),
            ("LEFTPADDING",(0,0),(-1,-1), 4),
            ("WORDWRAP",   (0,0),(-1,-1), True),
        ]
        for ri, (_, _, sc, _) in enumerate(d["sub_metrics"], start=1):
            sm_ts.append(("BACKGROUND", (0,ri),(-1,ri), maturity_light(sc)))
            sm_ts.append(("TEXTCOLOR",  (2,ri),(3,ri),  maturity_color(sc)))
            sm_ts.append(("FONTNAME",   (2,ri),(3,ri),  "Helvetica-Bold"))
        sm_table.setStyle(TableStyle(sm_ts))
        story.append(sm_table)
        story.append(PageBreak())

    # ── Section 7: Next Best Actions ───────────────────────────────────────────
    story.append(Paragraph("7. Next Best Actions — Digital Transformation Roadmap", S["h1"]))
    story.append(Paragraph(
        "Priority digital transformation initiatives based on the five-dimension DDMR diagnosis, "
        "sequenced by impact. Each initiative includes the expected benefit, cost of inaction, "
        "and the market window that makes timing critical.", S["body"]))

    for action in next_actions:
        items = [
            [Paragraph(f"<b>{action['rank']}: {action['title']}</b>",
                ParagraphStyle("ah", fontSize=10, fontName="Helvetica-Bold",
                textColor=WHITE, leading=14)),
             Paragraph(f"Addresses: {action['dimension']}",
                ParagraphStyle("ad", fontSize=8, textColor=WHITE, leading=12))],
        ]
        hdr_tbl = Table(items, colWidths=[W*0.65, W*0.35])
        hdr_tbl.setStyle(TableStyle([
            ("BACKGROUND", (0,0),(-1,-1), DARK_BLUE),
            ("VALIGN",     (0,0),(-1,-1), "MIDDLE"),
            ("TOPPADDING", (0,0),(-1,-1), 8),
            ("BOTTOMPADDING",(0,0),(-1,-1),8),
            ("LEFTPADDING",(0,0),(-1,-1), 8),
        ]))

        detail_rows = [
            [Paragraph("<b>Initiative</b>", ParagraphStyle("l",fontSize=8,fontName="Helvetica-Bold",textColor=WHITE)),
             Paragraph(action["initiative"], ParagraphStyle("v",fontSize=8,textColor=GRAY,leading=13))],
            [Paragraph("<b>Expected Benefit</b>", ParagraphStyle("l",fontSize=8,fontName="Helvetica-Bold",textColor=WHITE)),
             Paragraph(action["benefit"], ParagraphStyle("v",fontSize=8,textColor=colors.HexColor("#166534"),leading=13))],
            [Paragraph("<b>Cost of Inaction</b>", ParagraphStyle("l",fontSize=8,fontName="Helvetica-Bold",textColor=WHITE)),
             Paragraph(action["loss"], ParagraphStyle("v",fontSize=8,textColor=colors.HexColor("#991B1B"),leading=13))],
            [Paragraph("<b>Why Now</b>", ParagraphStyle("l",fontSize=8,fontName="Helvetica-Bold",textColor=WHITE)),
             Paragraph(action["why_now"], ParagraphStyle("v",fontSize=8,textColor=colors.HexColor("#1E3A5F"),leading=13))],
        ]
        detail_tbl = Table(detail_rows, colWidths=[2.8*cm, W-2.8*cm])
        detail_tbl.setStyle(TableStyle([
            ("BACKGROUND", (0,0),(0,0), MID_BLUE),
            ("BACKGROUND", (0,1),(0,1), GREEN),
            ("BACKGROUND", (0,2),(0,2), RED_),
            ("BACKGROUND", (0,3),(0,3), MID_BLUE),
            ("BACKGROUND", (1,0),(1,0), LIGHT_GRAY),
            ("BACKGROUND", (1,1),(1,1), LIGHT_GREEN),
            ("BACKGROUND", (1,2),(1,2), LIGHT_RED),
            ("BACKGROUND", (1,3),(1,3), LIGHT_BLUE2),
            ("GRID",       (0,0),(-1,-1), 0.5, colors.HexColor("#CCCCCC")),
            ("VALIGN",     (0,0),(-1,-1), "TOP"),
            ("TOPPADDING", (0,0),(-1,-1), 6),
            ("BOTTOMPADDING",(0,0),(-1,-1),6),
            ("LEFTPADDING",(0,0),(-1,-1), 6),
        ]))

        story.append(KeepTogether([hdr_tbl, detail_tbl, Spacer(1, 0.4*cm)]))

    story.append(PageBreak())

    # ── Section 12: Evidence & Citations ──────────────────────────────────────
    story.append(Paragraph("12. Evidence &amp; Source Citations", S["h1"]))
    story.append(Paragraph(
        "All scores are derived exclusively from publicly available information as of the report date. "
        "No management interviews, internal documents, or non-public data were accessed.", S["body"]))

    evidence = [
        ("MCA21 Public Records", "CIN U31103TZ1991PTC003409; incorporation 28 Aug 1991; ROC Coimbatore; 3 active directors; AGM Sept 2024. Source: Zaubacorp, Falcon eBiz, Company Vakil."),
        ("ICRA Credit Rating", "Multiple reaffirmations (2021–2025); Stable outlook; FY2025 revenue ₹527.8 Cr confirmed. Source: icra.in rationale reports."),
        ("Tracxn / Tofler", "FY2025 revenue ₹531 Cr; 662 employees; Auth. capital ₹30 Cr; paid-up ₹26.5 Cr. Source: tracxn.com, tofler.in."),
        ("Company Websites", "Product portfolio (3,500+ SKUs), 70,000 sqm plant, 42+ countries, USA/Germany subsidiaries, 17,000+ customers, 200 distribution partners, DSIR R&D, I4.0 / electric actuator / robotic product lines. Source: janatics.com, janaticspneumatics.com."),
        ("AIA India", "Janatics listed as member of Automation Industry Association India. Source: aia-india.org."),
        ("Digital Signals (Absence)", "No B2B e-commerce, active social media, cloud migration, ERP/MES disclosure, digital upskilling, cybersecurity framework, or formal DX roadmap identified from public sources."),
    ]
    for src, ev in evidence:
        story.append(Paragraph(f"• <b>{src}:</b> {ev}", S["bullet"]))

    story.append(Spacer(1, 0.4*cm))
    story.append(Paragraph("Disclaimer", S["h2"]))
    story.append(Paragraph(
        "This DDMR is prepared for informational purposes only. Scores are based on publicly available "
        "data and analytical judgment. They do not constitute investment advice, credit assessment, or a "
        "definitive representation of the company's digital capabilities.", S["body"]))

    # ── Footer on every page ───────────────────────────────────────────────────
    def add_footer(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(GRAY)
        canvas.drawString(2*cm, 1.2*cm, f"DDMR — {COMPANY} — Confidential — {DATE_STR}")
        canvas.drawRightString(A4[0]-2*cm, 1.2*cm, f"Page {doc.page}")
        canvas.restoreState()

    doc.build(story, onFirstPage=add_footer, onLaterPages=add_footer)
    print(f"Saved: {out}")

build()
