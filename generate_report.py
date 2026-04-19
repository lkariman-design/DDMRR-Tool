from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from datetime import date

COMPANY  = "Janatics India Private Limited"
CIN      = "U31103TZ1991PTC003409"
DATE_STR = date.today().strftime("%d %B %Y")
COMPOSITE = 75.7

# ── Helpers ────────────────────────────────────────────────────────────────────

def set_cell_bg(cell, hex_color):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    tcPr.append(shd)

def add_heading(doc, text, level=1, color="1F3864"):
    h = doc.add_heading(text, level=level)
    h.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for run in h.runs:
        run.font.color.rgb = RGBColor.from_string(color)
    return h

def score_label(s):
    if s >= 75: return "Strong"
    if s >= 50: return "Moderate"
    return "Weak"

# ── Document ───────────────────────────────────────────────────────────────────

doc = Document()

# Page margins
for section in doc.sections:
    section.top_margin    = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin   = Cm(2.5)
    section.right_margin  = Cm(2.5)

# ── Cover page ─────────────────────────────────────────────────────────────────

doc.add_paragraph()
t = doc.add_paragraph()
t.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = t.add_run("DIGITAL DIAGNOSTIC AND MATURITY REPORT (DDMR)")
run.font.bold  = True
run.font.size  = Pt(22)
run.font.color.rgb = RGBColor.from_string("1F3864")

doc.add_paragraph()
t2 = doc.add_paragraph()
t2.alignment = WD_ALIGN_PARAGRAPH.CENTER
r2 = t2.add_run(COMPANY)
r2.font.bold = True
r2.font.size = Pt(18)
r2.font.color.rgb = RGBColor.from_string("2E75B6")

doc.add_paragraph()
t3 = doc.add_paragraph()
t3.alignment = WD_ALIGN_PARAGRAPH.CENTER
r3 = t3.add_run(f"CIN: {CIN}   |   Report Date: {DATE_STR}")
r3.font.size = Pt(11)

doc.add_paragraph()
t4 = doc.add_paragraph()
t4.alignment = WD_ALIGN_PARAGRAPH.CENTER
r4 = t4.add_run(f"Composite DDMR Score: {COMPOSITE}/100  —  {score_label(COMPOSITE)}")
r4.font.bold  = True
r4.font.size  = Pt(14)
r4.font.color.rgb = RGBColor.from_string("00B050")

doc.add_page_break()

# ── Section 1: Executive Summary ───────────────────────────────────────────────

add_heading(doc, "1. Executive Summary")
para = doc.add_paragraph()
para.add_run(
    f"{COMPANY} (CIN: {CIN}) is a Coimbatore-based manufacturer of pneumatic components "
    f"and industrial automation products, incorporated in 1991 and operating since 1977. "
    f"With FY2025 revenue of ₹531 Crore and over 3,500 distinct products sold across 42+ countries "
    f"to 17,000+ customers, the company represents one of India's most established domestic "
    f"pneumatics brands. It holds ICRA credit ratings with a Stable outlook and has subsidiaries "
    f"in the USA and Germany."
)

doc.add_paragraph()
para2 = doc.add_paragraph()
para2.add_run(
    f"The DDMR composite score is "
).font.bold = False
run_score = para2.add_run(f"{COMPOSITE}/100 (Strong)")
run_score.font.bold = True
run_score.font.color.rgb = RGBColor.from_string("00B050")
para2.add_run(
    f", reflecting strong operational foundations, global presence, and brand maturity, "
    f"partially offset by below-industry revenue growth (4% vs 6.9% sector CAGR) and EBITDA "
    f"margin compression observed in FY2025."
)

# Summary table
doc.add_paragraph()
tbl = doc.add_table(rows=7, cols=4)
tbl.style = "Table Grid"
headers = ["Dimension", "Score /100", "Weight", "Rating"]
for i, h in enumerate(headers):
    cell = tbl.rows[0].cells[i]
    cell.text = h
    cell.paragraphs[0].runs[0].font.bold  = True
    cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
    set_cell_bg(cell, "1F3864")

rows_data = [
    ("Strategy & Leadership",      81, "25%", "Strong"),
    ("Sales & Marketing",          78, "20%", "Strong"),
    ("Operations & Supply Chain",  78, "20%", "Strong"),
    ("Finance",                    67, "25%", "Moderate"),
    ("Industry Context",           75, "10%", "Strong"),
    (f"COMPOSITE SCORE",           75.7, "100%", "Strong"),
]
for i, (dim, score, wt, rating) in enumerate(rows_data, start=1):
    row = tbl.rows[i]
    row.cells[0].text = dim
    row.cells[1].text = str(score)
    row.cells[2].text = wt
    row.cells[3].text = rating
    color = "00B050" if score >= 75 else ("FFC000" if score >= 50 else "FF0000")
    for j in range(4):
        set_cell_bg(row.cells[j], "D6E4F0" if i < 6 else "2E75B6")
    for j in range(4):
        p = row.cells[j].paragraphs[0]
        if p.runs:
            if i == 6:
                p.runs[0].font.bold = True
                p.runs[0].font.color.rgb = RGBColor(255,255,255)
            else:
                p.runs[0].font.color.rgb = RGBColor.from_string("1F3864")

doc.add_page_break()

# ── Sections 2–6: Dimension deep-dives ────────────────────────────────────────

dimension_narratives = {
    "SL": {
        "title": "2. Strategy & Leadership (Score: 81/100 — Strong)",
        "narrative": (
            "Janatics India demonstrates exceptional leadership tenure and strategic clarity. "
            "The company was established in 1977, giving it a 49-year operational track record — "
            "a rare distinction among domestic manufacturers. Its three-director board (Jaganathan Ganesh Kumar, "
            "G C Nageswaran, and Rajamani Ramesh) maintains active compliance with all DIN requirements and "
            "statutory filings with ROC Coimbatore.\n\n"
            "The strategic expansion beyond pneumatics into electric actuators, sensors, motors, and robotic "
            "systems demonstrates forward-looking leadership aligned with Industry 4.0 megatrends. The "
            "establishment of subsidiaries in the USA and Germany reflects deliberate global market strategy "
            "rather than opportunistic export. The company's DSIR-recognized R&D division signals commitment "
            "to proprietary innovation.\n\n"
            "The primary governance gap is the absence of independent director disclosures or ESG commitments "
            "in public records — a factor increasingly scrutinized by institutional buyers and export partners."
        ),
        "sub_metrics": [
            ("SL01", "Company/Founder Tenure", 90, "Established 1977; incorporated 1991 — 49-year operating history"),
            ("SL02", "Director Profile & DIN Compliance", 75, "3 active directors; all DINs current; AGM Sept 2024"),
            ("SL03", "Governance & Filing Compliance", 80, "Annual filings current; no ROC penalties noted in public records"),
            ("SL04", "Debt & Charge Structure", 78, "ICRA-rated; stable outlook; healthy debt protection metrics"),
            ("SL05", "Strategic Direction", 82, "Industry 4.0 pivot: electric actuators, robotics, sensors added"),
            ("SL06", "Global Presence", 88, "42+ countries; USA and Germany subsidiaries; 200 global distribution partners"),
        ]
    },
    "SM": {
        "title": "3. Sales & Marketing (Score: 78/100 — Strong)",
        "narrative": (
            "Janatics' sales and marketing profile is characterized by exceptional customer diversity and "
            "distribution breadth, partially tempered by revenue growth trailing industry pace.\n\n"
            "The company serves 17,000+ customers across seven end-use industries — pharmaceuticals, "
            "automotive, packaging, printing, food processing, medical equipment, and textiles. This "
            "diversification insulates revenues from sector-specific cyclicality. The global distribution "
            "network of 200+ partners enables market access without proportionate fixed-cost overhead.\n\n"
            "Revenue grew from ₹507.9 Cr (FY2024) to ₹527.8 Cr (FY2025), a 4% YoY increase — below the "
            "industry CAGR of 6.9%. This suggests either deliberate margin-over-volume positioning or "
            "competitive pressure from global players (SMC Corporation, Festo, Parker Hannifin) in the "
            "premium segment. The company operates under its own brand, Janatics, which commands recognition "
            "within AIA India member circles and among OEM procurement teams.\n\n"
            "A digital commerce and direct-to-customer channel strategy is not evidenced in public disclosures "
            "— a potential growth lever not yet activated."
        ),
        "sub_metrics": [
            ("SM01", "Revenue Growth vs Industry CAGR", 55, "4% YoY vs 6.9% sector CAGR — below market pace"),
            ("SM02", "Revenue Scale", 80, "₹531 Cr FY2025; consistent ₹500 Cr+ for 2 consecutive years"),
            ("SM03", "Customer Diversity", 90, "17,000+ customers; 7 industries; low single-customer concentration risk"),
            ("SM04", "Distribution Network", 85, "200 global partners; 3 manufacturing plants for regional reach"),
            ("SM05", "Brand Strength", 80, "49-year brand; AIA India member; DSIR-recognized R&D"),
        ]
    },
    "OSC": {
        "title": "4. Operations & Supply Chain (Score: 78/100 — Strong)",
        "narrative": (
            "Janatics operates one of the largest dedicated pneumatics manufacturing facilities in India — "
            "a 70,000 square meter plant in Coimbatore equipped with casting equipment, CNC machining centers, "
            "and surface treatment lines. Additional plants in Ahmedabad and Noida provide regional manufacturing "
            "flexibility and reduce logistics lead times for northern and western markets.\n\n"
            "The product portfolio exceeds 3,500 distinct SKUs spanning pneumatic cylinders, directional control "
            "valves, FRL (Filter-Regulator-Lubricator) units, solenoid valves, one-touch fittings, and didactic "
            "automation equipment. This breadth enables Janatics to serve as a single-source supplier for OEM "
            "customers — a significant procurement advantage.\n\n"
            "The DSIR-recognized R&D division underpins product development velocity and differentiates Janatics "
            "from purely assembly-based competitors. The company's expansion into electric actuators and robotics "
            "signals investment in manufacturing capabilities beyond legacy pneumatics.\n\n"
            "The principal operational risk is internal digitization transparency — ERP systems, supply chain "
            "visibility tools, and manufacturing execution systems are not disclosed, limiting the ability to "
            "score OSC05 (digitization) with confidence."
        ),
        "sub_metrics": [
            ("OSC01", "Manufacturing Scale", 85, "70,000 sqm Coimbatore HQ; Ahmedabad and Noida plants"),
            ("OSC02", "Product Portfolio Breadth", 90, "3,500+ distinct products across multiple automation categories"),
            ("OSC03", "R&D & Technology", 82, "DSIR-recognized R&D; CNC and casting; robotics expansion underway"),
            ("OSC04", "GST & Regulatory Compliance", 72, "Active GST entity; ICRA notes consistent compliance track record"),
            ("OSC05", "Supply Chain Digitization", 65, "Industry 4.0 product range present; internal digitization undisclosed"),
        ]
    },
    "FIN": {
        "title": "5. Finance (Score: 67/100 — Moderate)",
        "narrative": (
            "Janatics' financial profile is characterized by strong absolute scale and robust credit metrics, "
            "but flagged by EBITDA margin compression in FY2025 that warrants close monitoring.\n\n"
            "Revenue of ₹531 Crore in FY2025 positions the company solidly in the mid-large tier of "
            "domestic manufacturers, with employee productivity of approximately ₹80 lakh per employee — "
            "reasonable for a capital-intensive manufacturing operation. The capital structure is conservative: "
            "authorized capital of ₹30 Cr against paid-up capital of ₹26.5 Cr, with no external equity "
            "funding rounds identified.\n\n"
            "The ICRA credit rating with a Stable outlook and repeated reaffirmations over multiple years is "
            "the single most significant positive signal in this dimension — it validates balance sheet health, "
            "debt servicing capacity, and management quality from an independent third party.\n\n"
            "The primary concern is a reported 1-year EBITDA CAGR of -22%. While a single-year EBITDA "
            "decline can reflect input cost inflation, wage increases, or capital expenditure absorption, "
            "it represents a meaningful deviation from prior profitability levels. Full P&L and balance sheet "
            "data (not publicly available for unlisted private companies) would be needed for a definitive "
            "financial assessment."
        ),
        "sub_metrics": [
            ("FIN01", "Revenue Scale", 80, "₹531 Cr FY2025 — strong for unlisted domestic manufacturer"),
            ("FIN02", "Revenue Growth (4% YoY)", 55, "FY24→FY25 growth 4%; below industry but within inflation-adjusted range"),
            ("FIN03", "EBITDA Trajectory", 42, "EBITDA CAGR -22% (1-yr, Tracxn) — margin compression; requires full P&L review"),
            ("FIN04", "Credit Profile (ICRA)", 82, "ICRA-rated; stable; healthy debt protection metrics; multiple reaffirmations"),
            ("FIN05", "Capital Structure", 75, "Auth. ₹30 Cr; paid-up ₹26.5 Cr; no external equity; self-funded growth"),
            ("FIN06", "Employee Productivity", 70, "662 employees; ₹531 Cr revenue = ~₹80L per employee"),
        ]
    },
    "IC": {
        "title": "6. Industry Context (Score: 75/100 — Strong)",
        "narrative": (
            "The Indian pneumatics and industrial automation market presents a favourable medium-term "
            "backdrop for Janatics. The India pneumatic equipment market stood at $1.06 billion in 2024 "
            "and is projected to reach $1.94 billion by 2033 (CAGR: 6.9%), driven by manufacturing "
            "automation adoption, infrastructure build-out, and the Make in India policy framework.\n\n"
            "Janatics is well-positioned relative to these tailwinds. Its manufacturing base in India "
            "aligns with the government's Production Linked Incentive (PLI) schemes for advanced "
            "manufacturing. The China+1 sourcing strategy adopted by global OEMs creates incremental "
            "export demand for quality Indian component manufacturers — a direct opportunity for Janatics "
            "given its existing presence in 42 countries.\n\n"
            "Competitive intensity is the key contextual risk. SMC Corporation (Japan), Festo (Germany), "
            "and Parker Hannifin (USA) dominate the premium pneumatics segment globally and are present "
            "in India. Janatics competes primarily on price-quality positioning in the mid-market, where "
            "it has a strong installed base. Defending this position as global players intensify "
            "localization will be the central competitive challenge in the next 3–5 years."
        ),
        "sub_metrics": [
            ("IC01", "India Pneumatics Market Size", 70, "$1.06B (2024); growing to $1.94B by 2033 — mid-size addressable market"),
            ("IC02", "Industry Growth Rate", 78, "CAGR 6.9% (2025–2033) — healthy; aligned with industrial capex cycle"),
            ("IC03", "Competitive Position", 70, "Domestic leader in mid-market; competes with SMC, Festo, Parker in premium"),
            ("IC04", "Macro Tailwinds", 82, "Make in India, PLI schemes, China+1 sourcing, manufacturing automation"),
            ("IC05", "End-Market Diversification", 76, "7 end-use industries: pharma, auto, packaging, food, medical, textile, printing"),
        ]
    },
}

for code, info in dimension_narratives.items():
    add_heading(doc, info["title"], level=1)
    for para_text in info["narrative"].split("\n\n"):
        doc.add_paragraph(para_text)

    doc.add_paragraph()
    sm_tbl = doc.add_table(rows=len(info["sub_metrics"])+1, cols=3)
    sm_tbl.style = "Table Grid"
    for i, h in enumerate(["Sub-Metric", "Score", "Evidence"]):
        cell = sm_tbl.rows[0].cells[i]
        cell.text = h
        cell.paragraphs[0].runs[0].font.bold = True
        cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(255,255,255)
        set_cell_bg(cell, "2E75B6")
    for j, (code2, name, score, evidence) in enumerate(info["sub_metrics"], start=1):
        sm_tbl.rows[j].cells[0].text = f"{code2}: {name}"
        sm_tbl.rows[j].cells[1].text = f"{score}/100"
        sm_tbl.rows[j].cells[2].text = evidence
        color = "D5E8D4" if score >= 75 else ("FFF2CC" if score >= 50 else "F8CECC")
        for k in range(3):
            set_cell_bg(sm_tbl.rows[j].cells[k], color)

    sm_tbl.columns[0].width = Cm(5)
    sm_tbl.columns[1].width = Cm(2)
    sm_tbl.columns[2].width = Cm(10)
    doc.add_page_break()

# ── Section 12: Evidence Narrative ─────────────────────────────────────────────

add_heading(doc, "12. Section 12 — Evidence & Source Citations")
doc.add_paragraph(
    "All scores in this report are derived exclusively from publicly available information "
    "as of the report date. No management interviews, internal documents, or non-public "
    "financials were accessed."
)

evidence_entries = [
    ("MCA21 Public Records", "CIN U31103TZ1991PTC003409 verified; incorporation date 28 Aug 1991; "
     "ROC Coimbatore; 3 active directors; AGM Sept 30 2024 confirmed. "
     "Source: Zaubacorp, Falcon eBiz, Company Vakil (MCA mirror portals)."),
    ("ICRA Credit Rating", "Multiple ICRA rating reaffirmations (2021, 2022, 2023, 2025). "
     "Stable outlook. Healthy debt protection metrics confirmed. FY2025 revenue ₹527.8 Cr confirmed. "
     "Source: icra.in rationale reports (IDs 105171, 120842, 130532, 139580)."),
    ("Tracxn / Tofler", "FY2025 revenue ₹531 Cr; 662 employees; EBITDA CAGR -22% (1-yr). "
     "Authorized capital ₹30 Cr; paid-up ₹26.5 Cr. "
     "Source: tracxn.com, tofler.in company profiles."),
    ("Company Websites", "Product portfolio (3,500+ SKUs), manufacturing facilities (Coimbatore 70,000 sqm, "
     "Ahmedabad, Noida), export presence (42+ countries), subsidiaries (USA, Germany), "
     "customer base (17,000+), distribution network (200 partners), DSIR R&D recognition. "
     "Source: janatics.com, janaticspneumatics.com."),
    ("Industry Research", "India pneumatic equipment market: $1.06B (2024), CAGR 6.9% (2025–2033). "
     "Source: IMARC Group market reports, Fortune Business Insights, Grand View Research."),
    ("AIA India", "Janatics India Pvt Ltd listed as member of Automation Industry Association India. "
     "Source: aia-india.org/member-directory."),
]

for source, evidence in evidence_entries:
    p = doc.add_paragraph(style="List Bullet")
    run_s = p.add_run(f"{source}: ")
    run_s.font.bold = True
    p.add_run(evidence)

doc.add_paragraph()
add_heading(doc, "Disclaimer", level=2)
doc.add_paragraph(
    "This report is prepared for informational and due diligence purposes only. Scores are "
    "based on publicly available data and analytical judgment. They do not constitute investment "
    "advice, credit assessment, or a representation of the company's current financial position. "
    "Full financial statements (P&L, balance sheet, cash flow) should be obtained directly from "
    "the company for a complete financial assessment."
)

# ── Save ───────────────────────────────────────────────────────────────────────

out = "/Users/laaksandy/Documents/Claude-testing/DDMR/output/Janatics_DDMR_Report.docx"
doc.save(out)
print(f"Saved: {out}")
