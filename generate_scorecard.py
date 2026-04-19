import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.chart import RadarChart, Reference, BarChart
from openpyxl.utils import get_column_letter
from datetime import date

COMPANY = "Janatics India Private Limited"
CIN     = "U31103TZ1991PTC003409"
DATE    = date.today().strftime("%d %B %Y")

def maturity_label(s):
    if s >= 84: return "Future Ready"
    if s >= 67: return "Strategic"
    if s >= 34: return "Siloed"
    return "Legacy"

dimensions = [
    {
        "code": "ST", "name": "Strategy", "weight": 0.20, "score": 70,
        "sub_metrics": [
            ("ST01", "Digital Strategy Articulation", 65, "I4.0 and digital products mentioned but no formal DX roadmap publicly disclosed"),
            ("ST02", "Leadership Technology Vision", 72, "Management pivoted to electric actuators, robotics, sensors — signals digital awareness"),
            ("ST03", "Digital Investment Signals", 75, "New IoT-enabled product lines indicate technology capex commitment"),
            ("ST04", "Technology Partnerships", 68, "DSIR R&D recognition confirmed; no public tech alliance or digital ecosystem partnerships"),
            ("ST05", "Innovation Pipeline", 70, "Incremental product digitization underway; no IP filing trail or open innovation signals"),
        ]
    },
    {
        "code": "OSC", "name": "Operations & Supply Chain", "weight": 0.20, "score": 65,
        "sub_metrics": [
            ("OSC01", "Digital Manufacturing Readiness", 60, "70,000 sqm Coimbatore plant with CNC; no Industry 4.0 certification or smart factory disclosure"),
            ("OSC02", "Supply Chain Visibility Tools", 50, "No public disclosure of ERP, SCM platform, or supply chain digitization initiative"),
            ("OSC03", "Process Automation Adoption", 70, "CNC machining, casting automation, surface treatment lines — physical automation established"),
            ("OSC04", "Quality Management (Digital)", 72, "ICRA notes consistent regulatory compliance; ISO quality system assumed but unconfirmed"),
            ("OSC05", "Industry 4.0 Integration", 73, "I4.0 product portfolio exists; internal I4.0 adoption for own manufacturing not disclosed"),
        ]
    },
    {
        "code": "SM", "name": "Sales & Marketing", "weight": 0.20, "score": 55,
        "sub_metrics": [
            ("SM01", "Digital Presence & SEO", 60, "Two product websites (janatics.com, janaticspneumatics.com); SEO signals moderate"),
            ("SM02", "Social Media Engagement", 40, "Limited LinkedIn activity signal; no active social media marketing evidenced"),
            ("SM03", "E-commerce / Digital Ordering", 45, "No evidence of digital ordering portal, e-catalog with pricing, or B2B e-commerce"),
            ("SM04", "Digital Marketing Activity", 55, "Web presence established; no paid digital campaigns or content marketing evidenced"),
            ("SM05", "CRM & Customer Data Utilization", 75, "17,000+ customers across 7 industries — implies structured data; CRM platform undisclosed"),
        ]
    },
    {
        "code": "TA", "name": "Technology Adoption", "weight": 0.20, "score": 70,
        "sub_metrics": [
            ("TA01", "Website & Digital Infrastructure", 65, "Functional product websites; no modern tech stack, API integrations, or customer portal evidence"),
            ("TA02", "IoT / Smart Product Portfolio", 80, "Electric actuators, smart sensors, robotics, I4.0 didactic equipment launched — strong product digitization"),
            ("TA03", "Cloud & SaaS Adoption", 60, "No public disclosures of cloud migration, SaaS tools, or digital collaboration platforms"),
            ("TA04", "R&D in Digital / Smart Tech", 78, "DSIR R&D actively developing robotic systems, sensor products, and smart automation equipment"),
            ("TA05", "Cybersecurity & Data Practices", 67, "ICRA-compliant practices; no public cybersecurity framework, ISO 27001, or data policy disclosed"),
        ]
    },
    {
        "code": "SC", "name": "Skills & Capabilities", "weight": 0.20, "score": 60,
        "sub_metrics": [
            ("SC01", "Digital Talent Hiring Signals", 55, "662 employees primarily engineering/manufacturing; limited signals of software or data roles hired"),
            ("SC02", "Technical Workforce Depth", 72, "Strong mechanical and manufacturing engineering talent; DSIR R&D team; CNC and automation specialists"),
            ("SC03", "Digital Training Programs", 50, "No public evidence of L&D investment, digital upskilling programs, or online learning adoption"),
            ("SC04", "Leadership Digital Literacy", 65, "Management demonstrates I4.0 awareness via product decisions; no DX leadership hire signals"),
            ("SC05", "Innovation Culture Signals", 58, "DSIR R&D present; no public hackathon, innovation lab, or employee innovation program signals"),
        ]
    },
]

composite = sum(d["score"] * d["weight"] for d in dimensions)

def score_color(score):
    if score >= 84: return "2563EB"   # blue — Future Ready
    if score >= 67: return "00B050"   # green — Strategic
    if score >= 34: return "FFC000"   # amber — Siloed
    return "FF0000"                   # red — Legacy

def thin_border():
    s = Side(style='thin', color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

def header_fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

DARK_BLUE  = "1F3864"
MID_BLUE   = "2E75B6"
LIGHT_BLUE = "D6E4F0"
WHITE      = "FFFFFF"

wb = openpyxl.Workbook()

# ── Sheet 1: Summary ───────────────────────────────────────────────────────────

ws = wb.active
ws.title = "Scorecard Summary"
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 4
ws.column_dimensions["B"].width = 34
ws.column_dimensions["C"].width = 14
ws.column_dimensions["D"].width = 14
ws.column_dimensions["E"].width = 14
ws.column_dimensions["F"].width = 16
ws.row_dimensions[1].height = 8

for r in range(2, 6):
    ws.row_dimensions[r].height = 20

ws.merge_cells("B2:F2")
ws["B2"] = "DDMR — DIGITAL DIAGNOSTIC AND MATURITY REPORT"
ws["B2"].font = Font(name="Calibri", bold=True, size=16, color=WHITE)
ws["B2"].fill = header_fill(DARK_BLUE)
ws["B2"].alignment = Alignment(horizontal="center", vertical="center")

ws.merge_cells("B3:F3")
ws["B3"] = COMPANY
ws["B3"].font = Font(name="Calibri", bold=True, size=13, color=WHITE)
ws["B3"].fill = header_fill(MID_BLUE)
ws["B3"].alignment = Alignment(horizontal="center", vertical="center")

ws.merge_cells("B4:F4")
ws["B4"] = f"CIN: {CIN}   |   Report Date: {DATE}   |   Composite Score: {composite:.0f}/100   |   Maturity: {maturity_label(composite)}"
ws["B4"].font = Font(name="Calibri", size=10, color=DARK_BLUE)
ws["B4"].fill = header_fill(LIGHT_BLUE)
ws["B4"].alignment = Alignment(horizontal="center", vertical="center")

ws.row_dimensions[5].height = 8
ws.row_dimensions[6].height = 60
ws.merge_cells("B6:C6")
ws["B6"] = f"{composite:.0f}"
ws["B6"].font = Font(name="Calibri", bold=True, size=36, color=WHITE)
ws["B6"].fill = header_fill(score_color(composite))
ws["B6"].alignment = Alignment(horizontal="center", vertical="center")
ws.merge_cells("D6:F6")
ws["D6"] = f"COMPOSITE DDMR SCORE — {maturity_label(composite)}\n(Weighted across 5 Digital Dimensions)"
ws["D6"].font = Font(name="Calibri", bold=True, size=11, color=DARK_BLUE)
ws["D6"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

ws.row_dimensions[8].height = 22
for col, val in [
    ("B", "Dimension"), ("C", "Score /100"),
    ("D", "Weight"), ("E", "Weighted"), ("F", "Maturity Level")
]:
    ws[f"{col}8"] = val
    ws[f"{col}8"].font = Font(name="Calibri", bold=True, size=10, color=WHITE)
    ws[f"{col}8"].fill = header_fill(DARK_BLUE)
    ws[f"{col}8"].alignment = Alignment(horizontal="center", vertical="center")
    ws[f"{col}8"].border = thin_border()

for i, d in enumerate(dimensions, start=9):
    ws.row_dimensions[i].height = 20
    label = maturity_label(d["score"])
    vals = [d["name"], d["score"], f"{int(d['weight']*100)}%",
            round(d["score"] * d["weight"], 1), label]
    for col, val in zip(["B","C","D","E","F"], vals):
        cell = ws[f"{col}{i}"]
        cell.value = val
        cell.font = Font(name="Calibri", size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border()
        if col == "B":
            cell.alignment = Alignment(horizontal="left", vertical="center")
        if col in ("C", "E"):
            cell.fill = PatternFill("solid", fgColor=score_color(d["score"]) + "40")
        if col == "F":
            cell.font = Font(name="Calibri", bold=True, size=10, color=score_color(d["score"]))

tr = 9 + len(dimensions)
ws.row_dimensions[tr].height = 22
ws[f"B{tr}"] = "COMPOSITE SCORE"
ws[f"C{tr}"] = composite
ws[f"D{tr}"] = "100%"
ws[f"E{tr}"] = round(composite, 1)
ws[f"F{tr}"] = maturity_label(composite)
for col in ["B","C","D","E","F"]:
    cell = ws[f"{col}{tr}"]
    cell.font = Font(name="Calibri", bold=True, size=11, color=WHITE)
    cell.fill = header_fill(MID_BLUE)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thin_border()
ws[f"B{tr}"].alignment = Alignment(horizontal="left", vertical="center")

ws.conditional_formatting.add(
    f"C9:C{tr-1}",
    ColorScaleRule(
        start_type="num", start_value=0,   start_color="FF0000",
        mid_type="num",   mid_value=50,    mid_color="FFFF00",
        end_type="num",   end_value=100,   end_color="00B050"
    )
)

chart = RadarChart()
chart.type = "filled"
chart.title = "DDMR Digital Maturity — Dimension Scores"
chart.style = 10
chart.width = 14
chart.height = 12

labels = Reference(ws, min_col=2, min_row=9, max_row=9+len(dimensions)-1)
data   = Reference(ws, min_col=3, min_row=8, max_row=8+len(dimensions))
chart.add_data(data, titles_from_data=True)
chart.set_categories(labels)
ws.add_chart(chart, "B15")

# ── Sheets 2–6: Per-dimension ──────────────────────────────────────────────────

for d in dimensions:
    ws2 = wb.create_sheet(title=f"{d['code']} Detail")
    ws2.sheet_view.showGridLines = False
    ws2.column_dimensions["A"].width = 4
    ws2.column_dimensions["B"].width = 8
    ws2.column_dimensions["C"].width = 40
    ws2.column_dimensions["D"].width = 12
    ws2.column_dimensions["E"].width = 50

    ws2.row_dimensions[2].height = 24
    ws2.merge_cells("B2:E2")
    ws2["B2"] = f"{d['code']} — {d['name']}   |   Score: {d['score']}/100   |   Weight: {int(d['weight']*100)}%   |   {maturity_label(d['score'])}"
    ws2["B2"].font = Font(name="Calibri", bold=True, size=13, color=WHITE)
    ws2["B2"].fill = header_fill(MID_BLUE)
    ws2["B2"].alignment = Alignment(horizontal="center", vertical="center")

    ws2.row_dimensions[4].height = 20
    for col, val in [("B","Code"),("C","Sub-Metric"),("D","Score /100"),("E","Public Signal / Evidence")]:
        ws2[f"{col}4"] = val
        ws2[f"{col}4"].font = Font(name="Calibri", bold=True, size=10, color=WHITE)
        ws2[f"{col}4"].fill = header_fill(DARK_BLUE)
        ws2[f"{col}4"].alignment = Alignment(horizontal="center", vertical="center")
        ws2[f"{col}4"].border = thin_border()

    for j, (code, name, score, evidence) in enumerate(d["sub_metrics"], start=5):
        ws2.row_dimensions[j].height = 32
        for col, val in [("B",code),("C",name),("D",score),("E",evidence)]:
            cell = ws2[f"{col}{j}"]
            cell.value = val
            cell.font = Font(name="Calibri", size=9)
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            cell.border = thin_border()
            if col == "D":
                cell.font = Font(name="Calibri", bold=True, size=11, color=score_color(score))
                cell.alignment = Alignment(horizontal="center", vertical="center")
            if col == "B":
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = Font(name="Calibri", bold=True, size=9, color=MID_BLUE)

    last = 5 + len(d["sub_metrics"])
    ws2.row_dimensions[last].height = 22
    ws2[f"B{last}"] = "AVG"
    ws2[f"C{last}"] = "Dimension Average Score"
    ws2[f"D{last}"] = d["score"]
    ws2[f"E{last}"] = f"Maturity: {maturity_label(d['score'])}  |  Weighted contribution: {d['score'] * d['weight']:.1f} pts"
    for col in ["B","C","D","E"]:
        cell = ws2[f"{col}{last}"]
        cell.font = Font(name="Calibri", bold=True, size=10, color=WHITE)
        cell.fill = header_fill(MID_BLUE)
        cell.border = thin_border()
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws2[f"C{last}"].alignment = Alignment(horizontal="left", vertical="center")
    ws2[f"E{last}"].alignment = Alignment(horizontal="left", vertical="center")

    bar = BarChart()
    bar.type = "bar"
    bar.title = f"{d['code']}: {d['name']} — Sub-Metric Scores"
    bar.style = 10
    bar.width = 18
    bar.height = 10
    sm_labels = Reference(ws2, min_col=2, min_row=5, max_row=5+len(d["sub_metrics"])-1)
    sm_data   = Reference(ws2, min_col=4, min_row=4, max_row=4+len(d["sub_metrics"]))
    bar.add_data(sm_data, titles_from_data=True)
    bar.set_categories(sm_labels)
    ws2.add_chart(bar, "B13")

# ── Sheet 7: Evidence Index ────────────────────────────────────────────────────

ws3 = wb.create_sheet("Evidence Index")
ws3.sheet_view.showGridLines = False
ws3.column_dimensions["A"].width = 4
ws3.column_dimensions["B"].width = 10
ws3.column_dimensions["C"].width = 35
ws3.column_dimensions["D"].width = 20
ws3.column_dimensions["E"].width = 55

ws3.merge_cells("B2:E2")
ws3["B2"] = f"Evidence Index — {COMPANY} — Digital Diagnostic and Maturity Report"
ws3["B2"].font = Font(name="Calibri", bold=True, size=13, color=WHITE)
ws3["B2"].fill = header_fill(DARK_BLUE)
ws3["B2"].alignment = Alignment(horizontal="center", vertical="center")
ws3.row_dimensions[2].height = 24

for col, val in [("B","Code"),("C","Sub-Metric"),("D","Source"),("E","Public Signal / Evidence")]:
    ws3[f"{col}4"] = val
    ws3[f"{col}4"].font = Font(name="Calibri", bold=True, size=10, color=WHITE)
    ws3[f"{col}4"].fill = header_fill(MID_BLUE)
    ws3[f"{col}4"].border = thin_border()
    ws3[f"{col}4"].alignment = Alignment(horizontal="center", vertical="center")

sources = [
    ("ST01","Digital Strategy Articulation","Janatics website / MCA","No formal DX roadmap in public disclosures; I4.0 product lines signal intent"),
    ("ST02","Leadership Technology Vision","Janatics website","Management added electric actuators, robotics, smart sensors — visible strategic pivot"),
    ("ST03","Digital Investment Signals","Janatics website / ICRA","New IoT-enabled product categories launched; DSIR R&D investment ongoing"),
    ("ST04","Technology Partnerships","MCA / Public records","DSIR R&D confirmed; no public digital alliance, co-development, or tech partner disclosures"),
    ("ST05","Innovation Pipeline","Janatics website","Incremental product digitization; no IP filing, open innovation, or startup engagement public"),
    ("OSC01","Digital Manufacturing Readiness","janaticspneumatics.com","70,000 sqm Coimbatore plant; CNC machining; no I4.0 certification or smart factory claims"),
    ("OSC02","Supply Chain Visibility Tools","Public records","No ERP, SCM platform, or supply chain digitization initiative disclosed publicly"),
    ("OSC03","Process Automation Adoption","janaticspneumatics.com","CNC machining centers, casting, surface treatment — physical automation well established"),
    ("OSC04","Quality Management (Digital)","ICRA Rationale","ICRA notes consistent compliance; ISO quality framework assumed but not confirmed publicly"),
    ("OSC05","Industry 4.0 Integration","Janatics website","I4.0 didactic product line exists; internal adoption for own operations undisclosed"),
    ("SM01","Digital Presence & SEO","janatics.com / janaticspneumatics.com","Two product websites with online catalog; SEO signals moderate; no chatbot or digital engagement tools"),
    ("SM02","Social Media Engagement","LinkedIn / Public","Limited LinkedIn activity signal; no active social media marketing strategy evidenced"),
    ("SM03","E-commerce / Digital Ordering","Public research","No B2B e-commerce, digital ordering portal, or online pricing catalog identified"),
    ("SM04","Digital Marketing Activity","Web / Public","Websites established; no evidence of paid search, content marketing, or digital ad activity"),
    ("SM05","CRM & Customer Data","Janatics / ICRA","17,000+ customers across 7 industries implies structured data management; CRM tool undisclosed"),
    ("TA01","Website & Digital Infrastructure","janatics.com","Functional websites; no evidence of modern CMS, API layer, or customer self-service portal"),
    ("TA02","IoT / Smart Product Portfolio","Janatics website / ICRA","Electric actuators, smart sensors, robotic systems, I4.0 didactics — strong product-layer digitization"),
    ("TA03","Cloud & SaaS Adoption","Public records","No public cloud migration, SaaS adoption, or digital collaboration platform disclosures found"),
    ("TA04","R&D in Digital / Smart Tech","DSIR / Janatics website","DSIR-recognized R&D developing robotic systems, sensor products, and smart automation equipment"),
    ("TA05","Cybersecurity & Data Practices","ICRA / Public","ICRA-compliant; no public cybersecurity framework, ISO 27001 certification, or data policy found"),
    ("SC01","Digital Talent Hiring Signals","LinkedIn / Public","662 employees primarily engineering/manufacturing; limited visible software or data role hiring"),
    ("SC02","Technical Workforce Depth","Tracxn / ICRA","Strong mechanical and manufacturing engineering base; DSIR R&D team; CNC and automation specialists"),
    ("SC03","Digital Training Programs","Public records","No public L&D investment, digital upskilling program, or online learning platform adoption disclosed"),
    ("SC04","Leadership Digital Literacy","Janatics website / ICRA","Management I4.0 awareness demonstrated via product decisions; no digital leadership hire signals"),
    ("SC05","Innovation Culture Signals","Public research","DSIR R&D present; no hackathon, innovation lab, startup engagement, or employee innovation program"),
]

for j, row in enumerate(sources, start=5):
    ws3.row_dimensions[j].height = 28
    for col, val in zip(["B","C","D","E"], row):
        cell = ws3[f"{col}{j}"]
        cell.value = val
        cell.font = Font(name="Calibri", size=9)
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        cell.border = thin_border()
        if col == "B":
            cell.font = Font(name="Calibri", bold=True, size=9, color=MID_BLUE)
            cell.alignment = Alignment(horizontal="center", vertical="center")

out = "/Users/laaksandy/Documents/Claude-testing/DDMR/output/Janatics_DDMR_Scorecard.xlsx"
wb.save(out)
print(f"Saved: {out}")
