import { NextRequest, NextResponse } from "next/server";

export const maxDuration = 60;

// ── Palette (matching Python generators) ───────────────────────────────────────
const DARK_BLUE  = "1F3864";
const MID_BLUE   = "2E75B6";
const LIGHT_BLUE = "D6E4F0";

function maturityLabel(s: number): string {
  if (s >= 84) return "Future Ready";
  if (s >= 67) return "Strategic";
  if (s >= 34) return "Siloed";
  return "Legacy";
}
function scoreColor(s: number): string {
  if (s >= 84) return "2563EB";
  if (s >= 67) return "00B050";
  if (s >= 34) return "FFC000";
  return "FF0000";
}
function maturityLight(s: number): string {
  if (s >= 84) return "DAE8FC";
  if (s >= 67) return "D5E8D4";
  if (s >= 34) return "FFF2CC";
  return "F8CECC";
}
function safeFilename(name: string): string {
  return (name || "DDMR_Report").replace(/[^a-z0-9]/gi, "_").slice(0, 60);
}
function parseWeight(w: any): number {
  if (typeof w === "number") return w;
  if (typeof w === "string") return parseFloat(w) / 100 || 0.2;
  return 0.2;
}
function sanitize(text: string): string {
  return (text || "")
    .replace(/₹/g, "Rs.")
    .replace(/[""]/g, '"')
    .replace(/['']/g, "'")
    .replace(/[–—]/g, "-")
    .replace(/…/g, "...")
    .replace(/→/g, "->")
    .replace(/←/g, "<-")
    .replace(/•/g, "*")
    .replace(/™/g, "(TM)")
    .replace(/®/g, "(R)")
    .replace(/©/g, "(C)")
    .replace(/[^\x00-\xFF]/g, "?");
}

// ── XLSX (ExcelJS — full Calibri styling matching Python scorecard) ─────────────

async function buildXlsx(data: any): Promise<Buffer> {
  const ExcelJS = await import("exceljs");
  const wb = new ExcelJS.Workbook();
  wb.creator = "DDMR Tool";
  wb.created = new Date();

  const THIN: any = {
    top:    { style: "thin", color: { argb: "FFCCCCCC" } },
    left:   { style: "thin", color: { argb: "FFCCCCCC" } },
    bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
    right:  { style: "thin", color: { argb: "FFCCCCCC" } },
  };
  function fill(hex: string): any {
    return { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + hex } };
  }
  function cal(opts: any = {}): any {
    return { name: "Calibri", ...opts };
  }
  function center(wrapText = false): any {
    return { horizontal: "center", vertical: "middle", wrapText };
  }

  const dateStr = new Date().toLocaleDateString("en-IN", { day: "numeric", month: "long", year: "numeric" });
  const dims: any[] = data.dimensions || [];

  // ── Sheet 1: Scorecard Summary ──────────────────────────────────────────────
  const ws = wb.addWorksheet("Scorecard Summary");
  ws.views = [{ showGridLines: false }];
  ws.getColumn("A").width = 4;
  ws.getColumn("B").width = 34;
  ws.getColumn("C").width = 14;
  ws.getColumn("D").width = 14;
  ws.getColumn("E").width = 14;
  ws.getColumn("F").width = 16;

  ws.getRow(1).height = 8;

  // Row 2: DDMR title
  ws.getRow(2).height = 20;
  ws.mergeCells("B2:F2");
  const r2 = ws.getCell("B2");
  r2.value = "DDMR — DIGITAL DIAGNOSTIC AND MATURITY REPORT";
  r2.font  = cal({ bold: true, size: 16, color: { argb: "FFFFFFFF" } });
  r2.fill  = fill(DARK_BLUE);
  r2.alignment = center();

  // Row 3: Company
  ws.getRow(3).height = 20;
  ws.mergeCells("B3:F3");
  const r3 = ws.getCell("B3");
  r3.value = data.company || "";
  r3.font  = cal({ bold: true, size: 13, color: { argb: "FFFFFFFF" } });
  r3.fill  = fill(MID_BLUE);
  r3.alignment = center();

  // Row 4: Meta
  ws.getRow(4).height = 20;
  ws.mergeCells("B4:F4");
  const r4 = ws.getCell("B4");
  r4.value = `Report Date: ${dateStr}   |   Composite Score: ${data.composite}/100   |   Maturity: ${maturityLabel(data.composite)}`;
  r4.font  = cal({ size: 10, color: { argb: "FF" + DARK_BLUE } });
  r4.fill  = fill(LIGHT_BLUE);
  r4.alignment = center();

  ws.getRow(5).height = 8;

  // Row 6: Big score badge
  ws.getRow(6).height = 60;
  ws.mergeCells("B6:C6");
  const r6a = ws.getCell("B6");
  r6a.value = String(Math.round(data.composite));
  r6a.font  = cal({ bold: true, size: 36, color: { argb: "FFFFFFFF" } });
  r6a.fill  = fill(scoreColor(data.composite));
  r6a.alignment = center(true);

  ws.mergeCells("D6:F6");
  const r6b = ws.getCell("D6");
  r6b.value = `COMPOSITE DDMR SCORE — ${maturityLabel(data.composite)}\n(Weighted across ${dims.length} Digital Dimensions)`;
  r6b.font  = cal({ bold: true, size: 11, color: { argb: "FF" + DARK_BLUE } });
  r6b.alignment = { horizontal: "left", vertical: "middle", wrapText: true };

  ws.getRow(7).height = 8;

  // Row 8: Column headers
  ws.getRow(8).height = 22;
  const hdrCols = ["Dimension", "Score /100", "Weight", "Weighted", "Maturity Level"];
  ["B","C","D","E","F"].forEach((col, i) => {
    const c = ws.getCell(`${col}8`);
    c.value = hdrCols[i];
    c.font  = cal({ bold: true, size: 10, color: { argb: "FFFFFFFF" } });
    c.fill  = fill(DARK_BLUE);
    c.alignment = center();
    c.border = THIN;
  });

  // Rows 9+: Dimensions
  dims.forEach((d: any, i: number) => {
    const row = 9 + i;
    ws.getRow(row).height = 20;
    const wt = parseWeight(d.weight);
    const weighted = Math.round(d.score * wt * 10) / 10;
    const vals: any[] = [d.name, d.score, `${Math.round(wt * 100)}%`, weighted, maturityLabel(d.score)];
    ["B","C","D","E","F"].forEach((col, j) => {
      const c = ws.getCell(`${col}${row}`);
      c.value = vals[j];
      c.font  = col === "F"
        ? cal({ bold: true, size: 10, color: { argb: "FF" + scoreColor(d.score) } })
        : cal({ size: 10 });
      c.alignment = j === 0
        ? { horizontal: "left", vertical: "middle" }
        : center();
      c.border = THIN;
      if (col === "C" || col === "E") {
        c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "80" + scoreColor(d.score) } };
      }
    });
  });

  // Total row
  const tr = 9 + dims.length;
  ws.getRow(tr).height = 22;
  const totalVals: any[] = ["COMPOSITE SCORE", Math.round(data.composite), "100%", Math.round(data.composite * 10) / 10, maturityLabel(data.composite)];
  ["B","C","D","E","F"].forEach((col, j) => {
    const c = ws.getCell(`${col}${tr}`);
    c.value = totalVals[j];
    c.font  = cal({ bold: true, size: 11, color: { argb: "FFFFFFFF" } });
    c.fill  = fill(MID_BLUE);
    c.alignment = j === 0 ? { horizontal: "left", vertical: "middle" } : center();
    c.border = THIN;
  });

  // ── Sheets 2-6: Per-dimension ────────────────────────────────────────────────
  dims.forEach((d: any) => {
    const ws2 = wb.addWorksheet(`${d.code} Detail`);
    ws2.views = [{ showGridLines: false }];
    ws2.getColumn("A").width = 4;
    ws2.getColumn("B").width = 8;
    ws2.getColumn("C").width = 40;
    ws2.getColumn("D").width = 12;
    ws2.getColumn("E").width = 50;

    ws2.getRow(2).height = 24;
    ws2.mergeCells("B2:E2");
    const h2 = ws2.getCell("B2");
    h2.value = `${d.code} — ${d.name}   |   Score: ${d.score}/100   |   ${maturityLabel(d.score)}`;
    h2.font  = cal({ bold: true, size: 13, color: { argb: "FFFFFFFF" } });
    h2.fill  = fill(MID_BLUE);
    h2.alignment = center();

    ws2.getRow(4).height = 20;
    const smHdrs = ["Code", "Sub-Metric", "Score /100", "Public Signal / Evidence"];
    ["B","C","D","E"].forEach((col, i) => {
      const c = ws2.getCell(`${col}4`);
      c.value = smHdrs[i];
      c.font  = cal({ bold: true, size: 10, color: { argb: "FFFFFFFF" } });
      c.fill  = fill(DARK_BLUE);
      c.alignment = center();
      c.border = THIN;
    });

    const sms: any[] = d.submetrics || [];
    sms.forEach((sm: any, j: number) => {
      const row = 5 + j;
      ws2.getRow(row).height = 32;
      ["B","C","D","E"].forEach((col) => {
        const c = ws2.getCell(`${col}${row}`);
        c.border = THIN;
        c.alignment = { horizontal: "left", vertical: "middle", wrapText: true };
        if (col === "B") {
          c.value = sm.code;
          c.font  = cal({ bold: true, size: 9, color: { argb: "FF" + MID_BLUE } });
          c.alignment = center();
        } else if (col === "C") {
          c.value = sm.name;
          c.font  = cal({ size: 9 });
        } else if (col === "D") {
          c.value = sm.score;
          c.font  = cal({ bold: true, size: 11, color: { argb: "FF" + scoreColor(sm.score) } });
          c.alignment = center();
        } else {
          c.value = sm.evidence || "";
          c.font  = cal({ size: 9 });
        }
      });
    });

    // Average row
    const avgRow = 5 + sms.length;
    ws2.getRow(avgRow).height = 22;
    [["B","AVG"],["C","Dimension Average Score"],["D", d.score],["E",`Maturity: ${maturityLabel(d.score)}`]].forEach(([col, val]) => {
      const c = ws2.getCell(`${col}${avgRow}`);
      c.value = val;
      c.font  = cal({ bold: true, size: 10, color: { argb: "FFFFFFFF" } });
      c.fill  = fill(MID_BLUE);
      c.border = THIN;
      c.alignment = col === "D" ? center() : { horizontal: "left", vertical: "middle" };
    });
  });

  // ── Sheet 7: Evidence Index ──────────────────────────────────────────────────
  const wsE = wb.addWorksheet("Evidence Index");
  wsE.views = [{ showGridLines: false }];
  wsE.getColumn("A").width = 4;
  wsE.getColumn("B").width = 10;
  wsE.getColumn("C").width = 35;
  wsE.getColumn("D").width = 20;
  wsE.getColumn("E").width = 55;

  wsE.getRow(2).height = 24;
  wsE.mergeCells("B2:E2");
  const eH = wsE.getCell("B2");
  eH.value = `Evidence Index — ${data.company} — Digital Diagnostic and Maturity Report`;
  eH.font  = cal({ bold: true, size: 13, color: { argb: "FFFFFFFF" } });
  eH.fill  = fill(DARK_BLUE);
  eH.alignment = center();

  wsE.getRow(4).height = 20;
  ["Code","Sub-Metric","Source","Public Signal / Evidence"].forEach((lbl, i) => {
    const col = ["B","C","D","E"][i];
    const c = wsE.getCell(`${col}4`);
    c.value = lbl;
    c.font  = cal({ bold: true, size: 10, color: { argb: "FFFFFFFF" } });
    c.fill  = fill(MID_BLUE);
    c.border = THIN;
    c.alignment = center();
  });

  let evRow = 5;
  dims.forEach((d: any) => {
    (d.submetrics || []).forEach((sm: any) => {
      wsE.getRow(evRow).height = 28;
      ["B","C","D","E"].forEach((col) => {
        const c = wsE.getCell(`${col}${evRow}`);
        c.border = THIN;
        c.alignment = { horizontal: "left", vertical: "middle", wrapText: true };
        if (col === "B") {
          c.value = sm.code;
          c.font  = cal({ bold: true, size: 9, color: { argb: "FF" + MID_BLUE } });
          c.alignment = center();
        } else if (col === "C") { c.value = sm.name; c.font = cal({ size: 9 }); }
        else if (col === "D") { c.value = "Public Research"; c.font = cal({ size: 9 }); }
        else { c.value = sm.evidence || ""; c.font = cal({ size: 9 }); }
      });
      evRow++;
    });
  });

  const buf = await wb.xlsx.writeBuffer();
  return Buffer.from(buf as ArrayBuffer);
}

// ── DOCX (matching Python generate_report.py styling) ─────────────────────────

async function buildDocx(data: any): Promise<Buffer> {
  const {
    Document, Packer, Paragraph, TextRun, AlignmentType,
    Table, TableRow, TableCell, ShadingType, WidthType,
    PageBreak, convertInchesToTwip, BorderStyle, HeadingLevel,
  } = await import("docx");

  const Cm = (n: number) => convertInchesToTwip(n / 2.54);
  const Pt = (n: number) => n * 20;

  function darkCell(text: string, bold = true, fontSize = 18): any {
    return new TableCell({
      shading: { type: ShadingType.CLEAR, color: "auto", fill: DARK_BLUE },
      margins: { top: Cm(0.1), bottom: Cm(0.1), left: Cm(0.15), right: Cm(0.1) },
      children: [new Paragraph({
        children: [new TextRun({ text, bold, size: fontSize, color: "FFFFFF" })],
      })],
    });
  }
  function colorCell(text: string, bgHex: string, textHex = "1F3864", bold = false, fontSize = 18): any {
    return new TableCell({
      shading: { type: ShadingType.CLEAR, color: "auto", fill: bgHex },
      margins: { top: Cm(0.1), bottom: Cm(0.1), left: Cm(0.15), right: Cm(0.1) },
      children: [new Paragraph({
        children: [new TextRun({ text, bold, size: fontSize, color: textHex })],
      })],
    });
  }
  function midCell(text: string, bold = true, fontSize = 18): any {
    return new TableCell({
      shading: { type: ShadingType.CLEAR, color: "auto", fill: MID_BLUE },
      margins: { top: Cm(0.1), bottom: Cm(0.1), left: Cm(0.15), right: Cm(0.1) },
      children: [new Paragraph({
        children: [new TextRun({ text, bold, size: fontSize, color: "FFFFFF" })],
      })],
    });
  }
  function lightCell(text: string, bgHex: string, fontSize = 18): any {
    return new TableCell({
      shading: { type: ShadingType.CLEAR, color: "auto", fill: bgHex },
      margins: { top: Cm(0.1), bottom: Cm(0.1), left: Cm(0.15), right: Cm(0.1) },
      children: [new Paragraph({
        children: [new TextRun({ text, size: fontSize, color: "1F3864" })],
      })],
    });
  }

  const noBorder = {
    top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  };

  const dateStr = new Date().toLocaleDateString("en-IN", { day: "numeric", month: "long", year: "numeric" });
  const dims: any[] = data.dimensions || [];

  const children: any[] = [];

  // ── Cover page ──────────────────────────────────────────────────────────────
  children.push(new Paragraph({ children: [] }));
  children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 400, after: 160 },
    children: [new TextRun({ text: "DIGITAL DIAGNOSTIC AND MATURITY REPORT (DDMR)", bold: true, size: Pt(22), color: DARK_BLUE })],
  }));
  children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 160 },
    children: [new TextRun({ text: data.company || "", bold: true, size: Pt(18), color: MID_BLUE })],
  }));
  children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 160 },
    children: [new TextRun({ text: `Sector: ${data.sector || ""}   |   Report Date: ${dateStr}`, size: Pt(11) })],
  }));
  children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 400 },
    children: [new TextRun({
      text: `Composite DDMR Score: ${data.composite}/100  —  ${maturityLabel(data.composite)}`,
      bold: true, size: Pt(14), color: scoreColor(data.composite),
    })],
  }));
  children.push(new Paragraph({ children: [new PageBreak()] }));

  // ── Section 1: Executive Summary ────────────────────────────────────────────
  children.push(new Paragraph({
    text: "1. Executive Summary",
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 240, after: 120 },
  }));
  children.push(new Paragraph({
    spacing: { after: 120 },
    children: [new TextRun({ text: data.summary || "", size: Pt(11) })],
  }));

  // Dimension summary table
  const dimTableRows: any[] = [
    new TableRow({
      children: ["Dimension", "Score /100", "Weight", "Maturity Level"].map(h => darkCell(h, true, 20)),
    }),
    ...dims.map((d: any) => new TableRow({
      children: [
        lightCell(d.name, maturityLight(d.score)),
        lightCell(String(d.score), maturityLight(d.score)),
        lightCell(typeof d.weight === "string" ? d.weight : "20%", maturityLight(d.score)),
        lightCell(maturityLabel(d.score), maturityLight(d.score)),
      ],
    })),
    new TableRow({
      children: [
        midCell("COMPOSITE SCORE"),
        midCell(String(Math.round(data.composite))),
        midCell("100%"),
        midCell(maturityLabel(data.composite)),
      ],
    }),
  ];
  children.push(new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: dimTableRows,
  }));
  children.push(new Paragraph({ children: [new PageBreak()] }));

  // Maturity scale reference
  children.push(new Paragraph({
    text: "Maturity Scale Reference",
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 200, after: 100 },
  }));
  children.push(new Paragraph({
    spacing: { after: 200 },
    children: [new TextRun({
      text: "Legacy (0–33) — minimal digital presence; Siloed (34–66) — digital tools adopted in isolation; Strategic (67–83) — deliberate digital strategy with cross-functional integration; Future Ready (84–100) — continuous digital innovation, data-driven, platform-enabled.",
      size: Pt(10),
    })],
  }));
  children.push(new Paragraph({ children: [new PageBreak()] }));

  // ── Sections 2-6: Dimension analysis ────────────────────────────────────────
  const secNums = [2, 3, 4, 5, 6];
  dims.forEach((d: any, idx: number) => {
    children.push(new Paragraph({
      text: `${secNums[idx]}. ${d.name} (Score: ${d.score}/100 — ${maturityLabel(d.score)})`,
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 240, after: 120 },
    }));
    children.push(new Paragraph({
      spacing: { after: 160 },
      children: [new TextRun({ text: d.insight || "", size: Pt(11) })],
    }));

    // Sub-metrics table
    const smRows: any[] = [
      new TableRow({
        children: ["Sub-Metric", "Score", "Public Signal / Evidence"].map(h => midCell(h, true, 18)),
      }),
      ...(d.submetrics || []).map((sm: any) => {
        const bg = maturityLight(sm.score);
        return new TableRow({
          children: [
            lightCell(`${sm.code}: ${sm.name}`, bg),
            lightCell(`${sm.score}/100 (${maturityLabel(sm.score)})`, bg),
            lightCell(sm.evidence || "", bg),
          ],
        });
      }),
    ];
    children.push(new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      columnWidths: [Cm(5), Cm(3), Cm(9)],
      rows: smRows,
    }));
    children.push(new Paragraph({ children: [new PageBreak()] }));
  });

  // ── Section 7: Next Best Actions ─────────────────────────────────────────────
  children.push(new Paragraph({
    text: "7. Next Best Actions — Digital Transformation Roadmap",
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 240, after: 120 },
  }));
  children.push(new Paragraph({
    spacing: { after: 160 },
    children: [new TextRun({
      text: "Priority digital transformation initiatives based on the five-dimension DDMR diagnosis, sequenced by impact.",
      size: Pt(11),
    })],
  }));

  const ACTION_COLORS: Record<string, string> = {
    "Initiative": "D6E4F0",
    "Expected Benefit": "D5E8D4",
    "Cost of Inaction": "F8CECC",
    "Why Now": "DAE8FC",
  };

  (data.nextActions || []).forEach((a: any) => {
    children.push(new Paragraph({
      text: `#${a.rank} [${a.priority}] ${a.title}`,
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 240, after: 80 },
    }));
    children.push(new Paragraph({
      spacing: { after: 80 },
      children: [
        new TextRun({ text: "Addresses: ", bold: true, size: Pt(10) }),
        new TextRun({ text: a.dimension || "", size: Pt(10) }),
      ],
    }));

    const naRows: any[] = [
      ["Initiative", a.initiative || ""],
      ["Expected Benefit", a.benefit || ""],
      ["Cost of Inaction", a.loss || ""],
      ["Why Now", a.whyNow || ""],
    ].map(([label, content]) => new TableRow({
      children: [
        darkCell(label as string, true, 18),
        new TableCell({
          shading: { type: ShadingType.CLEAR, color: "auto", fill: ACTION_COLORS[label as string] || "FFFFFF" },
          margins: { top: Cm(0.1), bottom: Cm(0.1), left: Cm(0.15), right: Cm(0.1) },
          children: [new Paragraph({
            children: [new TextRun({ text: content as string, size: 18, color: "1F3864" })],
          })],
        }),
      ],
    }));

    // Blank separator row
    naRows.push(new TableRow({
      children: [
        new TableCell({ shading: { type: ShadingType.CLEAR, color: "auto", fill: "FFFFFF" }, children: [new Paragraph({ children: [] })], columnSpan: 2 }),
      ],
    }));

    children.push(new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      columnWidths: [Cm(3.5), Cm(13.5)],
      rows: naRows,
    }));
    children.push(new Paragraph({ spacing: { after: 80 }, children: [] }));
  });
  children.push(new Paragraph({ children: [new PageBreak()] }));

  // ── Section 12: Evidence & Citations ─────────────────────────────────────────
  children.push(new Paragraph({
    text: "12. Evidence & Source Citations",
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 240, after: 120 },
  }));
  children.push(new Paragraph({
    spacing: { after: 120 },
    children: [new TextRun({
      text: "All scores are derived exclusively from publicly available information. No management interviews, internal documents, or non-public data were accessed.",
      size: Pt(11),
    })],
  }));

  // Evidence bullets per sub-metric
  dims.forEach((d: any) => {
    (d.submetrics || []).forEach((sm: any) => {
      children.push(new Paragraph({
        bullet: { level: 0 },
        spacing: { after: 60 },
        children: [
          new TextRun({ text: `${sm.code} ${sm.name}: `, bold: true, size: Pt(10) }),
          new TextRun({ text: sm.evidence || "", size: Pt(10) }),
        ],
      }));
    });
  });

  children.push(new Paragraph({ spacing: { after: 160 }, children: [] }));
  children.push(new Paragraph({
    text: "Disclaimer",
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 200, after: 100 },
  }));
  children.push(new Paragraph({
    spacing: { after: 200 },
    children: [new TextRun({
      text: "This DDMR is prepared for informational purposes only. Scores are based on publicly available data and analytical judgment. They do not constitute investment advice, credit assessment, or a definitive representation of the company's digital capabilities.",
      size: Pt(10), italics: true, color: "888888",
    })],
  }));

  const doc = new Document({
    sections: [{
      properties: {
        page: {
          margin: {
            top: Cm(2), bottom: Cm(2),
            left: Cm(2.5), right: Cm(2.5),
          },
        },
      },
      children,
    }],
  });
  return await Packer.toBuffer(doc);
}

// ── PPTX (matching Python generate_deck.py) ─────────────────────────────────────

async function buildPptx(data: any): Promise<Buffer> {
  const pptxgen = (await import("pptxgenjs")).default;
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_WIDE"; // 13.33" × 7.5"

  const W = 13.33, H = 7.5;

  // Colors
  const DB = DARK_BLUE, MB = MID_BLUE, LB = LIGHT_BLUE;
  function sCol(s: number): string { return scoreColor(s); }

  function rect(slide: any, x: number, y: number, w: number, h: number, color: string, lineColor?: string) {
    slide.addShape(pptxgen.ShapeType.rect, {
      x, y, w, h,
      fill: { color },
      line: lineColor ? { color: lineColor, width: 0.5 } : { color, width: 0 },
    });
  }
  function txt(slide: any, text: string, opts: any) {
    slide.addText(text, { wrap: true, ...opts });
  }

  const dims: any[] = data.dimensions || [];
  const dateStr = new Date().toLocaleDateString("en-IN", { day: "numeric", month: "long", year: "numeric" });
  const ml = maturityLabel(data.composite);

  // ── Slide 1: Title ──────────────────────────────────────────────────────────
  const s1 = pptx.addSlide();
  rect(s1, 0, 0, W, H, DB);
  rect(s1, 0, 4.8, W, 2.7, MB);

  txt(s1, "DIGITAL DIAGNOSTIC AND MATURITY REPORT", {
    x: 0.8, y: 1.0, w: 11.7, h: 0.8,
    fontSize: 28, bold: true, color: "FFFFFF", align: "center",
  });
  txt(s1, data.company || "", {
    x: 0.8, y: 1.9, w: 11.7, h: 0.9,
    fontSize: 22, bold: true, color: LB, align: "center",
  });
  txt(s1, data.sector || "", {
    x: 0.8, y: 2.9, w: 11.7, h: 0.5,
    fontSize: 12, color: LB, align: "center",
  });

  // Score badge
  rect(s1, 4.9, 3.8, 3.5, 1.2, sCol(data.composite));
  txt(s1, `${data.composite}/100`, {
    x: 4.9, y: 3.8, w: 3.5, h: 0.7,
    fontSize: 32, bold: true, color: "FFFFFF", align: "center", valign: "middle",
  });
  txt(s1, `COMPOSITE DDMR SCORE — ${ml}`, {
    x: 4.9, y: 4.5, w: 3.5, h: 0.4,
    fontSize: 10, bold: true, color: "FFFFFF", align: "center",
  });

  txt(s1, `Report Date: ${dateStr}   |   Confidential`, {
    x: 0.8, y: 5.2, w: 11.7, h: 0.4,
    fontSize: 10, color: "FFFFFF", align: "center", italic: true,
  });

  // ── Slide 2: Company Snapshot ───────────────────────────────────────────────
  const s2 = pptx.addSlide();
  rect(s2, 0, 0, W, 1.1, DB);
  txt(s2, "Company Snapshot", { x: 0.3, y: 0.15, w: 12.7, h: 0.8, fontSize: 20, bold: true, color: "FFFFFF" });

  const meta = data.meta || {};
  const facts: [string, string][] = [
    ["Founded", meta.founded || ""],
    ["Headquarters", meta.hq || ""],
    ["Revenue", meta.revenue || ""],
    ["Employees", meta.employees || ""],
    ["Products", meta.products || ""],
    ["Customers", meta.customers || ""],
    ["Markets", meta.countries || ""],
    ["Sector", data.sector || ""],
  ];
  const cols2 = [facts.slice(0, 4), facts.slice(4)];
  cols2.forEach((col, ci) => {
    col.forEach(([label, val], ri) => {
      const lx = 0.4 + ci * 6.5;
      const ty = 1.3 + ri * 1.4;
      rect(s2, lx, ty, 2.3, 0.55, MB);
      txt(s2, label, { x: lx + 0.05, y: ty + 0.05, w: 2.2, h: 0.45, fontSize: 9, bold: true, color: "FFFFFF" });
      rect(s2, lx + 2.35, ty, 3.8, 0.55, LB);
      txt(s2, val, { x: lx + 2.4, y: ty + 0.05, w: 3.7, h: 0.45, fontSize: 10, color: DB });
    });
  });

  // ── Slide 3: DDMR Scorecard ─────────────────────────────────────────────────
  const s3 = pptx.addSlide();
  rect(s3, 0, 0, W, 1.1, DB);
  txt(s3, "DDMR Digital Maturity Scorecard", { x: 0.3, y: 0.15, w: 12.7, h: 0.8, fontSize: 20, bold: true, color: "FFFFFF" });

  // Dimension score panel on right
  rect(s3, 8.3, 1.1, 4.65, 0.45, MB);
  txt(s3, "Dimension Scores", { x: 8.3, y: 1.1, w: 4.65, h: 0.45, fontSize: 10, bold: true, color: "FFFFFF", align: "center" });

  dims.forEach((d: any, i: number) => {
    const ty = 1.6 + i * 1.1;
    rect(s3, 8.3, ty, 0.6, 0.7, sCol(d.score));
    txt(s3, String(d.score), { x: 8.3, y: ty, w: 0.6, h: 0.7, fontSize: 16, bold: true, color: "FFFFFF", align: "center", valign: "middle" });
    rect(s3, 8.95, ty, 4.0, 0.7, LB);
    txt(s3, `${d.code}: ${d.name}  [${maturityLabel(d.score)}]`, { x: 9.0, y: ty + 0.08, w: 3.9, h: 0.55, fontSize: 10, color: DB });
  });

  // Composite row
  rect(s3, 8.3, 7.0, 4.65, 0.45, sCol(data.composite));
  txt(s3, `Composite: ${data.composite}/100 — ${ml}`, {
    x: 8.3, y: 7.0, w: 4.65, h: 0.45, fontSize: 11, bold: true, color: "FFFFFF", align: "center",
  });

  // Score bars left panel
  txt(s3, "Score Summary", { x: 0.4, y: 1.15, w: 7.5, h: 0.4, fontSize: 12, bold: true, color: DB });
  dims.forEach((d: any, i: number) => {
    const by = 1.65 + i * 1.0;
    txt(s3, `${d.code}: ${d.name}`, { x: 0.4, y: by, w: 4.0, h: 0.35, fontSize: 10, color: DB });
    // Background bar
    rect(s3, 4.5, by + 0.05, 2.8, 0.28, LB);
    // Score bar
    const bw = 2.8 * d.score / 100;
    rect(s3, 4.5, by + 0.05, bw, 0.28, sCol(d.score));
    txt(s3, String(d.score), { x: 7.35, y: by, w: 0.5, h: 0.35, fontSize: 10, bold: true, color: sCol(d.score) });
    txt(s3, maturityLabel(d.score), { x: 0.4, y: by + 0.35, w: 7.5, h: 0.25, fontSize: 8, color: "666666" });
  });

  // ── Slides 4-8: Per dimension ───────────────────────────────────────────────
  dims.forEach((d: any) => {
    const slide = pptx.addSlide();
    const dc = sCol(d.score);

    // Header bar
    rect(slide, 0, 0, W, 1.1, dc);
    txt(slide, `${d.code}: ${d.name}`, { x: 0.3, y: 0.1, w: 9.0, h: 0.55, fontSize: 20, bold: true, color: "FFFFFF" });
    txt(slide, `Score: ${d.score}/100   |   Maturity: ${maturityLabel(d.score)}   |   Weight: ${typeof d.weight === "string" ? d.weight : "20%"}`,
      { x: 0.3, y: 0.6, w: 9.0, h: 0.4, fontSize: 12, color: "FFFFFF" });

    // Score badge
    rect(slide, 10.8, 0.1, 2.2, 0.9, DB);
    txt(slide, `${d.score}/100`, { x: 10.8, y: 0.1, w: 2.2, h: 0.9, fontSize: 24, bold: true, color: "FFFFFF", align: "center", valign: "middle" });

    // Insight
    txt(slide, d.insight || "", { x: 0.3, y: 1.15, w: 8.3, h: 0.6, fontSize: 10, color: "404040" });

    // Sub-metric bars (right panel)
    const px = 8.8;
    rect(slide, px, 1.1, 4.3, 0.4, MB);
    txt(slide, "Sub-Metric Scores", { x: px, y: 1.1, w: 4.3, h: 0.4, fontSize: 10, bold: true, color: "FFFFFF", align: "center" });

    const sms: any[] = d.submetrics || [];
    sms.forEach((sm: any, si: number) => {
      const sy = 1.6 + si * 0.65;
      txt(slide, sm.name, { x: px, y: sy, w: 2.0, h: 0.52, fontSize: 9, color: DB });
      // Background bar
      rect(slide, px + 2.05, sy + 0.12, 1.8, 0.28, LB);
      // Score bar
      const bw2 = 1.8 * sm.score / 100;
      rect(slide, px + 2.05, sy + 0.12, bw2, 0.28, sCol(sm.score));
      txt(slide, String(sm.score), { x: px + 3.9, y: sy, w: 0.4, h: 0.52, fontSize: 9, bold: true, color: sCol(sm.score) });
    });

    // Bullet points left
    const bullets = [
      ...(sms.slice(0, 3).map((sm: any) => `${sm.code}: ${sm.evidence || sm.name}`)),
      ...(d.highlights || []),
    ].filter(Boolean).slice(0, 5);

    bullets.forEach((b: string, bi: number) => {
      const by = 1.85 + bi * 0.87;
      rect(slide, 0.25, by + 0.15, 0.08, 0.22, dc);
      txt(slide, b, { x: 0.45, y: by, w: 8.1, h: 0.75, fontSize: 10, color: "404040" });
    });

    // Key strength / Risk footer
    const posH = data.highlights?.find((h: any) => h.type === "positive")?.text || d.insight?.split(".")[0] || "";
    const negH = data.highlights?.find((h: any) => h.type === "warning")?.text || "";
    rect(slide, 0, 6.85, W * 0.6, 0.55, "00B050");
    txt(slide, `Strength: ${posH}`, { x: 0.15, y: 6.87, w: 7.8, h: 0.45, fontSize: 9, bold: true, color: "FFFFFF" });
    rect(slide, W * 0.6, 6.85, W * 0.4, 0.55, "FF0000");
    txt(slide, `Risk: ${negH || "See dimension analysis"}`, { x: W * 0.6 + 0.1, y: 6.87, w: 5.1, h: 0.45, fontSize: 9, bold: true, color: "FFFFFF" });
  });

  // ── Slide 9: Next Best Actions ──────────────────────────────────────────────
  const s9 = pptx.addSlide();
  rect(s9, 0, 0, W, 1.1, DB);
  txt(s9, "Next Best Actions — Digital Transformation Roadmap",
    { x: 0.3, y: 0.15, w: 12.7, h: 0.8, fontSize: 20, bold: true, color: "FFFFFF" });

  const colW = 3.1, gap = 0.13, startX = 0.3;
  const actions: any[] = (data.nextActions || []).slice(0, 4);

  actions.forEach((a: any, ci: number) => {
    const x = startX + ci * (colW + gap);
    const ac = (a.priorityColor || "#dc2626").replace("#", "");

    // Rank badge
    rect(s9, x, 1.15, colW, 0.42, ac);
    txt(s9, `#${a.rank}  ${(a.priority || "").toUpperCase()}`, { x, y: 1.15, w: colW, h: 0.42, fontSize: 10, bold: true, color: "FFFFFF", align: "center" });

    // Title
    rect(s9, x, 1.6, colW, 0.52, MB);
    txt(s9, a.title || "", { x: x + 0.05, y: 1.62, w: colW - 0.1, h: 0.48, fontSize: 10, bold: true, color: "FFFFFF" });

    // Gap
    rect(s9, x, 2.15, colW, 0.32, LB);
    txt(s9, `Gap: ${a.dimension || ""}`, { x: x + 0.05, y: 2.16, w: colW - 0.1, h: 0.3, fontSize: 8, bold: true, color: DB });

    // 3 content rows
    const rows: [string, string, string, string][] = [
      ["Expected Benefit", a.benefit || "", "D5E8D4", "1F3864"],
      ["Cost of Inaction", a.loss || "", "F8CECC", "7F0000"],
      ["Why Now", a.whyNow || "", "DAE8FC", "1F3864"],
    ];
    rows.forEach(([label, text, bg, tc], ri) => {
      const rowColors: string[] = [ac, "FF0000", MB];
      const ry = 2.5 + ri * 1.55;
      rect(s9, x, ry, colW, 0.30, rowColors[ri]);
      txt(s9, label, { x: x + 0.05, y: ry, w: colW, h: 0.30, fontSize: 8, bold: true, color: "FFFFFF" });
      rect(s9, x, ry + 0.30, colW, 1.22, bg);
      txt(s9, text, { x: x + 0.06, y: ry + 0.33, w: colW - 0.12, h: 1.18, fontSize: 8, color: tc, valign: "top" });
    });
  });

  rect(s9, 0.3, 7.1, 12.7, 0.35, MB);
  txt(s9, `Maturity target after initiatives: ${maturityLabel(data.composite)} -> Strategic (67-83)`,
    { x: 0.3, y: 7.1, w: 12.7, h: 0.35, fontSize: 9, bold: true, color: "FFFFFF", align: "center" });

  // ── Slide 10: Executive Recommendation ─────────────────────────────────────
  const s10 = pptx.addSlide();
  rect(s10, 0, 0, W, 1.1, DB);
  txt(s10, "Executive Recommendation", { x: 0.3, y: 0.15, w: 12.7, h: 0.8, fontSize: 20, bold: true, color: "FFFFFF" });

  rect(s10, 0.4, 1.2, 12.5, 0.7, sCol(data.composite));
  txt(s10, `DDMR Composite: ${data.composite}/100 — ${ml.toUpperCase()}  |  ${data.summary?.split(".")[0] || ""}`,
    { x: 0.4, y: 1.2, w: 12.5, h: 0.7, fontSize: 12, bold: true, color: "FFFFFF", align: "center", valign: "middle" });

  const positives = (data.highlights || []).filter((h: any) => h.type === "positive").map((h: any) => h.text);
  const warnings  = (data.highlights || []).filter((h: any) => h.type === "warning").map((h: any) => h.text);

  const recPoints: [string, string][] = [
    ["Strengths", positives.slice(0, 2).join("; ") || "See dimension analysis"],
    ["Priority Gaps", warnings.slice(0, 2).join("; ") || "See dimension analysis"],
    ["Top Actions", (data.nextActions || []).slice(0, 2).map((a: any) => `${a.rank}. ${a.title}`).join("; ")],
    ["Opportunity", data.summary || ""],
  ];

  recPoints.forEach(([heading, body], ri) => {
    const ty = 2.1 + ri * 1.3;
    rect(s10, 0.4, ty, 2.8, 1.1, MB);
    txt(s10, heading, { x: 0.45, y: ty + 0.1, w: 2.7, h: 0.9, fontSize: 10, bold: true, color: "FFFFFF" });
    rect(s10, 3.3, ty, 9.6, 1.1, LB);
    txt(s10, body, { x: 3.4, y: ty + 0.05, w: 9.4, h: 1.0, fontSize: 9, color: DB });
  });

  txt(s10, `Generated: ${dateStr}  |  ${data.company}  |  DDMR v1.0`,
    { x: 0.4, y: 7.1, w: 12.5, h: 0.3, fontSize: 8, color: "888888", align: "center", italic: true });

  return await pptx.write({ outputType: "nodebuffer" }) as Buffer;
}

// ── PDF (pdf-lib with colored tables matching Python generate_pdf.py) ──────────

async function buildPdf(data: any): Promise<Buffer> {
  const { PDFDocument, rgb, StandardFonts } = await import("pdf-lib");

  function hex(h: string) {
    h = h.replace("#", "");
    return rgb(parseInt(h.slice(0, 2), 16) / 255, parseInt(h.slice(2, 4), 16) / 255, parseInt(h.slice(4, 6), 16) / 255);
  }

  const pdfDoc = await PDFDocument.create();
  const fontB = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
  const fontN = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const W = 595, H = 842, M = 50;

  const CBLUE = hex(DARK_BLUE), CMID = hex(MID_BLUE), CLIGHT = hex(LIGHT_BLUE);
  const CGRAY = hex("555555"), CTEXT = hex("1E293B"), CLGRAY = hex("F8F9FA");

  function newPage() {
    const page = pdfDoc.addPage([W, H]);
    return page;
  }

  function wrap(text: string, maxW: number, sz: number, f: any): string[] {
    const words = sanitize(text).split(" ");
    const lines: string[] = [];
    let line = "";
    for (const w of words) {
      const t = line ? `${line} ${w}` : w;
      if (f.widthOfTextAtSize(t, sz) > maxW && line) { lines.push(line); line = w; }
      else line = t;
    }
    if (line) lines.push(line);
    return lines;
  }

  function drawWrapped(page: any, text: string, x: number, y: number, sz: number, col: any, maxW: number, bold = false): number {
    const f = bold ? fontB : fontN;
    const lines = wrap(text, maxW, sz, f);
    for (const line of lines) {
      if (y < M + sz) break;
      page.drawText(line, { x, y, size: sz, font: f, color: col });
      y -= sz + 3;
    }
    return y;
  }

  function dt(page: any, text: string, opts: any) {
    page.drawText(sanitize(text), opts);
  }

  function colorRow(page: any, x: number, y: number, w: number, h: number, col: any) {
    page.drawRectangle({ x, y: y - h + 2, width: w, height: h, color: col });
  }

  function addFooter(page: any, pageNum: number) {
    page.drawRectangle({ x: 0, y: 0, width: W, height: M - 8, color: hex("F8F9FA") });
    dt(page, `DDMR - ${sanitize(data.company)} - Confidential - ${new Date().toLocaleDateString("en-IN")}`,
      { x: M, y: M - 26, size: 7, font: fontN, color: CGRAY });
    const pStr = `Page ${pageNum}`;
    dt(page, pStr, { x: W - M - fontB.widthOfTextAtSize(pStr, 7), y: M - 26, size: 7, font: fontB, color: CGRAY });
  }

  const dims: any[] = data.dimensions || [];
  let pageCount = 0;

  // ── Page 1: Cover ────────────────────────────────────────────────────────────
  {
    const page = newPage();
    pageCount++;

    // Header bar
    page.drawRectangle({ x: 0, y: H - 200, width: W, height: 200, color: CBLUE });

    dt(page, "DIGITAL DIAGNOSTIC AND MATURITY REPORT", { x: M, y: H - 38, size: 13, font: fontB, color: hex("FFFFFF") });
    dt(page, "(DDMR)", { x: M, y: H - 54, size: 9, font: fontN, color: hex("D6E4F0") });

    // Horizontal rule
    page.drawRectangle({ x: M, y: H - 70, width: W - M * 2, height: 2, color: CMID });

    let cy = H - 90;
    cy = drawWrapped(page, data.company || "", M, cy, 20, hex("FFFFFF"), W - M * 2, true) - 8;
    cy = drawWrapped(page, data.sector || "", M, cy, 10, hex("D6E4F0"), W - M * 2) - 4;

    // Score badge table
    const scoreW = W - M * 2;
    const col1 = 120, col2 = scoreW - col1;
    page.drawRectangle({ x: M, y: cy - 60, width: scoreW, height: 52, color: hex(scoreColor(data.composite)) });
    dt(page, `${data.composite}/100`, { x: M + 10, y: cy - 30, size: 28, font: fontB, color: hex("FFFFFF") });
    dt(page, `Composite DDMR Score`, { x: M + col1 + 10, y: cy - 22, size: 10, font: fontB, color: hex("FFFFFF") });
    dt(page, `${maturityLabel(data.composite)} Maturity Band`, { x: M + col1 + 10, y: cy - 38, size: 10, font: fontN, color: hex("FFFFFF") });
    cy -= 80;

    // Maturity scale bar
    const bands: [string, string, string][] = [
      ["Legacy 0-33", "FF0000", "FFFFFF"],
      ["Siloed 34-66", "FFC000", "FFFFFF"],
      ["Strategic 67-83", "00B050", "FFFFFF"],
      ["Future Ready 84-100", "2563EB", "FFFFFF"],
    ];
    const bW = (W - M * 2) / 4;
    bands.forEach(([lbl, bg, tc], i) => {
      page.drawRectangle({ x: M + i * bW, y: cy - 28, width: bW, height: 28, color: hex(bg) });
      dt(page, lbl, { x: M + i * bW + 4, y: cy - 18, size: 7, font: fontB, color: hex(tc) });
    });
    cy -= 48;

    // Company meta grid
    const metaItems: [string, string][] = [
      ["Founded", data.meta?.founded || ""],
      ["HQ", data.meta?.hq || ""],
      ["Revenue", data.meta?.revenue || ""],
      ["Employees", data.meta?.employees || ""],
      ["Products", data.meta?.products || ""],
      ["Customers", data.meta?.customers || ""],
    ].filter(([, v]) => v) as [string, string][];

    const mW = (W - M * 2) / 2;
    metaItems.forEach(([k, v], i) => {
      const mx = M + (i % 2) * mW;
      const my = cy - Math.floor(i / 2) * 24;
      page.drawRectangle({ x: mx, y: my - 20, width: mW - 4, height: 20, color: CLGRAY });
      dt(page, `${k}:`, { x: mx + 4, y: my - 14, size: 8, font: fontB, color: CTEXT });
      page.drawText(sanitize(v), { x: mx + 55, y: my - 14, size: 8, font: fontN, color: CGRAY, maxWidth: mW - 65 });
    });

    dt(page, `Report Date: ${new Date().toLocaleDateString("en-IN", { day: "numeric", month: "long", year: "numeric" })}`,
      { x: M, y: M, size: 8, font: fontN, color: CGRAY });

    addFooter(page, pageCount);
  }

  // ── Page 2: Executive Summary ─────────────────────────────────────────────────
  {
    const page = newPage();
    pageCount++;
    let y = H - M;

    dt(page, "1. Executive Summary", { x: M, y, size: 14, font: fontB, color: CBLUE }); y -= 22;
    y = drawWrapped(page, data.summary || "", M, y, 9, CGRAY, W - M * 2) - 14;

    // Dimension summary table
    dt(page, "Dimension Scores", { x: M, y, size: 11, font: fontB, color: CBLUE }); y -= 18;

    const tW = W - M * 2;
    const cW = [tW * 0.38, tW * 0.18, tW * 0.16, tW * 0.28];
    const hdrs = ["Dimension", "Score /100", "Weight", "Maturity Level"];

    // Header row
    page.drawRectangle({ x: M, y: y - 18, width: tW, height: 20, color: CBLUE });
    let cx2 = M;
    hdrs.forEach((h, i) => {
      dt(page, h, { x: cx2 + 4, y: y - 13, size: 8, font: fontB, color: hex("FFFFFF") });
      cx2 += cW[i];
    });
    y -= 20;

    // Dimension rows
    dims.forEach((d: any) => {
      const bg = hex(maturityLight(d.score));
      page.drawRectangle({ x: M, y: y - 16, width: tW, height: 18, color: bg });
      let cx3 = M;
      [d.name, String(d.score), typeof d.weight === "string" ? d.weight : "20%", maturityLabel(d.score)].forEach((v, i) => {
        const isScore = i === 1 || i === 3;
        dt(page, sanitize(v), { x: cx3 + 4, y: y - 12, size: 8, font: isScore ? fontB : fontN, color: isScore ? hex(scoreColor(d.score)) : CTEXT });
        cx3 += cW[i];
      });
      page.drawRectangle({ x: M, y: y - 16, width: tW, height: 0.5, color: hex("CCCCCC") });
      y -= 18;
    });

    // Composite row
    page.drawRectangle({ x: M, y: y - 18, width: tW, height: 20, color: CMID });
    let cx4 = M;
    [`COMPOSITE`, String(Math.round(data.composite)), "100%", maturityLabel(data.composite)].forEach((v, i) => {
      dt(page, v, { x: cx4 + 4, y: y - 13, size: 8, font: fontB, color: hex("FFFFFF") });
      cx4 += cW[i];
    });
    y -= 30;

    // Maturity reference
    dt(page, "Maturity Scale Reference", { x: M, y, size: 11, font: fontB, color: CBLUE }); y -= 16;
    const scaleText = "Legacy (0-33): minimal digital presence  |  Siloed (34-66): digital tools in isolation  |  Strategic (67-83): deliberate DX strategy  |  Future Ready (84-100): continuous innovation";
    y = drawWrapped(page, scaleText, M, y, 8, CGRAY, W - M * 2) - 10;

    addFooter(page, pageCount);
  }

  // ── Pages 3-7: Dimension deep-dives ──────────────────────────────────────────
  for (const d of dims) {
    const page = newPage();
    pageCount++;
    let y = H - M;

    // Header
    page.drawRectangle({ x: 0, y: H - 52, width: W, height: 52, color: hex("F8F9FA") });
    page.drawRectangle({ x: 0, y: H - 52, width: 6, height: 52, color: hex(scoreColor(d.score)) });
    dt(page, `${d.code}: ${d.name}`, { x: M, y: H - 30, size: 14, font: fontB, color: CBLUE });
    const scoreStr = `${d.score}/100 - ${maturityLabel(d.score)}`;
    dt(page, scoreStr, { x: W - M - fontB.widthOfTextAtSize(scoreStr, 11), y: H - 30, size: 11, font: fontB, color: hex(scoreColor(d.score)) });
    y = H - 66;
    y = drawWrapped(page, d.insight || "", M, y, 9, CGRAY, W - M * 2) - 14;

    // Sub-metrics table
    dt(page, "Sub-Metric Analysis", { x: M, y, size: 11, font: fontB, color: CBLUE }); y -= 18;

    const smCW = [(W - M * 2) * 0.12, (W - M * 2) * 0.30, (W - M * 2) * 0.10, (W - M * 2) * 0.15, (W - M * 2) * 0.33];
    const smHdrs = ["Code", "Sub-Metric", "Score", "Maturity", "Evidence"];

    // Header row
    page.drawRectangle({ x: M, y: y - 18, width: W - M * 2, height: 20, color: CMID });
    let smCx = M;
    smHdrs.forEach((h, i) => {
      dt(page, h, { x: smCx + 3, y: y - 13, size: 8, font: fontB, color: hex("FFFFFF") });
      smCx += smCW[i];
    });
    y -= 20;

    for (const sm of (d.submetrics || [])) {
      if (y < 100) break;
      const bg = hex(maturityLight(sm.score));
      const textLines = wrap(sm.evidence || "", smCW[4] - 8, 7.5, fontN);
      const rowH = Math.max(20, textLines.length * 11 + 8);

      page.drawRectangle({ x: M, y: y - rowH + 2, width: W - M * 2, height: rowH, color: bg });
      page.drawRectangle({ x: M, y: y - rowH + 2, width: W - M * 2, height: 0.5, color: hex("CCCCCC") });

      let smCx2 = M;
      const smVals: string[] = [sm.code, sm.name, `${sm.score}`, maturityLabel(sm.score), sm.evidence || ""];
      smVals.forEach((v, i) => {
        if (i === 4) {
          textLines.forEach((line, li) => {
            dt(page, line, { x: smCx2 + 3, y: y - 10 - li * 11, size: 7.5, font: fontN, color: CGRAY });
          });
        } else {
          dt(page, sanitize(v), {
            x: smCx2 + 3, y: y - 12, size: i === 0 || i === 2 ? 8.5 : 8,
            font: i === 2 || i === 0 ? fontB : fontN,
            color: i === 2 || i === 3 ? hex(scoreColor(sm.score)) : CTEXT,
          });
        }
        smCx2 += smCW[i];
      });
      y -= rowH;
    }

    addFooter(page, pageCount);
  }

  // ── Next Actions page ─────────────────────────────────────────────────────────
  {
    const page = newPage();
    pageCount++;

    page.drawRectangle({ x: 0, y: H - 52, width: W, height: 52, color: CBLUE });
    dt(page, "7. Next Best Actions - Digital Transformation Roadmap",
      { x: M, y: H - 32, size: 13, font: fontB, color: hex("FFFFFF") });
    let y = H - 68;

    const tW2 = W - M * 2;
    const lblW = 100;

    for (const a of (data.nextActions || [])) {
      if (y < 120) break;
      const pc = hex((a.priorityColor || "#dc2626").replace("#", ""));

      // Action header
      page.drawRectangle({ x: M, y: y - 22, width: tW2, height: 24, color: CBLUE });
      dt(page, sanitize(`#${a.rank} [${a.priority}] ${a.title || ""}`),
        { x: M + 6, y: y - 16, size: 10, font: fontB, color: hex("FFFFFF") });
      y -= 24;

      // 4 content rows
      const rows2: [string, string, string][] = [
        ["Initiative", a.initiative || "", "F8F9FA"],
        ["Expected Benefit", a.benefit || "", "D5E8D4"],
        ["Cost of Inaction", a.loss || "", "F8CECC"],
        ["Why Now", a.whyNow || "", "DAE8FC"],
      ];
      const lblColors = [CMID, hex("00B050"), hex("FF0000"), CMID];

      rows2.forEach(([label, content, bg], ri) => {
        const cLines = wrap(content, tW2 - lblW - 12, 8, fontN);
        const rowH2 = Math.max(20, cLines.length * 11 + 8);
        if (y < 60) return;

        page.drawRectangle({ x: M, y: y - rowH2 + 2, width: lblW, height: rowH2, color: lblColors[ri] });
        page.drawRectangle({ x: M + lblW, y: y - rowH2 + 2, width: tW2 - lblW, height: rowH2, color: hex(bg) });
        page.drawRectangle({ x: M, y: y - rowH2 + 2, width: tW2, height: 0.5, color: hex("CCCCCC") });

        dt(page, label, { x: M + 4, y: y - 12, size: 8, font: fontB, color: hex("FFFFFF") });
        cLines.forEach((line, li) => {
          dt(page, line, { x: M + lblW + 6, y: y - 10 - li * 11, size: 8, font: fontN, color: CTEXT });
        });
        y -= rowH2;
      });

      y -= 10;
    }

    dt(page, "Based on publicly available information - AI-generated - Verify independently before decisions",
      { x: M, y: M, size: 7.5, font: fontN, color: CGRAY });

    addFooter(page, pageCount);
  }

  return Buffer.from(await pdfDoc.save());
}

// ── Router ─────────────────────────────────────────────────────────────────────

export async function POST(req: NextRequest) {
  const { format, data } = await req.json();
  if (!data || !format) return NextResponse.json({ error: "Missing format or data." }, { status: 400 });
  const filename = safeFilename(data.company || "DDMR_Report");

  function respond(buf: Buffer, ct: string, ext: string) {
    return new NextResponse(new Uint8Array(buf), {
      headers: {
        "Content-Type": ct,
        "Content-Disposition": `attachment; filename="${filename}_DDMR.${ext}"`,
      },
    });
  }

  try {
    if (format === "xlsx") return respond(await buildXlsx(data), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "xlsx");
    if (format === "docx") return respond(await buildDocx(data), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "docx");
    if (format === "pptx") return respond(await buildPptx(data), "application/vnd.openxmlformats-officedocument.presentationml.presentation", "pptx");
    if (format === "pdf")  return respond(await buildPdf(data),  "application/pdf", "pdf");
    return NextResponse.json({ error: "Unsupported format." }, { status: 400 });
  } catch (err) {
    console.error("Export error:", err);
    return NextResponse.json({ error: "Export failed. Please try again." }, { status: 500 });
  }
}
