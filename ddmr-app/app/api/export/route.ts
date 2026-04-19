import { NextRequest, NextResponse } from "next/server";

export const maxDuration = 60;

function maturityLabel(s: number): string {
  if (s >= 84) return "Future Ready";
  if (s >= 67) return "Strategic";
  if (s >= 34) return "Siloed";
  return "Legacy";
}
function safeFilename(name: string): string {
  return (name || "DDMR_Report").replace(/[^a-z0-9]/gi, "_").slice(0, 60);
}

async function buildXlsx(data: any): Promise<Buffer> {
  const XLSX = await import("xlsx");
  const wb = XLSX.utils.book_new();

  const ws1 = XLSX.utils.aoa_to_sheet([
    ["DDMR Report — Digital Diagnostic & Maturity Report"],
    [],
    ["Company", data.company],
    ["Sector", data.sector],
    ["Composite Score", data.composite, "/ 100"],
    ["Maturity Level", maturityLabel(data.composite)],
    ["Report Date", new Date().toLocaleDateString("en-IN")],
    [],
    ["EXECUTIVE SUMMARY"],
    [data.summary],
    [],
    ["COMPANY PROFILE"],
    ["Founded", data.meta?.founded || ""],
    ["Headquarters", data.meta?.hq || ""],
    ["Revenue", data.meta?.revenue || ""],
    ["Employees", data.meta?.employees || ""],
    ["Products / Services", data.meta?.products || ""],
    ["Customer Segments", data.meta?.customers || ""],
    ["Markets", data.meta?.countries || ""],
  ]);
  ws1["!cols"] = [{ wch: 25 }, { wch: 70 }];
  XLSX.utils.book_append_sheet(wb, ws1, "Summary");

  const ws2 = XLSX.utils.aoa_to_sheet([
    ["Code", "Dimension", "Score", "Maturity", "Weight", "Trend", "Insight"],
    ...data.dimensions.map((d: any) => [d.code, d.name, d.score, maturityLabel(d.score), d.weight, d.trend, d.insight]),
  ]);
  ws2["!cols"] = [{ wch: 8 }, { wch: 28 }, { wch: 8 }, { wch: 14 }, { wch: 8 }, { wch: 8 }, { wch: 70 }];
  XLSX.utils.book_append_sheet(wb, ws2, "Dimensions");

  const ws3 = XLSX.utils.aoa_to_sheet([
    ["Dimension", "Code", "Sub-metric", "Score", "Maturity", "Evidence"],
    ...data.dimensions.flatMap((d: any) =>
      d.submetrics.map((sm: any) => [d.name, sm.code, sm.name, sm.score, maturityLabel(sm.score), sm.evidence])
    ),
  ]);
  ws3["!cols"] = [{ wch: 28 }, { wch: 8 }, { wch: 35 }, { wch: 8 }, { wch: 14 }, { wch: 80 }];
  XLSX.utils.book_append_sheet(wb, ws3, "Sub-metrics");

  const ws4 = XLSX.utils.aoa_to_sheet([
    ["Rank", "Priority", "Title", "Dimension", "Initiative", "Expected Benefit", "Cost of Inaction", "Why Now"],
    ...data.nextActions.map((a: any) => [a.rank, a.priority, a.title, a.dimension, a.initiative, a.benefit, a.loss, a.whyNow]),
  ]);
  ws4["!cols"] = Array(8).fill({ wch: 55 });
  XLSX.utils.book_append_sheet(wb, ws4, "Next Actions");

  const ws5 = XLSX.utils.aoa_to_sheet([
    ["Type", "Finding"],
    ...data.highlights.map((h: any) => [h.type === "positive" ? "Strength" : "Gap / Risk", h.text]),
  ]);
  ws5["!cols"] = [{ wch: 14 }, { wch: 90 }];
  XLSX.utils.book_append_sheet(wb, ws5, "Key Findings");

  return XLSX.write(wb, { type: "buffer", bookType: "xlsx" }) as Buffer;
}

async function buildDocx(data: any): Promise<Buffer> {
  const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType } = await import("docx");

  function h1(text: string) {
    return new Paragraph({ text, heading: HeadingLevel.HEADING_1, spacing: { before: 360, after: 120 } });
  }
  function h2(text: string) {
    return new Paragraph({ text, heading: HeadingLevel.HEADING_2, spacing: { before: 240, after: 80 } });
  }
  function para(text: string) {
    return new Paragraph({ children: [new TextRun({ text, size: 22 })], spacing: { after: 100 } });
  }
  function kv(label: string, value: string) {
    return new Paragraph({
      children: [new TextRun({ text: `${label}: `, bold: true, size: 22 }), new TextRun({ text: value, size: 22 })],
      spacing: { after: 60 },
    });
  }

  const children: any[] = [
    new Paragraph({
      children: [new TextRun({ text: "DDMR Report", bold: true, size: 52, color: "0f2442" })],
      alignment: AlignmentType.CENTER, spacing: { before: 800, after: 160 },
    }),
    new Paragraph({
      children: [new TextRun({ text: data.company, bold: true, size: 36 })],
      alignment: AlignmentType.CENTER, spacing: { after: 80 },
    }),
    new Paragraph({
      children: [new TextRun({ text: data.sector || "", size: 24, color: "555555" })],
      alignment: AlignmentType.CENTER, spacing: { after: 160 },
    }),
    new Paragraph({
      children: [new TextRun({ text: `Composite Score: ${data.composite}/100 — ${maturityLabel(data.composite)}`, bold: true, size: 28 })],
      alignment: AlignmentType.CENTER, spacing: { after: 80 },
    }),
    new Paragraph({
      children: [new TextRun({ text: `Generated: ${new Date().toLocaleDateString("en-IN", { day: "numeric", month: "long", year: "numeric" })}`, size: 20, color: "888888" })],
      alignment: AlignmentType.CENTER, spacing: { after: 600 },
    }),

    h1("Executive Summary"),
    para(data.summary || ""),

    h1("Company Profile"),
    ...[
      ["Founded", data.meta?.founded], ["Headquarters", data.meta?.hq],
      ["Revenue", data.meta?.revenue], ["Employees", data.meta?.employees],
      ["Products / Services", data.meta?.products], ["Customer Segments", data.meta?.customers],
      ["Markets Served", data.meta?.countries],
    ].filter(([, v]) => v).map(([l, v]) => kv(l as string, v as string)),

    h1("Key Findings"),
    ...data.highlights.map((h: any) =>
      new Paragraph({
        children: [
          new TextRun({ text: h.type === "positive" ? "✓  " : "⚠  ", color: h.type === "positive" ? "16a34a" : "d97706", bold: true }),
          new TextRun({ text: h.text, size: 22 }),
        ],
        spacing: { after: 80 },
      })
    ),

    h1("Dimension Analysis"),
    ...data.dimensions.flatMap((d: any) => [
      h2(`${d.code}: ${d.name} — ${d.score}/100 (${maturityLabel(d.score)})`),
      para(d.insight || ""),
      new Paragraph({ children: [new TextRun({ text: "Sub-metrics:", bold: true, size: 22 })], spacing: { after: 60 } }),
      ...d.submetrics.map((sm: any) =>
        new Paragraph({
          children: [
            new TextRun({ text: `${sm.code} ${sm.name}: ${sm.score}/100 — `, bold: true, size: 20 }),
            new TextRun({ text: sm.evidence || "", size: 20, color: "444444" }),
          ],
          spacing: { after: 60 }, indent: { left: 400 },
        })
      ),
    ]),

    h1("Next Best Actions"),
    ...data.nextActions.flatMap((a: any) => [
      h2(`#${a.rank} [${a.priority}] ${a.title}`),
      kv("Addresses", a.dimension || ""),
      para(a.initiative || ""),
      kv("Expected Benefit", a.benefit || ""),
      kv("Cost of Inaction", a.loss || ""),
      kv("Why Now", a.whyNow || ""),
    ]),

    new Paragraph({
      children: [new TextRun({ text: "Based on publicly available information · AI-generated · Verify independently before decisions", size: 18, color: "888888", italics: true })],
      spacing: { before: 600 },
    }),
  ];

  const doc = new Document({ sections: [{ children }] });
  return await Packer.toBuffer(doc);
}

async function buildPptx(data: any): Promise<Buffer> {
  const pptxgen = (await import("pptxgenjs")).default;
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_WIDE";

  const DARK = "0f2442";
  const BLUE = "2563eb";
  const scoreCol: Record<string, string> = {
    "Future Ready": "1d4ed8", "Strategic": "16a34a", "Siloed": "d97706", "Legacy": "dc2626",
  };

  // Slide 1: Cover
  const cover = pptx.addSlide();
  cover.background = { color: DARK };
  cover.addText("DDMR REPORT", { x: 0.5, y: 0.5, w: "90%", fontSize: 13, color: "60a5fa", bold: true, charSpacing: 4 });
  cover.addText("Digital Diagnostic & Maturity Report", { x: 0.5, y: 0.9, w: "90%", fontSize: 10, color: "93c5fd" });
  cover.addText(data.company, { x: 0.5, y: 1.6, w: "80%", fontSize: 30, color: "FFFFFF", bold: true });
  cover.addText(data.sector || "", { x: 0.5, y: 3.0, w: "80%", fontSize: 13, color: "93c5fd" });
  const ml = maturityLabel(data.composite);
  cover.addText(`${data.composite}/100  —  ${ml}`, {
    x: 0.5, y: 3.6, w: 3.5, h: 0.65, fontSize: 20, color: "FFFFFF", bold: true, align: "center",
    fill: { color: scoreCol[ml] || BLUE }, shape: pptxgen.ShapeType.roundRect, rectRadius: 0.06,
  });
  cover.addText(`Generated ${new Date().toLocaleDateString("en-IN", { day: "numeric", month: "long", year: "numeric" })}`,
    { x: 0.5, y: 6.8, w: "90%", fontSize: 9, color: "6b7280" });

  // Slide 2: Summary + scores overview
  const sumSlide = pptx.addSlide();
  sumSlide.addShape(pptxgen.ShapeType.rect, { x: 0, y: 0, w: "100%", h: 0.7, fill: { color: "f8fafc" } });
  sumSlide.addText("Executive Summary", { x: 0.5, y: 0.1, w: "90%", fontSize: 20, bold: true, color: DARK });
  sumSlide.addText(data.summary || "", { x: 0.5, y: 0.85, w: "90%", h: 1.3, fontSize: 11, color: "374151", valign: "top" });
  sumSlide.addText("Dimension Scores", { x: 0.5, y: 2.3, w: "90%", fontSize: 13, bold: true, color: DARK });
  const xs = [0.4, 2.85, 5.3, 7.75, 10.2];
  data.dimensions.forEach((d: any, i: number) => {
    const ml2 = maturityLabel(d.score);
    const c = scoreCol[ml2] || BLUE;
    sumSlide.addShape(pptxgen.ShapeType.roundRect, { x: xs[i], y: 2.65, w: 2.2, h: 1.4, fill: { color: "f8fafc" }, line: { color: "e2e8f0", width: 1 }, rectRadius: 0.06 });
    sumSlide.addText(d.score.toString(), { x: xs[i], y: 2.75, w: 2.2, fontSize: 26, bold: true, color: c, align: "center" });
    sumSlide.addText(d.name, { x: xs[i], y: 3.2, w: 2.2, fontSize: 8.5, color: "374151", align: "center" });
    sumSlide.addText(ml2, { x: xs[i], y: 3.5, w: 2.2, fontSize: 7.5, bold: true, color: c, align: "center" });
  });
  sumSlide.addText("Key Findings", { x: 0.5, y: 4.2, w: "90%", fontSize: 13, bold: true, color: DARK });
  data.highlights.slice(0, 5).forEach((h: any, i: number) => {
    sumSlide.addText(`${h.type === "positive" ? "✓" : "⚠"}  ${h.text}`,
      { x: 0.5, y: 4.6 + i * 0.4, w: "90%", fontSize: 9.5, color: h.type === "positive" ? "15803d" : "92400e" });
  });

  // Slides 3-7: One per dimension
  data.dimensions.forEach((d: any) => {
    const ml3 = maturityLabel(d.score);
    const c = scoreCol[ml3] || BLUE;
    const slide = pptx.addSlide();
    slide.addShape(pptxgen.ShapeType.rect, { x: 0, y: 0, w: "100%", h: 0.72, fill: { color: "f8fafc" } });
    slide.addText(`${d.code}  ·  ${d.name}`, { x: 0.5, y: 0.08, w: "72%", fontSize: 19, bold: true, color: DARK });
    slide.addText(`${d.score}/100  —  ${ml3}`, { x: "76%", y: 0.08, w: "22%", fontSize: 16, bold: true, color: c, align: "right" });
    slide.addText(d.insight || "", { x: 0.5, y: 0.85, w: "90%", fontSize: 11, color: "374151" });
    const rows: any[] = [
      [
        { text: "Code", options: { bold: true, fontSize: 9, fill: { color: "1e3a5f" }, color: "FFFFFF" } },
        { text: "Sub-metric", options: { bold: true, fontSize: 9, fill: { color: "1e3a5f" }, color: "FFFFFF" } },
        { text: "Score", options: { bold: true, fontSize: 9, fill: { color: "1e3a5f" }, color: "FFFFFF", align: "center" } },
        { text: "Evidence", options: { bold: true, fontSize: 9, fill: { color: "1e3a5f" }, color: "FFFFFF" } },
      ],
      ...d.submetrics.map((sm: any) => {
        const c2 = scoreCol[maturityLabel(sm.score)] || BLUE;
        return [
          { text: sm.code, options: { fontSize: 8.5 } },
          { text: sm.name, options: { fontSize: 8.5, bold: true } },
          { text: sm.score.toString(), options: { fontSize: 10, bold: true, color: c2, align: "center" } },
          { text: sm.evidence || "", options: { fontSize: 7.5, color: "374151" } },
        ];
      }),
    ];
    slide.addTable(rows, { x: 0.5, y: 1.35, w: 12.3, rowH: 0.5, colW: [0.9, 2.6, 0.7, 8.1], border: { type: "solid", color: "e2e8f0", pt: 0.5 } });
  });

  // Slide: Next Best Actions
  const nb = pptx.addSlide();
  nb.addShape(pptxgen.ShapeType.rect, { x: 0, y: 0, w: "100%", h: 0.72, fill: { color: DARK } });
  nb.addText("Next Best Actions", { x: 0.5, y: 0.1, w: "90%", fontSize: 21, bold: true, color: "FFFFFF" });
  const naXs = [0.3, 3.55, 6.8, 10.05];
  data.nextActions.slice(0, 4).forEach((a: any, i: number) => {
    const c2 = (a.priorityColor || "#dc2626").replace("#", "");
    nb.addShape(pptxgen.ShapeType.roundRect, { x: naXs[i], y: 0.88, w: 3.0, h: 5.9, fill: { color: "f8fafc" }, line: { color: "e2e8f0", width: 1 }, rectRadius: 0.07 });
    nb.addText(`#${a.rank}`, { x: naXs[i] + 0.1, y: 0.98, w: 2.8, fontSize: 10, color: c2, bold: true });
    nb.addText(a.priority?.toUpperCase() || "", { x: naXs[i] + 0.1, y: 1.2, w: 2.8, h: 0.28, fontSize: 8, color: "FFFFFF", bold: true, fill: { color: c2 }, align: "center" });
    nb.addText(a.title || "", { x: naXs[i] + 0.1, y: 1.56, w: 2.8, fontSize: 9.5, bold: true, color: "1e293b", wrap: true });
    nb.addText(a.initiative || "", { x: naXs[i] + 0.1, y: 2.3, w: 2.8, h: 1.3, fontSize: 7.5, color: "374151", wrap: true, valign: "top" });
    nb.addText("Expected Benefit", { x: naXs[i] + 0.1, y: 3.7, w: 2.8, fontSize: 7.5, bold: true, color: "15803d" });
    nb.addText(a.benefit || "", { x: naXs[i] + 0.1, y: 3.9, w: 2.8, h: 0.85, fontSize: 7, color: "374151", wrap: true, valign: "top" });
    nb.addText("Why Now", { x: naXs[i] + 0.1, y: 4.82, w: 2.8, fontSize: 7.5, bold: true, color: "1d4ed8" });
    nb.addText(a.whyNow || "", { x: naXs[i] + 0.1, y: 5.02, w: 2.8, h: 0.65, fontSize: 7, color: "374151", wrap: true, valign: "top" });
  });

  return await pptx.write({ outputType: "nodebuffer" }) as Buffer;
}

async function buildPdf(data: any): Promise<Buffer> {
  const { PDFDocument, rgb, StandardFonts } = await import("pdf-lib");

  // pdf-lib standard fonts only support Latin-1; replace common Unicode chars
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

  function hex(h: string) {
    const r = parseInt(h.slice(0, 2), 16) / 255;
    const g = parseInt(h.slice(2, 4), 16) / 255;
    const b = parseInt(h.slice(4, 6), 16) / 255;
    return rgb(r, g, b);
  }
  const scoreCol: Record<string, ReturnType<typeof rgb>> = {
    "Future Ready": hex("1d4ed8"), "Strategic": hex("16a34a"), "Siloed": hex("d97706"), "Legacy": hex("dc2626"),
  };

  const pdfDoc = await PDFDocument.create();
  const fontB = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
  const fontN = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const W = 595, H = 842, M = 48;
  const dark = hex("0f2442"), gray = hex("6b7280"), textDark = hex("1e293b");

  function newPage() {
    const page = pdfDoc.addPage([W, H]);
    return { page, cy: [H - M] };
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
  function drawWrapped(page: any, text: string, x: number, y: number, sz: number, col: ReturnType<typeof rgb>, maxW: number, bold = false): number {
    const f = bold ? fontB : fontN;
    for (const line of wrap(text, maxW, sz, f)) {
      if (y < M + 10) return y;
      page.drawText(line, { x, y, size: sz, font: f, color: col });
      y -= sz + 4;
    }
    return y;
  }
  function dt(page: any, text: string, opts: any) {
    page.drawText(sanitize(text), opts);
  }

  // Page 1: Cover
  {
    const { page } = newPage();
    page.drawRectangle({ x: 0, y: H - 210, width: W, height: 210, color: dark });
    dt(page, "DDMR REPORT", { x: M, y: H - 38, size: 12, font: fontB, color: hex("60a5fa") });
    dt(page, "Digital Diagnostic & Maturity Report", { x: M, y: H - 58, size: 9, font: fontN, color: hex("93c5fd") });
    let cy = H - 95;
    cy = drawWrapped(page, data.company, M, cy, 26, rgb(1, 1, 1), W - M * 2, true) - 8;
    cy = drawWrapped(page, data.sector || "", M, cy, 11, hex("93c5fd"), W - M * 2);
    const ml = maturityLabel(data.composite);
    const sc = scoreCol[ml] || hex("2563eb");
    page.drawRectangle({ x: M, y: cy - 50, width: 190, height: 38, color: sc });
    dt(page, `${data.composite}/100  -  ${ml}`, { x: M + 10, y: cy - 36, size: 14, font: fontB, color: rgb(1, 1, 1) });
    dt(page, `Report Date: ${new Date().toLocaleDateString("en-IN", { day: "numeric", month: "long", year: "numeric" })}`, { x: M, y: M, size: 8, font: fontN, color: gray });
  }

  // Page 2: Summary + Profile + Highlights
  {
    const { page } = newPage();
    let y = H - M;
    dt(page, "Executive Summary", { x: M, y, size: 15, font: fontB, color: dark }); y -= 22;
    y = drawWrapped(page, data.summary || "", M, y, 10, textDark, W - M * 2) - 12;
    dt(page, "Company Profile", { x: M, y, size: 13, font: fontB, color: dark }); y -= 18;
    for (const [k, v] of [["Founded", data.meta?.founded], ["HQ", data.meta?.hq], ["Revenue", data.meta?.revenue], ["Employees", data.meta?.employees], ["Products", data.meta?.products]].filter(([, v]) => v)) {
      dt(page, `${k}:`, { x: M, y, size: 9, font: fontB, color: textDark });
      dt(page, v as string, { x: M + 75, y, size: 9, font: fontN, color: textDark });
      y -= 14;
    }
    y -= 8;
    dt(page, "Key Findings", { x: M, y, size: 13, font: fontB, color: dark }); y -= 18;
    for (const h of data.highlights) {
      const c = h.type === "positive" ? hex("16a34a") : hex("d97706");
      dt(page, h.type === "positive" ? "+" : "!", { x: M, y, size: 10, font: fontB, color: c });
      y = drawWrapped(page, h.text, M + 14, y, 9.5, textDark, W - M * 2 - 14) - 4;
    }
    y -= 8;
    dt(page, "Dimension Scores", { x: M, y, size: 13, font: fontB, color: dark }); y -= 18;
    const colW2 = (W - M * 2) / 5;
    for (let i = 0; i < data.dimensions.length; i++) {
      const d = data.dimensions[i];
      const ml2 = maturityLabel(d.score);
      const c2 = scoreCol[ml2] || hex("2563eb");
      const x = M + i * colW2;
      page.drawRectangle({ x, y: y - 48, width: colW2 - 4, height: 50, color: hex("f8fafc") });
      dt(page, d.score.toString(), { x: x + colW2 / 2 - fontB.widthOfTextAtSize(d.score.toString(), 18) / 2, y: y - 14, size: 18, font: fontB, color: c2 });
      dt(page, d.name.split(" ")[0], { x: x + 4, y: y - 28, size: 7, font: fontN, color: textDark });
      dt(page, ml2, { x: x + 4, y: y - 40, size: 7, font: fontB, color: c2 });
    }
  }

  // Pages 3-7: Dimensions
  for (const d of data.dimensions) {
    const { page } = newPage();
    let y = H - M;
    const ml3 = maturityLabel(d.score);
    const c3 = scoreCol[ml3] || hex("2563eb");
    page.drawRectangle({ x: 0, y: H - 52, width: W, height: 52, color: hex("f8fafc") });
    dt(page, `${d.code}: ${d.name}`, { x: M, y: H - 32, size: 15, font: fontB, color: dark });
    dt(page, `${d.score}/100 - ${ml3}`, { x: W - M - 100, y: H - 32, size: 12, font: fontB, color: c3 });
    y = H - 66;
    y = drawWrapped(page, d.insight || "", M, y, 10, hex("374151"), W - M * 2) - 12;
    dt(page, "Sub-metric Analysis", { x: M, y, size: 12, font: fontB, color: dark }); y -= 18;
    for (const sm of d.submetrics) {
      if (y < 80) break;
      const ml4 = maturityLabel(sm.score);
      const c4 = scoreCol[ml4] || hex("2563eb");
      dt(page, `${sm.code}  ${sm.name}`, { x: M, y, size: 9.5, font: fontB, color: textDark });
      dt(page, `${sm.score}/100`, { x: W - M - 50, y, size: 9.5, font: fontB, color: c4 });
      y -= 13;
      y = drawWrapped(page, sm.evidence || "", M + 10, y, 8.5, gray, W - M * 2 - 10) - 6;
    }
  }

  // Next Actions page
  {
    const { page } = newPage();
    page.drawRectangle({ x: 0, y: H - 50, width: W, height: 50, color: dark });
    dt(page, "Next Best Actions", { x: M, y: H - 32, size: 15, font: fontB, color: rgb(1, 1, 1) });
    let y = H - 68;
    for (const a of data.nextActions) {
      if (y < 80) break;
      const c5 = a.priorityColor ? hex(a.priorityColor.replace("#", "")) : hex("dc2626");
      page.drawRectangle({ x: M - 3, y: y - 2, width: 5, height: 16, color: c5 });
      dt(page, `#${a.rank}  [${a.priority}]  ${a.title || ""}`, { x: M + 8, y, size: 11, font: fontB, color: dark }); y -= 16;
      y = drawWrapped(page, `Addresses: ${a.dimension || ""}`, M + 8, y, 8.5, gray, W - M * 2) - 3;
      y = drawWrapped(page, a.initiative || "", M + 8, y, 9, textDark, W - M * 2) - 3;
      y = drawWrapped(page, `Benefit: ${a.benefit || ""}`, M + 8, y, 8.5, hex("15803d"), W - M * 2) - 3;
      y = drawWrapped(page, `Why Now: ${a.whyNow || ""}`, M + 8, y, 8.5, hex("1d4ed8"), W - M * 2) - 10;
    }
    dt(page, "Based on publicly available information - AI-generated - Verify independently before decisions", { x: M, y: M, size: 8, font: fontN, color: gray });
  }

  return Buffer.from(await pdfDoc.save());
}

export async function POST(req: NextRequest) {
  const { format, data } = await req.json();
  if (!data || !format) return NextResponse.json({ error: "Missing format or data." }, { status: 400 });
  const filename = safeFilename(data.company || "DDMR_Report");
  try {
    function respond(buf: Buffer, ct: string, ext: string) {
      return new NextResponse(new Uint8Array(buf), { headers: { "Content-Type": ct, "Content-Disposition": `attachment; filename="${filename}_DDMR.${ext}"` } });
    }
    if (format === "xlsx") return respond(await buildXlsx(data), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "xlsx");
    if (format === "docx") return respond(await buildDocx(data), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "docx");
    if (format === "pptx") return respond(await buildPptx(data), "application/vnd.openxmlformats-officedocument.presentationml.presentation", "pptx");
    if (format === "pdf")  return respond(await buildPdf(data), "application/pdf", "pdf");
    return NextResponse.json({ error: "Unsupported format." }, { status: 400 });
  } catch (err) {
    console.error("Export error:", err);
    return NextResponse.json({ error: "Export failed. Please try again." }, { status: 500 });
  }
}
