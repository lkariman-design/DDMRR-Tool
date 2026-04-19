import { NextRequest, NextResponse } from "next/server";
import { existsSync, readFileSync } from "fs";
import { execSync } from "child_process";
import path from "path";

const CONFIG = {
  xlsx: {
    script: "generate_scorecard.py",
    file:   "Janatics_DDMR_Scorecard.xlsx",
    mime:   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  },
  docx: {
    script: "generate_report.py",
    file:   "Janatics_DDMR_Report.docx",
    mime:   "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  },
  pptx: {
    script: "generate_deck.py",
    file:   "Janatics_DDMR_Deck.pptx",
    mime:   "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  },
  pdf: {
    script: "generate_pdf.py",
    file:   "Janatics_DDMR_Report.pdf",
    mime:   "application/pdf",
  },
} as const;

export async function GET(req: NextRequest) {
  const format = req.nextUrl.searchParams.get("format") as keyof typeof CONFIG;

  if (!CONFIG[format]) {
    return NextResponse.json({ error: "Invalid format. Use xlsx, docx, pptx, or pdf." }, { status: 400 });
  }

  const { script, file, mime } = CONFIG[format];

  // On Vercel / production: serve pre-built static file from public/downloads
  const staticPath = path.join(process.cwd(), "public", "downloads", file);
  if (existsSync(staticPath)) {
    const buffer = readFileSync(staticPath);
    return new NextResponse(buffer, {
      headers: {
        "Content-Type": mime,
        "Content-Disposition": `attachment; filename="${file}"`,
        "Content-Length": String(buffer.length),
      },
    });
  }

  // Local dev fallback: regenerate via Python
  try {
    const ROOT = path.join(process.cwd(), "..");
    const scriptPath = path.join(ROOT, script);
    const outputPath = path.join(ROOT, "output", file);
    execSync(`python3 "${scriptPath}"`, { cwd: ROOT, timeout: 30000 });
    const buffer = readFileSync(outputPath);
    return new NextResponse(buffer, {
      headers: {
        "Content-Type": mime,
        "Content-Disposition": `attachment; filename="${file}"`,
        "Content-Length": String(buffer.length),
      },
    });
  } catch (err) {
    console.error("Download generation error:", err);
    return NextResponse.json({ error: "File generation failed." }, { status: 500 });
  }
}
