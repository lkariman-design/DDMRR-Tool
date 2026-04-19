import { NextRequest, NextResponse } from "next/server";
import { execSync } from "child_process";
import { readFileSync } from "fs";
import path from "path";

const ROOT = path.join(process.cwd(), "..");

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
} as const;

export async function GET(req: NextRequest) {
  const format = req.nextUrl.searchParams.get("format") as keyof typeof CONFIG;

  if (!CONFIG[format]) {
    return NextResponse.json({ error: "Invalid format. Use xlsx, docx, or pptx." }, { status: 400 });
  }

  const { script, file, mime } = CONFIG[format];
  const scriptPath = path.join(ROOT, script);
  const filePath   = path.join(ROOT, "output", file);

  try {
    execSync(`python3 "${scriptPath}"`, { cwd: ROOT, timeout: 30000 });
    const buffer = readFileSync(filePath);
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
