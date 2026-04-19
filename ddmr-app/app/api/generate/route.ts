import { NextRequest, NextResponse } from "next/server";
import Anthropic from "@anthropic-ai/sdk";

const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

const SYSTEM = `You are a DDMR (Digital Diagnostic and Maturity Report) analyst. You assess Indian industrial and manufacturing companies for digital transformation maturity using ONLY publicly available signals — company websites, LinkedIn, MCA filings, ICRA ratings, news, and job postings. You are rigorous, evidence-based, and concise.`;

function buildPrompt(company: string, website: string) {
  return `Analyse the company: ${company}${website ? `\nWebsite: ${website}` : ""}

Generate a comprehensive DDMR report assessing its digital transformation maturity across 5 dimensions using public signals. Base scores on what you know from public sources. If limited information is available, score conservatively and note the evidence gap.

Maturity scale: Legacy 0–33 | Siloed 34–66 | Strategic 67–83 | Future Ready 84–100

Return ONLY a valid JSON object — no markdown fences, no explanation, no trailing text. Use this exact structure:

{
  "company": "exact legal company name",
  "sector": "industry sector (e.g. Pneumatics & Industrial Automation)",
  "composite": <weighted average of 5 dimension scores, each weight 20%>,
  "summary": "2–3 sentence executive summary focused on digital maturity posture",
  "highlights": [
    {"type": "positive", "text": "specific positive finding"},
    {"type": "positive", "text": "specific positive finding"},
    {"type": "positive", "text": "specific positive finding"},
    {"type": "warning", "text": "specific gap or risk"},
    {"type": "warning", "text": "specific gap or risk"}
  ],
  "meta": {
    "founded": "year or decade",
    "hq": "city, state",
    "employees": "approximate headcount or range",
    "revenue": "approximate revenue if publicly known, else 'Not publicly disclosed'",
    "products": "main products or services",
    "customers": "customer types/segments",
    "countries": "markets served"
  },
  "dimensions": [
    {
      "code": "ST", "name": "Strategy", "score": <0-100>, "weight": "20%",
      "trend": "up"|"neutral"|"down",
      "insight": "1–2 sentence insight on digital strategy posture",
      "submetrics": [
        {"code":"ST01","name":"Digital Strategy Articulation","score":<0-100>,"evidence":"specific public signal or noted absence"},
        {"code":"ST02","name":"Leadership Technology Vision","score":<0-100>,"evidence":"..."},
        {"code":"ST03","name":"Digital Investment Signals","score":<0-100>,"evidence":"..."},
        {"code":"ST04","name":"Technology Partnerships","score":<0-100>,"evidence":"..."},
        {"code":"ST05","name":"Innovation Pipeline","score":<0-100>,"evidence":"..."}
      ]
    },
    {
      "code": "OSC", "name": "Operations & Supply Chain", "score": <0-100>, "weight": "20%",
      "trend": "up"|"neutral"|"down",
      "insight": "...",
      "submetrics": [
        {"code":"OSC01","name":"Digital Manufacturing Readiness","score":<0-100>,"evidence":"..."},
        {"code":"OSC02","name":"Supply Chain Visibility Tools","score":<0-100>,"evidence":"..."},
        {"code":"OSC03","name":"Process Automation Adoption","score":<0-100>,"evidence":"..."},
        {"code":"OSC04","name":"Quality Management (Digital)","score":<0-100>,"evidence":"..."},
        {"code":"OSC05","name":"Industry 4.0 Integration","score":<0-100>,"evidence":"..."}
      ]
    },
    {
      "code": "SM", "name": "Sales & Marketing", "score": <0-100>, "weight": "20%",
      "trend": "up"|"neutral"|"down",
      "insight": "...",
      "submetrics": [
        {"code":"SM01","name":"Digital Presence & SEO","score":<0-100>,"evidence":"..."},
        {"code":"SM02","name":"Social Media Engagement","score":<0-100>,"evidence":"..."},
        {"code":"SM03","name":"E-commerce / Digital Ordering","score":<0-100>,"evidence":"..."},
        {"code":"SM04","name":"Digital Marketing Activity","score":<0-100>,"evidence":"..."},
        {"code":"SM05","name":"CRM & Customer Data Utilization","score":<0-100>,"evidence":"..."}
      ]
    },
    {
      "code": "TA", "name": "Technology Adoption", "score": <0-100>, "weight": "20%",
      "trend": "up"|"neutral"|"down",
      "insight": "...",
      "submetrics": [
        {"code":"TA01","name":"Website & Digital Infrastructure","score":<0-100>,"evidence":"..."},
        {"code":"TA02","name":"IoT / Smart Product Portfolio","score":<0-100>,"evidence":"..."},
        {"code":"TA03","name":"Cloud & SaaS Adoption","score":<0-100>,"evidence":"..."},
        {"code":"TA04","name":"R&D in Digital / Smart Tech","score":<0-100>,"evidence":"..."},
        {"code":"TA05","name":"Cybersecurity & Data Practices","score":<0-100>,"evidence":"..."}
      ]
    },
    {
      "code": "SC", "name": "Skills & Capabilities", "score": <0-100>, "weight": "20%",
      "trend": "up"|"neutral"|"down",
      "insight": "...",
      "submetrics": [
        {"code":"SC01","name":"Digital Talent Hiring Signals","score":<0-100>,"evidence":"..."},
        {"code":"SC02","name":"Technical Workforce Depth","score":<0-100>,"evidence":"..."},
        {"code":"SC03","name":"Digital Training Programs","score":<0-100>,"evidence":"..."},
        {"code":"SC04","name":"Leadership Digital Literacy","score":<0-100>,"evidence":"..."},
        {"code":"SC05","name":"Innovation Culture Signals","score":<0-100>,"evidence":"..."}
      ]
    }
  ],
  "nextActions": [
    {
      "rank": "01", "priority": "High", "priorityColor": "#dc2626",
      "title": "initiative title",
      "dimension": "Dimension Name (sub-metric code: score/100)",
      "initiative": "what specifically to do — concrete and actionable",
      "benefit": "quantified or qualified expected benefit",
      "loss": "cost of inaction — competitive or operational risk",
      "whyNow": "market timing or urgency rationale"
    },
    {"rank":"02","priority":"High","priorityColor":"#dc2626","title":"...","dimension":"...","initiative":"...","benefit":"...","loss":"...","whyNow":"..."},
    {"rank":"03","priority":"Medium","priorityColor":"#d97706","title":"...","dimension":"...","initiative":"...","benefit":"...","loss":"...","whyNow":"..."},
    {"rank":"04","priority":"Medium","priorityColor":"#d97706","title":"...","dimension":"...","initiative":"...","benefit":"...","loss":"...","whyNow":"..."}
  ]
}`;
}

export async function POST(req: NextRequest) {
  if (!process.env.ANTHROPIC_API_KEY) {
    return NextResponse.json({ error: "ANTHROPIC_API_KEY not configured." }, { status: 503 });
  }

  const { company, website } = await req.json();
  if (!company?.trim()) {
    return NextResponse.json({ error: "Company name is required." }, { status: 400 });
  }

  try {
    const message = await client.messages.create({
      model: "claude-sonnet-4-6",
      max_tokens: 4096,
      system: SYSTEM,
      messages: [{ role: "user", content: buildPrompt(company.trim(), (website || "").trim()) }],
    });

    const raw = message.content[0].type === "text" ? message.content[0].text : "";

    // Strip any accidental markdown fences
    const cleaned = raw.replace(/^```(?:json)?\n?/,"").replace(/\n?```$/,"").trim();
    const data = JSON.parse(cleaned);

    // Validate minimal shape
    if (!data.company || !data.dimensions || data.dimensions.length !== 5) {
      throw new Error("Unexpected response shape from model");
    }

    return NextResponse.json(data);
  } catch (err) {
    console.error("Generate error:", err);
    return NextResponse.json({ error: "Report generation failed. Please try again." }, { status: 500 });
  }
}
