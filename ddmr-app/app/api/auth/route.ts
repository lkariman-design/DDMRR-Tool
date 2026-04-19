import { NextRequest, NextResponse } from "next/server";

const DEMO_EMAIL    = "demo@nichebrains.ai";
const DEMO_PASSWORD = "DDMR@2026";

export async function POST(req: NextRequest) {
  const { email, password } = await req.json();

  if (email?.trim().toLowerCase() !== DEMO_EMAIL || password !== DEMO_PASSWORD) {
    return NextResponse.json({ error: "Invalid email or password." }, { status: 401 });
  }

  const res = NextResponse.json({ ok: true });
  res.cookies.set("ddmr_auth", "ddmr-demo-authenticated", {
    httpOnly: true,
    secure: process.env.NODE_ENV === "production",
    sameSite: "lax",
    maxAge: 60 * 60 * 24, // 24 h
    path: "/",
  });
  return res;
}

export async function DELETE() {
  const res = NextResponse.json({ ok: true });
  res.cookies.set("ddmr_auth", "", { maxAge: 0, path: "/" });
  return res;
}
