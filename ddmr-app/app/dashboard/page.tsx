"use client";
import { useState } from "react";
import { useRouter } from "next/navigation";
import { Building2, Globe, Search, ChevronRight, LogOut, Sparkles } from "lucide-react";

const STEPS = [
  "Researching company profile...",
  "Analysing digital signals...",
  "Scoring 5 dimensions...",
  "Generating insights...",
  "Compiling report...",
];

export default function DashboardPage() {
  const router = useRouter();
  const [company, setCompany] = useState("");
  const [website, setWebsite] = useState("");
  const [error, setError] = useState("");
  const [generating, setGenerating] = useState(false);
  const [step, setStep] = useState(0);

  async function handleSignOut() {
    await fetch("/api/auth", { method: "DELETE" });
    router.push("/");
  }

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault();
    if (!company.trim()) { setError("Please enter a company name."); return; }
    setError("");
    setGenerating(true);
    setStep(0);

    const interval = setInterval(() => {
      setStep(s => (s < STEPS.length - 1 ? s + 1 : s));
    }, 2200);

    try {
      const res = await fetch("/api/generate", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ company: company.trim(), website: website.trim() }),
      });
      clearInterval(interval);
      if (res.ok) {
        const data = await res.json();
        sessionStorage.setItem("ddmr_generated", JSON.stringify(data));
        router.push("/report?generated=1");
      } else {
        const d = await res.json();
        setError(d.error || "Report generation failed. Please try again.");
        setGenerating(false);
      }
    } catch {
      clearInterval(interval);
      setError("Network error. Please try again.");
      setGenerating(false);
    }
  }

  if (generating) {
    return (
      <div className="min-h-screen bg-slate-50 flex items-center justify-center p-6">
        <div className="text-center max-w-sm w-full">
          <div className="w-16 h-16 rounded-2xl bg-blue-600 flex items-center justify-center mx-auto mb-6">
            <Building2 size={28} className="text-white"/>
          </div>
          <h2 className="text-2xl font-bold text-slate-900 mb-1">Generating Report</h2>
          <p className="text-slate-500 text-sm mb-8">
            AI is analysing <strong className="text-slate-700">{company}</strong>
          </p>

          <div className="space-y-2 text-left mb-8">
            {STEPS.map((s, i) => (
              <div key={i} className={`flex items-center gap-3 px-4 py-3 rounded-xl transition-all duration-300 ${
                i < step ? "opacity-40" :
                i === step ? "bg-white border border-blue-200 shadow-sm" :
                "opacity-20"
              }`}>
                {i < step ? (
                  <div className="w-5 h-5 rounded-full bg-green-500 flex items-center justify-center flex-shrink-0">
                    <span className="text-white text-xs font-bold">✓</span>
                  </div>
                ) : i === step ? (
                  <svg className="animate-spin h-5 w-5 text-blue-500 flex-shrink-0" viewBox="0 0 24 24" fill="none">
                    <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"/>
                    <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8H4z"/>
                  </svg>
                ) : (
                  <div className="w-5 h-5 rounded-full border-2 border-slate-200 flex-shrink-0"/>
                )}
                <span className={`text-sm ${i === step ? "font-semibold text-slate-800" : "text-slate-500"}`}>{s}</span>
              </div>
            ))}
          </div>

          <p className="text-slate-400 text-xs">This typically takes 20–40 seconds</p>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50">
      <nav className="bg-white border-b border-slate-100 px-6 py-4 flex items-center justify-between">
        <div className="flex items-center gap-2">
          <div className="w-8 h-8 rounded-lg bg-blue-600 flex items-center justify-center">
            <Building2 size={16} className="text-white"/>
          </div>
          <span className="font-bold text-slate-800">DDMR Tool</span>
        </div>
        <button onClick={handleSignOut}
          className="flex items-center gap-1.5 text-slate-500 hover:text-slate-700 text-sm transition">
          <LogOut size={15}/> Sign out
        </button>
      </nav>

      <div className="max-w-2xl mx-auto px-6 py-16">
        <div className="mb-10">
          <div className="flex items-center gap-2 mb-3">
            <Sparkles size={15} className="text-blue-500"/>
            <span className="text-blue-600 text-sm font-semibold">AI-Powered Assessment</span>
          </div>
          <h1 className="text-3xl font-bold text-slate-900 mb-2">New DDMR Report</h1>
          <p className="text-slate-500 text-sm leading-relaxed">
            Enter a company name and optional website. Claude AI will research public signals and generate a full 5-dimension digital maturity report in under 60 seconds.
          </p>
        </div>

        <div className="bg-white rounded-2xl shadow-sm border border-slate-100 p-8">
          <form onSubmit={handleSubmit} className="space-y-6">
            <div>
              <label className="block text-sm font-semibold text-slate-700 mb-2">
                Company Name <span className="text-red-500">*</span>
              </label>
              <input
                value={company}
                onChange={e => { setCompany(e.target.value); setError(""); }}
                placeholder="e.g. Kirloskar Electric Company Limited"
                className="w-full px-4 py-3 rounded-xl border border-slate-200 bg-slate-50 focus:bg-white focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm transition"
              />
              {error && <p className="text-red-500 text-xs mt-1.5">{error}</p>}
            </div>

            <div>
              <label className="block text-sm font-semibold text-slate-700 mb-2">
                Company Website <span className="text-slate-400 font-normal">(optional)</span>
              </label>
              <div className="relative">
                <Globe size={14} className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400"/>
                <input
                  value={website}
                  onChange={e => setWebsite(e.target.value)}
                  placeholder="https://www.company.com"
                  className="w-full pl-10 pr-4 py-3 rounded-xl border border-slate-200 bg-slate-50 focus:bg-white focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm transition"
                />
              </div>
              <p className="text-slate-400 text-xs mt-1.5">Providing the website improves analysis accuracy</p>
            </div>

            <div className="pt-2">
              <button type="submit"
                className="w-full py-4 rounded-xl font-semibold text-white transition flex items-center justify-center gap-2"
                style={{ background: "linear-gradient(135deg,#1a3a6e,#2563eb)" }}>
                <Search size={16}/>
                <span>Generate DDMR Report</span>
                <ChevronRight size={16}/>
              </button>
            </div>
          </form>
        </div>

        <div className="mt-8 p-5 rounded-2xl bg-blue-50 border border-blue-100">
          <div className="flex items-start gap-3">
            <Sparkles size={14} className="text-blue-500 flex-shrink-0 mt-0.5"/>
            <div>
              <p className="text-blue-800 text-sm font-semibold mb-1">What does the AI analyse?</p>
              <p className="text-blue-700 text-xs leading-relaxed">
                Strategy · Operations & Supply Chain · Sales & Marketing · Technology Adoption · Skills & Capabilities — all assessed from public signals: company website, LinkedIn, MCA filings, news, and job postings.
              </p>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
