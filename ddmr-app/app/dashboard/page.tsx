"use client";
import { useState, useEffect } from "react";
import { useRouter } from "next/navigation";
import { Building2, Globe, Search, ChevronRight, LogOut, Sparkles, FileText, Trash2, ExternalLink } from "lucide-react";

interface StoredReport {
  id: string;
  company: string;
  sector: string;
  composite: number;
  maturity: string;
  verdictColor: string;
  generatedAt: string;
  data: any;
}

function maturityLabel(s: number): string {
  if (s >= 84) return "Future Ready";
  if (s >= 67) return "Strategic";
  if (s >= 34) return "Siloed";
  return "Legacy";
}
function verdictColor(s: number): string {
  if (s >= 84) return "#1d4ed8";
  if (s >= 67) return "#16a34a";
  if (s >= 34) return "#d97706";
  return "#dc2626";
}
function genId(): string {
  return Date.now().toString(36) + Math.random().toString(36).slice(2);
}
function getReports(): StoredReport[] {
  try { return JSON.parse(localStorage.getItem("ddmr_reports") || "[]"); } catch { return []; }
}
function saveReport(data: any): string {
  const id = genId();
  const ml = data.maturity || maturityLabel(data.composite);
  const vc = data.verdictColor || verdictColor(data.composite);
  const reports = getReports();
  reports.unshift({ id, company: data.company, sector: data.sector, composite: data.composite, maturity: ml, verdictColor: vc, generatedAt: new Date().toISOString(), data });
  localStorage.setItem("ddmr_reports", JSON.stringify(reports.slice(0, 30)));
  return id;
}
function deleteReport(id: string) {
  localStorage.setItem("ddmr_reports", JSON.stringify(getReports().filter(r => r.id !== id)));
}

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
  const [generatingFor, setGeneratingFor] = useState("");
  const [reports, setReports] = useState<StoredReport[]>([]);

  useEffect(() => { setReports(getReports()); }, []);

  async function handleSignOut() {
    await fetch("/api/auth", { method: "DELETE" });
    router.push("/");
  }

  function handleDelete(id: string, e: React.MouseEvent) {
    e.stopPropagation();
    deleteReport(id);
    setReports(getReports());
  }

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault();
    if (!company.trim()) { setError("Please enter a company name."); return; }
    setError("");
    setGeneratingFor(company.trim());
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
        const id = saveReport(data);
        router.push(`/report?id=${id}`);
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
            AI is analysing <strong className="text-slate-700">{generatingFor}</strong>
          </p>
          <div className="space-y-2 text-left mb-8">
            {STEPS.map((s, i) => (
              <div key={i} className={`flex items-center gap-3 px-4 py-3 rounded-xl transition-all duration-300 ${
                i < step ? "opacity-40" : i === step ? "bg-white border border-blue-200 shadow-sm" : "opacity-20"
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

      <div className="max-w-2xl mx-auto px-6 py-12">
        <div className="mb-8">
          <div className="flex items-center gap-2 mb-2">
            <Sparkles size={14} className="text-blue-500"/>
            <span className="text-blue-600 text-sm font-semibold">AI-Powered Assessment</span>
          </div>
          <h1 className="text-3xl font-bold text-slate-900 mb-2">New DDMR Report</h1>
          <p className="text-slate-500 text-sm leading-relaxed">
            Enter a company name and optional website. Claude AI will research public signals and generate a full 5-dimension digital maturity report in under 60 seconds.
          </p>
        </div>

        <div className="bg-white rounded-2xl shadow-sm border border-slate-100 p-8 mb-10">
          <form onSubmit={handleSubmit} className="space-y-5">
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
            <button type="submit"
              className="w-full py-4 rounded-xl font-semibold text-white transition flex items-center justify-center gap-2"
              style={{ background: "linear-gradient(135deg,#1a3a6e,#2563eb)" }}>
              <Search size={16}/>
              <span>Generate DDMR Report</span>
              <ChevronRight size={16}/>
            </button>
          </form>
        </div>

        {/* Report history */}
        {reports.length > 0 && (
          <div>
            <h2 className="text-sm font-semibold text-slate-500 uppercase tracking-wider mb-3">Generated Reports</h2>
            <div className="space-y-2">
              {reports.map(r => (
                <div key={r.id}
                  className="bg-white rounded-xl border border-slate-100 px-5 py-4 flex items-center justify-between hover:border-blue-200 hover:shadow-sm transition group">
                  <div className="flex items-center gap-4 flex-1 min-w-0">
                    <div className="w-10 h-10 rounded-lg flex items-center justify-center flex-shrink-0" style={{ background: "#eff6ff" }}>
                      <FileText size={17} className="text-blue-600"/>
                    </div>
                    <div className="min-w-0">
                      <div className="font-semibold text-slate-800 text-sm truncate">{r.company}</div>
                      <div className="text-slate-400 text-xs truncate">{r.sector} · {new Date(r.generatedAt).toLocaleDateString("en-IN", { day: "numeric", month: "short", year: "numeric" })}</div>
                    </div>
                  </div>
                  <div className="flex items-center gap-2 flex-shrink-0 ml-4">
                    <span className="text-xs font-bold px-2.5 py-1 rounded-full text-white" style={{ background: r.verdictColor }}>{r.composite}/100</span>
                    <span className="text-xs text-slate-500 hidden sm:block">{r.maturity}</span>
                    <button
                      onClick={() => router.push(`/report?id=${r.id}`)}
                      className="p-1.5 rounded-lg text-blue-500 hover:bg-blue-50 transition"
                      title="View report">
                      <ExternalLink size={15}/>
                    </button>
                    <button
                      onClick={e => handleDelete(r.id, e)}
                      className="p-1.5 rounded-lg text-slate-300 hover:text-red-500 hover:bg-red-50 transition"
                      title="Delete report">
                      <Trash2 size={15}/>
                    </button>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
