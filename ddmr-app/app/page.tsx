"use client";
import { useState } from "react";
import { useRouter } from "next/navigation";
import { Building2, ArrowRight, Lock, Sparkles, Eye, EyeOff, KeyRound } from "lucide-react";

export default function LoginPage() {
  const router = useRouter();
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [showPwd, setShowPwd] = useState(false);
  const [loading, setLoading] = useState<string|null>(null);
  const [error, setError] = useState("");

  async function handleLogin(e: React.FormEvent) {
    e.preventDefault();
    setError("");
    setLoading("login");
    try {
      const res = await fetch("/api/auth", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ email, password }),
      });
      if (res.ok) {
        router.push("/dashboard");
      } else {
        const d = await res.json();
        setError(d.error || "Login failed.");
        setLoading(null);
      }
    } catch {
      setError("Network error. Please try again.");
      setLoading(null);
    }
  }

  function handleDemo(slug: string) {
    setLoading(slug);
    setTimeout(() => router.push(`/report?demo=${slug}`), 500);
  }

  function fillDemo() {
    setEmail("demo@nichebrains.ai");
    setPassword("DDMR@2026");
    setError("");
  }

  const demos = [
    {
      slug: "janatics",
      company: "Janatics India Pvt Ltd",
      sector: "Pneumatics & Industrial Automation",
      score: "64/100", maturity: "Siloed",
      revenue: "₹531 Cr", tag: "Domestic Manufacturer", tagColor: "#1d4ed8",
    },
    {
      slug: "messer",
      company: "Messer Cutting Systems India Pvt Ltd",
      sector: "Thermal & Laser Cutting Systems",
      score: "72/100", maturity: "Strategic",
      revenue: "₹185 Cr", tag: "MNC Subsidiary", tagColor: "#16a34a",
    },
  ];

  return (
    <div className="min-h-screen flex">
      {/* Left brand panel */}
      <div className="hidden lg:flex w-1/2 flex-col justify-between p-12"
        style={{background:"linear-gradient(150deg,#0f2442 0%,#1a3a6e 60%,#0f2442 100%)"}}>
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-lg bg-blue-500 flex items-center justify-center">
            <Building2 size={22} className="text-white"/>
          </div>
          <span className="text-white text-xl font-bold tracking-tight">DDMR Tool</span>
        </div>
        <div>
          <div className="text-blue-300 text-sm font-semibold uppercase tracking-widest mb-4">Digital Diagnostic and Maturity Report</div>
          <h1 className="text-white text-4xl font-bold leading-tight mb-6">Executive-grade<br/>digital maturity<br/>intelligence.</h1>
          <p className="text-blue-200 text-lg leading-relaxed mb-10">Enter any company name and website — get a complete AI-generated 5-dimension digital maturity report in under 60 seconds.</p>
          <div className="grid grid-cols-2 gap-4">
            {[{label:"Dimensions Scored",value:"5"},{label:"Sub-metrics",value:"25"},{label:"AI-Powered",value:"Yes"},{label:"Output Formats",value:"4"}].map(({label,value})=>(
              <div key={label} className="bg-white/10 rounded-xl p-4">
                <div className="text-white text-2xl font-bold">{value}</div>
                <div className="text-blue-300 text-sm mt-1">{label}</div>
              </div>
            ))}
          </div>
        </div>
        <p className="text-blue-400 text-sm">© 2026 DDMR Tool · Confidential Platform</p>
      </div>

      {/* Right auth panel */}
      <div className="flex-1 flex items-center justify-center p-8 bg-slate-50 overflow-y-auto">
        <div className="w-full max-w-md py-8">
          <h2 className="text-2xl font-bold text-slate-900 mb-1">Welcome back</h2>
          <p className="text-slate-500 text-sm mb-6">Sign in to generate AI-powered DDMR reports</p>

          {/* Demo credential hint */}
          <button onClick={fillDemo}
            className="w-full mb-5 p-3 rounded-xl border border-dashed border-blue-300 bg-blue-50/60 hover:bg-blue-50 transition text-left group">
            <div className="flex items-center gap-2 mb-1">
              <KeyRound size={13} className="text-blue-500"/>
              <span className="text-blue-700 text-xs font-semibold uppercase tracking-wider">Demo Credentials</span>
              <span className="ml-auto text-blue-500 text-xs group-hover:underline">Click to fill →</span>
            </div>
            <p className="text-slate-600 text-xs font-mono">demo@nichebrains.ai &nbsp;/&nbsp; DDMR@2026</p>
          </button>

          {/* Login form */}
          <form onSubmit={handleLogin} className="space-y-4 mb-6">
            <div>
              <label className="block text-sm font-medium text-slate-700 mb-1.5">Email</label>
              <input type="email" value={email} onChange={e=>{setEmail(e.target.value);setError("");}}
                placeholder="demo@nichebrains.ai"
                className="w-full px-4 py-3 rounded-xl border border-slate-200 bg-white text-slate-900 placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"/>
            </div>
            <div>
              <label className="block text-sm font-medium text-slate-700 mb-1.5">Password</label>
              <div className="relative">
                <input type={showPwd?"text":"password"} value={password}
                  onChange={e=>{setPassword(e.target.value);setError("");}}
                  placeholder="••••••••"
                  className="w-full px-4 py-3 pr-11 rounded-xl border border-slate-200 bg-white text-slate-900 placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"/>
                <button type="button" onClick={()=>setShowPwd(s=>!s)}
                  className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600">
                  {showPwd?<EyeOff size={16}/>:<Eye size={16}/>}
                </button>
              </div>
            </div>
            {error && (
              <p className="text-red-600 text-xs bg-red-50 border border-red-200 px-3 py-2 rounded-lg">{error}</p>
            )}
            <button type="submit" disabled={loading !== null}
              className="w-full py-3 rounded-xl bg-slate-900 hover:bg-slate-700 text-white font-semibold text-sm transition flex items-center justify-center gap-2 disabled:opacity-60">
              {loading === "login"
                ? <svg className="animate-spin h-4 w-4" viewBox="0 0 24 24" fill="none"><circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"/><path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8H4z"/></svg>
                : <><Lock size={15}/><span>Sign In & Generate Reports</span></>}
            </button>
          </form>

          {/* Divider */}
          <div className="flex items-center gap-3 mb-5">
            <div className="flex-1 h-px bg-slate-200"/>
            <span className="text-slate-400 text-xs">or browse live demos</span>
            <div className="flex-1 h-px bg-slate-200"/>
          </div>

          {/* Demo cards */}
          <div className="space-y-2.5">
            <div className="flex items-center gap-2 mb-2">
              <Sparkles size={13} className="text-blue-500"/>
              <span className="text-blue-700 text-xs font-semibold uppercase tracking-wider">Try Live Demos — No Login Required</span>
            </div>
            {demos.map(d=>(
              <button key={d.slug} onClick={()=>handleDemo(d.slug)} disabled={loading !== null}
                className="w-full p-3.5 rounded-xl border-2 border-blue-100 bg-blue-50/50 hover:bg-blue-50 hover:border-blue-300 transition-all group text-left disabled:opacity-60">
                <div className="flex items-center justify-between">
                  <div className="flex-1 min-w-0">
                    <span className="text-xs font-bold px-2 py-0.5 rounded-full text-white inline-block mb-1" style={{background:d.tagColor}}>{d.tag}</span>
                    <p className="text-slate-800 text-sm font-semibold truncate">{d.company}</p>
                    <p className="text-slate-500 text-xs">{d.sector}</p>
                    <div className="flex items-center gap-2 mt-1">
                      <span className="text-blue-700 font-bold text-xs">{d.score}</span>
                      <span className="text-slate-300 text-xs">·</span>
                      <span className="text-slate-500 text-xs">{d.maturity}</span>
                      <span className="text-slate-300 text-xs">·</span>
                      <span className="text-slate-500 text-xs">{d.revenue}</span>
                    </div>
                  </div>
                  {loading === d.slug
                    ? <svg className="animate-spin h-4 w-4 text-blue-500 flex-shrink-0" viewBox="0 0 24 24" fill="none"><circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"/><path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8H4z"/></svg>
                    : <ArrowRight size={15} className="text-blue-400 group-hover:translate-x-1 transition-transform flex-shrink-0"/>
                  }
                </div>
              </button>
            ))}
          </div>

          <p className="text-center text-slate-400 text-xs mt-8">Secure · Confidential · Enterprise-ready</p>
        </div>
      </div>
    </div>
  );
}
