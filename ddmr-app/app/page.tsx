"use client";
import { useState } from "react";
import { useRouter } from "next/navigation";
import { Building2, ArrowRight, Lock, Sparkles } from "lucide-react";

export default function LoginPage() {
  const router = useRouter();
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [loading, setLoading] = useState(false);

  function handleDemo() {
    setLoading(true);
    setTimeout(() => router.push("/report?demo=1"), 600);
  }
  function handleLogin(e: React.FormEvent) {
    e.preventDefault();
    setLoading(true);
    setTimeout(() => router.push("/dashboard"), 600);
  }

  return (
    <div className="min-h-screen flex">
      <div className="hidden lg:flex w-1/2 flex-col justify-between p-12"
        style={{background:"linear-gradient(150deg,#0f2442 0%,#1a3a6e 60%,#0f2442 100%)"}}>
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-lg bg-blue-500 flex items-center justify-center">
            <Building2 size={22} className="text-white"/>
          </div>
          <span className="text-white text-xl font-bold tracking-tight">DDMR Tool</span>
        </div>
        <div>
          <div className="text-blue-300 text-sm font-semibold uppercase tracking-widest mb-4">Due Diligence & Market Report</div>
          <h1 className="text-white text-4xl font-bold leading-tight mb-6">Executive-grade<br/>company intelligence<br/>in minutes.</h1>
          <p className="text-blue-200 text-lg leading-relaxed mb-10">Powered by public data, ICRA ratings, and AI scoring — get a complete 5-dimension report with evidence citations.</p>
          <div className="grid grid-cols-2 gap-4">
            {[{label:"Dimensions Scored",value:"5"},{label:"Sub-metrics",value:"25+"},{label:"Data Sources",value:"12+"},{label:"Output Formats",value:"3"}].map(({label,value})=>(
              <div key={label} className="bg-white/10 rounded-xl p-4">
                <div className="text-white text-2xl font-bold">{value}</div>
                <div className="text-blue-300 text-sm mt-1">{label}</div>
              </div>
            ))}
          </div>
        </div>
        <p className="text-blue-400 text-sm">© 2026 DDMR Tool · Confidential Platform</p>
      </div>

      <div className="flex-1 flex items-center justify-center p-8 bg-slate-50">
        <div className="w-full max-w-md">
          <h2 className="text-2xl font-bold text-slate-900 mb-1">Welcome back</h2>
          <p className="text-slate-500 text-sm mb-8">Sign in to your account to continue</p>

          <button onClick={handleDemo} disabled={loading}
            className="w-full mb-6 p-4 rounded-xl border-2 border-blue-200 bg-blue-50 hover:bg-blue-100 hover:border-blue-400 transition-all group text-left disabled:opacity-60">
            <div className="flex items-center justify-between">
              <div>
                <div className="flex items-center gap-2 mb-1">
                  <Sparkles size={16} className="text-blue-600"/>
                  <span className="text-blue-700 font-semibold text-sm">Try Live Demo</span>
                </div>
                <p className="text-slate-600 text-sm">Full DDMR report for <strong>Janatics India Pvt Ltd</strong></p>
                <p className="text-slate-400 text-xs mt-1">Score: 75.7/100 · ₹531 Cr Revenue · Pneumatics & Automation</p>
              </div>
              <ArrowRight size={18} className="text-blue-500 group-hover:translate-x-1 transition-transform"/>
            </div>
          </button>

          <div className="flex items-center gap-3 mb-6">
            <div className="flex-1 h-px bg-slate-200"/>
            <span className="text-slate-400 text-xs">or sign in</span>
            <div className="flex-1 h-px bg-slate-200"/>
          </div>

          <form onSubmit={handleLogin} className="space-y-4">
            <div>
              <label className="block text-sm font-medium text-slate-700 mb-1.5">Email</label>
              <input type="email" value={email} onChange={e=>setEmail(e.target.value)} placeholder="you@company.com"
                className="w-full px-4 py-3 rounded-xl border border-slate-200 bg-white text-slate-900 placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"/>
            </div>
            <div>
              <label className="block text-sm font-medium text-slate-700 mb-1.5">Password</label>
              <input type="password" value={password} onChange={e=>setPassword(e.target.value)} placeholder="••••••••"
                className="w-full px-4 py-3 rounded-xl border border-slate-200 bg-white text-slate-900 placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"/>
            </div>
            <button type="submit" disabled={loading}
              className="w-full py-3 rounded-xl bg-slate-900 hover:bg-slate-700 text-white font-semibold text-sm transition flex items-center justify-center gap-2 disabled:opacity-60">
              {loading
                ? <svg className="animate-spin h-4 w-4" viewBox="0 0 24 24" fill="none"><circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"/><path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8H4z"/></svg>
                : <><Lock size={15}/><span>Sign In</span></>}
            </button>
          </form>
          <p className="text-center text-slate-400 text-xs mt-8">Secure · Confidential · Enterprise-ready</p>
        </div>
      </div>
    </div>
  );
}
