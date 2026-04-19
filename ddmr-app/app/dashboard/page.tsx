"use client";
import { useState } from "react";
import { useRouter } from "next/navigation";
import { Building2, Search, ChevronRight, LogOut, FileText } from "lucide-react";

const SECTORS = ["Pneumatics & Automation","Manufacturing","IT & Software","Pharmaceuticals","FMCG","Retail","Financial Services","Infrastructure","Textiles","Automotive","Food Processing","Healthcare"];

export default function DashboardPage() {
  const router = useRouter();
  const [form, setForm] = useState({company:"",cin:"",sector:"",notes:""});
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  function set(k:string,v:string){setForm(f=>({...f,[k]:v}));setError("");}

  function handleSubmit(e:React.FormEvent){
    e.preventDefault();
    if(!form.company.trim()){setError("Please enter a company name.");return;}
    setLoading(true);
    const params=new URLSearchParams({company:form.company,cin:form.cin,sector:form.sector});
    setTimeout(()=>router.push(`/report?${params}`),800);
  }

  return (
    <div className="min-h-screen bg-slate-50">
      {/* Nav */}
      <nav className="bg-white border-b border-slate-100 px-6 py-4 flex items-center justify-between">
        <div className="flex items-center gap-2">
          <div className="w-8 h-8 rounded-lg bg-blue-600 flex items-center justify-center">
            <Building2 size={16} className="text-white"/>
          </div>
          <span className="font-bold text-slate-800">DDMR Tool</span>
        </div>
        <button onClick={()=>router.push("/")} className="flex items-center gap-1.5 text-slate-500 hover:text-slate-700 text-sm transition">
          <LogOut size={15}/> Sign out
        </button>
      </nav>

      <div className="max-w-2xl mx-auto px-6 py-16">
        <div className="mb-10">
          <h1 className="text-3xl font-bold text-slate-900 mb-2">New Report</h1>
          <p className="text-slate-500">Enter the company details below to generate a full DDMR assessment.</p>
        </div>

        <div className="bg-white rounded-2xl shadow-sm border border-slate-100 p-8">
          <form onSubmit={handleSubmit} className="space-y-6">
            <div>
              <label className="block text-sm font-semibold text-slate-700 mb-2">Company Name <span className="text-red-500">*</span></label>
              <input value={form.company} onChange={e=>set("company",e.target.value)}
                placeholder="e.g. Janatics India Private Limited"
                className="w-full px-4 py-3 rounded-xl border border-slate-200 bg-slate-50 focus:bg-white focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm transition"/>
              {error && <p className="text-red-500 text-xs mt-1.5">{error}</p>}
            </div>

            <div>
              <label className="block text-sm font-semibold text-slate-700 mb-2">CIN <span className="text-slate-400 font-normal">(optional)</span></label>
              <input value={form.cin} onChange={e=>set("cin",e.target.value.toUpperCase())}
                placeholder="e.g. U31103TZ1991PTC003409"
                className="w-full px-4 py-3 rounded-xl border border-slate-200 bg-slate-50 focus:bg-white focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm font-mono transition"/>
              <p className="text-slate-400 text-xs mt-1.5">Providing a CIN enables direct MCA data lookup</p>
            </div>

            <div>
              <label className="block text-sm font-semibold text-slate-700 mb-2">Industry Sector <span className="text-slate-400 font-normal">(optional)</span></label>
              <select value={form.sector} onChange={e=>set("sector",e.target.value)}
                className="w-full px-4 py-3 rounded-xl border border-slate-200 bg-slate-50 focus:bg-white focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm transition">
                <option value="">Select sector...</option>
                {SECTORS.map(s=><option key={s}>{s}</option>)}
              </select>
            </div>

            <div>
              <label className="block text-sm font-semibold text-slate-700 mb-2">Additional Context <span className="text-slate-400 font-normal">(optional)</span></label>
              <textarea value={form.notes} onChange={e=>set("notes",e.target.value)} rows={3}
                placeholder="Any specific focus areas, known concerns, or context for this assessment..."
                className="w-full px-4 py-3 rounded-xl border border-slate-200 bg-slate-50 focus:bg-white focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm resize-none transition"/>
            </div>

            <div className="pt-2">
              <button type="submit" disabled={loading}
                className="w-full py-4 rounded-xl font-semibold text-white transition flex items-center justify-center gap-2 disabled:opacity-60"
                style={{background:"linear-gradient(135deg,#1a3a6e,#2563eb)"}}>
                {loading
                  ? <><svg className="animate-spin h-4 w-4" viewBox="0 0 24 24" fill="none"><circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"/><path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8H4z"/></svg><span>Generating report...</span></>
                  : <><Search size={16}/><span>Generate DDMR Report</span><ChevronRight size={16}/></>}
              </button>
            </div>
          </form>
        </div>

        {/* Recent reports */}
        <div className="mt-10">
          <h2 className="text-sm font-semibold text-slate-500 uppercase tracking-wider mb-4">Recent Reports</h2>
          <div className="space-y-2">
            {[{name:"Janatics India Private Limited",score:75.7,date:"19 Apr 2026",sector:"Pneumatics & Automation",color:"#16a34a"}].map(r=>(
              <button key={r.name} onClick={()=>router.push("/report?demo=1")}
                className="w-full bg-white rounded-xl border border-slate-100 px-5 py-4 flex items-center justify-between hover:border-blue-200 hover:shadow-sm transition group">
                <div className="flex items-center gap-4">
                  <div className="w-10 h-10 rounded-lg flex items-center justify-center" style={{background:"#eff6ff"}}>
                    <FileText size={18} className="text-blue-600"/>
                  </div>
                  <div className="text-left">
                    <div className="font-semibold text-slate-800 text-sm">{r.name}</div>
                    <div className="text-slate-400 text-xs">{r.sector} · {r.date}</div>
                  </div>
                </div>
                <div className="flex items-center gap-3">
                  <span className="text-sm font-bold px-3 py-1 rounded-full text-white" style={{background:r.color}}>{r.score}/100</span>
                  <ChevronRight size={16} className="text-slate-300 group-hover:text-blue-500 transition"/>
                </div>
              </button>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}
