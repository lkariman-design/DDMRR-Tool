"use client";
import { useSearchParams, useRouter } from "next/navigation";
import { useState, useEffect } from "react";
import { Building2, Download, ChevronDown, ChevronUp, ArrowLeft, TrendingUp, TrendingDown, Minus, ExternalLink, FileSpreadsheet, FileText, Presentation } from "lucide-react";

const DEMO = {
  company: "Janatics India Private Limited",
  cin: "U31103TZ1991PTC003409",
  sector: "Pneumatics & Industrial Automation",
  composite: 75.7,
  verdict: "Strong",
  verdictColor: "#16a34a",
  summary: "Janatics India is one of India's most established domestic pneumatics manufacturers, operating since 1977 with ₹531 Cr revenue in FY2025, ICRA-rated credit, and a presence in 42+ countries. The composite score of 75.7/100 reflects strong operational foundations, brand maturity, and global reach — offset by below-industry revenue growth and EBITDA margin compression in FY2025.",
  highlights: [
    {type:"positive", text:"49-year brand with ICRA-validated Stable credit outlook"},
    {type:"positive", text:"17,000+ customers across 7 industries — low concentration risk"},
    {type:"positive", text:"42+ countries, subsidiaries in USA & Germany"},
    {type:"positive", text:"3,500+ product SKUs — single-source OEM supplier capability"},
    {type:"warning",  text:"Revenue growth 4% vs industry CAGR 6.9% — trailing market"},
    {type:"warning",  text:"EBITDA CAGR -22% (FY25) — margin compression, full P&L needed"},
  ],
  meta: {
    incorporated: "28 Aug 1991",
    founded: "1977",
    hq: "Coimbatore, Tamil Nadu",
    employees: "662",
    revenue: "₹531 Crore (FY2025)",
    rating: "ICRA — Stable",
    countries: "42+",
    products: "3,500+",
    customers: "17,000+",
    directors: "Jaganathan Ganesh Kumar, GC Nageswaran, Rajamani Ramesh",
  },
  dimensions: [
    {
      code:"SL", name:"Strategy & Leadership", score:81, weight:"25%", trend:"up",
      insight:"49-year brand, global subsidiaries, DSIR R&D, Industry 4.0 pivot underway.",
      submetrics:[
        {code:"SL01",name:"Company Tenure (est. 1977)",score:90,evidence:"Incorporated 1991; operating since 1977 — 49-year track record"},
        {code:"SL02",name:"Director Profile & DIN Compliance",score:75,evidence:"3 active directors; all DINs current; AGM Sept 2024"},
        {code:"SL03",name:"Governance & Filing Compliance",score:80,evidence:"Filings current with ROC Coimbatore; no penalties noted"},
        {code:"SL04",name:"Debt & Charge Structure",score:78,evidence:"ICRA rated; stable outlook; healthy debt protection metrics"},
        {code:"SL05",name:"Strategic Direction",score:82,evidence:"Expanding: electric actuators, robotics, sensors — Industry 4.0"},
        {code:"SL06",name:"Global Presence",score:88,evidence:"42+ countries; USA & Germany subsidiaries; 200 global partners"},
      ]
    },
    {
      code:"SM", name:"Sales & Marketing", score:78, weight:"20%", trend:"neutral",
      insight:"₹531 Cr revenue with 17,000+ customers and 200 global partners. Growth lags industry.",
      submetrics:[
        {code:"SM01",name:"Revenue Growth vs CAGR",score:55,evidence:"4% YoY vs 6.9% sector CAGR — below market pace"},
        {code:"SM02",name:"Revenue Scale",score:80,evidence:"₹531 Cr FY2025; consistent ₹500 Cr+ for 2 years"},
        {code:"SM03",name:"Customer Diversity",score:90,evidence:"17,000+ customers; 7 industries — low concentration"},
        {code:"SM04",name:"Distribution Network",score:85,evidence:"200 global partners; 3 plants for regional coverage"},
        {code:"SM05",name:"Brand Strength",score:80,evidence:"49-year brand; AIA India member; DSIR R&D recognised"},
      ]
    },
    {
      code:"OSC", name:"Operations & Supply Chain", score:78, weight:"20%", trend:"up",
      insight:"70,000 sqm plant, 3 locations, 3,500+ SKUs, DSIR-recognised R&D.",
      submetrics:[
        {code:"OSC01",name:"Manufacturing Scale",score:85,evidence:"70,000 sqm Coimbatore HQ; Ahmedabad and Noida plants"},
        {code:"OSC02",name:"Product Portfolio Breadth",score:90,evidence:"3,500+ distinct SKUs across automation categories"},
        {code:"OSC03",name:"R&D & Technology",score:82,evidence:"DSIR-recognised R&D; CNC machining; robotics expansion"},
        {code:"OSC04",name:"GST & Regulatory Compliance",score:72,evidence:"Active GST; ICRA notes consistent compliance track record"},
        {code:"OSC05",name:"Supply Chain Digitization",score:65,evidence:"Industry 4.0 product range; internal digitisation undisclosed"},
      ]
    },
    {
      code:"FIN", name:"Finance", score:67, weight:"25%", trend:"down",
      insight:"ICRA Stable but EBITDA CAGR -22% flags margin pressure. Full P&L not public.",
      submetrics:[
        {code:"FIN01",name:"Revenue Scale (₹531 Cr)",score:80,evidence:"Strong for unlisted domestic manufacturer in this segment"},
        {code:"FIN02",name:"Revenue Growth (4% YoY)",score:55,evidence:"FY24 ₹508 Cr → FY25 ₹528 Cr; 4% — below industry"},
        {code:"FIN03",name:"EBITDA Trajectory",score:42,evidence:"EBITDA CAGR -22% (1-yr, Tracxn); margin compression flagged"},
        {code:"FIN04",name:"Credit Profile (ICRA)",score:82,evidence:"ICRA rated; stable; multiple reaffirmations 2021–2025"},
        {code:"FIN05",name:"Capital Structure",score:75,evidence:"Auth ₹30 Cr; paid-up ₹26.5 Cr; no external equity"},
        {code:"FIN06",name:"Employee Productivity",score:70,evidence:"662 employees; ₹531 Cr → ~₹80L per employee"},
      ]
    },
    {
      code:"IC", name:"Industry Context", score:75, weight:"10%", trend:"up",
      insight:"India pneumatics market $1.06B (2024), CAGR 6.9%. Make in India + China+1 tailwinds.",
      submetrics:[
        {code:"IC01",name:"India Market Size",score:70,evidence:"$1.06B (2024) → $1.94B (2033) — mid-size TAM"},
        {code:"IC02",name:"Industry Growth (6.9% CAGR)",score:78,evidence:"Healthy; aligned with manufacturing capex cycle"},
        {code:"IC03",name:"Competitive Position",score:70,evidence:"Domestic leader in mid-market; global players in premium"},
        {code:"IC04",name:"Macro Tailwinds",score:82,evidence:"Make in India, PLI schemes, China+1 strategy — all favourable"},
        {code:"IC05",name:"End-Market Diversification",score:76,evidence:"7 industries: pharma, auto, packaging, food, medical, textile, print"},
      ]
    },
  ]
};

function scoreColor(s:number){
  if(s>=75) return {bg:"#dcfce7",text:"#16a34a",bar:"#16a34a"};
  if(s>=50) return {bg:"#fef9c3",text:"#b45309",bar:"#d97706"};
  return {bg:"#fee2e2",text:"#dc2626",bar:"#dc2626"};
}
function TrendIcon({t}:{t:string}){
  if(t==="up") return <TrendingUp size={14} className="text-green-500"/>;
  if(t==="down") return <TrendingDown size={14} className="text-red-500"/>;
  return <Minus size={14} className="text-slate-400"/>;
}

export default function ReportContent() {
  const params = useSearchParams();
  const router = useRouter();
  const isDemo = params.get("demo")==="1";
  const companyName = params.get("company") || DEMO.company;
  const data = DEMO;

  const [expanded, setExpanded] = useState<string|null>(null);
  const [loaded, setLoaded] = useState(false);

  useEffect(()=>{const t=setTimeout(()=>setLoaded(true),300);return()=>clearTimeout(t);},[]);

  return (
    <div className="min-h-screen bg-slate-50">
      {/* Nav */}
      <nav className="bg-white border-b border-slate-100 px-6 py-4 sticky top-0 z-10">
        <div className="max-w-6xl mx-auto flex items-center justify-between">
          <div className="flex items-center gap-4">
            <button onClick={()=>router.push(isDemo?"/":"/dashboard")}
              className="flex items-center gap-1.5 text-slate-500 hover:text-slate-700 text-sm transition">
              <ArrowLeft size={15}/> {isDemo?"Back to Login":"Back"}
            </button>
            <div className="w-px h-4 bg-slate-200"/>
            <div className="flex items-center gap-2">
              <div className="w-7 h-7 rounded bg-blue-600 flex items-center justify-center">
                <Building2 size={14} className="text-white"/>
              </div>
              <span className="font-bold text-slate-800 text-sm">DDMR Tool</span>
            </div>
          </div>
          <div className="flex items-center gap-2">
            {[
              {icon:FileSpreadsheet,label:"Excel",ext:"xlsx"},
              {icon:FileText,label:"Word",ext:"docx"},
              {icon:Presentation,label:"Deck",ext:"pptx"},
            ].map(({icon:Icon,label,ext})=>(
              <button key={ext}
                className="flex items-center gap-1.5 px-3 py-2 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 text-slate-600 text-xs font-medium transition"
                onClick={()=>alert(`Download feature: connect your Python generator to /api/download?format=${ext}`)}>
                <Icon size={13}/>{label}
              </button>
            ))}
          </div>
        </div>
      </nav>

      <div className="max-w-6xl mx-auto px-6 py-10">
        {/* Demo badge */}
        {isDemo && (
          <div className="mb-6 px-4 py-2.5 rounded-xl bg-blue-50 border border-blue-200 text-blue-700 text-sm flex items-center gap-2">
            <span className="w-2 h-2 rounded-full bg-blue-500 animate-pulse"/>
            <strong>Live Demo</strong> — Real company data from MCA, ICRA, and public sources. No API keys required.
          </div>
        )}

        {/* Company header */}
        <div className={`bg-white rounded-2xl border border-slate-100 p-8 mb-6 transition-all duration-500 ${loaded?"opacity-100 translate-y-0":"opacity-0 translate-y-4"}`}>
          <div className="flex flex-col lg:flex-row lg:items-start lg:justify-between gap-6">
            <div>
              <div className="flex items-center gap-3 mb-3">
                <div className="w-12 h-12 rounded-xl flex items-center justify-center" style={{background:"#eff6ff"}}>
                  <Building2 size={22} className="text-blue-600"/>
                </div>
                <div>
                  <h1 className="text-xl font-bold text-slate-900">{data.company}</h1>
                  <p className="text-slate-500 text-sm font-mono">{data.cin}</p>
                </div>
              </div>
              <p className="text-slate-600 text-sm leading-relaxed max-w-2xl">{data.summary}</p>
            </div>

            {/* Composite score */}
            <div className="flex-shrink-0 text-center px-8 py-6 rounded-2xl" style={{background:"linear-gradient(135deg,#0f2442,#1a3a6e)"}}>
              <div className="text-white/60 text-xs font-semibold uppercase tracking-widest mb-2">Composite Score</div>
              <div className="text-white text-5xl font-bold mb-1">{data.composite}</div>
              <div className="text-white/80 text-sm mb-3">out of 100</div>
              <div className="px-4 py-1.5 rounded-full text-sm font-bold text-white" style={{background:data.verdictColor}}>{data.verdict}</div>
            </div>
          </div>
        </div>

        {/* Meta grid */}
        <div className={`grid grid-cols-2 lg:grid-cols-5 gap-3 mb-6 transition-all duration-500 delay-100 ${loaded?"opacity-100 translate-y-0":"opacity-0 translate-y-4"}`}>
          {[
            {label:"Founded",val:data.meta.founded},
            {label:"HQ",val:data.meta.hq},
            {label:"Revenue",val:data.meta.revenue},
            {label:"Employees",val:data.meta.employees},
            {label:"Credit Rating",val:data.meta.rating},
            {label:"Countries",val:data.meta.countries},
            {label:"Products",val:data.meta.products},
            {label:"Customers",val:data.meta.customers},
            {label:"Directors",val:"3 Active"},
            {label:"Status",val:"Active"},
          ].map(({label,val})=>(
            <div key={label} className="bg-white rounded-xl border border-slate-100 px-4 py-3">
              <div className="text-slate-400 text-xs mb-0.5">{label}</div>
              <div className="text-slate-800 text-sm font-semibold truncate" title={val}>{val}</div>
            </div>
          ))}
        </div>

        {/* Highlights */}
        <div className={`bg-white rounded-2xl border border-slate-100 p-6 mb-6 transition-all duration-500 delay-150 ${loaded?"opacity-100 translate-y-0":"opacity-0 translate-y-4"}`}>
          <h2 className="text-sm font-semibold text-slate-500 uppercase tracking-wider mb-4">Key Findings</h2>
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-3">
            {data.highlights.map((h,i)=>(
              <div key={i} className={`flex items-start gap-3 p-3 rounded-xl ${h.type==="positive"?"bg-green-50":"bg-amber-50"}`}>
                <span className={`text-lg leading-none ${h.type==="positive"?"text-green-500":"text-amber-500"}`}>
                  {h.type==="positive"?"✓":"⚠"}
                </span>
                <span className={`text-sm ${h.type==="positive"?"text-green-800":"text-amber-800"}`}>{h.text}</span>
              </div>
            ))}
          </div>
        </div>

        {/* Dimension cards */}
        <div className="space-y-4">
          {data.dimensions.map((d,di)=>{
            const c=scoreColor(d.score);
            const open=expanded===d.code;
            return (
              <div key={d.code}
                className={`bg-white rounded-2xl border border-slate-100 overflow-hidden transition-all duration-500 score-card`}
                style={{transitionDelay:`${di*60}ms`, opacity:loaded?1:0, transform:loaded?"none":"translateY(16px)"}}>
                <button className="w-full px-6 py-5 flex items-center gap-5 text-left hover:bg-slate-50/50 transition"
                  onClick={()=>setExpanded(open?null:d.code)}>
                  {/* Score badge */}
                  <div className="flex-shrink-0 w-16 h-16 rounded-xl flex flex-col items-center justify-center" style={{background:c.bg}}>
                    <span className="text-2xl font-bold leading-none" style={{color:c.text}}>{d.score}</span>
                    <span className="text-xs font-medium mt-0.5" style={{color:c.text}}>/100</span>
                  </div>

                  {/* Info */}
                  <div className="flex-1 min-w-0">
                    <div className="flex items-center gap-2 mb-1">
                      <span className="text-xs font-bold text-slate-400 uppercase tracking-wider">{d.code}</span>
                      <span className="text-xs text-slate-400">·</span>
                      <span className="text-xs text-slate-400">Weight {d.weight}</span>
                      <TrendIcon t={d.trend}/>
                    </div>
                    <h3 className="font-semibold text-slate-800 text-base">{d.name}</h3>
                    <p className="text-slate-500 text-sm truncate">{d.insight}</p>
                  </div>

                  {/* Progress bar */}
                  <div className="hidden lg:block w-32">
                    <div className="h-2 rounded-full bg-slate-100 overflow-hidden">
                      <div className="h-full rounded-full transition-all duration-1000" style={{width:loaded?`${d.score}%`:"0%",background:c.bar}}/>
                    </div>
                  </div>

                  <div className="flex-shrink-0 text-slate-400">
                    {open?<ChevronUp size={18}/>:<ChevronDown size={18}/>}
                  </div>
                </button>

                {open && (
                  <div className="border-t border-slate-100 px-6 py-5 animate-fadeInUp">
                    <p className="text-slate-600 text-sm mb-5 leading-relaxed">{d.insight}</p>
                    <div className="space-y-3">
                      {d.submetrics.map(sm=>{
                        const sc=scoreColor(sm.score);
                        return (
                          <div key={sm.code} className="flex items-start gap-4 p-3 rounded-xl bg-slate-50">
                            <div className="flex-shrink-0 w-12 text-center">
                              <span className="text-lg font-bold" style={{color:sc.text}}>{sm.score}</span>
                              <div className="text-slate-400 text-xs">/100</div>
                            </div>
                            <div className="flex-1 min-w-0">
                              <div className="text-xs font-bold text-slate-400 uppercase mb-0.5">{sm.code}</div>
                              <div className="text-sm font-semibold text-slate-700 mb-0.5">{sm.name}</div>
                              <div className="text-slate-500 text-xs leading-relaxed">{sm.evidence}</div>
                            </div>
                            <div className="flex-shrink-0 w-20 pt-2">
                              <div className="h-1.5 rounded-full bg-slate-200 overflow-hidden">
                                <div className="h-full rounded-full" style={{width:`${sm.score}%`,background:sc.bar}}/>
                              </div>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                )}
              </div>
            );
          })}
        </div>

        {/* Sources footer */}
        <div className="mt-8 p-5 rounded-2xl bg-slate-100 border border-slate-200">
          <h3 className="text-xs font-semibold text-slate-500 uppercase tracking-wider mb-3">Data Sources</h3>
          <div className="flex flex-wrap gap-2">
            {["MCA21 / ROC Coimbatore","ICRA Rationale Reports","Tracxn / Tofler","Janatics.com","janaticspneumatics.com","IMARC Group","AIA India","DSIR"].map(s=>(
              <span key={s} className="px-3 py-1 rounded-full bg-white border border-slate-200 text-slate-600 text-xs font-medium">{s}</span>
            ))}
          </div>
          <p className="text-slate-400 text-xs mt-3">Report generated {new Date().toLocaleDateString("en-IN",{day:"numeric",month:"long",year:"numeric"})} · Based on publicly available information · Not investment advice</p>
        </div>
      </div>
    </div>
  );
}
