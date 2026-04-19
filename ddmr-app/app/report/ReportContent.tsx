"use client";
import { useSearchParams, useRouter } from "next/navigation";
import { useState, useEffect } from "react";
import { Building2, Download, ChevronDown, ChevronUp, ArrowLeft, TrendingUp, TrendingDown, Minus, ExternalLink, FileSpreadsheet, FileText, Presentation } from "lucide-react";

const DEMO = {
  company: "Janatics India Private Limited",
  cin: "U31103TZ1991PTC003409",
  sector: "Pneumatics & Industrial Automation",
  composite: 64,
  maturity: "Siloed",
  verdictColor: "#d97706",
  summary: "Janatics India Pvt Ltd (est. 1977) is assessed for digital transformation maturity using public signals across five dimensions. With ₹531 Cr revenue, ICRA Stable credit, and an emerging Industry 4.0 product portfolio, the company scores 64/100 — placing it in the Siloed maturity band. Digital awareness is evident in product strategy, but systematic adoption of digital tools in operations, sales, and workforce development remains nascent.",
  highlights: [
    {type:"positive", text:"Industry 4.0 product pivot: electric actuators, sensors, robotics added to portfolio"},
    {type:"positive", text:"DSIR-recognised R&D division signals structured technology investment"},
    {type:"positive", text:"Strong operational base: 70,000 sqm plant, 3,500+ SKUs, ICRA Stable credit"},
    {type:"positive", text:"Digital product awareness aligned with smart manufacturing trends"},
    {type:"warning",  text:"No evidence of digital sales channel, e-commerce, or CRM platform deployment"},
    {type:"warning",  text:"Internal digitization (ERP/MES/cloud) not publicly disclosed — key gap"},
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
      code:"ST", name:"Strategy", score:70, weight:"20%", trend:"up",
      maturity:"Siloed",
      insight:"Digital intent visible in I4.0 product launches and DSIR R&D. No formal digital transformation roadmap evidenced publicly.",
      submetrics:[
        {code:"ST01",name:"Digital Strategy Articulation",score:65,evidence:"I4.0 and digital products mentioned but no formal DX roadmap or digital strategy document publicly disclosed"},
        {code:"ST02",name:"Leadership Technology Vision",score:72,evidence:"Management pivoted to electric actuators, robotics, and sensors — signals awareness of digital megatrends"},
        {code:"ST03",name:"Digital Investment Signals",score:75,evidence:"New digital product lines (IoT-enabled actuators, smart sensors) indicate technology capex commitment"},
        {code:"ST04",name:"Technology Partnerships",score:68,evidence:"DSIR R&D recognition confirmed; no public tech alliance or digital ecosystem partnerships disclosed"},
        {code:"ST05",name:"Innovation Pipeline",score:70,evidence:"Incremental product digitization underway; no open innovation, startup collaboration, or IP filing trail public"},
      ]
    },
    {
      code:"OSC", name:"Operations & Supply Chain", score:65, weight:"20%", trend:"neutral",
      maturity:"Siloed",
      insight:"Strong physical manufacturing base. Digital operations layer (ERP, MES, supply chain visibility) not publicly evidenced.",
      submetrics:[
        {code:"OSC01",name:"Digital Manufacturing Readiness",score:60,evidence:"70,000 sqm Coimbatore plant with CNC machining; no Industry 4.0 certification or smart factory disclosure"},
        {code:"OSC02",name:"Supply Chain Visibility Tools",score:50,evidence:"No public disclosure of ERP, SCM platform, or supply chain digitization initiative"},
        {code:"OSC03",name:"Process Automation Adoption",score:70,evidence:"CNC machining, casting automation, surface treatment lines — physical automation well established"},
        {code:"OSC04",name:"Quality Management (Digital)",score:72,evidence:"ICRA notes consistent regulatory compliance; ISO/quality system assumed but not publicly confirmed"},
        {code:"OSC05",name:"Industry 4.0 Integration",score:73,evidence:"I4.0 product portfolio exists; internal I4.0 adoption for own manufacturing processes not disclosed"},
      ]
    },
    {
      code:"SM", name:"Sales & Marketing", score:55, weight:"20%", trend:"down",
      maturity:"Siloed",
      insight:"Functional web presence and large customer base. Digital sales channel, social engagement, and marketing automation are minimal.",
      submetrics:[
        {code:"SM01",name:"Digital Presence & SEO",score:60,evidence:"Two product websites (janatics.com, janaticspneumatics.com); catalog available online; SEO signals moderate"},
        {code:"SM02",name:"Social Media Engagement",score:40,evidence:"Limited LinkedIn activity signal; no active social media marketing evidence from public sources"},
        {code:"SM03",name:"E-commerce / Digital Ordering",score:45,evidence:"No evidence of digital ordering portal, e-catalog with pricing, or B2B e-commerce capability"},
        {code:"SM04",name:"Digital Marketing Activity",score:55,evidence:"Web presence established; no evidence of paid digital campaigns, content marketing, or SEO investment"},
        {code:"SM05",name:"CRM & Customer Data Utilization",score:75,evidence:"17,000+ customers managed across 7 industries — implies structured customer data; CRM platform undisclosed"},
      ]
    },
    {
      code:"TA", name:"Technology Adoption", score:70, weight:"20%", trend:"up",
      maturity:"Siloed",
      insight:"Strong digital product R&D (IoT, sensors, robotics). Internal technology infrastructure (cloud, data, cyber) not publicly evidenced.",
      submetrics:[
        {code:"TA01",name:"Website & Digital Infrastructure",score:65,evidence:"Functional product websites; no evidence of modern tech stack, API integrations, or customer portal"},
        {code:"TA02",name:"IoT / Smart Product Portfolio",score:80,evidence:"Electric actuators, smart sensors, robotic systems, Industry 4.0 didactic equipment launched — strong product digitization"},
        {code:"TA03",name:"Cloud & SaaS Adoption",score:60,evidence:"No public disclosures of cloud migration, SaaS tools, or digital collaboration platforms"},
        {code:"TA04",name:"R&D in Digital / Smart Tech",score:78,evidence:"DSIR-recognized R&D actively developing robotic systems, sensor products, and smart automation equipment"},
        {code:"TA05",name:"Cybersecurity & Data Practices",score:67,evidence:"ICRA-compliant operational practices; no public cybersecurity framework, ISO 27001, or data policy disclosed"},
      ]
    },
    {
      code:"SC", name:"Skills & Capabilities", score:60, weight:"20%", trend:"neutral",
      maturity:"Siloed",
      insight:"Engineering-depth workforce. No public evidence of digital upskilling, tech talent acquisition strategy, or learning programs.",
      submetrics:[
        {code:"SC01",name:"Digital Talent Hiring Signals",score:55,evidence:"662 employees primarily engineering/manufacturing; limited public signals of software, data, or digital roles hired"},
        {code:"SC02",name:"Technical Workforce Depth",score:72,evidence:"Strong mechanical and manufacturing engineering talent pool; DSIR R&D team; CNC and automation specialists"},
        {code:"SC03",name:"Digital Training Programs",score:50,evidence:"No public evidence of L&D investment, digital upskilling programs, or online learning platform adoption"},
        {code:"SC04",name:"Leadership Digital Literacy",score:65,evidence:"Management demonstrates awareness of I4.0 trends via product strategy decisions; no DX leadership hire signals"},
        {code:"SC05",name:"Innovation Culture Signals",score:58,evidence:"DSIR R&D present; no public hackathon, innovation lab, startup engagement, or employee innovation program signals"},
      ]
    },
  ],
  nextActions: [
    {rank:"01",priority:"High",priorityColor:"#dc2626",title:"Activate a Digital Sales Channel",dimension:"Sales & Marketing (SM03: 45/100)",initiative:"Launch a B2B digital ordering portal with searchable product catalog, real-time pricing, and quote-request workflows. Pair with SEO content and LinkedIn demand generation.",benefit:"Capture online-first industrial buyers; reduce dependency on distributor-led lead generation; unlock an estimated 15–20% incremental revenue from direct digital channel within 18 months.",loss:"Global competitors (SMC, Festo) already operate digital-first sales. Every quarter without a digital channel cedes share as India's B2B industrial procurement shifts online.",whyNow:"India's B2B e-commerce in industrial components is growing 25%+ annually. Janatics has the brand (49 years) and product depth (3,500+ SKUs) to become the dominant online destination for domestic pneumatics.",},
    {rank:"02",priority:"High",priorityColor:"#dc2626",title:"Deploy ERP and Supply Chain Visibility Platform",dimension:"Operations & Supply Chain (OSC02: 50/100)",initiative:"Implement an ERP (SAP/Oracle/Odoo) with connected MES and supply chain visibility layer. Publish a digitization roadmap to customers and ICRA as evidence of operational modernization.",benefit:"10–15% reduction in production inefficiencies; faster order-to-delivery; ability to pass global OEM digital supply chain audits — a prerequisite for large export contracts.",loss:"Global OEM customers increasingly mandate supply chain transparency as a vendor qualification criterion. Without digital traceability, Janatics risks disqualification from high-value export RFQs.",whyNow:"Make in India + China+1 is creating an influx of global OEM procurement inquiries. Janatics has a 12–18 month window to establish digital operational credibility before OEM qualification pipelines close.",},
    {rank:"03",priority:"Medium",priorityColor:"#d97706",title:"Build a Digital Workforce Capability Program",dimension:"Skills & Capabilities (SC03: 50/100, SC01: 55/100)",initiative:"Launch a structured digital upskilling program for all engineering and management staff. Hire 2–3 digital roles (data analyst, digital marketing manager, ERP lead). Partner with an L&D platform.",benefit:"Accelerates adoption of any digital tool investment; reduces transformation failure risk; builds internal champions who sustain change without consultant dependency.",loss:"Technology without people is stranded investment. Every digital initiative — ERP, e-commerce, analytics — will underperform if the workforce lacks capability and confidence to use it.",whyNow:"India's digital skills gap is widening. Early movers in manufacturing upskilling will have a 2–3 year head start on talent acquisition and retention.",},
    {rank:"04",priority:"Medium",priorityColor:"#d97706",title:"Turn the Coimbatore Plant into a Smart Factory Showcase",dimension:"Technology Adoption (TA01: 65/100) + Strategy (ST03: 75/100)",initiative:"Deploy Janatics' own I4.0 sensors, electric actuators, and robotic systems inside the Coimbatore manufacturing plant. Create a live demo environment for customer visits and virtual tours.",benefit:"Validates product claims with proof-by-example; accelerates B2B sales cycles; generates case study content for marketing; signals vendor digital maturity to global OEM evaluators.",loss:"Selling Industry 4.0 products while operating a non-digitized factory is a credibility gap that sophisticated buyers will notice during site audits and RFQ evaluations.",whyNow:"Janatics already owns the technology. Internal deployment requires no new vendor — it is the fastest and most credible path to a demonstrable digital maturity story.",},
  ],
};

const DEMO_MESSER = {
  company: "Messer Cutting Systems India Private Limited",
  cin: "U29294MH2004PTC145670",
  sector: "Thermal & Laser Cutting Systems (CNC Machinery)",
  composite: 72,
  maturity: "Strategic",
  verdictColor: "#16a34a",
  summary: "Messer Cutting Systems India Pvt Ltd is the Indian subsidiary of Messer Cutting Systems GmbH (Germany), a global leader in thermal cutting technology since 1898. Operating from Pune, Maharashtra, the company supplies plasma, laser, and oxy-fuel CNC cutting systems to India's steel, automotive, shipbuilding, and fabrication sectors. The DDMR composite score of 72/100 (Strategic) reflects strong technology adoption driven by the parent's inherently digital product portfolio, global R&D pipeline, and MNC-grade operational practices — partially offset by limited India-specific digital marketing and sales channel activation.",
  highlights: [
    {type:"positive", text:"Products are inherently digital: CNC plasma, fiber laser, and remote-diagnostic cutting systems"},
    {type:"positive", text:"Global parent (Messer GmbH, Germany) provides Innovation 4.0 roadmap and Connected Cutting IoT platform"},
    {type:"positive", text:"MNC-grade operational practices: SAP ERP, ISO quality systems, and German engineering standards"},
    {type:"positive", text:"Strong technical workforce with parent-backed L&D programs and certification pathways"},
    {type:"warning",  text:"India-specific digital marketing, SEO, and social media presence largely derivative of global parent"},
    {type:"warning",  text:"No India-specific DX roadmap or local digital investment disclosure evidenced publicly"},
  ],
  meta: {
    incorporated: "2004",
    founded: "1898 (Global) / 2004 (India)",
    hq: "Pune, Maharashtra",
    employees: "~320",
    revenue: "₹185 Crore (est. FY2025)",
    rating: "MNC Subsidiary — Parent-backed",
    countries: "India + SAARC exports",
    products: "Plasma, Laser, Oxy-fuel CNC Cutters",
    customers: "Steel mills, Automotive OEMs, Shipyards",
    directors: "3 Active Directors",
  },
  dimensions: [
    {
      code:"ST", name:"Strategy", score:75, weight:"20%", trend:"up",
      maturity:"Strategic",
      insight:"Global parent's Industry 4.0 roadmap cascades to India ops. No India-specific DX strategy publicly disclosed.",
      submetrics:[
        {code:"ST01",name:"Digital Strategy Articulation",score:78,evidence:"Global 'Messerworld' platform and Industry 4.0 strategy explicitly stated by parent; India-specific DX roadmap not separately disclosed"},
        {code:"ST02",name:"Leadership Technology Vision",score:80,evidence:"Products ARE the technology — CNC, plasma, fiber laser, automation cells signal deep digital leadership alignment"},
        {code:"ST03",name:"Digital Investment Signals",score:75,evidence:"India launches new fiber laser and plasma automation products — signals active technology investment pipeline"},
        {code:"ST04",name:"Technology Partnerships",score:72,evidence:"Parent partnerships with Siemens, Hypertherm, and Kjellberg for CNC controllers and plasma technology"},
        {code:"ST05",name:"Innovation Pipeline",score:70,evidence:"Parent-driven R&D pipeline; India-specific product innovation or local IP signals limited in public disclosures"},
      ]
    },
    {
      code:"OSC", name:"Operations & Supply Chain", score:72, weight:"20%", trend:"up",
      maturity:"Siloed",
      insight:"MNC-grade ERP (SAP) likely in place. Supply chain visibility spans global parent network. India-specific disclosures limited.",
      submetrics:[
        {code:"OSC01",name:"Digital Manufacturing Readiness",score:75,evidence:"Pune assembly plant for CNC cutting systems; MNC quality standards; SAP ERP assumed given German parent"},
        {code:"OSC02",name:"Supply Chain Visibility Tools",score:70,evidence:"Global supply chain from Germany and Europe; parent-level SAP integration likely; India-specific visibility undisclosed"},
        {code:"OSC03",name:"Process Automation Adoption",score:78,evidence:"Assembly of precision CNC cutting machinery inherently requires high process accuracy and automation"},
        {code:"OSC04",name:"Quality Management (Digital)",score:72,evidence:"ISO 9001 certification implied by MNC parent standards; digital quality management assumed but unconfirmed for India"},
        {code:"OSC05",name:"Industry 4.0 Integration",score:65,evidence:"Parent's Connected Cutting platform offers remote diagnostics; India-specific I4.0 internal adoption undisclosed"},
      ]
    },
    {
      code:"SM", name:"Sales & Marketing", score:62, weight:"20%", trend:"neutral",
      maturity:"Siloed",
      insight:"Global website with India page. Moderate LinkedIn presence. No B2B e-commerce or India-specific digital demand generation.",
      submetrics:[
        {code:"SM01",name:"Digital Presence & SEO",score:65,evidence:"messercuttingsystems.com India page exists; product specs online; SEO driven by global parent domain"},
        {code:"SM02",name:"Social Media Engagement",score:55,evidence:"LinkedIn company page active with moderate content; better than pure domestic players but below digital-native B2B standards"},
        {code:"SM03",name:"E-commerce / Digital Ordering",score:45,evidence:"Industrial machinery RFQ-based sales — no digital ordering, CPQ tool, or online pricing for India market"},
        {code:"SM04",name:"Digital Marketing Activity",score:65,evidence:"Trade show digital presence (IMTEX, India Essen); some content marketing via global parent; India-specific campaigns limited"},
        {code:"SM05",name:"CRM & Customer Data Utilization",score:80,evidence:"MNC-grade CRM (Salesforce or SAP CRM) assumed from parent standards; customer lifecycle management structured"},
      ]
    },
    {
      code:"TA", name:"Technology Adoption", score:80, weight:"20%", trend:"up",
      maturity:"Strategic",
      insight:"Products are digital technology. Connected Cutting IoT platform, fiber laser R&D, and cloud-enabled remote diagnostics place this dimension in Strategic band.",
      submetrics:[
        {code:"TA01",name:"Website & Digital Infrastructure",score:72,evidence:"Global-standard website with India page; decent CMS and product configurator; not fully localised for India buyers"},
        {code:"TA02",name:"IoT / Smart Product Portfolio",score:90,evidence:"Connected Cutting Suite: IoT-enabled remote diagnostics, predictive maintenance, and cloud-connected CNC systems — industry-leading"},
        {code:"TA03",name:"Cloud & SaaS Adoption",score:75,evidence:"Parent's 'Messer Connected Cutting' cloud platform; Microsoft 365 / cloud collaboration assumed at MNC level"},
        {code:"TA04",name:"R&D in Digital / Smart Tech",score:88,evidence:"Global R&D in fiber laser, AI-assisted cutting path optimisation, automation cells, and Industry 4.0 cutting workflows"},
        {code:"TA05",name:"Cybersecurity & Data Practices",score:75,evidence:"MNC cybersecurity standards (ISO 27001 at parent level); GDPR compliance for EU data flows; India-specific policy undisclosed"},
      ]
    },
    {
      code:"SC", name:"Skills & Capabilities", score:70, weight:"20%", trend:"up",
      maturity:"Siloed",
      insight:"German engineering culture and parent L&D programs provide strong technical base. India-specific digital talent development and innovation culture signals are limited.",
      submetrics:[
        {code:"SC01",name:"Digital Talent Hiring Signals",score:65,evidence:"LinkedIn shows engineering, sales, and service roles; some software/IoT roles visible; below tech-company hiring intensity"},
        {code:"SC02",name:"Technical Workforce Depth",score:82,evidence:"High-precision CNC machinery engineering requires deep technical expertise; German-standard quality training"},
        {code:"SC03",name:"Digital Training Programs",score:68,evidence:"Parent-backed global training programs and certifications (cutting technology, CNC programming); digital upskilling implied"},
        {code:"SC04",name:"Leadership Digital Literacy",score:72,evidence:"MNC management with global exposure; digital literacy at leadership level assumed given product portfolio"},
        {code:"SC05",name:"Innovation Culture Signals",score:63,evidence:"Parent-level innovation culture strong; India-specific innovation programs, hackathons, or startup engagement not evidenced"},
      ]
    },
  ],
  nextActions: [
    {
      rank:"01", priority:"High", priorityColor:"#dc2626",
      title:"Activate India-Specific Digital Marketing & SEO",
      dimension:"Sales & Marketing (SM01: 65/100, SM04: 65/100)",
      initiative:"Launch India-specific content marketing targeting industrial buyers — application case studies for steel, auto, and shipbuilding. Invest in Hindi/English SEO for cutting systems, plasma cutter, and laser cutting machine search terms. Run LinkedIn campaigns targeting plant engineers and procurement managers.",
      benefit:"Capture India's growing industrial machinery online search traffic (est. 30%+ CAGR); establish Messer as the default digital resource for CNC cutting knowledge; shorten enterprise sales cycles by 20–30% through inbound lead quality improvement.",
      loss:"Global competitors (ESAB, Lincoln Electric, Trumpf) are investing in India digital marketing. Without India-specific content, Messer's India page is invisible to most local industrial buyers who search in vernacular or India-specific terms.",
      whyNow:"India's manufacturing capex is at a 10-year high (PLI, infrastructure, defence). The window to become the dominant digital voice in India's cutting systems market is open now — before better-funded global players establish SEO authority.",
    },
    {
      rank:"02", priority:"High", priorityColor:"#dc2626",
      title:"Launch Connected Cutting IoT Upsell for India Installed Base",
      dimension:"Technology Adoption (TA03: 75/100) + Operations (OSC05: 65/100)",
      initiative:"Deploy the global 'Messer Connected Cutting' IoT platform as a structured India upsell program targeting the existing installed base of cutting systems. Offer remote diagnostics, predictive maintenance alerts, and consumables tracking via subscription model. Create an India-specific customer portal.",
      benefit:"Recurring revenue stream on hardware-sold installed base; improved customer retention and NPS; reduction in service call costs (est. 25–35%); data on India usage patterns to inform product roadmap; positions Messer India as 'smart manufacturing partner' vs. equipment vendor.",
      loss:"Every unconnected machine in the India installed base is a missed recurring revenue opportunity and a retention risk. Competitors offering IoT-enabled service contracts are more 'sticky' — customers who connect machines rarely switch brands.",
      whyNow:"India's industrial IoT adoption is accelerating. Early movers who connect India installed bases first will own the service relationship. The parent platform already exists — India activation requires localisation, not new development.",
    },
    {
      rank:"03", priority:"Medium", priorityColor:"#d97706",
      title:"Build India Digital Sales Capability (CPQ + Self-Service Portal)",
      dimension:"Sales & Marketing (SM03: 45/100)",
      initiative:"Implement a Configure-Price-Quote (CPQ) tool for standard cutting system configurations — allowing buyers to specify material type, thickness, and volume, and receive instant pricing guidance. Add a self-service spare parts and consumables ordering portal for existing customers.",
      benefit:"Reduce sales cycle length for standard configurations from weeks to days; capture consumables revenue digitally (high-margin, high-frequency); free up sales engineers for complex/custom deals; demonstrate digital maturity to MNC OEM procurement teams.",
      loss:"Industrial buyers are increasingly expecting digital self-service for standard purchases. Without CPQ or self-service, Messer India's sales team handles every enquiry manually — a capacity constraint that limits growth as order volumes scale.",
      whyNow:"India's manufacturing sector is entering a rapid automation adoption phase. Companies that build digital sales infrastructure now will process 3–5x the order volume without proportionate headcount growth — a structural competitive advantage.",
    },
    {
      rank:"04", priority:"Medium", priorityColor:"#d97706",
      title:"Establish India Innovation & Applications Lab",
      dimension:"Skills & Capabilities (SC05: 63/100) + Strategy (ST05: 70/100)",
      initiative:"Set up an India-specific Applications Lab in Pune where customers can test Messer cutting solutions on their actual materials before purchase. Use the lab to co-develop India-specific cutting parameters and publish application notes. Partner with IITs or CMTI for cutting technology research.",
      benefit:"Shortens sales cycles by enabling proof-of-concept at no customer risk; generates authentic India application content for marketing; builds local R&D capability; creates a talent pipeline through academic partnerships; signals long-term India commitment to global OEM customers.",
      loss:"Without a local applications lab, Messer India must rely on Germany-based demonstrations for complex customer evaluations — adding weeks to sales cycles and reducing win rates on time-sensitive RFQs from domestic manufacturers.",
      whyNow:"India's manufacturing ecosystem is maturing rapidly. Local applications capability is now a qualifier, not a differentiator, for global cutting systems buyers. Competitors are establishing India technical centres — Messer must match or exceed this investment.",
    },
  ],
};

function maturityLabel(s:number){
  if(s>=84) return "Future Ready";
  if(s>=67) return "Strategic";
  if(s>=34) return "Siloed";
  return "Legacy";
}
function scoreColor(s:number){
  if(s>=84) return {bg:"#dbeafe",text:"#1d4ed8",bar:"#2563eb"};
  if(s>=67) return {bg:"#dcfce7",text:"#16a34a",bar:"#16a34a"};
  if(s>=34) return {bg:"#fef9c3",text:"#b45309",bar:"#d97706"};
  return {bg:"#fee2e2",text:"#dc2626",bar:"#dc2626"};
}
function TrendIcon({t}:{t:string}){
  if(t==="up") return <TrendingUp size={14} className="text-green-500"/>;
  if(t==="down") return <TrendingDown size={14} className="text-red-500"/>;
  return <Minus size={14} className="text-slate-400"/>;
}

function getReportById(id: string): any | null {
  try {
    const reports: any[] = JSON.parse(localStorage.getItem("ddmr_reports") || "[]");
    return reports.find(r => r.id === id) ?? null;
  } catch { return null; }
}

export default function ReportContent() {
  const params = useSearchParams();
  const router = useRouter();
  const demoSlug = params.get("demo");
  const reportId = params.get("id");
  const isDemo = demoSlug !== null && demoSlug !== "";
  const isGenerated = reportId !== null;

  const [generatedData, setGeneratedData] = useState<any>(null);
  const [reportNotFound, setReportNotFound] = useState(false);
  const [expanded, setExpanded] = useState<string|null>(null);
  const [loaded, setLoaded] = useState(false);
  const [downloading, setDownloading] = useState<string|null>(null);
  const [downloadError, setDownloadError] = useState("");

  useEffect(()=>{
    if (reportId) {
      try {
        const stored = getReportById(reportId);
        if (stored) {
          const d = stored.data;
          setGeneratedData({
            ...d,
            cin: d.cin || "",
            maturity: d.maturity || maturityLabel(d.composite),
            verdictColor: d.verdictColor || scoreColor(d.composite).bar,
            meta: {
              incorporated: "",
              founded: d.meta?.founded || "Not disclosed",
              hq: d.meta?.hq || "Not disclosed",
              employees: d.meta?.employees || "Not disclosed",
              revenue: d.meta?.revenue || "Not disclosed",
              rating: "",
              countries: d.meta?.countries || "Not disclosed",
              products: d.meta?.products || "Not disclosed",
              customers: d.meta?.customers || "Not disclosed",
              directors: "",
            },
          });
        } else {
          setReportNotFound(true);
        }
      } catch { setReportNotFound(true); }
    }
    const t=setTimeout(()=>setLoaded(true),300);
    return()=>clearTimeout(t);
  },[reportId]);

  if (isGenerated && reportNotFound) {
    return (
      <div className="min-h-screen bg-slate-50 flex items-center justify-center p-6">
        <div className="text-center max-w-sm">
          <div className="w-14 h-14 rounded-2xl bg-slate-100 flex items-center justify-center mx-auto mb-4">
            <Building2 size={24} className="text-slate-400"/>
          </div>
          <h2 className="text-xl font-bold text-slate-800 mb-2">Report not found</h2>
          <p className="text-slate-500 text-sm mb-6">This report is stored in the browser where it was generated. Reports cannot be shared via URL.</p>
          <button onClick={()=>router.push("/dashboard")}
            className="px-6 py-2.5 rounded-xl bg-blue-600 text-white text-sm font-semibold hover:bg-blue-700 transition">
            Back to Dashboard
          </button>
        </div>
      </div>
    );
  }

  if (isGenerated && !generatedData) {
    return (
      <div className="min-h-screen bg-slate-50 flex items-center justify-center">
        <div className="text-center">
          <svg className="animate-spin h-8 w-8 text-blue-500 mx-auto mb-4" viewBox="0 0 24 24" fill="none">
            <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"/>
            <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8H4z"/>
          </svg>
          <p className="text-slate-600 text-sm">Loading report...</p>
        </div>
      </div>
    );
  }

  const data = isGenerated ? generatedData : (demoSlug === "messer" ? DEMO_MESSER : DEMO);

  async function downloadGenerated(format: string) {
    setDownloading(format);
    setDownloadError("");
    try {
      const res = await fetch("/api/export", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ format, data }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: "Export failed" }));
        setDownloadError(err.error || "Export failed. Please try again.");
        setDownloading(null);
        return;
      }
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${(data.company || "DDMR").replace(/[^a-z0-9]/gi, "_")}_DDMR.${format}`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 2000);
    } catch (err) {
      setDownloadError("Download failed. Please try again.");
    }
    setDownloading(null);
  }

  return (
    <div className="min-h-screen bg-slate-50">
      {/* Nav */}
      <nav className="bg-white border-b border-slate-100 px-6 py-4 sticky top-0 z-10">
        <div className="max-w-6xl mx-auto flex items-center justify-between">
          <div className="flex items-center gap-4">
            <button onClick={()=>router.push(isGenerated?"/dashboard":(isDemo?"/":"/dashboard"))}
              className="flex items-center gap-1.5 text-slate-500 hover:text-slate-700 text-sm transition">
              <ArrowLeft size={15}/> {isGenerated ? "Back to Dashboard" : (isDemo ? "Back to Login" : "Back")}
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
              {icon:FileSpreadsheet,label:"Excel",ext:"xlsx",color:"text-emerald-600"},
              {icon:FileText,label:"Word",ext:"docx",color:"text-blue-600"},
              {icon:Presentation,label:"Deck",ext:"pptx",color:"text-orange-500"},
              {icon:FileText,label:"PDF",ext:"pdf",color:"text-red-500"},
            ].map(({icon:Icon,label,ext,color})=>(
              isGenerated ? (
                <button key={ext} onClick={()=>downloadGenerated(ext)} disabled={downloading !== null}
                  className={`flex items-center gap-1.5 px-3 py-2 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 ${color} text-xs font-medium transition disabled:opacity-50`}>
                  {downloading === ext
                    ? <svg className="animate-spin h-3 w-3" viewBox="0 0 24 24" fill="none"><circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"/><path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8H4z"/></svg>
                    : <Icon size={13}/>}
                  {label}
                </button>
              ) : (
                <a key={ext} href={`/api/download?format=${ext}`} download
                  className={`flex items-center gap-1.5 px-3 py-2 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 ${color} text-xs font-medium transition`}>
                  <Icon size={13}/>{label}
                </a>
              )
            ))}
          </div>
        </div>
      </nav>

      <div className="max-w-6xl mx-auto px-6 py-10">
        {downloadError && (
          <div className="mb-4 px-4 py-2.5 rounded-xl bg-red-50 border border-red-200 text-red-700 text-sm flex items-center justify-between">
            <span>{downloadError}</span>
            <button onClick={()=>setDownloadError("")} className="text-red-400 hover:text-red-600 ml-4 text-lg leading-none">×</button>
          </div>
        )}
        {isGenerated && (
          <div className="mb-6 px-4 py-2.5 rounded-xl bg-green-50 border border-green-200 text-green-700 text-sm flex items-center gap-2">
            <span className="w-2 h-2 rounded-full bg-green-500 animate-pulse"/>
            <strong>AI Generated</strong> — Report created by Claude AI using public signals. Verify findings independently before decision-making.
          </div>
        )}
        {isDemo && !isGenerated && (
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
                  {data.cin && <p className="text-slate-500 text-sm font-mono">{data.cin}</p>}
                  {data.sector && <p className="text-slate-400 text-xs mt-0.5">{data.sector}</p>}
                </div>
              </div>
              <p className="text-slate-600 text-sm leading-relaxed max-w-2xl">{data.summary}</p>
            </div>

            {/* Composite score */}
            <div className="flex-shrink-0 text-center px-8 py-6 rounded-2xl" style={{background:"linear-gradient(135deg,#0f2442,#1a3a6e)"}}>
              <div className="text-white/60 text-xs font-semibold uppercase tracking-widest mb-2">Composite Score</div>
              <div className="text-white text-5xl font-bold mb-1">{data.composite}</div>
              <div className="text-white/80 text-sm mb-3">out of 100</div>
              <div className="px-4 py-1.5 rounded-full text-sm font-bold text-white" style={{background:data.verdictColor}}>{data.maturity}</div>
            </div>
          </div>
        </div>

        {/* Meta grid */}
        <div className={`grid grid-cols-2 lg:grid-cols-5 gap-3 mb-6 transition-all duration-500 delay-100 ${loaded?"opacity-100 translate-y-0":"opacity-0 translate-y-4"}`}>
          {([
            {label:"Founded",val:data.meta.founded},
            {label:"HQ",val:data.meta.hq},
            {label:"Revenue",val:data.meta.revenue},
            {label:"Employees",val:data.meta.employees},
            data.meta.rating ? {label:"Credit Rating",val:data.meta.rating} : null,
            {label:"Countries",val:data.meta.countries},
            {label:"Products",val:data.meta.products},
            {label:"Customers",val:data.meta.customers},
            data.meta.directors ? {label:"Directors",val:data.meta.directors} : null,
            !isGenerated ? {label:"Status",val:"Active"} : null,
          ] as ({label:string,val:string}|null)[]).filter((x): x is {label:string,val:string} => x !== null && !!x.val).map(({label,val})=>(
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
            {data.highlights.map((h:any,i:number)=>(
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
          {data.dimensions.map((d:any,di:number)=>{
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
                      <span className="text-xs font-semibold px-1.5 py-0.5 rounded" style={{background:scoreColor(d.score).bg,color:scoreColor(d.score).text}}>{maturityLabel(d.score)}</span>
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
                      {d.submetrics.map((sm:any)=>{
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

        {/* Next Best Actions */}
        <div className={`mt-8 transition-all duration-500 delay-300 ${loaded?"opacity-100 translate-y-0":"opacity-0 translate-y-4"}`}>
          <div className="flex items-center gap-3 mb-5">
            <div className="w-8 h-8 rounded-lg bg-slate-900 flex items-center justify-center flex-shrink-0">
              <span className="text-white text-sm font-bold">→</span>
            </div>
            <div>
              <h2 className="text-lg font-bold text-slate-900">Next Best Actions</h2>
              <p className="text-slate-500 text-sm">Priority digital transformation initiatives based on this diagnosis</p>
            </div>
          </div>

          <div className="space-y-4">
            {data.nextActions.map((item:any, idx:number) => (
              <div key={idx} className="bg-white rounded-2xl border border-slate-100 overflow-hidden">
                <div className="flex items-stretch">
                  {/* Rank column */}
                  <div className="flex-shrink-0 w-16 flex flex-col items-center justify-center bg-slate-900 p-4 gap-1">
                    <span className="text-white/40 text-xs font-bold">#{item.rank}</span>
                    <div className="w-px h-6 bg-white/20"/>
                    <span className="text-white text-xs font-bold" style={{writingMode:"vertical-lr",transform:"rotate(180deg)",letterSpacing:"0.1em"}}>{item.priority}</span>
                  </div>

                  <div className="flex-1 p-6">
                    <div className="flex items-start justify-between gap-4 mb-4">
                      <div>
                        <span className="text-xs font-semibold px-2 py-0.5 rounded-full text-white mb-2 inline-block" style={{background:item.priorityColor}}>{item.priority} Priority</span>
                        <h3 className="text-base font-bold text-slate-900 mt-1">{item.title}</h3>
                        <p className="text-slate-400 text-xs mt-0.5">Addresses: {item.dimension}</p>
                      </div>
                    </div>

                    <p className="text-slate-600 text-sm leading-relaxed mb-4 p-3 bg-slate-50 rounded-xl">{item.initiative}</p>

                    <div className="grid grid-cols-1 lg:grid-cols-3 gap-3">
                      <div className="p-3 rounded-xl bg-green-50 border border-green-100">
                        <div className="text-green-700 text-xs font-bold uppercase tracking-wider mb-1.5">Expected Benefit</div>
                        <p className="text-green-800 text-xs leading-relaxed">{item.benefit}</p>
                      </div>
                      <div className="p-3 rounded-xl bg-red-50 border border-red-100">
                        <div className="text-red-700 text-xs font-bold uppercase tracking-wider mb-1.5">Cost of Inaction</div>
                        <p className="text-red-800 text-xs leading-relaxed">{item.loss}</p>
                      </div>
                      <div className="p-3 rounded-xl bg-blue-50 border border-blue-100">
                        <div className="text-blue-700 text-xs font-bold uppercase tracking-wider mb-1.5">Why Now</div>
                        <p className="text-blue-800 text-xs leading-relaxed">{item.whyNow}</p>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            ))}
          </div>
        </div>

        {/* Sources footer */}
        <div className="mt-8 p-5 rounded-2xl bg-slate-100 border border-slate-200">
          <h3 className="text-xs font-semibold text-slate-500 uppercase tracking-wider mb-3">Data Sources</h3>
          <div className="flex flex-wrap gap-2">
            {(isGenerated
              ? ["Claude AI (claude-sonnet-4-6)","Public Company Website","LinkedIn","MCA21 Filings","News & Press Releases","Job Postings","Industry Reports"]
              : ["MCA21 / ROC Coimbatore","ICRA Rationale Reports","Tracxn / Tofler","Janatics.com","janaticspneumatics.com","IMARC Group","AIA India","DSIR"]
            ).map(s=>(
              <span key={s} className="px-3 py-1 rounded-full bg-white border border-slate-200 text-slate-600 text-xs font-medium">{s}</span>
            ))}
          </div>
          <p className="text-slate-400 text-xs mt-3">
            Report generated {new Date().toLocaleDateString("en-IN",{day:"numeric",month:"long",year:"numeric"})} · Based on publicly available information ·{" "}
            {isGenerated ? "AI-generated — verify independently before decisions" : "Not investment advice"}
          </p>
        </div>
      </div>
    </div>
  );
}
