"use client";
import { Suspense } from "react";
import ReportContent from "./ReportContent";

export default function ReportPage() {
  return <Suspense fallback={<div className="min-h-screen bg-slate-50 flex items-center justify-center"><div className="animate-spin rounded-full h-8 w-8 border-b-2 border-blue-600"/></div>}><ReportContent/></Suspense>;
}
