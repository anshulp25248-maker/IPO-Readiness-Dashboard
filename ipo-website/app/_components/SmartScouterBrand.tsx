"use client";

import { useEffect, useState } from "react";

export function SmartScouterLogo({ compact = false }: { compact?: boolean }) {
  return (
    <div className="flex items-center gap-3">
      <div className="relative flex h-12 w-12 shrink-0 items-center justify-center overflow-hidden rounded-2xl bg-slate-950 shadow-[0_16px_34px_rgba(15,23,42,0.24)]">
        <svg viewBox="0 0 72 72" className="h-11 w-11" aria-hidden="true">
          <path d="M36 9c9 7 14 16 14 27 0 11-5 20-14 27-9-7-14-16-14-27 0-11 5-20 14-27Z" fill="#f8fafc" />
          <path d="M36 9c9 7 14 16 14 27H22c0-11 5-20 14-27Z" fill="#e2f4bb" />
          <circle cx="36" cy="32" r="9" fill="#0f172a" />
          <circle cx="36" cy="32" r="5" fill="#34d399" />
          <path d="M24 43 12 55l15-4" fill="#f59e0b" />
          <path d="M48 43 60 55l-15-4" fill="#f59e0b" />
          <path d="M30 59c2 4 4 6 6 8 2-2 4-4 6-8H30Z" fill="#f97316" />
          <path d="M20 56h32" stroke="#34d399" strokeWidth="3" strokeLinecap="round" />
        </svg>
      </div>
      <div className={compact ? "hidden sm:block" : ""}>
        <p className="font-serif text-2xl font-black leading-none tracking-normal text-slate-950">Smart Scouter</p>
        <p className="mt-1 text-[10px] font-black uppercase tracking-[0.22em] text-slate-800/65">Investment Console</p>
      </div>
    </div>
  );
}

function IntroRocketScene() {
  return (
    <div className="intro-scene" aria-hidden="true">
      <div className="intro-board">
        <div className="intro-board-header">
          <span />
          <span />
          <span />
        </div>
        <div className="intro-screen intro-screen-left">
          <span />
          <span />
          <span />
        </div>
        <div className="intro-screen intro-screen-right">
          <span />
          <span />
          <span />
        </div>
        <svg viewBox="0 0 360 360" className="intro-rocket" role="img">
          <path d="M180 38c46 36 70 84 70 144 0 54-26 100-70 140-44-40-70-86-70-140 0-60 24-108 70-144Z" fill="#f8fafc" />
          <path d="M180 38c46 36 70 84 70 144H110c0-60 24-108 70-144Z" fill="#dbeafe" />
          <path d="M117 216 56 292l75-24" fill="#f59e0b" />
          <path d="M243 216 304 292l-75-24" fill="#f59e0b" />
          <circle cx="180" cy="150" r="42" fill="#0f172a" />
          <circle cx="180" cy="150" r="32" fill="#93c5fd" />
          <circle cx="180" cy="142" r="10" fill="#2563eb" />
          <path d="M160 170c9-12 31-12 40 0" stroke="#2563eb" strokeWidth="8" strokeLinecap="round" />
          <path d="M150 248h60l-30 66-30-66Z" fill="#ef4444" />
          <path d="M163 292c4 13 10 23 17 31 7-8 13-18 17-31h-34Z" fill="#fb923c" />
          <path d="M140 103h80" stroke="#0f172a" strokeWidth="8" strokeLinecap="round" opacity="0.16" />
        </svg>
        <div className="intro-flame" />
        <div className="intro-graph intro-graph-one" />
        <div className="intro-graph intro-graph-two" />
        <div className="intro-company-dot intro-dot-one" />
        <div className="intro-company-dot intro-dot-two" />
        <div className="intro-company-dot intro-dot-three" />
      </div>
    </div>
  );
}

export function LaunchIntro() {
  const [visible, setVisible] = useState(false);

  useEffect(() => {
    const forceIntro = new URLSearchParams(window.location.search).has("intro");
    if (!forceIntro && window.sessionStorage.getItem("smart-scouter-intro-seen") === "true") return;
    setVisible(true);
    if (!forceIntro) window.sessionStorage.setItem("smart-scouter-intro-seen", "true");
    const timer = window.setTimeout(() => setVisible(false), 4700);
    return () => window.clearTimeout(timer);
  }, []);

  if (!visible) return null;

  return (
    <div className="intro-overlay" role="presentation">
      <div className="intro-logo">
        <SmartScouterLogo />
      </div>
      <IntroRocketScene />
      <div className="intro-copy">
        <p className="font-serif text-4xl font-black tracking-normal text-slate-950 sm:text-6xl">Smart Scouter</p>
        <p className="mt-3 text-sm font-black uppercase tracking-[0.28em] text-slate-800/75">
          Investor launch sequence
        </p>
      </div>
    </div>
  );
}
