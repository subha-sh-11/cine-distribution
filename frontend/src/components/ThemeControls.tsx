"use client";

import { useEffect, useState } from "react";

type Theme = "light" | "dark";

export default function ThemeControls() {
  const [theme, setTheme] = useState<Theme>("light");

  useEffect(() => {
    const t = (localStorage.getItem("svf.theme") as Theme) || "light";
    setTheme(t);
  }, []);

  function applyTheme(t: Theme) {
    setTheme(t);
    localStorage.setItem("svf.theme", t);
    document.documentElement.setAttribute("data-theme", t);
  }

  return (
    <div className="flex rounded-lg bg-white/15 p-0.5">
      {(["light", "dark"] as Theme[]).map((t) => (
        <button
          key={t}
          onClick={() => applyTheme(t)}
          className={`flex flex-1 items-center justify-center gap-1.5 rounded-md py-1.5 text-xs font-medium capitalize transition ${
            theme === t
              ? "bg-white/25 text-white shadow-sm"
              : "text-white/70 hover:text-white"
          }`}
        >
          {t === "light" ? (
            <svg viewBox="0 0 24 24" className="h-3.5 w-3.5" fill="none" stroke="currentColor" strokeWidth={2}>
              <circle cx="12" cy="12" r="4" />
              <path d="M12 2v2m0 16v2M4 12H2m20 0h-2M5 5l1.5 1.5M17.5 17.5 19 19M19 5l-1.5 1.5M6.5 17.5 5 19" />
            </svg>
          ) : (
            <svg viewBox="0 0 24 24" className="h-3.5 w-3.5" fill="none" stroke="currentColor" strokeWidth={2}>
              <path d="M21 12.8A9 9 0 1 1 11.2 3a7 7 0 0 0 9.8 9.8Z" />
            </svg>
          )}
          {t}
        </button>
      ))}
    </div>
  );
}
