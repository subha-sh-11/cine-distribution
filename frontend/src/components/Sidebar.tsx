"use client";

import Link from "next/link";
import { usePathname } from "next/navigation";
import ThemeControls from "./ThemeControls";

const NAV = [
  {
    href: "/admin/dashboard",
    label: "Dashboard",
    icon: (
      <path d="M3 13h8V3H3v10Zm0 8h8v-6H3v6Zm10 0h8V11h-8v10Zm0-18v6h8V3h-8Z" />
    ),
  },
  {
    href: "/admin/theatres",
    label: "Theatres",
    icon: (
      <path d="M4 4h16v4H4V4Zm0 6h16v10H4V10Zm3 3v4m5-4v4m5-4v4" />
    ),
  },
  {
    href: "/admin/representatives",
    label: "Representatives",
    icon: (
      <path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2M9 11a4 4 0 1 0 0-8 4 4 0 0 0 0 8Zm13 10v-2a4 4 0 0 0-3-3.87M16 3.13A4 4 0 0 1 16 11" />
    ),
  },
  {
    href: "/admin/users",
    label: "Users",
    icon: (
      <path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2M9 11a4 4 0 1 0 0-8 4 4 0 0 0 0 8Zm10 4 2 2 4-4" />
    ),
  },
  {
    href: "/admin/history",
    label: "History",
    icon: (
      <>
        <path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8" />
        <path d="M3 3v5h5" />
        <path d="M12 7v5l4 2" />
      </>
    ),
  },
];

export default function Sidebar() {
  const pathname = usePathname();

  return (
    <aside
      className="sticky top-0 flex h-screen w-60 shrink-0 flex-col text-white/90"
      style={{ backgroundImage: "linear-gradient(180deg,#f97316 0%,#ea580c 55%,#c2410c 100%)" }}
    >
      <div className="flex h-16 items-center gap-3 px-5">
        <span
          className="grid h-9 w-9 place-items-center rounded-md bg-white text-sm font-bold tracking-tight text-[#ea580c] shadow-sm"
        >
          SVF
        </span>
        <div className="leading-tight">
          <p className="text-sm font-semibold text-white">Distribution</p>
          <p className="text-[11px] text-white/70">Nizam Territory</p>
        </div>
      </div>

      <nav className="flex-1 space-y-0.5 px-3 py-3">
        {NAV.map((item) => {
          const active = pathname.startsWith(item.href);
          return (
            <Link
              key={item.href}
              href={item.href}
              className={`flex items-center gap-3 rounded-md px-3 py-2 text-sm transition ${active
                  ? "bg-white/20 font-semibold text-white shadow-sm"
                  : "text-white/85 hover:bg-white/10 hover:text-white"
                }`}
            >
              <svg
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth={1.7}
                strokeLinecap="round"
                strokeLinejoin="round"
                className="h-[18px] w-[18px]"
              >
                {item.icon}
              </svg>
              {item.label}
            </Link>
          );
        })}
      </nav>

      <div className="space-y-4 border-t border-white/20 p-4">
        <ThemeControls />
        <div className="flex items-center gap-3">
          <span className="grid h-8 w-8 place-items-center rounded-full bg-white/20 text-xs font-semibold text-white">
            AD
          </span>
          <div className="min-w-0 flex-1 leading-tight">
            <p className="text-xs font-medium text-white">Admin</p>
            <p className="truncate text-[11px] text-white/70">admin@svf.in</p>
          </div>
          <button
            onClick={async () => {
              await fetch("/api/auth/logout", { method: "POST" });
              window.location.href = "/login";
            }}
            title="Sign out"
            className="rounded-md p-1.5 text-white/80 hover:bg-white/20 hover:text-white"
          >
            <svg viewBox="0 0 24 24" className="h-4 w-4" fill="none" stroke="currentColor" strokeWidth={1.8} strokeLinecap="round" strokeLinejoin="round">
              <path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4M16 17l5-5-5-5M21 12H9" />
            </svg>
          </button>
        </div>
      </div>
    </aside>
  );
}
