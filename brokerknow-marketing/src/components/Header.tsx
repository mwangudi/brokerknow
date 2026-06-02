import { useState } from "react";
import { Link, NavLink } from "react-router";
import { NAV, SITE } from "../data/site";

export default function Header() {
  const [open, setOpen] = useState(false);

  return (
    <header className="sticky top-0 z-40 border-b border-slate-200 bg-white/90 backdrop-blur">
      <div className="mx-auto flex max-w-6xl items-center justify-between px-4 py-3 lg:px-6">
        <Link to="/" className="flex items-center gap-3" onClick={() => setOpen(false)}>
          {/* TODO: replace with high-res Cedar logo SVG once provided */}
          <img src="/images/logo.png" alt="Cedar Capital" className="h-10 w-auto" />
          <span className="hidden font-semibold text-brand-800 sm:block">{SITE.name}</span>
        </Link>

        <nav className="hidden items-center gap-1 lg:flex">
          {NAV.map((n) => (
            <NavLink
              key={n.to}
              to={n.to}
              end={n.to === "/"}
              className={({ isActive }) =>
                `rounded-md px-3 py-2 text-sm font-medium transition ${
                  isActive
                    ? "bg-brand-50 text-brand-800"
                    : "text-slate-700 hover:bg-slate-100 hover:text-brand-700"
                }`
              }
            >
              {n.label}
            </NavLink>
          ))}
          <a
            href={SITE.portalUrl}
            className="ml-2 rounded-md bg-brand-600 px-4 py-2 text-sm font-semibold text-white shadow-sm transition hover:bg-brand-700"
          >
            Client Portal
          </a>
        </nav>

        <button
          aria-label="Menu"
          className="rounded-md p-2 text-slate-700 hover:bg-slate-100 lg:hidden"
          onClick={() => setOpen((o) => !o)}
        >
          <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            {open ? <path d="M6 6l12 12M6 18L18 6" /> : <path d="M4 6h16M4 12h16M4 18h16" />}
          </svg>
        </button>
      </div>

      {open && (
        <nav className="border-t border-slate-200 bg-white lg:hidden">
          <div className="mx-auto flex max-w-6xl flex-col px-4 py-3">
            {NAV.map((n) => (
              <NavLink
                key={n.to}
                to={n.to}
                end={n.to === "/"}
                onClick={() => setOpen(false)}
                className={({ isActive }) =>
                  `rounded-md px-3 py-2 text-base font-medium ${
                    isActive ? "bg-brand-50 text-brand-800" : "text-slate-700 hover:bg-slate-100"
                  }`
                }
              >
                {n.label}
              </NavLink>
            ))}
            <a
              href={SITE.portalUrl}
              className="mt-2 rounded-md bg-brand-600 px-3 py-2 text-center text-base font-semibold text-white"
            >
              Client Portal
            </a>
          </div>
        </nav>
      )}
    </header>
  );
}
