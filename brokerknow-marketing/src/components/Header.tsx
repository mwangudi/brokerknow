import { useState } from "react";
import { Link, NavLink } from "react-router";
import { NAV, SITE } from "../data/site";

export default function Header() {
  const [open, setOpen] = useState(false);

  return (
    <header className="sticky top-0 z-40 border-b border-[#e1e1e1] bg-white">
      <div className="mx-auto flex max-w-[1320px] items-center justify-between px-4 py-6 lg:px-6">
        <Link to="/" className="flex items-center gap-3" onClick={() => setOpen(false)}>
          {/* TODO: replace with high-res Cedar logo SVG once provided */}
          <img src="/images/logo.png" alt="Cedar Capital" className="h-14 w-auto" />
          <span className="hidden font-secondary text-lg font-bold text-dark sm:block">{SITE.name}</span>
        </Link>

        <nav className="hidden items-center gap-1 lg:flex">
          {NAV.map((n) => (
            <NavLink
              key={n.to}
              to={n.to}
              end={n.to === "/"}
              className={({ isActive }) => `nav-link ${isActive ? "active" : ""}`}
            >
              {n.label}
            </NavLink>
          ))}
          <a href={SITE.portalUrl} className="btn btn-primary ml-3">
            Client Portal
          </a>
        </nav>

        <button
          aria-label="Menu"
          className="rounded-md p-2 text-dark hover:bg-theme-light lg:hidden"
          onClick={() => setOpen((o) => !o)}
        >
          <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            {open ? <path d="M6 6l12 12M6 18L18 6" /> : <path d="M4 6h16M4 12h16M4 18h16" />}
          </svg>
        </button>
      </div>

      {open && (
        <nav className="border-t border-[#e1e1e1] bg-white lg:hidden">
          <div className="mx-auto flex max-w-[1320px] flex-col gap-1 px-4 py-3">
            {NAV.map((n) => (
              <NavLink
                key={n.to}
                to={n.to}
                end={n.to === "/"}
                onClick={() => setOpen(false)}
                className={({ isActive }) => `nav-link ${isActive ? "active" : ""}`}
              >
                {n.label}
              </NavLink>
            ))}
            <a href={SITE.portalUrl} className="btn btn-primary mt-2 w-full">
              Client Portal
            </a>
          </div>
        </nav>
      )}
    </header>
  );
}
