import { useEffect, useState } from "react";
import { Link } from "react-router";
import { SITE } from "../data/site";

type Slide = {
  eyebrow: string;
  title: string;
  body: string;
  primary?: { label: string; href: string; external?: boolean };
  secondary?: { label: string; href: string; external?: boolean };
  gradient: string;
};

const SLIDES: Slide[] = [
  {
    eyebrow: "Welcome to",
    title: "Cedar Capital",
    body: "Malawi's leading private financial services firm — stockbroking, research and corporate advisory, headquartered in Livingstone Towers, Blantyre.",
    primary: { label: "Open the Client Portal", href: SITE.portalUrl, external: true },
    secondary: { label: "Open an Account", href: "/forms" },
    gradient:
      "radial-gradient(circle at 20% 20%, rgba(198,157,78,0.35), transparent 45%), radial-gradient(circle at 80% 60%, rgba(255,255,255,0.15), transparent 50%)",
  },
  {
    eyebrow: "Market intelligence",
    title: "Weekly market reports",
    body: "Eight years of unbroken weekly commentary on the Malawi Stock Exchange — equities, fixed income and macro insight delivered every Friday.",
    primary: { label: "Browse weekly reports", href: "/weekly-reports" },
    secondary: { label: "Read research notes", href: "/research" },
    gradient:
      "radial-gradient(circle at 70% 30%, rgba(198,157,78,0.45), transparent 50%), radial-gradient(circle at 20% 75%, rgba(255,255,255,0.18), transparent 55%)",
  },
  {
    eyebrow: "Corporate advisory",
    title: "Capital raising & M&A",
    body: "Trusted by issuers and institutional investors for IPOs, rights issues, bond listings and corporate actions across the Malawi capital markets.",
    primary: { label: "About Cedar Capital", href: "/about" },
    secondary: { label: "Talk to us", href: "/contact" },
    gradient:
      "radial-gradient(circle at 50% 30%, rgba(198,157,78,0.40), transparent 55%), radial-gradient(circle at 85% 80%, rgba(255,255,255,0.16), transparent 50%)",
  },
];

const INTERVAL_MS = 6000;

export default function HeroCarousel() {
  const [index, setIndex] = useState(0);
  const [paused, setPaused] = useState(false);

  useEffect(() => {
    if (paused) return;
    const t = setInterval(() => setIndex((i) => (i + 1) % SLIDES.length), INTERVAL_MS);
    return () => clearInterval(t);
  }, [paused]);

  const go = (i: number) => setIndex((i + SLIDES.length) % SLIDES.length);

  return (
    <section
      className="relative isolate overflow-hidden border-b border-[#e1e1e1] bg-white text-text"
      onMouseEnter={() => setPaused(true)}
      onMouseLeave={() => setPaused(false)}
      aria-roledescription="carousel"
      aria-label="Cedar Capital highlights"
    >
      <div className="relative">
        {SLIDES.map((s, i) => (
          <div
            key={s.title}
            className={`transition-opacity duration-700 ${
              i === index ? "opacity-100" : "pointer-events-none absolute inset-0 opacity-0"
            }`}
            aria-hidden={i !== index}
          >
            <div
              className="absolute inset-0 opacity-60"
              style={{ backgroundImage: s.gradient }}
            />
            <div className="relative mx-auto max-w-[1320px] px-4 py-24 lg:px-6 lg:py-32">
              <p className="mb-4 text-sm font-bold uppercase tracking-widest text-primary">
                {s.eyebrow}
              </p>
              <h1 className="banner-title max-w-4xl text-4xl leading-[1.05] sm:text-5xl lg:text-[64px]">
                <strong>{s.title}</strong>
              </h1>
              <p className="mt-6 max-w-2xl text-lg text-text lg:text-xl">{s.body}</p>
              <div className="mt-8 flex flex-wrap gap-4">
                {s.primary &&
                  (s.primary.external ? (
                    <a href={s.primary.href} className="btn btn-primary">
                      {s.primary.label}
                    </a>
                  ) : (
                    <Link to={s.primary.href} className="btn btn-primary">
                      {s.primary.label}
                    </Link>
                  ))}
                {s.secondary &&
                  (s.secondary.external ? (
                    <a href={s.secondary.href} className="btn btn-outline-primary">
                      {s.secondary.label}
                    </a>
                  ) : (
                    <Link to={s.secondary.href} className="btn btn-outline-primary">
                      {s.secondary.label}
                    </Link>
                  ))}
              </div>
            </div>
          </div>
        ))}
      </div>

      <button
        type="button"
        onClick={() => go(index - 1)}
        aria-label="Previous slide"
        className="absolute left-3 top-1/2 z-10 hidden -translate-y-1/2 rounded-full bg-primary p-3 text-white transition hover:opacity-85 sm:block"
      >
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><path d="M15 18l-6-6 6-6"/></svg>
      </button>
      <button
        type="button"
        onClick={() => go(index + 1)}
        aria-label="Next slide"
        className="absolute right-3 top-1/2 z-10 hidden -translate-y-1/2 rounded-full bg-primary p-3 text-white transition hover:opacity-85 sm:block"
      >
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><path d="M9 6l6 6-6 6"/></svg>
      </button>

      <div className="absolute bottom-6 left-0 right-0 z-10 flex justify-center gap-2">
        {SLIDES.map((s, i) => (
          <button
            key={s.title}
            type="button"
            onClick={() => go(i)}
            aria-label={`Go to slide ${i + 1}`}
            aria-current={i === index}
            className={`h-2 rounded-full transition-all ${
              i === index ? "w-8 bg-primary" : "w-2 bg-black/15 hover:bg-black/30"
            }`}
          />
        ))}
      </div>
    </section>
  );
}
