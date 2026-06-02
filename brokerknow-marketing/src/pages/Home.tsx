import { Link } from "react-router";
import Page from "../components/Page";
import HeroCarousel from "../components/HeroCarousel";
import { SERVICES, SITE } from "../data/site";

export default function Home() {
  return (
    <Page
      title="Stockbroking, Research & Corporate Advisory in Malawi"
      description="Cedar Capital — the leading private financial services firm in Malawi, headed by Armstrong Kamphoni and specialising in stockbroking, research and corporate advisory."
    >
      {/* ---------- HERO ---------- */}
      <HeroCarousel />

      {/* ---------- ABOUT TEASER ---------- */}
      <section className="mx-auto max-w-6xl px-4 py-16 lg:px-6 lg:py-24">
        <div className="grid items-center gap-12 lg:grid-cols-2">
          <div>
            <h2 className="text-3xl font-bold text-brand-800">
              Trusted partner to institutional and high-net-worth investors
            </h2>
            <p className="mt-4 text-slate-700">
              Situated in the prestigious Livingstone Towers in Blantyre, Malawi&rsquo;s
              financial capital, Cedar Capital is headed by veteran stockbroker
              Armstrong Kamphoni. We specialise in corporate actions, stockbroking and
              the management of portfolios for foreign investors and high-net-worth
              Malawians.
            </p>
            <p className="mt-4 text-slate-700">
              Our processes, procedures and the combined experience of our energetic,
              client-centric dealing staff are integral to delivering world-class
              stockbroking services in a price-sensitive and confidential manner.
            </p>
            <Link
              to="/about"
              className="mt-6 inline-flex items-center text-sm font-semibold text-brand-700 hover:text-brand-900"
            >
              More about us &rarr;
            </Link>
          </div>
          <div className="relative flex aspect-[4/3] items-end overflow-hidden rounded-lg bg-gradient-to-br from-brand-700 via-brand-800 to-brand-900 p-8 text-white shadow-lg">
            <div
              className="absolute inset-0 opacity-25"
              style={{
                backgroundImage:
                  "radial-gradient(circle at 30% 30%, rgba(198,157,78,0.6), transparent 50%), radial-gradient(circle at 75% 70%, rgba(255,255,255,0.2), transparent 55%)",
              }}
            />
            <div className="relative">
              <p className="text-sm font-semibold uppercase tracking-widest text-gold-500">
                Livingstone Towers
              </p>
              <p className="mt-2 text-2xl font-semibold leading-snug">
                Glyn Jones Road, Blantyre &mdash; the heart of Malawi&rsquo;s
                financial district.
              </p>
            </div>
          </div>
        </div>
      </section>

      {/* ---------- SERVICES ---------- */}
      <section className="bg-slate-50 py-16 lg:py-24">
        <div className="mx-auto max-w-6xl px-4 lg:px-6">
          <div className="mb-12 text-center">
            <h2 className="text-3xl font-bold text-brand-800">What we do</h2>
            <p className="mt-3 text-slate-600">
              Full-service capital-markets coverage on the Malawi Stock Exchange.
            </p>
          </div>
          <div className="grid gap-6 sm:grid-cols-2 lg:grid-cols-4">
            {SERVICES.map((s) => (
              <div
                key={s.title}
                className="rounded-lg border border-slate-200 bg-white p-6 shadow-sm"
              >
                <h3 className="text-lg font-semibold text-brand-800">{s.title}</h3>
                <p className="mt-3 text-sm text-slate-600">{s.body}</p>
              </div>
            ))}
          </div>
        </div>
      </section>

      {/* ---------- CTA ---------- */}
      <section className="bg-brand-700 py-16 text-white">
        <div className="mx-auto flex max-w-6xl flex-col items-center justify-between gap-6 px-4 text-center sm:flex-row sm:text-left lg:px-6">
          <div>
            <h3 className="text-2xl font-bold">Ready to invest on the MSE?</h3>
            <p className="mt-1 text-brand-100">
              Open an account in minutes through our secure online portal.
            </p>
          </div>
          <a
            href={SITE.portalUrl}
            className="rounded-md bg-white px-6 py-3 text-sm font-semibold text-brand-800 shadow-sm transition hover:bg-brand-50"
          >
            Get started &rarr;
          </a>
        </div>
      </section>
    </Page>
  );
}
