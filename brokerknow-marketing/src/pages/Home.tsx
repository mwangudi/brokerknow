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
      <section className="mx-auto max-w-[1320px] px-4 py-16 lg:px-6 lg:py-24">
        <div className="grid items-center gap-12 lg:grid-cols-2">
          <div>
            <h2 className="section-title inline-block text-3xl text-dark">
              Trusted partner to institutional and high-net-worth investors
            </h2>
            <p className="mt-4 text-text">
              Situated in the prestigious Livingstone Towers in Blantyre, Malawi&rsquo;s
              financial capital, Cedar Capital is headed by veteran stockbroker
              Armstrong Kamphoni. We specialise in corporate actions, stockbroking and
              the management of portfolios for foreign investors and high-net-worth
              Malawians.
            </p>
            <p className="mt-4 text-text">
              Our processes, procedures and the combined experience of our energetic,
              client-centric dealing staff are integral to delivering world-class
              stockbroking services in a price-sensitive and confidential manner.
            </p>
            <Link
              to="/about"
              className="mt-6 inline-flex items-center font-secondary text-sm font-bold text-primary hover:opacity-80"
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
      <section className="bg-theme-light py-16 lg:py-24">
        <div className="mx-auto max-w-[1320px] px-4 lg:px-6">
          <div className="mb-12 text-center">
            <h2 className="section-title inline-block text-3xl text-dark">What we do</h2>
            <p className="mt-3 text-text">
              Full-service capital-markets coverage on the Malawi Stock Exchange.
            </p>
          </div>
          <div className="grid gap-6 sm:grid-cols-2 lg:grid-cols-4">
            {SERVICES.map((s) => (
              <div
                key={s.title}
                className="rounded-lg border border-[#e1e1e1] bg-white p-6 transition hover:shadow-lg"
              >
                <h3 className="text-lg font-bold text-dark">{s.title}</h3>
                <p className="mt-3 text-sm text-text">{s.body}</p>
              </div>
            ))}
          </div>
        </div>
      </section>

      {/* ---------- CTA ---------- */}
      <section className="bg-primary py-16 text-white">
        <div className="mx-auto flex max-w-[1320px] flex-col items-center justify-between gap-6 px-4 text-center sm:flex-row sm:text-left lg:px-6">
          <div>
            <h3 className="text-2xl font-bold text-white">Ready to invest on the MSE?</h3>
            <p className="mt-1 text-white/80">
              Open an account in minutes through our secure online portal.
            </p>
          </div>
          <a href={SITE.portalUrl} className="btn bg-white text-primary hover:opacity-90">
            Get started &rarr;
          </a>
        </div>
      </section>
    </Page>
  );
}
