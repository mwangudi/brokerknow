import Page from "../components/Page";
import { RESEARCH } from "../data/research";

export default function Research() {
  return (
    <Page
      title="Research"
      description="Quarterly and thematic research on the Malawi Stock Exchange, listed companies and the broader macro environment."
    >
      <section className="bg-brand-50 py-12 lg:py-16">
        <div className="mx-auto max-w-6xl px-4 lg:px-6">
          <h1 className="text-4xl font-bold text-brand-800 lg:text-5xl">Research</h1>
          <p className="mt-3 max-w-3xl text-brand-700">
            In-depth coverage of listed counters on the Malawi Stock Exchange and the
            macro environment that shapes them.
          </p>
        </div>
      </section>

      <section className="mx-auto max-w-4xl px-4 py-12 lg:px-6">
        <ul className="divide-y divide-slate-200 rounded-lg border border-slate-200 bg-white shadow-sm">
          {RESEARCH.map((r) => (
            <li key={r.url}>
              <a
                href={r.url}
                target="_blank"
                rel="noopener noreferrer"
                className="flex items-center justify-between gap-4 px-5 py-4 hover:bg-brand-50"
              >
                <div className="flex items-start gap-3">
                  <svg
                    className="mt-0.5 flex-shrink-0 text-brand-500"
                    width="20" height="20" viewBox="0 0 24 24"
                    fill="none" stroke="currentColor" strokeWidth="2"
                  >
                    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
                    <polyline points="14 2 14 8 20 8" />
                  </svg>
                  <div>
                    <p className="font-medium text-slate-900">{r.title}</p>
                    {r.date && (
                      <p className="text-xs text-slate-500">{r.date}</p>
                    )}
                  </div>
                </div>
                <span className="text-sm font-semibold text-brand-700">PDF &rarr;</span>
              </a>
            </li>
          ))}
        </ul>
      </section>
    </Page>
  );
}
