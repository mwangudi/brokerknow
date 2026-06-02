import Page from "../components/Page";
import { FORMS } from "../data/forms";
import { SITE } from "../data/site";

export default function Forms() {
  return (
    <Page
      title="Account Opening Forms"
      description="Download Cedar Capital's account opening, KYC, order and authorisation forms."
    >
      <section className="bg-brand-50 py-12 lg:py-16">
        <div className="mx-auto max-w-6xl px-4 lg:px-6">
          <h1 className="text-4xl font-bold text-brand-800 lg:text-5xl">Forms</h1>
          <p className="mt-3 max-w-3xl text-brand-700">
            All the documents you need to open an account, place orders and manage
            your relationship with Cedar Capital.
          </p>
          <p className="mt-4 text-sm text-brand-700">
            Prefer to do it online? &nbsp;
            <a className="font-semibold text-brand-800 underline" href={SITE.portalUrl}>
              Open an account in the Client Portal
            </a>
            .
          </p>
        </div>
      </section>

      <section className="mx-auto max-w-4xl px-4 py-12 lg:px-6">
        <ul className="grid gap-3 sm:grid-cols-2">
          {FORMS.map((f) => (
            <li key={f.url}>
              <a
                href={f.url}
                target="_blank"
                rel="noopener noreferrer"
                className="flex h-full items-start gap-3 rounded-lg border border-slate-200 bg-white p-5 shadow-sm transition hover:border-brand-400 hover:shadow-md"
              >
                <svg
                  className="mt-0.5 flex-shrink-0 text-brand-500"
                  width="24" height="24" viewBox="0 0 24 24"
                  fill="none" stroke="currentColor" strokeWidth="2"
                >
                  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
                  <polyline points="14 2 14 8 20 8" />
                </svg>
                <div>
                  <p className="font-semibold text-slate-900">{f.title}</p>
                  <p className="mt-1 text-xs text-brand-700">Download PDF &rarr;</p>
                </div>
              </a>
            </li>
          ))}
        </ul>
      </section>
    </Page>
  );
}
