import { useMemo, useState } from "react";
import Page from "../components/Page";
import { WEEKLY } from "../data/weekly";

const MONTH_NAMES = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December",
];

export default function WeeklyReports() {
  const [year, setYear] = useState<string>("all");
  const [search, setSearch] = useState("");

  const grouped = useMemo(() => {
    const filtered = WEEKLY
      .filter((r) => year === "all" || r.date.startsWith(year))
      .filter((r) => !search || r.label.toLowerCase().includes(search.toLowerCase()) || (r.note ?? "").toLowerCase().includes(search.toLowerCase()))
      .sort((a, b) => b.date.localeCompare(a.date));

    const byYear: Record<string, Record<string, typeof filtered>> = {};
    for (const r of filtered) {
      const [y, m] = r.date.split("-");
      (byYear[y] ??= {})[m] ??= [];
      byYear[y][m].push(r);
    }
    return byYear;
  }, [year, search]);

  const years = useMemo(
    () => Array.from(new Set(WEEKLY.map((r) => r.date.slice(0, 4)))).sort((a, b) => b.localeCompare(a)),
    []
  );

  return (
    <Page
      title="Weekly Market Reports"
      description="Cedar Capital's weekly commentary on the Malawi Stock Exchange — market statistics, listed-company updates and economic overviews."
    >
      <section className="bg-brand-50 py-12 lg:py-16">
        <div className="mx-auto max-w-6xl px-4 lg:px-6">
          <h1 className="text-4xl font-bold text-brand-800 lg:text-5xl">Weekly Reports</h1>
          <p className="mt-3 max-w-3xl text-brand-700">
            A continuous archive of our weekly market commentary, going back to 2018.
          </p>
        </div>
      </section>

      <section className="mx-auto max-w-6xl px-4 py-10 lg:px-6">
        <div className="mb-8 flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
          <div className="flex flex-wrap items-center gap-2">
            <button
              onClick={() => setYear("all")}
              className={`rounded-full px-3 py-1 text-sm font-medium transition ${
                year === "all"
                  ? "bg-brand-600 text-white"
                  : "bg-slate-100 text-slate-700 hover:bg-slate-200"
              }`}
            >
              All
            </button>
            {years.map((y) => (
              <button
                key={y}
                onClick={() => setYear(y)}
                className={`rounded-full px-3 py-1 text-sm font-medium transition ${
                  year === y
                    ? "bg-brand-600 text-white"
                    : "bg-slate-100 text-slate-700 hover:bg-slate-200"
                }`}
              >
                {y}
              </button>
            ))}
          </div>
          <input
            type="search"
            placeholder="Search reports…"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            className="w-full rounded-md border border-slate-300 px-3 py-2 text-sm focus:border-brand-500 focus:outline-none focus:ring-2 focus:ring-brand-200 sm:w-64"
          />
        </div>

        {Object.keys(grouped).length === 0 ? (
          <p className="rounded-md border border-dashed border-slate-300 p-8 text-center text-slate-500">
            No reports match your search.
          </p>
        ) : (
          Object.entries(grouped)
            .sort(([a], [b]) => b.localeCompare(a))
            .map(([y, months]) => (
              <div key={y} className="mb-10">
                <h2 className="mb-4 text-2xl font-bold text-brand-800">{y}</h2>
                <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-3">
                  {Object.entries(months)
                    .sort(([a], [b]) => b.localeCompare(a))
                    .map(([m, reports]) => (
                      <div
                        key={m}
                        className="rounded-lg border border-slate-200 bg-white p-5 shadow-sm"
                      >
                        <h3 className="mb-3 text-sm font-semibold uppercase tracking-wide text-slate-500">
                          {MONTH_NAMES[parseInt(m, 10) - 1]}
                        </h3>
                        <ul className="space-y-2">
                          {reports.map((r) => (
                            <li key={r.url}>
                              <a
                                href={r.url}
                                target="_blank"
                                rel="noopener noreferrer"
                                className="flex items-start gap-2 text-sm text-brand-700 hover:text-brand-900 hover:underline"
                              >
                                <svg
                                  className="mt-0.5 flex-shrink-0 text-brand-500"
                                  width="16" height="16" viewBox="0 0 24 24"
                                  fill="none" stroke="currentColor" strokeWidth="2"
                                >
                                  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
                                  <polyline points="14 2 14 8 20 8" />
                                </svg>
                                <span>
                                  <span className="font-medium">{r.label}</span>
                                  {r.note && (
                                    <span className="ml-1 text-xs text-slate-500">— {r.note}</span>
                                  )}
                                </span>
                              </a>
                            </li>
                          ))}
                        </ul>
                      </div>
                    ))}
                </div>
              </div>
            ))
        )}
      </section>
    </Page>
  );
}
