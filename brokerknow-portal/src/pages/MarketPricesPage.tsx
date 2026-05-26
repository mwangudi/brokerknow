import { useEffect, useMemo, useState } from "react";
import api from "../lib/api";

export interface MarketQuote {
  securityDpa: number;
  securityCode: string;
  securityName: string;
  exchange: string;
  quoteDate: string;
  open: number | null;
  high: number | null;
  low: number | null;
  close: number | null;
  previousClose: number | null;
  bid: number | null;
  offer: number | null;
  volume: number | null;
  changeAbs: number | null;
  changePct: number | null;
}

function fmt(n: number | null, dp = 2) {
  if (n === null || n === undefined) return "—";
  return Number(n).toLocaleString("en", { minimumFractionDigits: dp, maximumFractionDigits: dp });
}

function fmtVol(n: number | null) {
  if (n === null || n === undefined) return "—";
  return Number(n).toLocaleString();
}

function fmtDate(iso: string) {
  return new Date(iso).toLocaleDateString("en-GB", {
    weekday: "long", day: "2-digit", month: "long", year: "numeric",
  });
}

export default function MarketPricesPage() {
  const [quotes, setQuotes] = useState<MarketQuote[] | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [search, setSearch] = useState("");
  const [filter, setFilter] = useState<"all" | "gainers" | "losers" | "unchanged">("all");

  useEffect(() => {
    setLoading(true);
    api
      .get<{ quoteDate: string | null; items: MarketQuote[] }>("/market-quotes/latest")
      .then((r) => setQuotes(r.data.items ?? []))
      .catch(() => setError("Could not load market prices."))
      .finally(() => setLoading(false));
  }, []);

  const asAt = quotes && quotes.length > 0 ? quotes[0].quoteDate : null;

  const filtered = useMemo(() => {
    if (!quotes) return [];
    const term = search.trim().toLowerCase();
    return quotes.filter((q) => {
      if (term && !q.securityCode.toLowerCase().includes(term) && !q.securityName.toLowerCase().includes(term)) {
        return false;
      }
      const ch = q.changeAbs ?? 0;
      if (filter === "gainers" && ch <= 0) return false;
      if (filter === "losers" && ch >= 0) return false;
      if (filter === "unchanged" && ch !== 0) return false;
      return true;
    });
  }, [quotes, search, filter]);

  return (
    <div>
      <div className="mb-6 overflow-hidden rounded-2xl bg-gradient-to-r from-emerald-600 to-teal-700 p-6 text-white shadow-lg">
        <p className="text-sm text-emerald-100">Today's market</p>
        <h2 className="mt-1 text-2xl font-bold">Market Prices</h2>
        <p className="mt-2 text-sm text-emerald-100">
          {asAt ? `Quotes as at ${fmtDate(asAt)}` : "Latest closing prices from the exchange"}
        </p>
      </div>

      <div className="mb-4 rounded-xl border border-amber-200 bg-amber-50 p-3 text-xs text-amber-900">
        <strong>Disclaimer:</strong> Prices shown are end-of-day published quotes
        {asAt ? <> as at <strong>{fmtDate(asAt)}</strong></> : null}.
        They are <strong>indicative only</strong> and may differ from the live order-book.
        Please confirm with your broker before placing trades.
      </div>

      {loading ? (
        <p className="text-gray-500">Loading prices...</p>
      ) : error ? (
        <p className="text-red-600">{error}</p>
      ) : !quotes || quotes.length === 0 ? (
        <div className="rounded-2xl border border-dashed border-gray-300 bg-white p-10 text-center text-sm text-gray-500">
          No market data has been published yet.
        </div>
      ) : (
        <>
          {/* Search + filter toolbar */}
          <div className="mb-3 flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
            <div className="relative max-w-xs flex-1">
              <svg className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-gray-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" />
              </svg>
              <input
                type="search"
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                placeholder="Search by code or name…"
                className="w-full rounded-lg border border-gray-300 bg-white py-2 pl-9 pr-3 text-sm placeholder-gray-400 focus:border-emerald-500 focus:outline-none focus:ring-1 focus:ring-emerald-500"
              />
            </div>
            <div className="flex flex-wrap items-center gap-1 rounded-lg bg-gray-100 p-1 text-xs font-medium">
              {([
                { key: "all", label: `All (${quotes.length})` },
                { key: "gainers", label: "Gainers" },
                { key: "losers", label: "Losers" },
                { key: "unchanged", label: "Unchanged" },
              ] as const).map((opt) => (
                <button
                  key={opt.key}
                  type="button"
                  onClick={() => setFilter(opt.key)}
                  className={`rounded-md px-3 py-1.5 transition ${
                    filter === opt.key
                      ? "bg-white text-emerald-700 shadow-sm"
                      : "text-gray-600 hover:text-gray-900"
                  }`}
                >
                  {opt.label}
                </button>
              ))}
            </div>
          </div>

          <div className="overflow-hidden rounded-2xl border border-gray-200 bg-white shadow-sm">
            <table className="min-w-full divide-y divide-gray-200 text-sm">
              <thead className="bg-gray-50 text-left text-xs uppercase tracking-wider text-gray-500">
                <tr>
                  <th className="px-4 py-3">Code</th>
                  <th className="px-4 py-3">Security</th>
                  <th className="px-4 py-3 text-right">Open</th>
                  <th className="px-4 py-3 text-right">High</th>
                  <th className="px-4 py-3 text-right">Low</th>
                  <th className="px-4 py-3 text-right">Close</th>
                  <th className="px-4 py-3 text-right">Change</th>
                  <th className="px-4 py-3 text-right">Volume</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-100">
                {filtered.length === 0 ? (
                  <tr>
                    <td colSpan={8} className="px-4 py-10 text-center text-sm text-gray-400">
                      No securities match your search.
                    </td>
                  </tr>
                ) : filtered.map((q) => {
                  const up = (q.changeAbs ?? 0) > 0;
                  const down = (q.changeAbs ?? 0) < 0;
                  const colour = up ? "text-emerald-600" : down ? "text-red-600" : "text-gray-500";
                  return (
                    <tr key={q.securityDpa} className="hover:bg-gray-50">
                      <td className="px-4 py-2.5 font-semibold text-gray-900">{q.securityCode}</td>
                      <td className="px-4 py-2.5 text-gray-700">{q.securityName}</td>
                      <td className="px-4 py-2.5 text-right tabular-nums">{fmt(q.open)}</td>
                      <td className="px-4 py-2.5 text-right tabular-nums">{fmt(q.high)}</td>
                      <td className="px-4 py-2.5 text-right tabular-nums">{fmt(q.low)}</td>
                      <td className="px-4 py-2.5 text-right font-semibold tabular-nums text-gray-900">{fmt(q.close)}</td>
                      <td className={`px-4 py-2.5 text-right font-medium tabular-nums ${colour}`}>
                        {q.changePct == null ? "—" : (
                          <>{up ? "▲" : down ? "▼" : "•"} {Math.abs(q.changePct).toFixed(2)}%</>
                        )}
                      </td>
                      <td className="px-4 py-2.5 text-right tabular-nums text-gray-600">{fmtVol(q.volume)}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>

          <p className="mt-3 text-xs text-gray-500">
            Showing <strong>{filtered.length}</strong> of {quotes.length} securities.
          </p>
        </>
      )}
    </div>
  );
}
