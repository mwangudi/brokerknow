import { useEffect, useState } from "react";
import api from "../lib/api";
import DatePicker from "../components/form/DatePicker";
import Icon from "../components/ui/Icon";

interface StatementRow {
  source: string;
  id: number;
  date: string;
  reference?: string;
  particulars: string;
  debit: number;
  credit: number;
  balance: number;
  pending: boolean;
  kind: string;
}

interface StatementData {
  total: number;
  rows: StatementRow[];
  summary: {
    totalDebit: number;
    totalCredit: number;
    closingBalance: number;
    pendingCount: number;
  };
}

function fmt(n: number) {
  return Math.abs(n).toLocaleString("en", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

export default function StatementPage() {
  const [from, setFrom] = useState("");
  const [to, setTo] = useState("");
  const [appliedFrom, setAppliedFrom] = useState("");
  const [appliedTo, setAppliedTo] = useState("");
  const [page, setPage] = useState(1);
  const [pageSize] = useState(15);
  const [data, setData] = useState<StatementData | null>(null);
  const [loading, setLoading] = useState(true);
  const [pdfError, setPdfError] = useState<string | null>(null);

  function fetchStatement(f: string, t: string, p: number) {
    setLoading(true);
    const params = new URLSearchParams();
    if (f) params.set("from", f);
    if (t) params.set("to", t);
    params.set("page", p.toString());
    params.set("pageSize", pageSize.toString());
    api
      .get(`/portal/statement?${params}`)
      .then((r) => setData(r.data))
      .catch(() => {})
      .finally(() => setLoading(false));
  }

  useEffect(() => {
    fetchStatement(appliedFrom, appliedTo, page);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [appliedFrom, appliedTo, page]);

  function search() {
    setAppliedFrom(from);
    setAppliedTo(to);
    setPage(1);
  }

  function downloadPdf() {
    const params = new URLSearchParams();
    if (appliedFrom) params.set("from", appliedFrom);
    if (appliedTo) params.set("to", appliedTo);
    setPdfError(null);
    api
      .get(`/portal/statement.pdf?${params}`, { responseType: "blob" })
      .then((r) => {
        const url = window.URL.createObjectURL(
          new Blob([r.data], { type: "application/pdf" }),
        );
        const link = document.createElement("a");
        link.href = url;
        link.setAttribute(
          "download",
          `statement-${new Date().toISOString().slice(0, 10)}.pdf`,
        );
        document.body.appendChild(link);
        link.click();
        link.remove();
        window.URL.revokeObjectURL(url);
      })
      .catch(() =>
        setPdfError(
          "Failed to download statement PDF. Please try again in a moment.",
        ),
      );
  }

  const totalPages = data ? Math.max(1, Math.ceil(data.total / pageSize)) : 1;

  return (
    <div className="space-y-6">
      <div>
        <h1 className="font-display text-2xl font-semibold text-primary">
          Statement
        </h1>
        <p className="text-sm text-on-surface-variant">
          Your full transaction ledger and running balance.
        </p>
      </div>

      {/* Summary bento */}
      {data?.summary && (
        <div className="grid grid-cols-1 gap-4 md:grid-cols-4">
          <SummaryCard
            label="Total Debit"
            value={`MWK ${fmt(data.summary.totalDebit)}`}
            icon="trending_down"
            tone="error"
          />
          <SummaryCard
            label="Total Credit"
            value={`MWK ${fmt(data.summary.totalCredit)}`}
            icon="trending_up"
            tone="secondary"
          />
          <div className="overflow-hidden rounded-xl bg-primary-container p-5 shadow-lg">
            <div className="mb-3 flex items-start justify-between">
              <span className="text-[11px] font-semibold uppercase tracking-wide text-on-primary-container">
                Closing Balance
              </span>
              <Icon
                name="account_balance_wallet"
                size={18}
                className="text-secondary-fixed"
              />
            </div>
            <div className="font-display text-xl font-semibold text-white">
              MWK {fmt(data.summary.closingBalance)}{" "}
              <span className="text-sm font-medium text-on-primary-container">
                {data.summary.closingBalance >= 0 ? "Cr" : "Dr"}
              </span>
            </div>
          </div>
          <SummaryCard
            label="Pending Clearance"
            value={`${data.summary.pendingCount} ${data.summary.pendingCount === 1 ? "item" : "items"}`}
            icon="schedule"
            tone="muted"
          />
        </div>
      )}

      {/* Ledger */}
      <div className="overflow-hidden rounded-xl border border-outline-variant bg-surface-container-lowest shadow-[0px_4px_12px_rgba(15,23,42,0.03)]">
        {/* Filters */}
        <div className="flex flex-wrap items-center justify-between gap-3 border-b border-outline-variant px-5 py-3">
          <div className="flex flex-wrap items-end gap-2">
            <div className="w-40">
              <DatePicker value={from} onChange={setFrom} placeholder="From date" maxDate={null} />
            </div>
            <div className="w-40">
              <DatePicker value={to} onChange={setTo} placeholder="To date" maxDate={null} />
            </div>
            <button
              onClick={search}
              className="rounded-lg bg-primary px-4 py-2 text-xs font-semibold uppercase tracking-wide text-on-primary transition-opacity hover:opacity-90"
            >
              Search
            </button>
          </div>
          <button
            onClick={downloadPdf}
            className="inline-flex items-center gap-1.5 rounded-lg border border-outline-variant bg-surface-container-lowest px-4 py-2 text-xs font-semibold text-on-surface transition-colors hover:bg-surface-container-low"
          >
            <Icon name="picture_as_pdf" size={16} />
            Export PDF
          </button>
        </div>

        {pdfError && (
          <div className="border-b border-axis-error/20 bg-axis-error/5 px-5 py-2 text-sm text-axis-error">
            {pdfError}
          </div>
        )}

        {/* Table */}
        <div className="overflow-x-auto">
          <table className="w-full text-left text-sm">
            <thead>
              <tr className="border-b border-outline-variant bg-surface-container text-[11px] uppercase tracking-wider text-on-surface-variant">
                <th className="px-5 py-3 font-semibold">Date</th>
                <th className="px-5 py-3 font-semibold">Type</th>
                <th className="px-5 py-3 font-semibold">Reference</th>
                <th className="px-5 py-3 font-semibold">Particulars</th>
                <th className="px-5 py-3 text-right font-semibold">Debit</th>
                <th className="px-5 py-3 text-right font-semibold">Credit</th>
                <th className="px-5 py-3 text-right font-semibold">Balance</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-outline-variant">
              {loading ? (
                <tr>
                  <td colSpan={7} className="px-5 py-10 text-center text-on-surface-variant">
                    Loading…
                  </td>
                </tr>
              ) : data?.rows.length === 0 ? (
                <tr>
                  <td colSpan={7} className="px-5 py-10 text-center text-on-surface-variant">
                    No transactions.
                  </td>
                </tr>
              ) : (
                data?.rows.map((r) => (
                  <tr
                    key={`${r.source}-${r.id}`}
                    className="transition-colors hover:bg-surface-container-low"
                  >
                    <td className="whitespace-nowrap px-5 py-3 text-on-surface">
                      {r.source === "opening"
                        ? ""
                        : new Date(r.date).toLocaleDateString("en-GB", {
                            day: "2-digit",
                            month: "short",
                            year: "numeric",
                          })}
                    </td>
                    <td className="px-5 py-3">
                      <span
                        className={`rounded-full px-2 py-0.5 text-[10px] font-semibold uppercase ${
                          r.pending
                            ? "bg-amber-100 text-amber-700"
                            : "bg-primary/5 text-on-surface-variant"
                        }`}
                      >
                        {r.kind}
                        {r.pending ? " *" : ""}
                      </span>
                    </td>
                    <td className="px-5 py-3 font-mono text-xs text-outline">
                      {r.reference || "—"}
                    </td>
                    <td className="px-5 py-3 text-on-surface">{r.particulars}</td>
                    <td className="whitespace-nowrap px-5 py-3 text-right tabular-nums text-axis-error">
                      {r.debit > 0 ? fmt(r.debit) : "—"}
                    </td>
                    <td className="whitespace-nowrap px-5 py-3 text-right tabular-nums text-secondary">
                      {r.credit > 0 ? fmt(r.credit) : "—"}
                    </td>
                    <td className="whitespace-nowrap px-5 py-3 text-right font-bold tabular-nums text-on-surface">
                      {fmt(r.balance)}{" "}
                      <span className="text-[10px] font-medium text-on-surface-variant">
                        {r.balance >= 0 ? "Cr" : "Dr"}
                      </span>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>

        {data && data.total > pageSize && (
          <div className="flex items-center justify-between border-t border-outline-variant px-5 py-3">
            <span className="text-xs text-on-surface-variant">
              Page {page} of {totalPages} · {data.total} entries
            </span>
            <div className="flex gap-1">
              <button
                onClick={() => setPage(Math.max(1, page - 1))}
                disabled={page === 1}
                className="rounded-lg border border-outline-variant px-3 py-1.5 text-xs font-medium text-on-surface-variant transition-colors hover:bg-surface-container disabled:opacity-40"
              >
                Previous
              </button>
              <button
                onClick={() => setPage(Math.min(totalPages, page + 1))}
                disabled={page === totalPages}
                className="rounded-lg border border-outline-variant px-3 py-1.5 text-xs font-medium text-on-surface-variant transition-colors hover:bg-surface-container disabled:opacity-40"
              >
                Next
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

function SummaryCard({
  label,
  value,
  icon,
  tone,
}: {
  label: string;
  value: string;
  icon: string;
  tone: "error" | "secondary" | "muted";
}) {
  const iconWrap =
    tone === "error"
      ? "bg-axis-error-container/50 text-axis-error"
      : tone === "secondary"
        ? "bg-secondary-container/20 text-secondary"
        : "bg-surface-container text-outline";
  return (
    <div className="rounded-xl border border-outline-variant bg-surface-container-lowest p-5 shadow-[0px_4px_12px_rgba(15,23,42,0.03)]">
      <div className="mb-3 flex items-start justify-between">
        <span className="text-[11px] font-semibold uppercase tracking-wide text-on-surface-variant">
          {label}
        </span>
        <span className={`flex h-7 w-7 items-center justify-center rounded ${iconWrap}`}>
          <Icon name={icon} size={18} />
        </span>
      </div>
      <div className="font-display text-xl font-semibold text-primary">{value}</div>
    </div>
  );
}
