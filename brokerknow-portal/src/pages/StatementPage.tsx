import { useEffect, useState } from "react";
import api from "../lib/api";
import DatePicker from "../components/form/DatePicker";

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
  summary: { totalDebit: number; totalCredit: number; closingBalance: number; pendingCount: number };
}

function fmt(n: number) {
  return Math.abs(n).toLocaleString("en", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
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

  function fetchStatement(f: string, t: string, p: number) {
    setLoading(true);
    const params = new URLSearchParams();
    if (f) params.set("from", f);
    if (t) params.set("to", t);
    params.set("page", p.toString());
    params.set("pageSize", pageSize.toString());
    api.get(`/portal/statement?${params}`)
      .then((r) => setData(r.data))
      .catch(() => {})
      .finally(() => setLoading(false));
  }

  useEffect(() => {
    fetchStatement(appliedFrom, appliedTo, page);
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
    api
      .get(`/portal/statement.pdf?${params}`, { responseType: "blob" })
      .then((r) => {
        const url = window.URL.createObjectURL(new Blob([r.data], { type: "application/pdf" }));
        const link = document.createElement("a");
        link.href = url;
        link.setAttribute("download", `statement-${new Date().toISOString().slice(0, 10)}.pdf`);
        document.body.appendChild(link);
        link.click();
        link.remove();
        window.URL.revokeObjectURL(url);
      })
      .catch(() => alert("Failed to download statement PDF."));
  }

  const totalPages = data ? Math.max(1, Math.ceil(data.total / pageSize)) : 1;

  return (
    <div>
      <h2 className="mb-6 text-2xl font-bold text-gray-900">My Statement</h2>

      {/* Summary cards (top) */}
      {data?.summary && (
        <div className="mb-5 grid grid-cols-2 gap-4 lg:grid-cols-4">
          <Stat label="Total Debit" value={`MWK ${fmt(data.summary.totalDebit)}`} />
          <Stat label="Total Credit" value={`MWK ${fmt(data.summary.totalCredit)}`} />
          <Stat label="Closing Balance" value={`MWK ${fmt(data.summary.closingBalance)} ${data.summary.closingBalance >= 0 ? "Cr" : "Dr"}`} />
          <Stat label="Pending" value={data.summary.pendingCount.toString()} />
        </div>
      )}

      {/* Filters / Download */}
      <div className="mb-5 flex flex-wrap items-end justify-end gap-4 rounded-xl border border-gray-200 bg-white p-4 dark:border-gray-800 dark:bg-white/[0.03]">
        <div>
          <div className="w-44">
            <DatePicker value={from} onChange={setFrom} placeholder="Filter from date" maxDate={null} />
          </div>
        </div>
        <div>
          <div className="w-44">
            <DatePicker value={to} onChange={setTo} placeholder="Filter to date" maxDate={null} />
          </div>
        </div>
        <button onClick={search}
          className="rounded-lg bg-brand-500 px-4 py-2 text-sm font-medium text-white hover:bg-brand-600">
          Search
        </button>
        <button onClick={downloadPdf}
          className="rounded-lg bg-red-600 px-4 py-2 text-sm font-medium text-white hover:bg-red-700">
          PDF
        </button>
      </div>

      {/* Table */}
      <div className="rounded-xl border border-gray-200 bg-white">
        <div className="overflow-x-auto">
          <table className="min-w-full text-sm">
            <thead>
              <tr className="border-b border-gray-100 text-left text-xs font-medium uppercase text-gray-500">
                <th className="px-4 py-3">Date</th>
                <th className="px-4 py-3">Type</th>
                <th className="px-4 py-3">Reference</th>
                <th className="px-4 py-3">Particulars</th>
                <th className="px-4 py-3 text-right">Debit</th>
                <th className="px-4 py-3 text-right">Credit</th>
                <th className="px-4 py-3 text-right">Balance</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {loading ? (
                <tr><td colSpan={7} className="px-4 py-8 text-center text-gray-500">Loading...</td></tr>
              ) : data?.rows.length === 0 ? (
                <tr><td colSpan={7} className="px-4 py-8 text-center text-gray-500">No transactions.</td></tr>
              ) : data?.rows.map((r, i) => (
                <tr key={`${r.source}-${r.id}`} className={i % 2 === 0 ? "bg-blue-50/50" : ""}>
                  <td className="whitespace-nowrap px-4 py-2.5 text-gray-700">
                    {r.source === "opening" ? "" : new Date(r.date).toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" })}
                  </td>
                  <td className="px-4 py-2.5">
                    <span className={`inline-flex rounded-full px-2 py-0.5 text-xs font-medium ${r.pending ? "bg-amber-100 text-amber-700" : "bg-gray-100 text-gray-700"}`}>
                      {r.kind}{r.pending ? " *" : ""}
                    </span>
                  </td>
                  <td className="px-4 py-2.5 text-gray-500">{r.reference || "—"}</td>
                  <td className="px-4 py-2.5">{r.particulars}</td>
                  <td className="whitespace-nowrap px-4 py-2.5 text-right font-medium">{r.debit > 0 ? fmt(r.debit) : "—"}</td>
                  <td className="whitespace-nowrap px-4 py-2.5 text-right font-medium">{r.credit > 0 ? fmt(r.credit) : "—"}</td>
                  <td className={`whitespace-nowrap px-4 py-2.5 text-right font-semibold ${r.balance >= 0 ? "text-green-600" : "text-red-600"}`}>
                    {fmt(r.balance)} {r.balance >= 0 ? "Cr" : "Dr"}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        {data && data.total > pageSize && (
          <div className="flex items-center justify-between border-t border-gray-100 px-4 py-3">
            <span className="text-sm text-gray-500">Page {page} of {totalPages} ({data.total} entries)</span>
            <div className="flex gap-1">
              <button onClick={() => setPage(Math.max(1, page - 1))} disabled={page === 1}
                className="rounded-lg border px-3 py-1.5 text-sm disabled:opacity-40">Previous</button>
              <button onClick={() => setPage(Math.min(totalPages, page + 1))} disabled={page === totalPages}
                className="rounded-lg border px-3 py-1.5 text-sm disabled:opacity-40">Next</button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

function Stat({ label, value }: { label: string; value: string }) {
  return (
    <div className="rounded-xl border border-gray-200 bg-white p-4">
      <p className="text-xs font-medium uppercase tracking-wider text-gray-500">{label}</p>
      <p className="mt-1.5 text-lg font-semibold text-gray-900">{value}</p>
    </div>
  );
}
