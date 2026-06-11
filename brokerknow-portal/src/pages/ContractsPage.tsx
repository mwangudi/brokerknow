import { useEffect, useState } from "react";
import api from "../lib/api";
import { brand } from "../lib/brand";

interface ContractRow {
  contractDpa: number;
  contractNumber: string | null;
  settlementDate: string | null;
  side: string | null;
  symbol: string | null;
  securityName: string | null;
  quantity: number;
  gross: number;
  tradeDate: string | null;
}

const fmt = (n: number) =>
  n.toLocaleString("en", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

const fmtDate = (iso: string | null) =>
  iso
    ? new Date(iso).toLocaleDateString("en-GB", {
        day: "2-digit",
        month: "short",
        year: "numeric",
      })
    : "—";

export default function ContractsPage() {
  const [rows, setRows] = useState<ContractRow[]>([]);
  const [total, setTotal] = useState(0);
  const [page, setPage] = useState(1);
  const [loading, setLoading] = useState(true);
  const [downloadError, setDownloadError] = useState<string | null>(null);
  const pageSize = 15;

  useEffect(() => {
    setLoading(true);
    api
      .get(`/portal/contracts?page=${page}&pageSize=${pageSize}`)
      .then((r) => {
        setRows(r.data.rows);
        setTotal(r.data.total);
      })
      .catch(() => {
        /* leave the list empty */
      })
      .finally(() => setLoading(false));
  }, [page]);

  function downloadNote(id: number, label: string) {
    setDownloadError(null);
    api
      .get(`/portal/contracts/${id}/note.pdf`, { responseType: "blob" })
      .then((r) => {
        const url = window.URL.createObjectURL(
          new Blob([r.data], { type: "application/pdf" }),
        );
        const link = document.createElement("a");
        link.href = url;
        link.setAttribute("download", `contract-note-${label}.pdf`);
        document.body.appendChild(link);
        link.click();
        link.remove();
        window.URL.revokeObjectURL(url);
      })
      .catch(() =>
        setDownloadError("Failed to download contract note. Please try again."),
      );
  }

  const totalPages = Math.max(1, Math.ceil(total / pageSize));

  const sideColor: Record<string, string> = {
    Purchase: "bg-green-100 text-green-700",
    Sale: "bg-red-100 text-red-700",
  };

  return (
    <div>
      <div className="mb-6">
        <h2 className="text-2xl font-bold text-gray-900">Contract Notes</h2>
        <p className="mt-1 text-sm text-gray-600">
          Every trade executed on your account generates a contract note.
          Download the PDF for your records.
        </p>
      </div>

      {downloadError && (
        <div className="mb-3 rounded-lg border border-rose-200 bg-rose-50 px-3 py-2 text-sm text-rose-800">
          {downloadError}
        </div>
      )}

      <div className="rounded-xl border border-gray-200 bg-white">
        <div className="overflow-x-auto">
          <table className="min-w-full text-sm">
            <thead>
              <tr className="border-b border-gray-100 text-left text-xs font-medium uppercase text-gray-500">
                <th className="px-4 py-3">Contract #</th>
                <th className="px-4 py-3">Trade Date</th>
                <th className="px-4 py-3">Settlement</th>
                <th className="px-4 py-3">Side</th>
                <th className="px-4 py-3">Security</th>
                <th className="px-4 py-3 text-right">Quantity</th>
                <th className="px-4 py-3 text-right">Gross ({brand.currency})</th>
                <th className="px-4 py-3"></th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {loading ? (
                <tr>
                  <td colSpan={8} className="px-4 py-8 text-center text-gray-500">
                    Loading…
                  </td>
                </tr>
              ) : rows.length === 0 ? (
                <tr>
                  <td colSpan={8} className="px-4 py-8 text-center text-gray-500">
                    You have no contract notes yet.
                  </td>
                </tr>
              ) : (
                rows.map((r, i) => (
                  <tr key={r.contractDpa} className={i % 2 === 0 ? "bg-blue-50/50" : ""}>
                    <td className="px-4 py-2.5 font-medium text-gray-900">
                      {r.contractNumber ?? `#${r.contractDpa}`}
                    </td>
                    <td className="px-4 py-2.5 text-gray-700">{fmtDate(r.tradeDate)}</td>
                    <td className="px-4 py-2.5 text-gray-700">{fmtDate(r.settlementDate)}</td>
                    <td className="px-4 py-2.5">
                      <span
                        className={`inline-flex rounded-full px-2 py-0.5 text-xs font-medium ${
                          sideColor[r.side ?? ""] ?? "bg-gray-100 text-gray-700"
                        }`}
                      >
                        {r.side ?? "—"}
                      </span>
                    </td>
                    <td className="px-4 py-2.5 text-gray-700">
                      {r.symbol ? (
                        <>
                          <span className="font-medium">{r.symbol}</span>
                          {r.securityName ? (
                            <span className="text-gray-500"> — {r.securityName}</span>
                          ) : null}
                        </>
                      ) : (
                        "—"
                      )}
                    </td>
                    <td className="px-4 py-2.5 text-right text-gray-700">
                      {r.quantity.toLocaleString()}
                    </td>
                    <td className="px-4 py-2.5 text-right font-medium text-gray-900">
                      {fmt(r.gross)}
                    </td>
                    <td className="px-4 py-2.5 text-right">
                      <button
                        type="button"
                        onClick={() =>
                          downloadNote(
                            r.contractDpa,
                            r.contractNumber ?? String(r.contractDpa),
                          )
                        }
                        className="text-sm font-medium text-blue-600 hover:underline"
                      >
                        Download PDF
                      </button>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>

        {totalPages > 1 && (
          <div className="flex items-center justify-between border-t border-gray-100 px-4 py-3 text-sm">
            <span className="text-gray-500">
              Page {page} of {totalPages} · {total} contract(s)
            </span>
            <div className="flex gap-2">
              <button
                type="button"
                onClick={() => setPage((p) => Math.max(1, p - 1))}
                disabled={page === 1}
                className="rounded-md border border-gray-200 px-3 py-1 text-xs text-gray-700 hover:bg-gray-50 disabled:opacity-50"
              >
                Previous
              </button>
              <button
                type="button"
                onClick={() => setPage((p) => Math.min(totalPages, p + 1))}
                disabled={page === totalPages}
                className="rounded-md border border-gray-200 px-3 py-1 text-xs text-gray-700 hover:bg-gray-50 disabled:opacity-50"
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
