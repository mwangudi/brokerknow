import { useEffect, useState } from "react";
import api from "../lib/api";

interface MyApplication {
  id: number;
  applicationRef: string | null;
  sharesApplied: number;
  amount: number;
  status: string;
}

interface OfferRow {
  id: number;
  offerType: string;
  name: string;
  securityName: string | null;
  price: number;
  minimumApplication: number | null;
  applicationMultiple: number | null;
  openDate: string | null;
  closeDate: string | null;
  notes: string | null;
  myApplication: MyApplication | null;
}

interface ApplicationRow {
  id: number;
  applicationRef: string | null;
  applicationDate: string;
  offeringName: string;
  offerType: string;
  sharesApplied: number;
  amount: number;
  palNumber: string | null;
  allottedRights: number | null;
  acceptedRights: number | null;
  status: string;
  paidAt: string | null;
}

const fmt = (n: number | null | undefined, dp = 2) =>
  (n ?? 0).toLocaleString("en", { minimumFractionDigits: dp, maximumFractionDigits: dp });

const fmtDate = (iso: string | null) =>
  iso
    ? new Date(iso).toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" })
    : "-";

const statusClass = (s: string) =>
  s === "Paid"
    ? "bg-emerald-50 text-emerald-700"
    : s === "Cancelled"
      ? "bg-rose-50 text-rose-700"
      : "bg-amber-50 text-amber-700";

/**
 * Clients browse open IPO / Rights offers and apply. Applications are recorded
 * as Pending; the broker confirms the money and posts the receipt, so a client
 * can never mark their own application paid.
 */
export default function OfferingsPage() {
  const [offers, setOffers] = useState<OfferRow[]>([]);
  const [applications, setApplications] = useState<ApplicationRow[]>([]);
  const [loading, setLoading] = useState(true);

  const [applyTo, setApplyTo] = useState<OfferRow | null>(null);
  const [shares, setShares] = useState("");
  const [palNumber, setPalNumber] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  const [success, setSuccess] = useState("");

  function load() {
    setLoading(true);
    Promise.all([
      api.get<OfferRow[]>("/portal/ipo/offers"),
      api.get<ApplicationRow[]>("/portal/ipo/applications"),
    ])
      .then(([o, a]) => {
        setOffers(o.data);
        setApplications(a.data);
      })
      .catch(() => {/* leave empty */})
      .finally(() => setLoading(false));
  }

  useEffect(load, []);

  function openApply(offer: OfferRow) {
    setShares(offer.minimumApplication ? String(offer.minimumApplication) : "");
    setPalNumber("");
    setError("");
    setApplyTo(offer);
  }

  const amountDue = (() => {
    const n = Number(shares);
    if (!applyTo || !Number.isFinite(n) || n <= 0) return null;
    return Math.round(n * applyTo.price * 100) / 100;
  })();

  async function submit() {
    if (!applyTo) return;
    setBusy(true);
    setError("");
    try {
      const res = await api.post("/portal/ipo/apply", {
        offeringId: applyTo.id,
        sharesApplied: Number(shares),
        palNumber: palNumber.trim() || null,
      });
      setApplyTo(null);
      setSuccess(
        `Application ${res.data.applicationRef} received. Please pay ${fmt(res.data.amount)} to complete it.`,
      );
      load();
    } catch (err: unknown) {
      const ex = err as { response?: { data?: { error?: string } } };
      setError(ex.response?.data?.error ?? "Could not submit your application.");
    } finally {
      setBusy(false);
    }
  }

  return (
    <div>
      <div className="mb-6">
        <h2 className="text-2xl font-bold text-gray-900">IPOs &amp; Rights Issues</h2>
        <p className="mt-1 text-sm text-gray-500">
          Offers currently open for application, and the applications you have already made.
        </p>
      </div>

      {success && (
        <div className="mb-4 rounded-lg border border-emerald-200 bg-emerald-50 px-4 py-3 text-sm text-emerald-700">
          {success}
        </div>
      )}

      <h3 className="mb-3 text-sm font-semibold uppercase tracking-wide text-gray-500">
        Open offers
      </h3>

      {loading ? (
        <div className="rounded-xl border border-gray-200 bg-white px-4 py-8 text-center text-gray-500">
          Loading...
        </div>
      ) : offers.length === 0 ? (
        <div className="rounded-xl border border-gray-200 bg-white px-4 py-8 text-center text-gray-500">
          There are no offers open for application at the moment.
        </div>
      ) : (
        <div className="grid gap-4 sm:grid-cols-2">
          {offers.map((o) => (
            <div key={o.id} className="rounded-xl border border-gray-200 bg-white p-5">
              <div className="flex items-start justify-between gap-3">
                <div>
                  <p className="font-semibold text-gray-900">{o.name}</p>
                  <p className="text-xs text-gray-500">
                    {o.offerType}
                    {o.securityName ? ` · ${o.securityName}` : ""}
                  </p>
                </div>
                <span className="rounded-full bg-blue-50 px-2.5 py-1 text-xs font-medium text-blue-700">
                  {fmt(o.price)} per share
                </span>
              </div>

              <dl className="mt-4 space-y-1 text-sm text-gray-600">
                <div className="flex justify-between">
                  <dt>Closes</dt>
                  <dd className="font-medium text-gray-900">{fmtDate(o.closeDate)}</dd>
                </div>
                {o.minimumApplication ? (
                  <div className="flex justify-between">
                    <dt>Minimum</dt>
                    <dd className="font-medium text-gray-900">
                      {fmt(o.minimumApplication, 0)} shares
                    </dd>
                  </div>
                ) : null}
                {o.applicationMultiple ? (
                  <div className="flex justify-between">
                    <dt>In multiples of</dt>
                    <dd className="font-medium text-gray-900">
                      {fmt(o.applicationMultiple, 0)} shares
                    </dd>
                  </div>
                ) : null}
              </dl>

              {o.notes && <p className="mt-3 text-xs text-gray-500">{o.notes}</p>}

              <div className="mt-4">
                {o.myApplication ? (
                  <div className="rounded-lg bg-gray-50 px-3 py-2 text-sm">
                    <span className="text-gray-600">
                      You applied for {fmt(o.myApplication.sharesApplied, 0)} shares
                    </span>
                    <span
                      className={`ml-2 rounded-full px-2 py-0.5 text-xs font-medium ${statusClass(o.myApplication.status)}`}
                    >
                      {o.myApplication.status}
                    </span>
                  </div>
                ) : (
                  <button
                    type="button"
                    onClick={() => openApply(o)}
                    className="w-full rounded-lg bg-blue-600 px-4 py-2 text-sm font-medium text-white hover:bg-blue-700"
                  >
                    Apply
                  </button>
                )}
              </div>
            </div>
          ))}
        </div>
      )}

      <h3 className="mb-3 mt-8 text-sm font-semibold uppercase tracking-wide text-gray-500">
        My applications
      </h3>
      <div className="rounded-xl border border-gray-200 bg-white">
        <div className="overflow-x-auto">
          <table className="min-w-full text-sm">
            <thead>
              <tr className="border-b border-gray-100 text-left text-xs font-medium uppercase text-gray-500">
                <th className="px-4 py-3">Reference</th>
                <th className="px-4 py-3">Offer</th>
                <th className="px-4 py-3">Date</th>
                <th className="px-4 py-3 text-right">Shares</th>
                <th className="px-4 py-3 text-right">Amount</th>
                <th className="px-4 py-3">Status</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {applications.length === 0 ? (
                <tr>
                  <td colSpan={6} className="px-4 py-8 text-center text-gray-500">
                    You have not applied for any offers yet.
                  </td>
                </tr>
              ) : (
                applications.map((a, i) => (
                  <tr key={a.id} className={i % 2 === 0 ? "bg-blue-50/50" : ""}>
                    <td className="px-4 py-2.5 font-medium text-gray-900">
                      {a.applicationRef ?? `#${a.id}`}
                    </td>
                    <td className="px-4 py-2.5 text-gray-700">
                      {a.offeringName}
                      <span className="block text-xs text-gray-500">{a.offerType}</span>
                    </td>
                    <td className="px-4 py-2.5 text-gray-700">{fmtDate(a.applicationDate)}</td>
                    <td className="px-4 py-2.5 text-right tabular-nums text-gray-700">
                      {fmt(a.sharesApplied, 0)}
                    </td>
                    <td className="px-4 py-2.5 text-right tabular-nums font-medium text-gray-900">
                      {fmt(a.amount)}
                    </td>
                    <td className="px-4 py-2.5">
                      <span
                        className={`rounded-full px-2 py-0.5 text-xs font-medium ${statusClass(a.status)}`}
                      >
                        {a.status}
                      </span>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>

      {applyTo && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4">
          <div className="w-full max-w-md rounded-xl bg-white p-6 shadow-xl">
            <h3 className="text-lg font-semibold text-gray-900">Apply for {applyTo.name}</h3>
            <p className="mt-1 text-sm text-gray-500">
              Shares are offered at {fmt(applyTo.price)} each. Your broker will confirm your
              payment before the application is complete.
            </p>

            <label className="mt-5 block text-sm font-medium text-gray-700">
              Number of shares
            </label>
            <input
              className="mt-1 w-full rounded-lg border border-gray-300 px-3 py-2 text-right tabular-nums focus:border-blue-500 focus:outline-none"
              inputMode="numeric"
              value={shares}
              onChange={(e) => setShares(e.target.value.replace(/[^\d]/g, ""))}
            />
            {applyTo.minimumApplication ? (
              <p className="mt-1 text-xs text-gray-500">
                Minimum {fmt(applyTo.minimumApplication, 0)} shares
                {applyTo.applicationMultiple
                  ? `, in multiples of ${fmt(applyTo.applicationMultiple, 0)}`
                  : ""}
              </p>
            ) : null}

            {applyTo.offerType === "Rights" && (
              <>
                <label className="mt-4 block text-sm font-medium text-gray-700">
                  PAL number
                </label>
                <input
                  className="mt-1 w-full rounded-lg border border-gray-300 px-3 py-2 focus:border-blue-500 focus:outline-none"
                  value={palNumber}
                  maxLength={40}
                  placeholder="From your provisional allotment letter"
                  onChange={(e) => setPalNumber(e.target.value)}
                />
              </>
            )}

            <div className="mt-4 flex items-center justify-between rounded-lg bg-gray-50 px-4 py-3 text-sm">
              <span className="text-gray-600">Amount due</span>
              <span className="font-semibold tabular-nums text-gray-900">
                {amountDue != null ? fmt(amountDue) : "-"}
              </span>
            </div>

            {error && <p className="mt-3 text-sm text-rose-600">{error}</p>}

            <div className="mt-5 flex justify-end gap-2">
              <button
                type="button"
                onClick={() => setApplyTo(null)}
                disabled={busy}
                className="rounded-lg border border-gray-300 px-4 py-2 text-sm font-medium text-gray-700 hover:bg-gray-50 disabled:opacity-50"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={submit}
                disabled={busy || !shares}
                className="rounded-lg bg-blue-600 px-4 py-2 text-sm font-medium text-white hover:bg-blue-700 disabled:opacity-50"
              >
                {busy ? "Submitting..." : "Submit application"}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
