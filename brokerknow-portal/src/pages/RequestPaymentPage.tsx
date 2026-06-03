import { useEffect, useState, type FormEvent } from "react";
import api from "../lib/api";

interface PaymentRequestRow {
  id: number;
  requestType: string;
  amount: number;
  reference: string | null;
  narrative: string | null;
  status: string;
  rejectReason: string | null;
  createdAt: string;
  processedAt: string | null;
  clientBankAccountDpa: number | null;
  clientBankAccount: string | null;
  hasProof: boolean;
}

interface BalanceResponse {
  openingBalance: number;
  balance: number;
  outstanding: number;
  creditLimit: number;
  withdrawable: number;
}

interface BankAccountOption {
  id: number;
  bankName: string | null;
  accountNumber: string | null;
  accountName: string | null;
  branch: string | null;
  display: string;
}

function fmtMoney(n: number) {
  return n.toLocaleString("en", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function fmtDate(iso: string) {
  return new Date(iso).toLocaleString("en-GB", {
    day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit",
  });
}

function statusBadge(status: string) {
  const map: Record<string, string> = {
    Pending:  "bg-amber-100 text-amber-800 ring-amber-200",
    Approved: "bg-emerald-100 text-emerald-800 ring-emerald-200",
    Rejected: "bg-rose-100 text-rose-800 ring-rose-200",
  };
  const cls = map[status] ?? "bg-gray-100 text-gray-700 ring-gray-200";
  return (
    <span className={`inline-flex items-center rounded-full px-2.5 py-0.5 text-xs font-medium ring-1 ring-inset ${cls}`}>
      {status}
    </span>
  );
}

export default function RequestPaymentPage() {
  const [rows, setRows] = useState<PaymentRequestRow[] | null>(null);
  const [balance, setBalance] = useState<BalanceResponse | null>(null);
  const [bankAccounts, setBankAccounts] = useState<BankAccountOption[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  // Form state
  const [requestType, setRequestType] = useState<"Deposit" | "Withdrawal">("Deposit");
  const [amount, setAmount] = useState("");
  const [reference, setReference] = useState("");
  const [narrative, setNarrative] = useState("");
  const [bankAccountId, setBankAccountId] = useState<string>("");
  const [proofFile, setProofFile] = useState<File | null>(null);
  const [submitting, setSubmitting] = useState(false);
  const [submitError, setSubmitError] = useState<string | null>(null);
  const [submitOk, setSubmitOk] = useState<string | null>(null);

  const parsedAmount = parseFloat(amount);
  const overLimit =
    requestType === "Withdrawal" &&
    balance != null &&
    Number.isFinite(parsedAmount) &&
    parsedAmount > balance.withdrawable;

  // Withdrawals require a bank account when the client has one or more on
  // file. If they have none, the back office will use the free-text reference.
  const needsBankAccount =
    requestType === "Withdrawal" && bankAccounts.length > 0;
  const missingBankAccount = needsBankAccount && !bankAccountId;

  async function refresh() {
    setLoading(true);
    setError(null);
    try {
      const [reqs, bal, accts] = await Promise.all([
        api.get<PaymentRequestRow[]>("/portal/payment-requests"),
        api.get<BalanceResponse>("/portal/balance"),
        api.get<BankAccountOption[]>("/portal/bank-accounts"),
      ]);
      setRows(reqs.data);
      setBalance(bal.data);
      setBankAccounts(accts.data);
      // Pre-select the first account if there's only one — saves a click.
      if (accts.data.length === 1) {
        setBankAccountId(String(accts.data[0].id));
      }
    } catch {
      setError("Could not load your payment requests.");
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => { void refresh(); }, []);

  async function viewProof(id: number) {
    try {
      const resp = await api.get(`/portal/payment-requests/${id}/proof`, { responseType: "blob" });
      const url = window.URL.createObjectURL(resp.data as Blob);
      window.open(url, "_blank", "noopener,noreferrer");
      // Revoke shortly after to free memory; the new tab keeps its own ref.
      setTimeout(() => window.URL.revokeObjectURL(url), 60_000);
    } catch {
      setError("Could not open the proof file.");
    }
  }

  async function handleSubmit(e: FormEvent) {
    e.preventDefault();
    setSubmitError(null);
    setSubmitOk(null);

    const amt = parseFloat(amount);
    if (!Number.isFinite(amt) || amt <= 0) {
      setSubmitError("Please enter an amount greater than zero.");
      return;
    }
    if (reference.length > 50) {
      setSubmitError("Reference can be at most 50 characters.");
      return;
    }
    if (narrative.length > 500) {
      setSubmitError("Narrative can be at most 500 characters.");
      return;
    }
    // Pre-flight balance check for withdrawals — the server will re-check, but
    // catching it here is faster and clearer for the client.
    if (requestType === "Withdrawal" && balance && amt > balance.withdrawable) {
      setSubmitError(
        `You can withdraw at most MWK ${fmtMoney(balance.withdrawable)} right now.`,
      );
      return;
    }
    if (missingBankAccount) {
      setSubmitError("Please choose the bank account where the funds should be sent.");
      return;
    }

    setSubmitting(true);
    try {
      const created = await api.post<{ id: number }>("/portal/payment-requests", {
        requestType,
        amount: amt,
        reference: reference || null,
        narrative: narrative || null,
        clientBankAccountDpa:
          requestType === "Withdrawal" && bankAccountId ? Number(bankAccountId) : null,
      });
      // J5 — upload the proof of deposit if attached.
      if (requestType === "Deposit" && proofFile && created.data?.id) {
        const fd = new FormData();
        fd.append("file", proofFile);
        try {
          await api.post(`/portal/payment-requests/${created.data.id}/proof`, fd, {
            headers: { "Content-Type": "multipart/form-data" },
          });
        } catch (uploadErr) {
          const ex = uploadErr as { response?: { data?: { error?: string } } };
          setSubmitError(
            "Request saved, but the proof could not be uploaded: " +
              (ex.response?.data?.error ?? "please try again from the history list."),
          );
        }
      }
      setAmount("");
      setReference("");
      setNarrative("");
      setProofFile(null);
      if (bankAccounts.length !== 1) setBankAccountId("");
      setSubmitOk(
        requestType === "Deposit"
          ? "Deposit request submitted. Your account will be credited once the broker verifies the funds."
          : "Withdrawal request submitted. The broker will contact you when funds are released.",
      );
      await refresh();
    } catch (err) {
      const ex = err as { response?: { data?: { error?: string } } };
      setSubmitError(ex.response?.data?.error ?? "Could not submit your request. Please try again.");
    } finally {
      setSubmitting(false);
    }
  }

  return (
    <div>
      <div className="mb-6 overflow-hidden rounded-2xl bg-gradient-to-r from-amber-500 to-orange-600 p-6 text-white shadow-lg">
        <p className="text-sm text-amber-100">Cash management</p>
        <h2 className="mt-1 text-2xl font-bold">Request a Payment</h2>
        <p className="mt-2 text-sm text-amber-100">
          Submit a deposit notice or request a withdrawal of available funds. Your broker
          will review every request and update your statement once processed.
        </p>
      </div>

      {balance && (
        <div className="mb-6 grid grid-cols-1 gap-3 sm:grid-cols-3">
          <div className="rounded-xl border border-gray-200 bg-white p-4">
            <p className="text-xs uppercase tracking-wide text-gray-500">Available balance</p>
            <p className={`mt-1 text-xl font-bold ${balance.balance >= 0 ? "text-emerald-700" : "text-rose-700"}`}>
              MWK {fmtMoney(Math.abs(balance.balance))} {balance.balance < 0 ? "Dr" : "Cr"}
            </p>
          </div>
          <div className="rounded-xl border border-gray-200 bg-white p-4">
            <p className="text-xs uppercase tracking-wide text-gray-500">Withdrawable now</p>
            <p className={`mt-1 text-xl font-bold ${balance.withdrawable > 0 ? "text-emerald-700" : "text-gray-400"}`}>
              MWK {fmtMoney(balance.withdrawable)}
            </p>
            <p className="mt-0.5 text-[11px] text-gray-500">
              Cash available after other pending withdrawals.
            </p>
          </div>
          <div className="rounded-xl border border-gray-200 bg-white p-4">
            <p className="text-xs uppercase tracking-wide text-gray-500">Credit limit</p>
            <p className="mt-1 text-xl font-bold text-gray-800">MWK {fmtMoney(balance.creditLimit)}</p>
            <p className="mt-0.5 text-[11px] text-gray-500">
              Trading facility — not available for withdrawal.
            </p>
          </div>
        </div>
      )}

      <div className="grid grid-cols-1 gap-6 lg:grid-cols-2">
        {/* ── Form ────────────────────────────────────────── */}
        <form
          onSubmit={handleSubmit}
          className="rounded-2xl border border-gray-200 bg-white p-6 shadow-sm"
        >
          <h3 className="mb-4 text-base font-semibold text-gray-800">New request</h3>

          {submitOk && (
            <div className="mb-4 rounded-lg border border-emerald-200 bg-emerald-50 px-3 py-2 text-sm text-emerald-800">
              {submitOk}
            </div>
          )}
          {submitError && (
            <div className="mb-4 rounded-lg border border-rose-200 bg-rose-50 px-3 py-2 text-sm text-rose-800">
              {submitError}
            </div>
          )}

          <div className="space-y-4">
            <div>
              <label className="mb-1.5 block text-sm font-medium text-gray-700">Request type</label>
              <div className="inline-flex rounded-lg border border-gray-200 bg-gray-50 p-1">
                {(["Deposit", "Withdrawal"] as const).map((t) => (
                  <button
                    key={t}
                    type="button"
                    onClick={() => setRequestType(t)}
                    className={`rounded-md px-4 py-1.5 text-sm font-medium transition-colors ${
                      requestType === t
                        ? "bg-amber-500 text-white shadow-sm"
                        : "text-gray-600 hover:text-gray-800"
                    }`}
                  >
                    {t}
                  </button>
                ))}
              </div>
              <p className="mt-1.5 text-xs text-gray-500">
                {requestType === "Deposit"
                  ? "Tell us you have sent funds (cash, cheque or bank transfer) so we can credit your account."
                  : "Ask us to release available cash from your trading account to your bank."}
              </p>
            </div>

            <div>
              <label className="mb-1.5 block text-sm font-medium text-gray-700">Amount (MWK) *</label>
              <input
                type="number"
                step="0.01"
                min={0.01}
                value={amount}
                onChange={(e) => setAmount(e.target.value)}
                placeholder="0.00"
                className="w-full rounded-lg border border-gray-300 bg-white px-4 py-2.5 text-sm text-gray-700 focus:border-amber-400 focus:outline-none focus:ring-4 focus:ring-amber-200"
                required
              />
              {requestType === "Withdrawal" && balance && (
                <p className={`mt-1.5 text-xs ${overLimit ? "text-rose-600" : "text-gray-500"}`}>
                  You can withdraw up to <strong>MWK {fmtMoney(balance.withdrawable)}</strong> right now.
                  {balance.withdrawable <= 0 && " There is no cash available to withdraw."}
                </p>
              )}
            </div>

            {requestType === "Withdrawal" && (
              <div>
                <label className="mb-1.5 block text-sm font-medium text-gray-700">
                  Send funds to *
                </label>
                {bankAccounts.length === 0 ? (
                  <p className="rounded-lg border border-amber-200 bg-amber-50 px-3 py-2 text-xs text-amber-800">
                    We don't have any of your bank accounts on file. Please contact
                    your broker to register one — for now, type the destination
                    bank and account number in the reference field below.
                  </p>
                ) : (
                  <>
                    <select
                      value={bankAccountId}
                      onChange={(e) => setBankAccountId(e.target.value)}
                      className="w-full rounded-lg border border-gray-300 bg-white px-4 py-2.5 text-sm text-gray-700 focus:border-amber-400 focus:outline-none focus:ring-4 focus:ring-amber-200"
                      required
                    >
                      <option value="">— Select your bank account —</option>
                      {bankAccounts.map((a) => (
                        <option key={a.id} value={a.id}>
                          {a.display}
                          {a.accountName ? ` (${a.accountName})` : ""}
                        </option>
                      ))}
                    </select>
                    <p className="mt-1.5 text-xs text-gray-500">
                      The broker will release the funds to this account on file.
                    </p>
                  </>
                )}
              </div>
            )}

            <div>
              <label className="mb-1.5 block text-sm font-medium text-gray-700">
                {requestType === "Deposit" ? "Bank / cheque reference" : "Bank account / reference"}
              </label>
              <input
                value={reference}
                onChange={(e) => setReference(e.target.value)}
                placeholder={requestType === "Deposit"
                  ? "e.g. NBM transfer ref TX12345"
                  : "e.g. NBM 0123456789"}
                maxLength={50}
                className="w-full rounded-lg border border-gray-300 bg-white px-4 py-2.5 text-sm text-gray-700 focus:border-amber-400 focus:outline-none focus:ring-4 focus:ring-amber-200"
              />
            </div>

            <div>
              <label className="mb-1.5 block text-sm font-medium text-gray-700">Notes</label>
              <textarea
                value={narrative}
                onChange={(e) => setNarrative(e.target.value)}
                rows={3}
                maxLength={500}
                placeholder={requestType === "Deposit"
                  ? "Any extra information to help us match the funds (date sent, paying bank, etc.)"
                  : "Reason for the withdrawal (optional)."}
                className="w-full rounded-lg border border-gray-300 bg-white px-4 py-2.5 text-sm text-gray-700 focus:border-amber-400 focus:outline-none focus:ring-4 focus:ring-amber-200"
              />
            </div>

            {requestType === "Deposit" && (
              <div>
                <label className="mb-1.5 block text-sm font-medium text-gray-700">
                  Proof of deposit
                </label>
                <input
                  type="file"
                  accept=".pdf,.jpg,.jpeg,.png,.webp"
                  onChange={(e) => setProofFile(e.target.files?.[0] ?? null)}
                  className="block w-full text-sm text-gray-700 file:mr-3 file:rounded-md file:border-0 file:bg-amber-50 file:px-3 file:py-2 file:text-sm file:font-medium file:text-amber-700 hover:file:bg-amber-100"
                />
                <p className="mt-1.5 text-xs text-gray-500">
                  Optional. Attach a bank slip or transfer receipt (PDF, JPG, PNG or WEBP, max 10 MB).
                </p>
              </div>
            )}

            <div className="flex justify-end">
              <button
                type="submit"
                disabled={submitting || overLimit || missingBankAccount || (requestType === "Withdrawal" && balance != null && balance.withdrawable <= 0)}
                className="rounded-lg bg-amber-500 px-5 py-2.5 text-sm font-medium text-white shadow-sm hover:bg-amber-600 disabled:cursor-not-allowed disabled:opacity-50"
              >
                {submitting ? "Submitting..." : "Submit request"}
              </button>
            </div>
          </div>
        </form>

        {/* ── History ─────────────────────────────────────── */}
        <div className="rounded-2xl border border-gray-200 bg-white p-6 shadow-sm">
          <h3 className="mb-4 text-base font-semibold text-gray-800">Your requests</h3>

          {loading ? (
            <p className="text-sm text-gray-500">Loading...</p>
          ) : error ? (
            <p className="text-sm text-rose-600">{error}</p>
          ) : !rows || rows.length === 0 ? (
            <div className="rounded-lg border border-dashed border-gray-300 bg-gray-50 p-6 text-center text-sm text-gray-500">
              You haven't raised any payment requests yet.
            </div>
          ) : (
            <ul className="divide-y divide-gray-100">
              {rows.map((r) => (
                <li key={r.id} className="py-3">
                  <div className="flex items-start justify-between gap-3">
                    <div className="min-w-0">
                      <div className="flex flex-wrap items-center gap-2">
                        <span className="text-sm font-semibold text-gray-800">
                          {r.requestType} · MWK {fmtMoney(r.amount)}
                        </span>
                        {statusBadge(r.status)}
                      </div>
                      <p className="mt-0.5 text-xs text-gray-500">
                        Submitted {fmtDate(r.createdAt)}
                        {r.processedAt && (
                          <> · Processed {fmtDate(r.processedAt)}</>
                        )}
                      </p>
                      {r.reference && (
                        <p className="mt-1 text-xs text-gray-600">Ref: {r.reference}</p>
                      )}
                      {r.clientBankAccount && (
                        <p className="mt-1 text-xs text-gray-600">To: {r.clientBankAccount}</p>
                      )}
                      {r.narrative && (
                        <p className="mt-1 text-xs text-gray-600">{r.narrative}</p>
                      )}
                      {r.status === "Rejected" && r.rejectReason && (
                        <p className="mt-1 text-xs text-rose-700">
                          Rejected: {r.rejectReason}
                        </p>
                      )}
                      {r.hasProof && (
                        <p className="mt-1">
                          <button
                            type="button"
                            onClick={() => void viewProof(r.id)}
                            className="text-xs font-medium text-amber-700 hover:underline"
                          >
                            View proof of deposit
                          </button>
                        </p>
                      )}
                    </div>
                  </div>
                </li>
              ))}
            </ul>
          )}
        </div>
      </div>
    </div>
  );
}
