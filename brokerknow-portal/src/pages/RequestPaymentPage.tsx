import { useEffect, useState, type FormEvent } from "react";
import api from "../lib/api";
import Icon from "../components/ui/Icon";
import { brand } from "../lib/brand";

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
  return n.toLocaleString("en", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function fmtDate(iso: string) {
  return new Date(iso).toLocaleString("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });
}

function statusTone(status: string) {
  switch (status) {
    case "Approved":
      return "bg-secondary/10 text-secondary";
    case "Rejected":
      return "bg-axis-error/10 text-axis-error";
    case "Pending":
      return "bg-amber-100 text-amber-700";
    default:
      return "bg-surface-container text-on-surface-variant";
  }
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
  // file. If they have none, the back office uses the free-text reference.
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

  useEffect(() => {
    void refresh();
  }, []);

  async function viewProof(id: number) {
    try {
      const resp = await api.get(`/portal/payment-requests/${id}/proof`, {
        responseType: "blob",
      });
      const url = window.URL.createObjectURL(resp.data as Blob);
      window.open(url, "_blank", "noopener,noreferrer");
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
    if (requestType === "Withdrawal" && balance && amt > balance.withdrawable) {
      setSubmitError(
        `You can withdraw at most ${brand.currency} ${fmtMoney(balance.withdrawable)} right now.`,
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
          requestType === "Withdrawal" && bankAccountId
            ? Number(bankAccountId)
            : null,
      });
      // Upload the proof of deposit if attached.
      if (requestType === "Deposit" && proofFile && created.data?.id) {
        const fd = new FormData();
        fd.append("file", proofFile);
        try {
          await api.post(
            `/portal/payment-requests/${created.data.id}/proof`,
            fd,
            { headers: { "Content-Type": "multipart/form-data" } },
          );
        } catch (uploadErr) {
          const ex = uploadErr as { response?: { data?: { error?: string } } };
          setSubmitError(
            "Request saved, but the proof could not be uploaded: " +
              (ex.response?.data?.error ??
                "please try again from the history list."),
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
      setSubmitError(
        ex.response?.data?.error ??
          "Could not submit your request. Please try again.",
      );
    } finally {
      setSubmitting(false);
    }
  }

  const inputCls =
    "w-full rounded-lg border border-outline-variant bg-surface-container-lowest px-3 py-2.5 text-sm text-on-surface focus:border-secondary focus:outline-none focus:ring-1 focus:ring-secondary";
  const labelCls =
    "mb-1.5 block text-[11px] font-semibold uppercase tracking-wide text-on-surface-variant";

  const submitDisabled =
    submitting ||
    overLimit ||
    missingBankAccount ||
    (requestType === "Withdrawal" &&
      balance != null &&
      balance.withdrawable <= 0);

  return (
    <div className="space-y-6">
      <div>
        <h1 className="font-display text-2xl font-semibold text-primary">
          Payments
        </h1>
        <p className="text-sm text-on-surface-variant">
          Submit a deposit notice or request a withdrawal of available funds.
        </p>
      </div>

      {/* Balance summary bento */}
      {balance && (
        <div className="grid grid-cols-1 gap-4 md:grid-cols-3">
          <SummaryCard
            label="Account Balance"
            value={`${brand.currency} ${fmtMoney(Math.abs(balance.balance))} ${balance.balance < 0 ? "Dr" : "Cr"}`}
            sub="Current account position"
            icon="account_balance_wallet"
            tone={balance.balance < 0 ? "error" : "secondary"}
          />
          <SummaryCard
            label="Withdrawable Now"
            value={`${brand.currency} ${fmtMoney(balance.withdrawable)}`}
            sub="Cash available to withdraw"
            icon="payments"
            tone={balance.withdrawable > 0 ? "secondary" : "muted"}
          />
          <SummaryCard
            label="Credit Limit"
            value={`${brand.currency} ${fmtMoney(balance.creditLimit)}`}
            sub="Trading facility — not withdrawable"
            icon="credit_card"
            tone="muted"
          />
        </div>
      )}

      <div className="grid grid-cols-12 gap-6">
        {/* Form */}
        <form
          onSubmit={handleSubmit}
          className="col-span-12 rounded-xl border border-outline-variant bg-surface-container-lowest p-6 shadow-[0px_4px_12px_rgba(15,23,42,0.03)] lg:col-span-6"
        >
          <div className="mb-5 flex items-center justify-between border-b border-outline-variant pb-4">
            <h3 className="font-display text-lg font-semibold text-primary">
              New Request
            </h3>
            <Icon name="currency_exchange" size={20} className="text-secondary" />
          </div>

          {submitOk && (
            <div className="mb-4 flex items-start gap-2 rounded-lg border border-secondary/30 bg-secondary/5 px-3 py-2 text-sm text-secondary">
              <Icon name="check_circle" size={18} filled className="mt-0.5 shrink-0" />
              <span>{submitOk}</span>
            </div>
          )}
          {submitError && (
            <div className="mb-4 rounded-lg border border-axis-error/30 bg-axis-error/5 px-3 py-2 text-sm text-axis-error">
              {submitError}
            </div>
          )}

          <div className="space-y-5">
            {/* Type toggle */}
            <div className="grid grid-cols-2 gap-4">
              {(["Deposit", "Withdrawal"] as const).map((t) => {
                const active = requestType === t;
                return (
                  <button
                    key={t}
                    type="button"
                    onClick={() => setRequestType(t)}
                    className={`flex items-center justify-center gap-2 rounded-lg border-2 py-3 transition-all ${
                      active
                        ? "border-secondary bg-secondary-container/10"
                        : "border-outline-variant hover:border-secondary/50"
                    }`}
                  >
                    <Icon
                      name={t === "Deposit" ? "south_east" : "north_east"}
                      size={20}
                      className={active ? "text-secondary" : "text-on-surface-variant"}
                    />
                    <span
                      className={`text-xs font-semibold uppercase tracking-widest ${
                        active ? "text-secondary" : "text-on-surface-variant"
                      }`}
                    >
                      {t}
                    </span>
                  </button>
                );
              })}
            </div>
            <p className="text-xs text-on-surface-variant">
              {requestType === "Deposit"
                ? "Tell us you have sent funds (cash, cheque or bank transfer) so we can credit your account."
                : "Ask us to release available cash from your trading account to your bank."}
            </p>

            {/* Amount */}
            <div>
              <label className={labelCls}>Amount ({brand.currency})</label>
              <input
                type="number"
                step="0.01"
                min={0.01}
                value={amount}
                onChange={(e) => setAmount(e.target.value)}
                placeholder="0.00"
                className={`${inputCls} text-right`}
                required
              />
              {requestType === "Withdrawal" && balance && (
                <p className={`mt-1.5 text-xs ${overLimit ? "text-axis-error" : "text-on-surface-variant"}`}>
                  You can withdraw up to{" "}
                  <strong>{brand.currency} {fmtMoney(balance.withdrawable)}</strong> right now.
                  {balance.withdrawable <= 0 && " There is no cash available to withdraw."}
                </p>
              )}
            </div>

            {/* Bank account (withdrawal) */}
            {requestType === "Withdrawal" && (
              <div>
                <label className={labelCls}>Send Funds To</label>
                {bankAccounts.length === 0 ? (
                  <p className="rounded-lg border border-amber-300 bg-amber-50 px-3 py-2 text-xs text-amber-800">
                    We don't have any of your bank accounts on file. Please contact
                    your broker to register one — for now, type the destination bank
                    and account number in the reference field below.
                  </p>
                ) : (
                  <>
                    <select
                      value={bankAccountId}
                      onChange={(e) => setBankAccountId(e.target.value)}
                      className={inputCls}
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
                    <p className="mt-1.5 text-xs text-on-surface-variant">
                      The broker will release the funds to this account on file.
                    </p>
                  </>
                )}
              </div>
            )}

            {/* Reference */}
            <div>
              <label className={labelCls}>
                {requestType === "Deposit"
                  ? "Bank / Cheque Reference"
                  : "Bank Account / Reference"}
              </label>
              <input
                value={reference}
                onChange={(e) => setReference(e.target.value)}
                placeholder={
                  requestType === "Deposit"
                    ? "e.g. NBM transfer ref TX12345"
                    : "e.g. NBM 0123456789"
                }
                maxLength={50}
                className={inputCls}
              />
            </div>

            {/* Notes */}
            <div>
              <label className={labelCls}>Notes</label>
              <textarea
                value={narrative}
                onChange={(e) => setNarrative(e.target.value)}
                rows={3}
                maxLength={500}
                placeholder={
                  requestType === "Deposit"
                    ? "Any extra information to help us match the funds (date sent, paying bank, etc.)"
                    : "Reason for the withdrawal (optional)."
                }
                className={inputCls}
              />
            </div>

            {/* Proof (deposit) */}
            {requestType === "Deposit" && (
              <div>
                <label className={labelCls}>Proof of Deposit</label>
                <label className="flex cursor-pointer items-center gap-3 rounded-lg border border-dashed border-outline-variant bg-surface-container-low px-3 py-3 transition-colors hover:border-secondary">
                  <Icon name="upload" size={20} className="text-on-surface-variant" />
                  <span className="min-w-0 flex-1 truncate text-sm text-on-surface-variant">
                    {proofFile ? proofFile.name : "Attach a bank slip or receipt…"}
                  </span>
                  <input
                    type="file"
                    accept=".pdf,.jpg,.jpeg,.png,.webp"
                    onChange={(e) => setProofFile(e.target.files?.[0] ?? null)}
                    className="hidden"
                  />
                </label>
                <p className="mt-1.5 text-xs text-on-surface-variant">
                  Optional. PDF, JPG, PNG or WEBP, max 10 MB.
                </p>
              </div>
            )}

            <button
              type="submit"
              disabled={submitDisabled}
              className="w-full rounded-lg bg-primary py-3 text-sm font-semibold uppercase tracking-wide text-on-primary transition-opacity hover:opacity-90 disabled:cursor-not-allowed disabled:opacity-50"
            >
              {submitting ? "Submitting…" : "Submit Request"}
            </button>
          </div>
        </form>

        {/* History */}
        <div className="col-span-12 overflow-hidden rounded-xl border border-outline-variant bg-surface-container-lowest shadow-[0px_4px_12px_rgba(15,23,42,0.03)] lg:col-span-6">
          <div className="flex items-center justify-between border-b border-outline-variant px-5 py-4">
            <h3 className="font-display text-lg font-semibold text-primary">
              Your Requests
            </h3>
            {rows && rows.length > 0 && (
              <span className="text-xs font-semibold text-on-surface-variant">
                {rows.length} total
              </span>
            )}
          </div>

          <div className="p-5">
            {loading ? (
              <p className="text-sm text-on-surface-variant">Loading…</p>
            ) : error ? (
              <p className="text-sm text-axis-error">{error}</p>
            ) : !rows || rows.length === 0 ? (
              <div className="rounded-lg border border-dashed border-outline-variant bg-surface-container-low px-4 py-10 text-center text-sm text-on-surface-variant">
                You haven't raised any payment requests yet.
              </div>
            ) : (
              <ul className="space-y-3">
                {rows.map((r) => (
                  <li
                    key={r.id}
                    className="rounded-lg border border-outline-variant p-4 transition-colors hover:border-secondary/50"
                  >
                    <div className="flex items-start justify-between gap-3">
                      <div className="min-w-0">
                        <div className="flex flex-wrap items-center gap-2">
                          <span className="flex items-center gap-1.5 text-sm font-semibold text-on-surface">
                            <Icon
                              name={r.requestType === "Deposit" ? "south_east" : "north_east"}
                              size={16}
                              className="text-on-surface-variant"
                            />
                            {r.requestType} · {brand.currency} {fmtMoney(r.amount)}
                          </span>
                          <span
                            className={`rounded-full px-2 py-0.5 text-[10px] font-semibold uppercase ${statusTone(r.status)}`}
                          >
                            {r.status}
                          </span>
                        </div>
                        <p className="mt-1 text-[11px] text-on-surface-variant">
                          Submitted {fmtDate(r.createdAt)}
                          {r.processedAt && <> · Processed {fmtDate(r.processedAt)}</>}
                        </p>
                        {r.reference && (
                          <p className="mt-1 text-xs text-on-surface-variant">
                            Ref: {r.reference}
                          </p>
                        )}
                        {r.clientBankAccount && (
                          <p className="mt-1 text-xs text-on-surface-variant">
                            To: {r.clientBankAccount}
                          </p>
                        )}
                        {r.narrative && (
                          <p className="mt-1 text-xs text-on-surface-variant">
                            {r.narrative}
                          </p>
                        )}
                        {r.status === "Rejected" && r.rejectReason && (
                          <p className="mt-1 text-xs text-axis-error">
                            Rejected: {r.rejectReason}
                          </p>
                        )}
                        {r.hasProof && (
                          <button
                            type="button"
                            onClick={() => void viewProof(r.id)}
                            className="mt-2 inline-flex items-center gap-1 text-xs font-semibold text-secondary hover:underline"
                          >
                            <Icon name="visibility" size={14} />
                            View proof of deposit
                          </button>
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
    </div>
  );
}

function SummaryCard({
  label,
  value,
  sub,
  icon,
  tone,
}: {
  label: string;
  value: string;
  sub: string;
  icon: string;
  tone: "secondary" | "error" | "muted";
}) {
  const toneCls =
    tone === "secondary"
      ? "text-secondary"
      : tone === "error"
        ? "text-axis-error"
        : "text-on-surface-variant";
  return (
    <div className="rounded-xl border border-outline-variant bg-surface-container-lowest p-5 shadow-[0px_4px_12px_rgba(15,23,42,0.03)]">
      <div className="mb-2 flex items-start justify-between">
        <span className="text-[11px] font-semibold uppercase tracking-wider text-on-surface-variant">
          {label}
        </span>
        <Icon name={icon} size={20} className={toneCls} />
      </div>
      <div className="font-display text-xl font-semibold text-primary">{value}</div>
      <p className="mt-1 text-xs text-on-surface-variant">{sub}</p>
    </div>
  );
}
