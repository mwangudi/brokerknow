import { useEffect, useState, type FormEvent } from "react";
import { useNavigate } from "react-router";
import api from "../lib/api";
import { useMarketToday } from "../hooks/useMarketToday";
import Icon from "../components/ui/Icon";
import OrderTabs from "../components/orders/OrderTabs";

interface LookupOption {
  value: number;
  label: string;
}

interface OrderOptions {
  orderTypes: LookupOption[];
  orderSecTypes: LookupOption[];
  securities: LookupOption[];
}

const ORDER_TYPE_PURCHASE = 1;

function fmtMoney(n: number) {
  return n.toLocaleString("en", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

export default function PlaceOrderPage() {
  const navigate = useNavigate();
  const [options, setOptions] = useState<OrderOptions | null>(null);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState<string | null>(null);

  const [orderTypeDpa, setOrderTypeDpa] = useState<string>(String(ORDER_TYPE_PURCHASE));
  const [orderSecTypeDpa, setOrderSecTypeDpa] = useState<string>("");
  const [securitySearch, setSecuritySearch] = useState("");
  const [securityDpa, setSecurityDpa] = useState<string>("");
  const [securityPickerOpen, setSecurityPickerOpen] = useState(false);
  const [best, setBest] = useState(false);
  const [quantity, setQuantity] = useState("");
  const [price, setPrice] = useState("");
  const [amount, setAmount] = useState("");
  const [reference, setReference] = useState("");
  const [remarks, setRemarks] = useState("");
  const [submitting, setSubmitting] = useState(false);
  const [submitError, setSubmitError] = useState<string | null>(null);

  // Portal orders are dated today, so warn before submit when the exchange
  // isn't open (weekend or public holiday).
  const market = useMarketToday();

  useEffect(() => {
    api
      .get<OrderOptions>("/portal/order-options")
      .then((r) => {
        setOptions(r.data);
        if (r.data.orderSecTypes.length > 0) {
          setOrderSecTypeDpa(String(r.data.orderSecTypes[0].value));
        }
      })
      .catch(() => setLoadError("Could not load order options. Please try again."))
      .finally(() => setLoading(false));
  }, []);

  const isPurchase = orderTypeDpa === String(ORDER_TYPE_PURCHASE);
  const selectedSecurity = options?.securities.find(
    (s) => String(s.value) === securityDpa,
  );

  const filteredSecurities = (() => {
    if (!options) return [];
    const term = securitySearch.trim().toLowerCase();
    if (!term) return options.securities.slice(0, 50);
    return options.securities
      .filter((s) => s.label.toLowerCase().includes(term))
      .slice(0, 50);
  })();

  // Live order-summary estimate from the form inputs.
  const qtyNum = parseInt(quantity, 10);
  const priceNum = parseFloat(price);
  const exposure =
    !best && Number.isFinite(qtyNum) && Number.isFinite(priceNum)
      ? qtyNum * priceNum
      : best && amount
        ? parseFloat(amount)
        : null;

  async function handleSubmit(e: FormEvent) {
    e.preventDefault();
    setSubmitError(null);

    if (!securityDpa) {
      setSubmitError("Please choose a security.");
      return;
    }
    if (!orderSecTypeDpa) {
      setSubmitError("Please choose the security type.");
      return;
    }
    const qty = best ? 0 : parseInt(quantity, 10);
    if (!best && (!Number.isFinite(qty) || qty <= 0)) {
      setSubmitError("Please enter a quantity greater than zero.");
      return;
    }
    if (!best && !price.trim()) {
      setSubmitError("Please enter a price.");
      return;
    }
    const amt = amount ? parseFloat(amount) : null;
    if (best && isPurchase && (!amt || amt <= 0)) {
      setSubmitError("Please enter an amount for a Best Purchase.");
      return;
    }
    if (reference.length > 100) {
      setSubmitError("Reference can be at most 100 characters.");
      return;
    }
    if (remarks.length > 50) {
      setSubmitError("Remarks can be at most 50 characters.");
      return;
    }

    setSubmitting(true);
    try {
      await api.post("/portal/orders", {
        orderTypeDpa: parseInt(orderTypeDpa, 10),
        orderSecTypeDpa: parseInt(orderSecTypeDpa, 10),
        securityDpa: parseInt(securityDpa, 10),
        quantity: qty,
        price: best ? null : price.trim(),
        best,
        amount: amt,
        reference: reference || null,
        remarks: remarks || null,
      });
      navigate("/orders");
    } catch (err) {
      const ex = err as { response?: { data?: { error?: string } } };
      setSubmitError(
        ex.response?.data?.error ?? "Could not place the order. Please try again.",
      );
    } finally {
      setSubmitting(false);
    }
  }

  const inputCls =
    "w-full rounded-lg border border-outline-variant bg-surface-container-lowest px-3 py-2.5 text-sm text-on-surface focus:border-secondary focus:outline-none focus:ring-1 focus:ring-secondary";
  const labelCls =
    "mb-1.5 block text-[11px] font-semibold uppercase tracking-wide text-on-surface-variant";

  if (loading) {
    return <p className="text-sm text-on-surface-variant">Loading…</p>;
  }
  if (loadError) {
    return <p className="text-sm text-axis-error">{loadError}</p>;
  }

  return (
    <div className="space-y-6">
      {/* Header + toggle */}
      <div className="flex flex-col items-start justify-between gap-4 md:flex-row md:items-end">
        <div>
          <h1 className="font-display text-2xl font-semibold text-primary">
            Order Management
          </h1>
          <p className="text-sm text-on-surface-variant">
            Execute and monitor your market activity.
          </p>
        </div>
        <OrderTabs active="place" />
      </div>

      {!market.isOpen && (
        <div className="flex items-start gap-2 rounded-lg border border-amber-300 bg-amber-50 px-4 py-3 text-sm text-amber-800">
          <Icon name="info" size={18} className="mt-0.5 shrink-0" />
          <div>
            <strong>The market is closed today.</strong>{" "}
            {market.reason === "weekend"
              ? "Orders can only be submitted on business days (Mon–Fri, excluding public holidays)."
              : `Today is a public holiday${market.holidayName ? ` (${market.holidayName})` : ""}. Orders can only be submitted on business days.`}
            {market.nextOpen && (
              <>
                {" "}You'll be able to place this order on{" "}
                <strong>
                  {new Date(market.nextOpen).toLocaleDateString("en-GB", {
                    weekday: "long",
                    day: "2-digit",
                    month: "short",
                    year: "numeric",
                  })}
                </strong>
                .
              </>
            )}
          </div>
        </div>
      )}

      <form onSubmit={handleSubmit} className="grid grid-cols-12 gap-6">
        {/* Order specification */}
        <div className="col-span-12 rounded-xl border border-outline-variant bg-surface-container-lowest p-6 shadow-[0px_4px_12px_rgba(15,23,42,0.03)] lg:col-span-7">
          <div className="mb-5 flex items-center justify-between border-b border-outline-variant pb-4">
            <h3 className="font-display text-lg font-semibold text-primary">
              Order Specification
            </h3>
            {selectedSecurity && (
              <span className="font-mono text-xs text-on-surface-variant">
                {selectedSecurity.label}
              </span>
            )}
          </div>

          {submitError && (
            <div className="mb-4 rounded-lg border border-axis-error/30 bg-axis-error/5 px-3 py-2 text-sm text-axis-error">
              {submitError}
            </div>
          )}

          <div className="space-y-5">
            {/* Side */}
            <div className="grid grid-cols-2 gap-4">
              {options?.orderTypes.map((ot) => {
                const active = orderTypeDpa === String(ot.value);
                const sale = ot.label.toLowerCase().startsWith("sale");
                const accent = sale ? "error" : "secondary";
                return (
                  <button
                    key={ot.value}
                    type="button"
                    onClick={() => setOrderTypeDpa(String(ot.value))}
                    className={`flex items-center justify-center gap-2 rounded-lg border-2 py-3 transition-all ${
                      active
                        ? accent === "error"
                          ? "border-axis-error bg-axis-error/5"
                          : "border-secondary bg-secondary-container/10"
                        : "border-outline-variant hover:border-secondary/50"
                    }`}
                  >
                    <Icon
                      name={sale ? "sell" : "add_shopping_cart"}
                      size={20}
                      className={accent === "error" ? "text-axis-error" : "text-secondary"}
                    />
                    <span
                      className={`text-xs font-semibold uppercase tracking-widest ${
                        active
                          ? accent === "error"
                            ? "text-axis-error"
                            : "text-secondary"
                          : "text-on-surface-variant"
                      }`}
                    >
                      {ot.label}
                    </span>
                  </button>
                );
              })}
            </div>

            {/* Security + type */}
            <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
              <div className="relative">
                <label className={labelCls}>Security</label>
                {selectedSecurity && !securityPickerOpen ? (
                  <div className="flex items-center justify-between rounded-lg border border-secondary/40 bg-secondary-container/10 px-3 py-2.5 text-sm">
                    <span className="truncate font-medium text-on-surface">
                      {selectedSecurity.label}
                    </span>
                    <button
                      type="button"
                      onClick={() => {
                        setSecurityPickerOpen(true);
                        setSecuritySearch("");
                      }}
                      className="ml-2 shrink-0 text-xs font-semibold text-secondary hover:underline"
                    >
                      Change
                    </button>
                  </div>
                ) : (
                  <>
                    <div className="relative">
                      <Icon
                        name="search"
                        size={18}
                        className="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2 text-outline"
                      />
                      <input
                        value={securitySearch}
                        onChange={(e) => setSecuritySearch(e.target.value)}
                        onFocus={() => setSecurityPickerOpen(true)}
                        placeholder="Ticker or name…"
                        className={`${inputCls} pl-9`}
                      />
                    </div>
                    {securityPickerOpen && (
                      <div className="absolute z-20 mt-1 max-h-72 w-full overflow-auto rounded-lg border border-outline-variant bg-surface-container-lowest shadow-lg">
                        {filteredSecurities.length === 0 ? (
                          <p className="px-3 py-2 text-sm text-on-surface-variant">
                            No securities match.
                          </p>
                        ) : (
                          <ul>
                            {filteredSecurities.map((s) => (
                              <li key={s.value}>
                                <button
                                  type="button"
                                  onClick={() => {
                                    setSecurityDpa(String(s.value));
                                    setSecurityPickerOpen(false);
                                    setSecuritySearch("");
                                  }}
                                  className="block w-full px-3 py-2 text-left text-sm text-on-surface hover:bg-surface-container-low"
                                >
                                  {s.label}
                                </button>
                              </li>
                            ))}
                          </ul>
                        )}
                        {!securitySearch.trim() &&
                          options &&
                          options.securities.length > 50 && (
                            <p className="border-t border-outline-variant px-3 py-1.5 text-[11px] text-on-surface-variant">
                              Showing first 50 of {options.securities.length}. Type
                              to filter.
                            </p>
                          )}
                      </div>
                    )}
                  </>
                )}
              </div>

              <div>
                <label className={labelCls}>Security Type</label>
                <select
                  value={orderSecTypeDpa}
                  onChange={(e) => setOrderSecTypeDpa(e.target.value)}
                  className={inputCls}
                  required
                >
                  <option value="">— Select —</option>
                  {options?.orderSecTypes.map((o) => (
                    <option key={o.value} value={o.value}>
                      {o.label}
                    </option>
                  ))}
                </select>
              </div>
            </div>

            {/* Quantity + execution type */}
            <div className="grid grid-cols-12 gap-4">
              <div className="col-span-7">
                <label className={labelCls}>Quantity</label>
                <input
                  type="number"
                  min={1}
                  value={quantity}
                  onChange={(e) => setQuantity(e.target.value)}
                  disabled={best}
                  placeholder="0.00"
                  className={`${inputCls} text-right disabled:opacity-50`}
                />
              </div>
              <div className="col-span-5">
                <label className={labelCls}>Execution Type</label>
                <button
                  type="button"
                  onClick={() => setBest(!best)}
                  className="flex w-full items-center justify-between rounded-lg border border-outline-variant bg-surface-container-low px-3 py-2.5"
                >
                  <span className="text-sm text-on-surface">
                    {best ? "Best (market)" : "Limit price"}
                  </span>
                  <span
                    className={`relative h-4 w-8 rounded-full transition-colors ${
                      best ? "bg-secondary" : "bg-outline"
                    }`}
                  >
                    <span
                      className={`absolute top-0.5 h-3 w-3 rounded-full bg-white transition-transform ${
                        best ? "translate-x-4" : "translate-x-0.5"
                      }`}
                    />
                  </span>
                </button>
              </div>
            </div>

            {/* Price (limit) or amount (best purchase) */}
            {!best ? (
              <div>
                <label className={labelCls}>Price (MWK)</label>
                <input
                  type="text"
                  value={price}
                  onChange={(e) => setPrice(e.target.value)}
                  placeholder="e.g. 1250.00"
                  className={`${inputCls} text-right`}
                />
              </div>
            ) : (
              isPurchase && (
                <div>
                  <label className={labelCls}>Maximum Amount (MWK)</label>
                  <input
                    type="number"
                    step="0.01"
                    min={0.01}
                    value={amount}
                    onChange={(e) => setAmount(e.target.value)}
                    placeholder="0.00"
                    className={`${inputCls} text-right`}
                  />
                  <p className="mt-1.5 text-xs text-on-surface-variant">
                    The broker will not buy more than this value.
                  </p>
                </div>
              )
            )}

            {/* Reference + remarks */}
            <div>
              <label className={labelCls}>External Reference (Optional)</label>
              <input
                value={reference}
                onChange={(e) => setReference(e.target.value)}
                maxLength={100}
                className={inputCls}
              />
            </div>
            <div>
              <label className={labelCls}>Execution Remarks</label>
              <textarea
                value={remarks}
                onChange={(e) => setRemarks(e.target.value)}
                maxLength={50}
                rows={2}
                placeholder="Instructions for the trading desk…"
                className={inputCls}
              />
            </div>
          </div>
        </div>

        {/* Order summary */}
        <div className="col-span-12 space-y-6 lg:col-span-5">
          <div className="overflow-hidden rounded-xl border border-primary-container bg-primary-container p-6 text-on-primary-container shadow-lg">
            <h3 className="mb-5 border-b border-on-primary-container/40 pb-4 font-display text-lg font-semibold text-primary-fixed-dim">
              Order Summary
            </h3>
            <div className="mb-6 space-y-3">
              <Row label="Side" value={isPurchase ? "Purchase" : "Sale"} />
              <Row
                label="Estimated Price"
                value={best ? "Market" : price ? `MWK ${price}` : "—"}
              />
              <Row
                label="Total Quantity"
                value={
                  best
                    ? "Best"
                    : Number.isFinite(qtyNum)
                      ? qtyNum.toLocaleString()
                      : "—"
                }
              />
              <div className="mt-4 border-t border-dashed border-on-primary-container/40 pt-4">
                <div className="flex items-end justify-between">
                  <span className="text-[11px] font-semibold uppercase tracking-widest text-on-primary-container">
                    {best ? "Max. Exposure" : "Est. Total Exposure"}
                  </span>
                  <span className="font-display text-xl font-bold text-secondary-fixed">
                    {exposure != null ? `MWK ${fmtMoney(exposure)}` : "—"}
                  </span>
                </div>
              </div>
            </div>
            <button
              type="submit"
              disabled={submitting || !market.isOpen}
              title={
                !market.isOpen
                  ? "The market is closed today — orders can only be placed on business days."
                  : undefined
              }
              className="w-full rounded-lg bg-secondary-fixed py-3.5 text-sm font-bold uppercase tracking-widest text-on-secondary-fixed transition-all hover:brightness-110 active:scale-95 disabled:cursor-not-allowed disabled:opacity-50"
            >
              {submitting ? "Submitting…" : "Submit Order"}
            </button>
            <p className="mt-3 text-center text-[11px] italic text-on-primary-container/60">
              Execution remains subject to market liquidity and desk approval.
            </p>
          </div>

          <button
            type="button"
            onClick={() => navigate("/orders")}
            className="w-full rounded-lg border border-outline-variant bg-surface-container-lowest py-2.5 text-sm font-medium text-on-surface-variant transition-colors hover:bg-surface-container-low"
          >
            Cancel
          </button>
        </div>
      </form>
    </div>
  );
}

function Row({ label, value }: { label: string; value: string }) {
  return (
    <div className="flex items-center justify-between">
      <span className="text-sm text-on-primary-container">{label}</span>
      <span className="text-sm font-medium text-white">{value}</span>
    </div>
  );
}
