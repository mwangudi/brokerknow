import { useEffect, useState, type FormEvent } from "react";
import { useNavigate } from "react-router";
import api from "../lib/api";

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

  useEffect(() => {
    api.get<OrderOptions>("/portal/order-options")
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
  const selectedSecurity = options?.securities.find((s) => String(s.value) === securityDpa);

  const filteredSecurities = (() => {
    if (!options) return [];
    const term = securitySearch.trim().toLowerCase();
    if (!term) return options.securities.slice(0, 50);
    return options.securities
      .filter((s) => s.label.toLowerCase().includes(term))
      .slice(0, 50);
  })();

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
      setSubmitError(ex.response?.data?.error ?? "Could not place the order. Please try again.");
    } finally {
      setSubmitting(false);
    }
  }

  if (loading) {
    return <p className="text-sm text-gray-500">Loading...</p>;
  }
  if (loadError) {
    return <p className="text-sm text-rose-600">{loadError}</p>;
  }

  return (
    <div>
      <div className="mb-6 overflow-hidden rounded-2xl bg-gradient-to-r from-blue-600 to-indigo-600 p-6 text-white shadow-lg">
        <p className="text-sm text-blue-100">Trading</p>
        <h2 className="mt-1 text-2xl font-bold">Place an order</h2>
        <p className="mt-2 text-sm text-blue-100">
          Submit a buy or sell instruction. Your broker reviews every portal
          order before releasing it onto the exchange.
        </p>
      </div>

      <form
        onSubmit={handleSubmit}
        className="max-w-2xl space-y-5 rounded-2xl border border-gray-200 bg-white p-6 shadow-sm"
      >
        {submitError && (
          <div className="rounded-lg border border-rose-200 bg-rose-50 px-3 py-2 text-sm text-rose-800">
            {submitError}
          </div>
        )}

        <div>
          <label className="mb-1.5 block text-sm font-medium text-gray-700">Side *</label>
          <div className="inline-flex rounded-lg border border-gray-200 bg-gray-50 p-1">
            {options?.orderTypes.map((ot) => {
              const active = orderTypeDpa === String(ot.value);
              return (
                <button
                  key={ot.value}
                  type="button"
                  onClick={() => setOrderTypeDpa(String(ot.value))}
                  className={`rounded-md px-4 py-1.5 text-sm font-medium transition-colors ${
                    active
                      ? "bg-blue-600 text-white shadow-sm"
                      : "text-gray-600 hover:text-gray-800"
                  }`}
                >
                  {ot.label}
                </button>
              );
            })}
          </div>
        </div>

        <div>
          <label className="mb-1.5 block text-sm font-medium text-gray-700">Security type *</label>
          <select
            value={orderSecTypeDpa}
            onChange={(e) => setOrderSecTypeDpa(e.target.value)}
            className="w-full rounded-lg border border-gray-300 bg-white px-4 py-2.5 text-sm text-gray-700 focus:border-blue-400 focus:outline-none focus:ring-4 focus:ring-blue-200"
            required
          >
            <option value="">— Select —</option>
            {options?.orderSecTypes.map((o) => (
              <option key={o.value} value={o.value}>{o.label}</option>
            ))}
          </select>
        </div>

        <div className="relative">
          <label className="mb-1.5 block text-sm font-medium text-gray-700">Security *</label>
          {selectedSecurity && !securityPickerOpen ? (
            <div className="flex items-center justify-between rounded-lg border border-blue-200 bg-blue-50 px-3 py-2.5 text-sm">
              <span className="font-medium text-blue-900">{selectedSecurity.label}</span>
              <button
                type="button"
                onClick={() => { setSecurityPickerOpen(true); setSecuritySearch(""); }}
                className="text-xs font-medium text-blue-700 hover:underline"
              >
                Change
              </button>
            </div>
          ) : (
            <>
              <input
                value={securitySearch}
                onChange={(e) => setSecuritySearch(e.target.value)}
                onFocus={() => setSecurityPickerOpen(true)}
                placeholder="Type a code or name to search…"
                className="w-full rounded-lg border border-gray-300 bg-white px-4 py-2.5 text-sm text-gray-700 focus:border-blue-400 focus:outline-none focus:ring-4 focus:ring-blue-200"
              />
              {securityPickerOpen && (
                <div className="absolute z-10 mt-1 max-h-72 w-full overflow-auto rounded-lg border border-gray-200 bg-white shadow-lg">
                  {filteredSecurities.length === 0 ? (
                    <p className="px-3 py-2 text-sm text-gray-500">No securities match.</p>
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
                            className="block w-full px-3 py-2 text-left text-sm text-gray-700 hover:bg-blue-50"
                          >
                            {s.label}
                          </button>
                        </li>
                      ))}
                    </ul>
                  )}
                  {!securitySearch.trim() && options && options.securities.length > 50 && (
                    <p className="border-t border-gray-100 px-3 py-1.5 text-[11px] text-gray-500">
                      Showing first 50 of {options.securities.length}. Type to filter.
                    </p>
                  )}
                </div>
              )}
            </>
          )}
        </div>

        <div>
          <label className="inline-flex items-center gap-2 text-sm text-gray-700">
            <input
              type="checkbox"
              checked={best}
              onChange={(e) => setBest(e.target.checked)}
              className="h-4 w-4 rounded border-gray-300"
            />
            <span>Best market price (broker decides quantity/price)</span>
          </label>
        </div>

        {!best && (
          <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
            <div>
              <label className="mb-1.5 block text-sm font-medium text-gray-700">Quantity *</label>
              <input
                type="number"
                min={1}
                value={quantity}
                onChange={(e) => setQuantity(e.target.value)}
                className="w-full rounded-lg border border-gray-300 bg-white px-4 py-2.5 text-sm text-gray-700 focus:border-blue-400 focus:outline-none focus:ring-4 focus:ring-blue-200"
                required
              />
            </div>
            <div>
              <label className="mb-1.5 block text-sm font-medium text-gray-700">Price (MWK) *</label>
              <input
                type="text"
                value={price}
                onChange={(e) => setPrice(e.target.value)}
                placeholder="e.g. 1250.00"
                className="w-full rounded-lg border border-gray-300 bg-white px-4 py-2.5 text-sm text-gray-700 focus:border-blue-400 focus:outline-none focus:ring-4 focus:ring-blue-200"
                required
              />
            </div>
          </div>
        )}

        {best && isPurchase && (
          <div>
            <label className="mb-1.5 block text-sm font-medium text-gray-700">Maximum amount (MWK) *</label>
            <input
              type="number"
              step="0.01"
              min={0.01}
              value={amount}
              onChange={(e) => setAmount(e.target.value)}
              placeholder="0.00"
              className="w-full rounded-lg border border-gray-300 bg-white px-4 py-2.5 text-sm text-gray-700 focus:border-blue-400 focus:outline-none focus:ring-4 focus:ring-blue-200"
              required
            />
            <p className="mt-1.5 text-xs text-gray-500">Broker will not buy more than this value.</p>
          </div>
        )}

        <div>
          <label className="mb-1.5 block text-sm font-medium text-gray-700">Reference</label>
          <input
            value={reference}
            onChange={(e) => setReference(e.target.value)}
            maxLength={100}
            placeholder="Your own reference (optional)"
            className="w-full rounded-lg border border-gray-300 bg-white px-4 py-2.5 text-sm text-gray-700 focus:border-blue-400 focus:outline-none focus:ring-4 focus:ring-blue-200"
          />
        </div>

        <div>
          <label className="mb-1.5 block text-sm font-medium text-gray-700">Remarks</label>
          <input
            value={remarks}
            onChange={(e) => setRemarks(e.target.value)}
            maxLength={50}
            placeholder="Short note for the broker (optional)"
            className="w-full rounded-lg border border-gray-300 bg-white px-4 py-2.5 text-sm text-gray-700 focus:border-blue-400 focus:outline-none focus:ring-4 focus:ring-blue-200"
          />
        </div>

        <div className="flex justify-end gap-2 pt-2">
          <button
            type="button"
            onClick={() => navigate("/orders")}
            className="rounded-lg border border-gray-300 bg-white px-5 py-2.5 text-sm font-medium text-gray-700 hover:bg-gray-50"
          >
            Cancel
          </button>
          <button
            type="submit"
            disabled={submitting}
            className="rounded-lg bg-blue-600 px-5 py-2.5 text-sm font-medium text-white shadow-sm hover:bg-blue-700 disabled:cursor-not-allowed disabled:opacity-50"
          >
            {submitting ? "Submitting..." : "Submit order"}
          </button>
        </div>
      </form>
    </div>
  );
}
