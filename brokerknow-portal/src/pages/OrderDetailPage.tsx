import { useEffect, useState } from "react";
import { Link, useParams } from "react-router";
import api from "../lib/api";
import Icon from "../components/ui/Icon";
import { brand } from "../lib/brand";

interface OrderItem {
  security: string;
  ordDetailQty: number;
  ordDetailPrice: string;
  best: boolean;
  amount: number | null;
  filledQty: number;
}

interface OrderDetail {
  orderDpa: number;
  orderRef?: string | null;
  orderDate: string;
  remarks?: string | null;
  orderType: string;
  secType: string;
  status: string;
  orderCompounded?: boolean;
  isCustodian?: boolean;
  interBank?: boolean;
  items: OrderItem[];
}

function sideTone(type: string) {
  const t = (type || "").toLowerCase();
  if (t.startsWith("purchase") || t.startsWith("buy"))
    return "bg-secondary-container/20 text-on-secondary-container";
  if (t.startsWith("sale") || t.startsWith("sell"))
    return "bg-axis-error-container/40 text-on-axis-error-container";
  return "bg-surface-container text-on-surface-variant";
}

function statusTone(status: string) {
  switch (status) {
    case "Traded":
      return "bg-secondary/10 text-secondary";
    case "Held":
      return "bg-amber-100 text-amber-700";
    case "Canceled":
      return "bg-axis-error/10 text-axis-error";
    case "Released":
      return "bg-primary/5 text-on-surface-variant";
    default:
      return "bg-surface-container text-on-surface-variant";
  }
}

function fmt(n: number) {
  return n.toLocaleString("en", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

export default function OrderDetailPage() {
  const { id } = useParams();
  const [order, setOrder] = useState<OrderDetail | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    setLoading(true);
    api
      .get<OrderDetail>(`/portal/orders/${id}`)
      .then((r) => setOrder(r.data))
      .catch(() => setError("Could not load this order."))
      .finally(() => setLoading(false));
  }, [id]);

  if (loading) return <p className="text-sm text-on-surface-variant">Loading…</p>;
  if (error || !order)
    return (
      <p className="text-sm text-axis-error">{error ?? "Order not found."}</p>
    );

  const totalQty = order.items.reduce(
    (s, it) => s + (it.best ? 0 : it.ordDetailQty),
    0,
  );
  const filledQty = order.items.reduce((s, it) => s + it.filledQty, 0);
  const grossValue = order.items.reduce(
    (s, it) =>
      s +
      (it.amount ??
        (it.best ? 0 : it.ordDetailQty * parseFloat(it.ordDetailPrice || "0"))),
    0,
  );

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="flex flex-col items-start justify-between gap-3 md:flex-row md:items-center">
        <div className="flex flex-wrap items-center gap-2">
          <h1 className="font-display text-2xl font-semibold text-primary">
            Order #{order.orderDpa}
          </h1>
          <span
            className={`rounded px-2 py-0.5 text-[10px] font-semibold uppercase ${sideTone(order.orderType)}`}
          >
            {order.orderType}
          </span>
          <span
            className={`rounded px-2 py-0.5 text-[10px] font-semibold uppercase ${statusTone(order.status)}`}
          >
            {order.status}
          </span>
          {order.isCustodian && (
            <span className="rounded bg-primary/5 px-2 py-0.5 text-[10px] font-semibold uppercase text-on-surface-variant">
              Custodian
            </span>
          )}
          {order.interBank && (
            <span className="rounded bg-primary/5 px-2 py-0.5 text-[10px] font-semibold uppercase text-on-surface-variant">
              Interbank
            </span>
          )}
        </div>
        <Link
          to="/orders"
          className="inline-flex items-center gap-1 text-sm font-semibold text-secondary hover:underline"
        >
          <Icon name="arrow_back" size={16} />
          Back to orders
        </Link>
      </div>

      {/* Summary tiles */}
      <div className="grid grid-cols-2 gap-4 lg:grid-cols-4">
        <Tile label="Items" value={String(order.items.length)} icon="list" />
        <Tile
          label="Total Qty"
          value={totalQty.toLocaleString("en")}
          icon="numbers"
        />
        <Tile
          label="Filled"
          value={`${filledQty.toLocaleString("en")} / ${totalQty.toLocaleString("en")}`}
          sub={totalQty ? `${Math.round((filledQty / totalQty) * 100)}% filled` : undefined}
          icon="task_alt"
        />
        <Tile
          label="Gross Value"
          value={`${brand.currency} ${fmt(grossValue)}`}
          icon="payments"
          emphasis
        />
      </div>

      {/* Detail grid */}
      <div className="grid grid-cols-1 gap-x-6 gap-y-5 rounded-xl border border-outline-variant bg-surface-container-lowest p-6 shadow-[0px_4px_12px_rgba(15,23,42,0.03)] sm:grid-cols-2 lg:grid-cols-4">
        <Detail
          label="Date"
          value={new Date(order.orderDate).toLocaleDateString("en-GB", {
            day: "2-digit",
            month: "short",
            year: "numeric",
          })}
        />
        <Detail label="Side" value={order.orderType} />
        <Detail label="Security Type" value={order.secType} />
        <Detail label="Status" value={order.status} />
        <Detail label="Reference" value={order.orderRef || "—"} />
        <div className="sm:col-span-2 lg:col-span-3">
          <dt className="mb-1 text-[11px] font-semibold uppercase tracking-wider text-outline">
            Remarks
          </dt>
          <dd className="text-sm text-on-surface">{order.remarks || "—"}</dd>
        </div>
      </div>

      {/* Items table */}
      <div className="overflow-hidden rounded-xl border border-outline-variant bg-surface-container-lowest shadow-[0px_4px_12px_rgba(15,23,42,0.03)]">
        <div className="border-b border-outline-variant px-5 py-4">
          <h3 className="font-display text-base font-semibold text-primary">
            Order Items
          </h3>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-left text-sm">
            <thead>
              <tr className="border-b border-outline-variant bg-surface-container-low text-[11px] uppercase tracking-wider text-on-surface-variant">
                <th className="px-5 py-3 font-semibold">Security</th>
                <th className="px-5 py-3 text-right font-semibold">Quantity</th>
                <th className="px-5 py-3 text-right font-semibold">Price</th>
                <th className="px-5 py-3 text-right font-semibold">Max Amount</th>
                <th className="px-5 py-3 text-right font-semibold">Filled</th>
                <th className="px-5 py-3 font-semibold">Best</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-outline-variant">
              {order.items.map((it, i) => (
                <tr key={i} className="transition-colors hover:bg-surface-container-low">
                  <td className="px-5 py-3 font-medium text-on-surface">
                    {it.security}
                  </td>
                  <td className="px-5 py-3 text-right tabular-nums text-on-surface">
                    {it.best ? "Best" : it.ordDetailQty.toLocaleString()}
                  </td>
                  <td className="px-5 py-3 text-right tabular-nums text-on-surface">
                    {it.best ? "Market" : it.ordDetailPrice}
                  </td>
                  <td className="px-5 py-3 text-right tabular-nums text-on-surface">
                    {it.amount != null
                      ? it.amount.toLocaleString(undefined, {
                          minimumFractionDigits: 2,
                        })
                      : "—"}
                  </td>
                  <td className="px-5 py-3 text-right tabular-nums text-on-surface">
                    {it.filledQty.toLocaleString()}
                  </td>
                  <td className="px-5 py-3">
                    {it.best ? (
                      <span className="rounded px-2 py-0.5 text-[10px] font-semibold uppercase bg-secondary/10 text-secondary">
                        Best
                      </span>
                    ) : (
                      <span className="text-outline">—</span>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function Tile({
  label,
  value,
  sub,
  icon,
  emphasis,
}: {
  label: string;
  value: string;
  sub?: string;
  icon: string;
  emphasis?: boolean;
}) {
  return (
    <div
      className={`rounded-xl border p-4 shadow-[0px_4px_12px_rgba(15,23,42,0.03)] ${
        emphasis
          ? "border-secondary/30 bg-secondary-container/10"
          : "border-outline-variant bg-surface-container-lowest"
      }`}
    >
      <div className="mb-2 flex items-center justify-between">
        <span className="text-[11px] font-semibold uppercase tracking-wide text-on-surface-variant">
          {label}
        </span>
        <Icon
          name={icon}
          size={18}
          className={emphasis ? "text-secondary" : "text-outline"}
        />
      </div>
      <p className="font-display text-lg font-semibold text-primary">{value}</p>
      {sub && <p className="mt-0.5 text-xs text-on-surface-variant">{sub}</p>}
    </div>
  );
}

function Detail({ label, value }: { label: string; value: string }) {
  return (
    <div>
      <dt className="mb-1 text-[11px] font-semibold uppercase tracking-wider text-outline">
        {label}
      </dt>
      <dd className="text-sm text-on-surface">{value}</dd>
    </div>
  );
}
