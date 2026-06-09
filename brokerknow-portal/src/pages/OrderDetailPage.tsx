import { useEffect, useState } from "react";
import { Link, useParams } from "react-router";
import api from "../lib/api";
import Badge from "../components/ui/badge/Badge";

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

/** Buy/Sell side colour: Purchase = green, Sale = red — matches the back office. */
function sideColor(type: string): "success" | "error" | "primary" {
  const t = (type || "").toLowerCase();
  if (t.startsWith("purchase") || t.startsWith("buy")) return "success";
  if (t.startsWith("sale") || t.startsWith("sell")) return "error";
  return "primary";
}

function statusColor(status: string): "success" | "error" | "warning" | "info" | "light" {
  switch (status) {
    case "Canceled": return "error";
    case "Held": return "warning";
    case "Traded": return "success";
    case "Released": return "info";
    default: return "light";
  }
}

function fmt(n: number) {
  return n.toLocaleString("en", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function StatTile({ label, value, sub, emphasis }: { label: string; value: string; sub?: string; emphasis?: boolean }) {
  return (
    <div className={`rounded-xl border p-4 ${emphasis ? "border-blue-200 bg-blue-50" : "border-gray-200 bg-white"}`}>
      <p className="text-xs font-medium uppercase tracking-wide text-gray-500">{label}</p>
      <p className={`mt-1 text-lg font-semibold ${emphasis ? "text-blue-700" : "text-gray-900"}`}>{value}</p>
      {sub && <p className="mt-0.5 text-xs text-gray-500">{sub}</p>}
    </div>
  );
}

export default function OrderDetailPage() {
  const { id } = useParams();
  const [order, setOrder] = useState<OrderDetail | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    setLoading(true);
    api.get<OrderDetail>(`/portal/orders/${id}`)
      .then((r) => setOrder(r.data))
      .catch(() => setError("Could not load this order."))
      .finally(() => setLoading(false));
  }, [id]);

  if (loading) return <p className="text-sm text-gray-500">Loading...</p>;
  if (error || !order) return <p className="text-sm text-rose-600">{error ?? "Order not found."}</p>;

  const totalQty = order.items.reduce((s, it) => s + (it.best ? 0 : it.ordDetailQty), 0);
  const filledQty = order.items.reduce((s, it) => s + it.filledQty, 0);
  const grossValue = order.items.reduce(
    (s, it) => s + (it.amount ?? (it.best ? 0 : it.ordDetailQty * parseFloat(it.ordDetailPrice || "0"))),
    0,
  );

  return (
    <div>
      <div className="mb-6 flex items-center justify-between">
        <div className="flex flex-wrap items-center gap-2">
          <h2 className="text-2xl font-bold text-gray-900">Order #{order.orderDpa}</h2>
          <Badge size="sm" color={sideColor(order.orderType)}>{order.orderType}</Badge>
          <Badge size="sm" color={statusColor(order.status)}>{order.status}</Badge>
          {order.isCustodian && <Badge size="sm" color="info">Custodian</Badge>}
          {order.interBank && <Badge size="sm" color="info">Interbank</Badge>}
        </div>
        <Link to="/orders" className="text-sm font-medium text-blue-600 hover:underline">
          ← Back to orders
        </Link>
      </div>

      {/* Summary tiles */}
      <div className="mb-6 grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-4">
        <StatTile label="Items" value={String(order.items.length)} />
        <StatTile label="Total Qty" value={totalQty.toLocaleString("en")} />
        <StatTile
          label="Filled Qty"
          value={`${filledQty.toLocaleString("en")} / ${totalQty.toLocaleString("en")}`}
          sub={totalQty ? `${Math.round((filledQty / totalQty) * 100)}% filled` : undefined}
        />
        <StatTile label="Gross Value" value={`MWK ${fmt(grossValue)}`} emphasis />
      </div>

      <div className="mb-6 grid grid-cols-1 gap-4 rounded-xl border border-gray-200 bg-white p-6 sm:grid-cols-2 lg:grid-cols-4">
        <div>
          <p className="text-xs font-medium uppercase text-gray-500">Date</p>
          <p className="mt-1 text-sm text-gray-900">
            {new Date(order.orderDate).toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" })}
          </p>
        </div>
        <div>
          <p className="text-xs font-medium uppercase text-gray-500">Side</p>
          <p className="mt-1"><Badge size="sm" color={sideColor(order.orderType)}>{order.orderType}</Badge></p>
        </div>
        <div>
          <p className="text-xs font-medium uppercase text-gray-500">Security type</p>
          <p className="mt-1 text-sm text-gray-900">{order.secType}</p>
        </div>
        <div>
          <p className="text-xs font-medium uppercase text-gray-500">Status</p>
          <p className="mt-1"><Badge size="sm" color={statusColor(order.status)}>{order.status}</Badge></p>
        </div>
        <div>
          <p className="text-xs font-medium uppercase text-gray-500">Reference</p>
          <p className="mt-1 text-sm text-gray-900">{order.orderRef || "—"}</p>
        </div>
        <div className="sm:col-span-2 lg:col-span-3">
          <p className="text-xs font-medium uppercase text-gray-500">Remarks</p>
          <p className="mt-1 text-sm text-gray-900">{order.remarks || "—"}</p>
        </div>
      </div>

      <div className="rounded-xl border border-gray-200 bg-white">
        <div className="overflow-x-auto">
          <table className="min-w-full text-sm">
            <thead>
              <tr className="border-b border-gray-100 text-left text-xs font-medium uppercase text-gray-500">
                <th className="px-4 py-3">Security</th>
                <th className="px-4 py-3">Quantity</th>
                <th className="px-4 py-3">Price</th>
                <th className="px-4 py-3">Max amount</th>
                <th className="px-4 py-3">Filled</th>
                <th className="px-4 py-3">Best</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {order.items.map((it, i) => (
                <tr key={i} className={i % 2 === 0 ? "bg-blue-50/50" : ""}>
                  <td className="px-4 py-2.5 text-gray-900">{it.security}</td>
                  <td className="px-4 py-2.5 text-gray-700">
                    {it.best ? "Best (market)" : it.ordDetailQty.toLocaleString()}
                  </td>
                  <td className="px-4 py-2.5 text-gray-700">{it.best ? "Best" : it.ordDetailPrice}</td>
                  <td className="px-4 py-2.5 text-gray-700">
                    {it.amount != null ? it.amount.toLocaleString(undefined, { minimumFractionDigits: 2 }) : "—"}
                  </td>
                  <td className="px-4 py-2.5 text-gray-700">{it.filledQty.toLocaleString()}</td>
                  <td className="px-4 py-2.5">
                    {it.best ? <Badge size="sm" color="success">Best</Badge> : <span className="text-gray-400">—</span>}
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
