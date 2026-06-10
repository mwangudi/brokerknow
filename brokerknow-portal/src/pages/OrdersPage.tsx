import { useEffect, useState } from "react";
import { Link } from "react-router";
import api from "../lib/api";
import Icon from "../components/ui/Icon";
import OrderTabs from "../components/orders/OrderTabs";

interface OrderRow {
  orderDpa: number;
  orderRef?: string;
  orderDate: string;
  orderType: string;
  status: string;
  security?: string | null;
  quantity?: number | null;
  price?: string | null;
  best?: boolean;
}

function sideTone(type: string) {
  const t = (type || "").toLowerCase();
  if (t.startsWith("purchase") || t.startsWith("buy"))
    return "bg-secondary-container/20 text-on-secondary-container";
  if (t.startsWith("sale") || t.startsWith("sell"))
    return "bg-axis-error-container/40 text-on-axis-error-container";
  return "bg-surface-container text-on-surface-variant";
}

function statusDot(status: string) {
  switch (status) {
    case "Traded":
      return "bg-secondary";
    case "Held":
      return "bg-amber-500";
    case "Canceled":
      return "bg-axis-error";
    case "Released":
      return "bg-primary-fixed-dim";
    default:
      return "bg-outline";
  }
}

type Filter = "all" | "filled" | "pending";

export default function OrdersPage() {
  const [orders, setOrders] = useState<OrderRow[]>([]);
  const [total, setTotal] = useState(0);
  const [page, setPage] = useState(1);
  const [loading, setLoading] = useState(true);
  const [filter, setFilter] = useState<Filter>("all");
  const pageSize = 15;

  useEffect(() => {
    setLoading(true);
    api
      .get(`/portal/orders?page=${page}&pageSize=${pageSize}`)
      .then((r) => {
        setOrders(r.data.rows);
        setTotal(r.data.total);
      })
      .catch(() => {})
      .finally(() => setLoading(false));
  }, [page]);

  const totalPages = Math.max(1, Math.ceil(total / pageSize));

  const visible = orders.filter((o) => {
    if (filter === "filled") return o.status === "Traded";
    if (filter === "pending") return o.status === "Held" || o.status === "Released";
    return true;
  });

  return (
    <div className="space-y-6">
      {/* Header + toggle */}
      <div className="flex flex-col items-start justify-between gap-4 md:flex-row md:items-end">
        <div>
          <h1 className="font-display text-2xl font-semibold text-primary">
            Order Management
          </h1>
          <p className="text-sm text-on-surface-variant">
            Track and monitor your market activity.
          </p>
        </div>
        <OrderTabs active="view" />
      </div>

      {/* Orders table */}
      <div className="overflow-hidden rounded-xl border border-outline-variant bg-surface-container-lowest shadow-[0px_4px_12px_rgba(15,23,42,0.03)]">
        <div className="flex items-center gap-2 border-b border-outline-variant bg-surface-container-low px-4 py-3">
          {(["all", "filled", "pending"] as Filter[]).map((f) => (
            <button
              key={f}
              onClick={() => setFilter(f)}
              className={`rounded px-3 py-1 text-xs font-semibold uppercase tracking-wide transition-colors ${
                filter === f
                  ? "bg-primary text-on-primary"
                  : "text-on-surface-variant hover:bg-surface-container"
              }`}
            >
              {f === "all" ? "All Orders" : f === "filled" ? "Filled" : "Pending"}
            </button>
          ))}
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left text-sm">
            <thead>
              <tr className="border-b border-outline-variant bg-surface-container-low text-[11px] uppercase tracking-wider text-on-surface-variant">
                <th className="px-4 py-3 font-semibold">Order #</th>
                <th className="px-4 py-3 font-semibold">Date</th>
                <th className="px-4 py-3 font-semibold">Security</th>
                <th className="px-4 py-3 font-semibold">Side</th>
                <th className="px-4 py-3 text-right font-semibold">Qty</th>
                <th className="px-4 py-3 text-right font-semibold">Price</th>
                <th className="px-4 py-3 font-semibold">Status</th>
                <th className="px-4 py-3"></th>
              </tr>
            </thead>
            <tbody className="divide-y divide-outline-variant">
              {loading ? (
                <tr>
                  <td colSpan={8} className="px-4 py-10 text-center text-on-surface-variant">
                    Loading…
                  </td>
                </tr>
              ) : visible.length === 0 ? (
                <tr>
                  <td colSpan={8} className="px-4 py-10 text-center text-on-surface-variant">
                    No orders found.
                  </td>
                </tr>
              ) : (
                visible.map((o) => (
                  <tr
                    key={o.orderDpa}
                    className="transition-colors hover:bg-surface-container-low"
                  >
                    <td className="px-4 py-3 font-mono text-xs text-on-surface-variant">
                      #{o.orderDpa}
                    </td>
                    <td className="whitespace-nowrap px-4 py-3 text-on-surface-variant">
                      {new Date(o.orderDate).toLocaleDateString("en-GB", {
                        day: "2-digit",
                        month: "short",
                        year: "numeric",
                      })}
                    </td>
                    <td className="px-4 py-3 font-medium text-on-surface">
                      {o.security || "—"}
                      {o.orderRef && (
                        <span className="ml-2 text-[11px] text-outline">{o.orderRef}</span>
                      )}
                    </td>
                    <td className="px-4 py-3">
                      <span
                        className={`rounded px-2 py-0.5 text-[10px] font-semibold uppercase ${sideTone(o.orderType)}`}
                      >
                        {o.orderType}
                      </span>
                    </td>
                    <td className="px-4 py-3 text-right tabular-nums text-on-surface">
                      {o.best ? "Best" : o.quantity ? o.quantity.toLocaleString() : "—"}
                    </td>
                    <td className="px-4 py-3 text-right tabular-nums text-on-surface">
                      {o.best ? "Market" : o.price || "—"}
                    </td>
                    <td className="px-4 py-3">
                      <span className="flex items-center gap-2 text-on-surface-variant">
                        <span className={`h-2 w-2 rounded-full ${statusDot(o.status)}`} />
                        <span className="text-xs font-medium">{o.status}</span>
                      </span>
                    </td>
                    <td className="px-4 py-3 text-right">
                      <Link
                        to={`/orders/${o.orderDpa}`}
                        className="inline-flex items-center gap-1 text-xs font-semibold text-secondary hover:underline"
                      >
                        View
                        <Icon name="chevron_right" size={16} />
                      </Link>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>

        {total > pageSize && (
          <div className="flex items-center justify-between border-t border-outline-variant px-4 py-3">
            <span className="text-xs text-on-surface-variant">
              Page {page} of {totalPages} · {total} orders
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
