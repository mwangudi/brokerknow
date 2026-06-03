import { useEffect, useState } from "react";
import { Link, useParams } from "react-router";
import api from "../lib/api";

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
  items: OrderItem[];
}

const statusColor: Record<string, string> = {
  Released: "bg-green-100 text-green-700",
  Held: "bg-amber-100 text-amber-700",
  Canceled: "bg-red-100 text-red-700",
  Pending: "bg-gray-100 text-gray-700",
};

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

  return (
    <div>
      <div className="mb-6 flex items-center justify-between">
        <h2 className="text-2xl font-bold text-gray-900">Order #{order.orderDpa}</h2>
        <Link to="/orders" className="text-sm font-medium text-blue-600 hover:underline">
          ← Back to orders
        </Link>
      </div>

      <div className="mb-6 grid grid-cols-1 gap-4 rounded-xl border border-gray-200 bg-white p-6 sm:grid-cols-2 lg:grid-cols-4">
        <div>
          <p className="text-xs font-medium uppercase text-gray-500">Date</p>
          <p className="mt-1 text-sm text-gray-900">
            {new Date(order.orderDate).toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" })}
          </p>
        </div>
        <div>
          <p className="text-xs font-medium uppercase text-gray-500">Type</p>
          <p className="mt-1 text-sm text-gray-900">{order.orderType}</p>
        </div>
        <div>
          <p className="text-xs font-medium uppercase text-gray-500">Security type</p>
          <p className="mt-1 text-sm text-gray-900">{order.secType}</p>
        </div>
        <div>
          <p className="text-xs font-medium uppercase text-gray-500">Status</p>
          <span className={`mt-1 inline-flex rounded-full px-2 py-0.5 text-xs font-medium ${statusColor[order.status] || "bg-gray-100 text-gray-700"}`}>
            {order.status}
          </span>
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
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}
