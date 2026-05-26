import { useEffect, useState } from "react";
import api from "../lib/api";

interface OrderRow {
  orderDpa: number;
  orderRef?: string;
  orderDate: string;
  orderType: string;
  status: string;
}

export default function OrdersPage() {
  const [orders, setOrders] = useState<OrderRow[]>([]);
  const [total, setTotal] = useState(0);
  const [page, setPage] = useState(1);
  const [loading, setLoading] = useState(true);
  const pageSize = 15;

  useEffect(() => {
    setLoading(true);
    api.get(`/portal/orders?page=${page}&pageSize=${pageSize}`)
      .then((r) => { setOrders(r.data.rows); setTotal(r.data.total); })
      .catch(() => {})
      .finally(() => setLoading(false));
  }, [page]);

  const totalPages = Math.max(1, Math.ceil(total / pageSize));

  const statusColor: Record<string, string> = {
    Released: "bg-green-100 text-green-700",
    Held: "bg-amber-100 text-amber-700",
    Canceled: "bg-red-100 text-red-700",
    Pending: "bg-gray-100 text-gray-700",
  };

  return (
    <div>
      <h2 className="mb-6 text-2xl font-bold text-gray-900">My Orders</h2>

      <div className="rounded-xl border border-gray-200 bg-white">
        <div className="overflow-x-auto">
          <table className="min-w-full text-sm">
            <thead>
              <tr className="border-b border-gray-100 text-left text-xs font-medium uppercase text-gray-500">
                <th className="px-4 py-3">Order #</th>
                <th className="px-4 py-3">Date</th>
                <th className="px-4 py-3">Type</th>
                <th className="px-4 py-3">Reference</th>
                <th className="px-4 py-3">Status</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {loading ? (
                <tr><td colSpan={5} className="px-4 py-8 text-center text-gray-500">Loading...</td></tr>
              ) : orders.length === 0 ? (
                <tr><td colSpan={5} className="px-4 py-8 text-center text-gray-500">No orders found.</td></tr>
              ) : orders.map((o, i) => (
                <tr key={o.orderDpa} className={i % 2 === 0 ? "bg-blue-50/50" : ""}>
                  <td className="px-4 py-2.5 font-medium text-gray-900">{o.orderDpa}</td>
                  <td className="px-4 py-2.5 text-gray-700">
                    {new Date(o.orderDate).toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" })}
                  </td>
                  <td className="px-4 py-2.5">{o.orderType}</td>
                  <td className="px-4 py-2.5 text-gray-500">{o.orderRef || "—"}</td>
                  <td className="px-4 py-2.5">
                    <span className={`inline-flex rounded-full px-2 py-0.5 text-xs font-medium ${statusColor[o.status] || "bg-gray-100 text-gray-700"}`}>
                      {o.status}
                    </span>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        {total > pageSize && (
          <div className="flex items-center justify-between border-t border-gray-100 px-4 py-3">
            <span className="text-sm text-gray-500">Page {page} of {totalPages}</span>
            <div className="flex gap-1">
              <button onClick={() => setPage(Math.max(1, page - 1))} disabled={page === 1}
                className="rounded-lg border px-3 py-1.5 text-sm disabled:opacity-40">Previous</button>
              <button onClick={() => setPage(Math.min(totalPages, page + 1))} disabled={page === totalPages}
                className="rounded-lg border px-3 py-1.5 text-sm disabled:opacity-40">Next</button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
