import { useEffect, useState } from "react";
import { Link } from "react-router";
import api from "../lib/api";
import { useAuth } from "../context/AuthContext";
import { useMarketToday } from "../hooks/useMarketToday";
import Icon from "../components/ui/Icon";
import { brand } from "../lib/brand";

interface Balance {
  openingBalance: number;
  balance: number;
  outstanding: number;
  creditLimit: number;
}

interface MarketQuoteLite {
  securityDpa: number;
  securityCode: string;
  securityName: string;
  quoteDate: string;
  close: number | null;
  changePct: number | null;
  changeAbs: number | null;
  volume: number | null;
}

interface OrderRow {
  orderDpa: number;
  orderRef?: string;
  orderDate: string;
  orderType: string;
  status: string;
  security?: string | null;
  quantity?: number | null;
  price?: string | null;
}

function fmt(n: number) {
  return Math.abs(n).toLocaleString("en", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function sideTone(side: string) {
  const s = side.toLowerCase();
  if (s.includes("purchase") || s.includes("buy"))
    return "bg-secondary/10 text-secondary";
  if (s.includes("sale") || s.includes("sell"))
    return "bg-axis-error/10 text-axis-error";
  return "bg-primary/5 text-on-surface-variant";
}

function statusTone(status: string) {
  const s = status.toLowerCase();
  if (s === "traded" || s === "settled") return "bg-secondary/10 text-secondary";
  if (s === "held" || s === "pending") return "bg-amber-100 text-amber-700";
  if (s === "canceled" || s === "cancelled" || s === "rejected")
    return "bg-axis-error/10 text-axis-error";
  return "bg-surface-container text-on-surface-variant";
}

export default function DashboardPage() {
  const { user } = useAuth();
  const market = useMarketToday();
  const [balance, setBalance] = useState<Balance | null>(null);
  const [ordersTotal, setOrdersTotal] = useState<number | null>(null);
  const [recent, setRecent] = useState<OrderRow[]>([]);
  const [quotes, setQuotes] = useState<MarketQuoteLite[] | null>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    Promise.all([
      api.get("/portal/balance").then((r) => setBalance(r.data)).catch(() => {}),
      api
        .get("/portal/orders?page=1&pageSize=5")
        .then((r) => {
          setOrdersTotal(r.data.total);
          setRecent(r.data.rows ?? []);
        })
        .catch(() => {}),
      api
        .get("/market-quotes/latest")
        .then((r) => setQuotes(r.data?.items ?? []))
        .catch(() => {}),
    ]).finally(() => setLoading(false));
  }, []);

  const movers = (() => {
    if (!quotes || quotes.length === 0)
      return { gainers: [], losers: [], asAt: null as string | null };
    const withChange = quotes.filter((q) => q.changePct != null);
    const sorted = [...withChange].sort(
      (a, b) => (b.changePct ?? 0) - (a.changePct ?? 0),
    );
    return {
      gainers: sorted.filter((q) => (q.changeAbs ?? 0) > 0).slice(0, 4),
      losers: sorted
        .filter((q) => (q.changeAbs ?? 0) < 0)
        .slice(-4)
        .reverse(),
      asAt: quotes[0]?.quoteDate ?? null,
    };
  })();

  return (
    <div className="space-y-6">
      {/* Portfolio summary bento */}
      <section className="grid grid-cols-12 gap-4 lg:gap-6">
        <div className="col-span-12 grid grid-cols-1 gap-4 md:grid-cols-3 lg:col-span-8 lg:gap-6">
          <StatCard
            label="Account Balance"
            value={`${brand.currency} ${balance ? fmt(balance.balance) : "\u2014"}`}
            sub={
              balance
                ? balance.balance >= 0
                  ? "Credit"
                  : "Debit"
                : "Liquid reserves"
            }
            tone={balance && balance.balance < 0 ? "error" : "secondary"}
            icon="account_balance_wallet"
          />
          <StatCard
            label="Outstanding"
            value={`${brand.currency} ${balance ? fmt(balance.outstanding) : "\u2014"}`}
            sub="Settlement due"
            tone="error"
            icon="warning"
          />
          <StatCard
            label="Credit Limit"
            value={`${brand.currency} ${balance ? fmt(balance.creditLimit) : "\u2014"}`}
            sub="Approved facility"
            tone="muted"
            icon="verified_user"
          />
        </div>

        {/* Trade desk / market status (dark card) */}
        <div className="col-span-12 flex flex-col justify-between overflow-hidden rounded-xl border border-primary-container bg-primary-container p-6 text-on-primary-container lg:col-span-4">
          <div>
            <div className="mb-3 flex items-center justify-between">
              {market.loading ? (
                <span className="rounded bg-white/10 px-2 py-1 text-[11px] font-semibold uppercase tracking-wide">
                  Checking…
                </span>
              ) : market.isOpen ? (
                <span className="rounded bg-secondary-container px-2 py-1 text-[11px] font-bold uppercase tracking-wide text-on-secondary-container">
                  Market Open
                </span>
              ) : (
                <span className="rounded bg-white/10 px-2 py-1 text-[11px] font-bold uppercase tracking-wide text-on-primary-container">
                  Market Closed
                </span>
              )}
              <span className="text-xs text-on-primary-container">
                {new Date().toLocaleTimeString("en-GB", {
                  hour: "2-digit",
                  minute: "2-digit",
                })}
              </span>
            </div>
            <h3 className="mb-1 font-display text-lg font-bold text-white">
              Trade Desk
            </h3>
            <p className="text-sm text-on-primary-container/80">
              {market.loading
                ? "Loading market session…"
                : market.isOpen
                  ? "The MSE is open. Place a buy or sell order with the trading desk."
                  : market.reason === "holiday"
                    ? `Closed${market.holidayName ? ` — ${market.holidayName}` : ""}.${market.nextOpen ? ` Reopens ${new Date(market.nextOpen).toLocaleDateString("en-GB", { day: "2-digit", month: "short" })}.` : ""}`
                    : "The market is currently closed. You can still queue an order."}
            </p>
          </div>
          <Link
            to="/orders/new"
            className="mt-6 flex items-center justify-center gap-2 rounded-lg bg-secondary-fixed py-3 font-bold text-on-secondary-fixed transition-transform hover:brightness-110 active:scale-95"
          >
            Place Order
            <Icon name="chevron_right" />
          </Link>
        </div>
      </section>

      {/* Market movers + recent activity */}
      <section className="grid grid-cols-12 gap-4 lg:gap-6">
        {/* Market watch */}
        <div className="col-span-12 flex flex-col overflow-hidden rounded-xl border border-outline-variant bg-surface-container-lowest shadow-[0px_4px_12px_rgba(15,23,42,0.03)] xl:col-span-5">
          <div className="flex items-center justify-between border-b border-outline-variant px-5 py-4">
            <h3 className="font-display text-base font-semibold text-primary">
              Market Watch
            </h3>
            <Link
              to="/market-prices"
              className="text-xs font-semibold text-secondary hover:underline"
            >
              View all
            </Link>
          </div>
          {loading ? (
            <div className="p-6 text-sm text-on-surface-variant">Loading…</div>
          ) : !quotes || quotes.length === 0 ? (
            <div className="p-6 text-sm text-on-surface-variant">
              No market data available.
            </div>
          ) : (
            <div className="grid grid-cols-2 divide-x divide-outline-variant">
              <MoverList title="Gainers" rows={movers.gainers} up />
              <MoverList title="Losers" rows={movers.losers} up={false} />
            </div>
          )}
          {movers.asAt && (
            <p className="border-t border-outline-variant px-5 py-2 text-[11px] text-on-surface-variant">
              As at{" "}
              {new Date(movers.asAt).toLocaleDateString("en-GB", {
                day: "2-digit",
                month: "long",
                year: "numeric",
              })}{" "}
              · Indicative only
            </p>
          )}
        </div>

        {/* Recent activity */}
        <div className="col-span-12 flex flex-col overflow-hidden rounded-xl border border-outline-variant bg-surface-container-lowest shadow-[0px_4px_12px_rgba(15,23,42,0.03)] xl:col-span-7">
          <div className="flex items-center justify-between border-b border-outline-variant px-5 py-4">
            <h3 className="font-display text-base font-semibold text-primary">
              Recent Activity
            </h3>
            <Link
              to="/orders"
              className="text-xs font-semibold text-secondary hover:underline"
            >
              View all orders
            </Link>
          </div>
          <div className="overflow-x-auto">
            {loading ? (
              <div className="p-6 text-sm text-on-surface-variant">Loading…</div>
            ) : recent.length === 0 ? (
              <div className="p-6 text-sm text-on-surface-variant">
                No recent orders. Place your first order from the Trade Desk.
              </div>
            ) : (
              <table className="w-full text-left text-sm">
                <thead>
                  <tr className="border-b border-outline-variant text-[11px] uppercase tracking-wider text-on-surface-variant">
                    <th className="px-5 py-3 font-semibold">Date</th>
                    <th className="px-5 py-3 font-semibold">Security</th>
                    <th className="px-5 py-3 font-semibold">Side</th>
                    <th className="px-5 py-3 text-right font-semibold">Qty</th>
                    <th className="px-5 py-3 font-semibold">Status</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-outline-variant">
                  {recent.map((o) => (
                    <tr
                      key={o.orderDpa}
                      className="transition-colors hover:bg-surface-container-low"
                    >
                      <td className="px-5 py-3 text-on-surface-variant">
                        {new Date(o.orderDate).toLocaleDateString("en-GB", {
                          day: "2-digit",
                          month: "short",
                          year: "numeric",
                        })}
                      </td>
                      <td className="px-5 py-3 font-medium text-on-surface">
                        {o.security ?? "—"}
                        {o.orderRef && (
                          <span className="ml-2 text-[11px] text-outline">
                            {o.orderRef}
                          </span>
                        )}
                      </td>
                      <td className="px-5 py-3">
                        <span
                          className={`rounded-full px-2 py-0.5 text-[10px] font-semibold uppercase ${sideTone(o.orderType)}`}
                        >
                          {o.orderType}
                        </span>
                      </td>
                      <td className="px-5 py-3 text-right tabular-nums text-on-surface">
                        {o.quantity != null ? o.quantity.toLocaleString() : "—"}
                      </td>
                      <td className="px-5 py-3">
                        <span
                          className={`rounded-full px-2 py-0.5 text-[10px] font-semibold uppercase ${statusTone(o.status)}`}
                        >
                          {o.status}
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
          {ordersTotal != null && ordersTotal > recent.length && (
            <p className="border-t border-outline-variant px-5 py-2 text-[11px] text-on-surface-variant">
              Showing latest {recent.length} of {ordersTotal} orders
            </p>
          )}
        </div>
      </section>

      {/* Signed-in footer line */}
      <p className="text-xs text-on-surface-variant">
        Signed in as{" "}
        <span className="font-medium text-on-surface">
          {user?.firstName} {user?.lastName}
        </span>{" "}
        ·{" "}
        {new Date().toLocaleDateString("en-GB", {
          weekday: "long",
          day: "2-digit",
          month: "long",
          year: "numeric",
        })}
      </p>
    </div>
  );
}

function StatCard({
  label,
  value,
  sub,
  tone,
  icon,
}: {
  label: string;
  value: string;
  sub: string;
  tone: "secondary" | "error" | "muted";
  icon: string;
}) {
  const toneCls =
    tone === "secondary"
      ? "text-secondary"
      : tone === "error"
        ? "text-axis-error"
        : "text-on-surface-variant";
  return (
    <div className="group rounded-xl border border-outline-variant bg-surface-container-lowest p-5 shadow-[0px_4px_12px_rgba(15,23,42,0.03)] transition-colors hover:border-secondary">
      <div className="mb-2 flex items-start justify-between">
        <span className="text-[11px] font-semibold uppercase tracking-wider text-on-surface-variant">
          {label}
        </span>
        <Icon
          name={icon}
          size={20}
          className={`opacity-0 transition-opacity group-hover:opacity-100 ${toneCls}`}
        />
      </div>
      <h2
        className="truncate font-display text-2xl font-semibold text-primary"
        title={value}
      >
        {value}
      </h2>
      <p className={`mt-1 text-xs ${toneCls}`}>{sub}</p>
    </div>
  );
}

function MoverList({
  title,
  rows,
  up,
}: {
  title: string;
  rows: MarketQuoteLite[];
  up: boolean;
}) {
  const tone = up ? "text-secondary" : "text-axis-error";
  return (
    <div className="min-w-0">
      <div className="border-b border-outline-variant bg-surface-container-low px-4 py-2 text-[10px] font-bold uppercase tracking-widest text-on-surface-variant">
        {title}
      </div>
      <ul className="divide-y divide-outline-variant/60">
        {rows.length === 0 ? (
          <li className="px-4 py-6 text-center text-xs text-outline">None</li>
        ) : (
          rows.map((q) => (
            <li
              key={q.securityDpa}
              className="flex items-center justify-between gap-2 px-4 py-2.5"
            >
              <div className="min-w-0">
                <p className="truncate text-sm font-bold text-primary">
                  {q.securityCode}
                </p>
                <p className="truncate text-[10px] text-on-surface-variant/70">
                  {q.securityName}
                </p>
              </div>
              <div className="text-right">
                <p className="text-xs font-semibold tabular-nums text-on-surface">
                  {(q.close ?? 0).toLocaleString("en", {
                    minimumFractionDigits: 2,
                    maximumFractionDigits: 2,
                  })}
                </p>
                <p className={`text-[11px] font-bold tabular-nums ${tone}`}>
                  {up ? "▲" : "▼"} {Math.abs(q.changePct ?? 0).toFixed(2)}%
                </p>
              </div>
            </li>
          ))
        )}
      </ul>
    </div>
  );
}
