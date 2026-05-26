import { useEffect, useState } from "react";
import { Link } from "react-router";
import api from "../lib/api";
import { useAuth } from "../context/AuthContext";

interface Balance {
  openingBalance: number;
  balance: number;
  outstanding: number;
  creditLimit: number;
}

interface OrdersSummary {
  total: number;
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

function fmt(n: number) {
  return Math.abs(n).toLocaleString("en", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

export default function DashboardPage() {
  const { user } = useAuth();
  const [balance, setBalance] = useState<Balance | null>(null);
  const [orders, setOrders] = useState<OrdersSummary | null>(null);
  const [quotes, setQuotes] = useState<MarketQuoteLite[] | null>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    Promise.all([
      api.get("/portal/balance").then((r) => setBalance(r.data)).catch(() => {}),
      api.get("/portal/orders?page=1&pageSize=1").then((r) => setOrders({ total: r.data.total })).catch(() => {}),
      api.get("/market-quotes/latest").then((r) => setQuotes(r.data?.items ?? [])).catch(() => {}),
    ]).finally(() => setLoading(false));
  }, []);

  const greeting = (() => {
    const h = new Date().getHours();
    if (h < 12) return "Good morning";
    if (h < 17) return "Good afternoon";
    return "Good evening";
  })();

  return (
    <div>
      {/* Hero card with gradient */}
      <div className="mb-6 overflow-hidden rounded-2xl bg-gradient-to-r from-blue-600 to-indigo-700 p-6 text-white shadow-lg">
        <div className="flex items-center justify-between gap-4">
          <div>
            <p className="text-sm text-blue-100">{greeting},</p>
            <h2 className="mt-1 text-2xl font-bold">{user?.firstName} {user?.lastName}</h2>
            <p className="mt-2 text-sm text-blue-100">
              {new Date().toLocaleDateString("en-GB", { weekday: "long", day: "2-digit", month: "long", year: "numeric" })}
            </p>
          </div>
          <div className="hidden md:block">
            {/* Decorative chart icon */}
            <svg className="h-20 w-20 text-blue-300/40" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M3 13.125C3 12.504 3.504 12 4.125 12h2.25c.621 0 1.125.504 1.125 1.125v6.75C7.5 20.496 6.996 21 6.375 21h-2.25A1.125 1.125 0 013 19.875v-6.75zM9.75 8.625c0-.621.504-1.125 1.125-1.125h2.25c.621 0 1.125.504 1.125 1.125v11.25c0 .621-.504 1.125-1.125 1.125h-2.25a1.125 1.125 0 01-1.125-1.125V8.625zM16.5 4.125c0-.621.504-1.125 1.125-1.125h2.25C20.496 3 21 3.504 21 4.125v15.75c0 .621-.504 1.125-1.125 1.125h-2.25a1.125 1.125 0 01-1.125-1.125V4.125z" />
            </svg>
          </div>
        </div>
      </div>

      {loading ? (
        <p className="text-gray-500">Loading...</p>
      ) : (
        <>
          {/* Stat cards */}
          <div className="mb-6 grid grid-cols-1 gap-5 sm:grid-cols-2 lg:grid-cols-4">
            <StatCard
              label="Account Balance"
              value={`MWK ${balance ? fmt(balance.balance) : "—"}`}
              suffix={balance ? (balance.balance >= 0 ? "Cr" : "Dr") : ""}
              accent={balance && balance.balance >= 0 ? "green" : "red"}
              icon={
                <svg className="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 8c-1.657 0-3 .895-3 2s1.343 2 3 2 3 .895 3 2-1.343 2-3 2m0-8c1.11 0 2.08.402 2.599 1M12 8V7m0 1v8m0 0v1m0-1c-1.11 0-2.08-.402-2.599-1M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
                </svg>
              }
            />
            <StatCard
              label="Outstanding"
              value={`MWK ${balance ? fmt(balance.outstanding) : "—"}`}
              accent="amber"
              icon={
                <svg className="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                </svg>
              }
            />
            <StatCard
              label="Credit Limit"
              value={`MWK ${balance ? fmt(balance.creditLimit) : "—"}`}
              accent="purple"
              icon={
                <svg className="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M3 10h18M7 15h1m4 0h1m-7 4h12a3 3 0 003-3V8a3 3 0 00-3-3H6a3 3 0 00-3 3v8a3 3 0 003 3z" />
                </svg>
              }
            />
            <StatCard
              label="My Orders"
              value={orders ? orders.total.toString() : "—"}
              suffix="trades"
              accent="blue"
              icon={
                <svg className="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-6 9l2 2 4-4" />
                </svg>
              }
            />
          </div>

          {/* Today's market prices — eye-catching movers strip */}
          {quotes && quotes.length > 0 && (() => {
            const sortedByChange = [...quotes].filter(q => q.changePct != null).sort((a, b) => (b.changePct ?? 0) - (a.changePct ?? 0));
            const gainers = sortedByChange.filter(q => (q.changeAbs ?? 0) > 0).slice(0, 3);
            const losers  = sortedByChange.filter(q => (q.changeAbs ?? 0) < 0).slice(-3).reverse();
            const mostActive = [...quotes].sort((a, b) => (b.volume ?? 0) - (a.volume ?? 0)).slice(0, 3);
            const asAtLabel = new Date(quotes[0].quoteDate).toLocaleDateString("en-GB", { day: "2-digit", month: "long", year: "numeric" });

            return (
              <div className="mb-6">
                <div className="mb-3 flex items-end justify-between">
                  <div>
                    <h3 className="text-lg font-semibold text-gray-900">Today's Market</h3>
                    <p className="text-xs text-gray-500">As at <strong>{asAtLabel}</strong> · Indicative only</p>
                  </div>
                  <Link to="/market-prices" className="text-xs font-semibold text-blue-600 hover:underline">
                    View all {quotes.length} prices →
                  </Link>
                </div>

                <div className="grid grid-cols-1 gap-4 lg:grid-cols-3">
                  {/* Gainers */}
                  <div className="overflow-hidden rounded-2xl border border-emerald-200 bg-gradient-to-br from-emerald-50 to-white shadow-sm">
                    <div className="flex items-center justify-between border-b border-emerald-100 bg-emerald-100/60 px-4 py-2.5">
                      <div className="flex items-center gap-2">
                        <span className="flex h-7 w-7 items-center justify-center rounded-full bg-emerald-500 text-white">
                          <svg className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2.5} d="M13 7h8m0 0v8m0-8l-8 8-4-4-6 6" /></svg>
                        </span>
                        <h4 className="text-sm font-semibold text-emerald-900">Top Gainers</h4>
                      </div>
                      <span className="text-[10px] font-medium uppercase tracking-wider text-emerald-700">{gainers.length} up</span>
                    </div>
                    <ul className="divide-y divide-emerald-100/70">
                      {gainers.length === 0 ? (
                        <li className="px-4 py-6 text-center text-xs text-gray-400">No gainers today</li>
                      ) : gainers.map(q => (
                        <li key={q.securityDpa} className="flex items-center justify-between px-4 py-2.5">
                          <div className="min-w-0">
                            <p className="text-sm font-bold text-gray-900">{q.securityCode}</p>
                            <p className="truncate text-[11px] text-gray-500">{q.securityName}</p>
                          </div>
                          <div className="text-right">
                            <p className="text-sm font-semibold tabular-nums text-gray-900">{(q.close ?? 0).toLocaleString("en", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</p>
                            <p className="text-xs font-bold tabular-nums text-emerald-600">▲ {Math.abs(q.changePct ?? 0).toFixed(2)}%</p>
                          </div>
                        </li>
                      ))}
                    </ul>
                  </div>

                  {/* Losers */}
                  <div className="overflow-hidden rounded-2xl border border-red-200 bg-gradient-to-br from-red-50 to-white shadow-sm">
                    <div className="flex items-center justify-between border-b border-red-100 bg-red-100/60 px-4 py-2.5">
                      <div className="flex items-center gap-2">
                        <span className="flex h-7 w-7 items-center justify-center rounded-full bg-red-500 text-white">
                          <svg className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2.5} d="M13 17h8m0 0v-8m0 8l-8-8-4 4-6-6" /></svg>
                        </span>
                        <h4 className="text-sm font-semibold text-red-900">Top Losers</h4>
                      </div>
                      <span className="text-[10px] font-medium uppercase tracking-wider text-red-700">{losers.length} down</span>
                    </div>
                    <ul className="divide-y divide-red-100/70">
                      {losers.length === 0 ? (
                        <li className="px-4 py-6 text-center text-xs text-gray-400">No losers today</li>
                      ) : losers.map(q => (
                        <li key={q.securityDpa} className="flex items-center justify-between px-4 py-2.5">
                          <div className="min-w-0">
                            <p className="text-sm font-bold text-gray-900">{q.securityCode}</p>
                            <p className="truncate text-[11px] text-gray-500">{q.securityName}</p>
                          </div>
                          <div className="text-right">
                            <p className="text-sm font-semibold tabular-nums text-gray-900">{(q.close ?? 0).toLocaleString("en", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</p>
                            <p className="text-xs font-bold tabular-nums text-red-600">▼ {Math.abs(q.changePct ?? 0).toFixed(2)}%</p>
                          </div>
                        </li>
                      ))}
                    </ul>
                  </div>

                  {/* Most active */}
                  <div className="overflow-hidden rounded-2xl border border-blue-200 bg-gradient-to-br from-blue-50 to-white shadow-sm">
                    <div className="flex items-center justify-between border-b border-blue-100 bg-blue-100/60 px-4 py-2.5">
                      <div className="flex items-center gap-2">
                        <span className="flex h-7 w-7 items-center justify-center rounded-full bg-blue-500 text-white">
                          <svg className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2.5} d="M13 10V3L4 14h7v7l9-11h-7z" /></svg>
                        </span>
                        <h4 className="text-sm font-semibold text-blue-900">Most Active</h4>
                      </div>
                      <span className="text-[10px] font-medium uppercase tracking-wider text-blue-700">by volume</span>
                    </div>
                    <ul className="divide-y divide-blue-100/70">
                      {mostActive.map(q => {
                        const ch = q.changeAbs ?? 0;
                        const colour = ch > 0 ? "text-emerald-600" : ch < 0 ? "text-red-600" : "text-gray-500";
                        const arrow = ch > 0 ? "▲" : ch < 0 ? "▼" : "•";
                        return (
                          <li key={q.securityDpa} className="flex items-center justify-between px-4 py-2.5">
                            <div className="min-w-0">
                              <p className="text-sm font-bold text-gray-900">{q.securityCode}</p>
                              <p className="truncate text-[11px] text-gray-500">Vol: {(q.volume ?? 0).toLocaleString()}</p>
                            </div>
                            <div className="text-right">
                              <p className="text-sm font-semibold tabular-nums text-gray-900">{(q.close ?? 0).toLocaleString("en", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</p>
                              <p className={`text-xs font-bold tabular-nums ${colour}`}>
                                {q.changePct == null ? "—" : <>{arrow} {Math.abs(q.changePct).toFixed(2)}%</>}
                              </p>
                            </div>
                          </li>
                        );
                      })}
                    </ul>
                  </div>
                </div>
              </div>
            );
          })()}

          {/* Quick actions */}
          <h3 className="mb-4 text-lg font-semibold text-gray-900">Quick Actions</h3>
          <div className="grid grid-cols-1 gap-5 sm:grid-cols-2 lg:grid-cols-4">
            <ActionCard
              to="/market-prices"
              title="Market Prices"
              desc="Live MSE board prices, gainers and losers"
              gradient="from-amber-500 to-orange-500"
              icon={
                <svg className="h-7 w-7" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M3 3v18h18M7 14l3-3 4 4 5-5" />
                </svg>
              }
            />
            <ActionCard
              to="/statement"
              title="View Statement"
              desc="Download or view your full transaction history"
              gradient="from-blue-500 to-cyan-500"
              icon={
                <svg className="h-7 w-7" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 17v-2m3 2v-4m3 4v-6m2 10H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                </svg>
              }
            />
            <ActionCard
              to="/orders"
              title="My Orders"
              desc="Track your buy and sell orders"
              gradient="from-emerald-500 to-teal-500"
              icon={
                <svg className="h-7 w-7" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13 7h8m0 0v8m0-8l-8 8-4-4-6 6" />
                </svg>
              }
            />
            <ActionCard
              to="/profile"
              title="My Profile"
              desc="View and manage your personal details"
              gradient="from-purple-500 to-pink-500"
              icon={
                <svg className="h-7 w-7" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M16 7a4 4 0 11-8 0 4 4 0 018 0zM12 14a7 7 0 00-7 7h14a7 7 0 00-7-7z" />
                </svg>
              }
            />
          </div>
        </>
      )}
    </div>
  );
}

const ACCENT_CLASSES: Record<string, { iconBg: string; iconText: string; valueText: string }> = {
  green: { iconBg: "bg-green-100", iconText: "text-green-600", valueText: "text-green-700" },
  red: { iconBg: "bg-red-100", iconText: "text-red-600", valueText: "text-red-700" },
  amber: { iconBg: "bg-amber-100", iconText: "text-amber-600", valueText: "text-gray-900" },
  purple: { iconBg: "bg-purple-100", iconText: "text-purple-600", valueText: "text-gray-900" },
  blue: { iconBg: "bg-blue-100", iconText: "text-blue-600", valueText: "text-gray-900" },
};

function StatCard({
  label,
  value,
  suffix,
  accent = "blue",
  icon,
}: {
  label: string;
  value: string;
  suffix?: string;
  accent?: "green" | "red" | "amber" | "purple" | "blue";
  icon: React.ReactNode;
}) {
  const c = ACCENT_CLASSES[accent];
  return (
    <div className="rounded-2xl border border-gray-200 bg-white p-5 shadow-sm transition-shadow hover:shadow-md">
      <div className="flex items-start justify-between">
        <div className="min-w-0">
          <p className="text-xs font-medium uppercase tracking-wider text-gray-500">{label}</p>
          <p className={`mt-2 truncate text-xl font-bold ${c.valueText}`} title={value}>
            {value}
            {suffix && <span className="ml-1.5 text-sm font-medium text-gray-500">{suffix}</span>}
          </p>
        </div>
        <div className={`flex h-11 w-11 shrink-0 items-center justify-center rounded-lg ${c.iconBg} ${c.iconText}`}>
          {icon}
        </div>
      </div>
    </div>
  );
}

function ActionCard({
  to, title, desc, icon, gradient,
}: {
  to: string;
  title: string;
  desc: string;
  icon: React.ReactNode;
  gradient: string;
}) {
  return (
    <Link
      to={to}
      className="group relative overflow-hidden rounded-2xl border border-gray-200 bg-white p-5 shadow-sm transition-all hover:shadow-md hover:-translate-y-0.5"
    >
      <div className={`mb-3 inline-flex h-12 w-12 items-center justify-center rounded-xl bg-gradient-to-br ${gradient} text-white shadow-sm`}>
        {icon}
      </div>
      <h4 className="font-semibold text-gray-900">{title}</h4>
      <p className="mt-1 text-sm text-gray-500">{desc}</p>
      <div className="mt-3 flex items-center text-sm font-medium text-blue-600">
        Open
        <svg className="ml-1 h-4 w-4 transition-transform group-hover:translate-x-0.5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 5l7 7-7 7" />
        </svg>
      </div>
    </Link>
  );
}
