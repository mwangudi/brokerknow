import { useEffect, useState } from "react";
import api from "../lib/api";

interface HolidayDate {
  holidayDate: string;
  description: string;
  recurring: boolean;
}

export interface MarketDayInfo {
  isOpen: boolean;
  reason: "weekend" | "holiday" | null;
  holidayName: string | null;
  /** Next open business day as ISO yyyy-mm-dd (or null while loading). */
  nextOpen: string | null;
  /** True while the holiday list is loading. */
  loading: boolean;
}

function toISODate(d: Date): string {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

/**
 * Tells the caller whether the market is open today. Portal orders are
 * always implicitly dated "today", so this is enough for the heads-up banner
 * on the Place Order page.
 */
export function useMarketToday(): MarketDayInfo {
  const [holidays, setHolidays] = useState<HolidayDate[] | null>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    let alive = true;
    const year = new Date().getFullYear();
    Promise.all([
      api.get<HolidayDate[]>(`/holidays/dates?year=${year}`),
      api.get<HolidayDate[]>(`/holidays/dates?year=${year + 1}`),
    ])
      .then(([a, b]) => {
        if (alive) setHolidays([...(a.data ?? []), ...(b.data ?? [])]);
      })
      .catch(() => {
        if (alive) setHolidays([]);
      })
      .finally(() => {
        if (alive) setLoading(false);
      });
    return () => {
      alive = false;
    };
  }, []);

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const dow = today.getDay();
  const todayISO = toISODate(today);

  if (loading) {
    return { isOpen: true, reason: null, holidayName: null, nextOpen: null, loading: true };
  }

  function findHoliday(d: Date): string | null {
    if (!holidays) return null;
    const iso = toISODate(d);
    const mmdd = `${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
    for (const h of holidays) {
      const dt = new Date(h.holidayDate);
      const hISO = toISODate(dt);
      const hMmdd = `${String(dt.getMonth() + 1).padStart(2, "0")}-${String(dt.getDate()).padStart(2, "0")}`;
      if (hISO === iso) return h.description;
      if (h.recurring && hMmdd === mmdd) return h.description;
    }
    return null;
  }

  const holidayName = findHoliday(today);
  const isWeekend = dow === 0 || dow === 6;

  if (!isWeekend && !holidayName) {
    return { isOpen: true, reason: null, holidayName: null, nextOpen: todayISO, loading: false };
  }

  // Walk forward to find the next open day.
  let nextOpen: string | null = null;
  const cursor = new Date(today);
  for (let i = 0; i < 14; i++) {
    cursor.setDate(cursor.getDate() + 1);
    const cDow = cursor.getDay();
    if (cDow === 0 || cDow === 6) continue;
    if (findHoliday(cursor)) continue;
    nextOpen = toISODate(cursor);
    break;
  }

  return {
    isOpen: false,
    reason: isWeekend ? "weekend" : "holiday",
    holidayName: isWeekend ? null : holidayName,
    nextOpen,
    loading: false,
  };
}
