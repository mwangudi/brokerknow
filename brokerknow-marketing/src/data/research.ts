export interface ResearchDoc {
  title: string;
  url: string;
  date?: string;
}

// Pulled from cedarcapital.mw/research/ — June 2026.
// Sorted newest first.
export const RESEARCH: ResearchDoc[] = [
  { title: "1st Quarter Market Performance Report — 2025", date: "2025-05", url: "http://cedarcapital.mw/content/uploads/2025/05/1ST-QUARTER-MARKET-PERFORMANCE-REPORT-2025-1.pdf" },
  { title: "Listed Banks FY24 Results",                    date: "2025-04", url: "http://cedarcapital.mw/content/uploads/2025/04/Listed-Banks-FY24-Results.pdf" },
  { title: "Cedar Capital Trading Updates — March 2022",   date: "2022-06", url: "http://cedarcapital.mw/content/uploads/2022/06/Cedar-Capital-Trading-Updates-March-2022.pdf" },
  { title: "Illovo Sugar Malawi Report — January 2022",    date: "2022-02", url: "http://cedarcapital.mw/content/uploads/2022/02/Illovo-Sugar-Malawi-Report-January-2022.pdf" },
  { title: "Annual Market Performance Report — 2021",      date: "2022-01", url: "http://cedarcapital.mw/content/uploads/2022/01/ANNUAL-MARKET-PERFORMANCE-REPORT-2021.pdf" },
];
