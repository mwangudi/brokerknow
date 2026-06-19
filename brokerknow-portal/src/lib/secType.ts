/**
 * Maps the legacy order security-type code to a human-friendly label.
 * The database stores OrderSecTypeDescription as "F" (fixed income) or
 * "S" (shares/stock); clients should see "Bonds" / "Equity".
 */
export function secTypeLabel(code: string | null | undefined): string {
  const c = (code ?? "").trim().toUpperCase();
  if (c === "F") return "Bonds";
  if (c === "S") return "Equity";
  return code ?? "";
}
