/**
 * White-label brand configuration.
 *
 * All values default to Cedar Capital (Malawi) so the existing production
 * build is unchanged. A second tenant (e.g. the DRC demo) is produced purely
 * by overriding these VITE_* vars at build time in an .env file — no code
 * changes per client. See .env.production / .env.drc.
 */
export const brand = {
  /** Client/company display name shown in the sidebar, login and copy. */
  name: (import.meta.env.VITE_BRAND_NAME as string) || "Cedar Capital",
  /** Small caption under the brand name (e.g. "Institutional Portal"). */
  tagline: (import.meta.env.VITE_BRAND_TAGLINE as string) || "Institutional Portal",
  /** Path to the client logo (served from /public/images/logo/...). */
  logo: (import.meta.env.VITE_BRAND_LOGO as string) || "/images/logo/cedar-logo.png",
  /** Support contact surfaced in the sidebar. */
  supportEmail:
    (import.meta.env.VITE_SUPPORT_EMAIL as string) || "support@cedarcapital.mw",
  /** Currency code/prefix shown before monetary amounts (e.g. MWK, FC, USD). */
  currency: (import.meta.env.VITE_CURRENCY as string) || "MWK",
} as const;
