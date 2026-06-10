import { ICON_CODEPOINTS } from "./material-symbols-codepoints";

type IconProps = {
  /** Material Symbols Outlined name, e.g. "dashboard", "swap_horiz". */
  name: string;
  /** Render the filled variant. */
  filled?: boolean;
  /** Pixel size (defaults to inherited font-size). */
  size?: number;
  className?: string;
};

/**
 * Renders a Material Symbols Outlined glyph from the SELF-HOSTED subset
 * (src/fonts/material-symbols-subset.woff2 — same-origin, ~14 KB, no Google
 * CDN). We render the glyph by its codepoint rather than the ligature name so
 * the font can be subset to just the icons in use; an unknown name renders
 * nothing instead of leaking the literal text. To add an icon, add it to
 * scripts/subset-icons.py, re-run it, and rebuild (see scripts/README.md).
 */
export default function Icon({
  name,
  filled = false,
  size,
  className = "",
}: IconProps) {
  const glyph = ICON_CODEPOINTS[name] ?? "";
  return (
    <span
      className={`material-symbols-outlined ${className}`}
      aria-hidden="true"
      style={{
        fontSize: size ? `${size}px` : undefined,
        fontVariationSettings: `'FILL' ${filled ? 1 : 0}`,
      }}
    >
      {glyph}
    </span>
  );
}
