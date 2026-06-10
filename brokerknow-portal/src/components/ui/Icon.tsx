type IconProps = {
  /** Material Symbols Outlined ligature name, e.g. "dashboard", "swap_horiz". */
  name: string;
  /** Render the filled variant. */
  filled?: boolean;
  /** Pixel size (defaults to inherited font-size). */
  size?: number;
  weight?: number;
  className?: string;
};

/**
 * Thin wrapper over the Material Symbols Outlined icon font used by the Axis
 * design language. Loaded once from index.html.
 */
export default function Icon({
  name,
  filled = false,
  size,
  weight = 400,
  className = "",
}: IconProps) {
  return (
    <span
      className={`material-symbols-outlined ${className}`}
      aria-hidden="true"
      style={{
        fontSize: size ? `${size}px` : undefined,
        fontVariationSettings: `'FILL' ${filled ? 1 : 0}, 'wght' ${weight}, 'GRAD' 0, 'opsz' 24`,
      }}
    >
      {name}
    </span>
  );
}
