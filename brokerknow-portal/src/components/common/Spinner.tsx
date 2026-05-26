import type { ReactNode } from "react";

interface SpinnerProps {
  /** Tailwind size classes — defaults to small (h-5 w-5). */
  size?: "sm" | "md" | "lg";
  /** Color of the spinner ring (defaults to brand). */
  tone?: "brand" | "muted" | "white";
  className?: string;
  /** Aria label */
  label?: string;
}

const SIZES: Record<NonNullable<SpinnerProps["size"]>, string> = {
  sm: "h-5 w-5 border-2",
  md: "h-8 w-8 border-[3px]",
  lg: "h-12 w-12 border-4",
};

const TONES: Record<NonNullable<SpinnerProps["tone"]>, string> = {
  brand: "border-brand-200 border-t-brand-500 dark:border-brand-500/20 dark:border-t-brand-400",
  muted: "border-gray-200 border-t-gray-500 dark:border-gray-700 dark:border-t-gray-300",
  white: "border-white/20 border-t-white",
};

export default function Spinner({
  size = "sm",
  tone = "brand",
  className = "",
  label = "Loading",
}: SpinnerProps) {
  return (
    <span
      role="status"
      aria-label={label}
      className={`inline-block animate-spin rounded-full ${SIZES[size]} ${TONES[tone]} ${className}`}
    />
  );
}

interface InlineLoaderProps {
  /** Optional text shown next to the spinner. */
  children?: ReactNode;
  size?: SpinnerProps["size"];
  tone?: SpinnerProps["tone"];
}

/** Convenience: spinner + label, centered. Use inside loading panels. */
export function InlineLoader({
  children = "Loading...",
  size = "md",
  tone = "muted",
}: InlineLoaderProps) {
  return (
    <div className="flex items-center justify-center gap-3 text-sm text-gray-500 dark:text-gray-400">
      <Spinner size={size} tone={tone} />
      <span>{children}</span>
    </div>
  );
}
