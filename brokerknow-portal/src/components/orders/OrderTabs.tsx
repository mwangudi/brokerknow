import { Link } from "react-router";

/**
 * The "Place Order / View Orders" segmented toggle used on the Orders and
 * Place Order pages. Navigates between /orders/new and /orders.
 */
export default function OrderTabs({ active }: { active: "place" | "view" }) {
  return (
    <div className="inline-flex rounded-lg border border-outline-variant bg-surface-container p-1">
      <Link
        to="/orders/new"
        className={`rounded-md px-5 py-2 text-xs font-semibold uppercase tracking-wide transition-all ${
          active === "place"
            ? "bg-surface-container-lowest text-primary shadow-sm"
            : "text-on-surface-variant hover:text-primary"
        }`}
      >
        Place Order
      </Link>
      <Link
        to="/orders"
        className={`rounded-md px-5 py-2 text-xs font-semibold uppercase tracking-wide transition-all ${
          active === "view"
            ? "bg-surface-container-lowest text-primary shadow-sm"
            : "text-on-surface-variant hover:text-primary"
        }`}
      >
        View Orders
      </Link>
    </div>
  );
}
