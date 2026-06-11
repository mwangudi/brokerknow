import { Link, useLocation } from "react-router";
import { useSidebar } from "../context/SidebarContext";
import { useAuth } from "../context/AuthContext";
import Icon from "../components/ui/Icon";
import { brand } from "../lib/brand";

type NavItem = { name: string; icon: string; path: string };

const navItems: NavItem[] = [
  { name: "Dashboard", icon: "dashboard", path: "/" },
  { name: "Orders", icon: "swap_horiz", path: "/orders" },
  { name: "Statement", icon: "receipt_long", path: "/statement" },
  { name: "Payments", icon: "payments", path: "/request-payment" },
  { name: "Profile", icon: "person", path: "/profile" },
];

const AppSidebar: React.FC = () => {
  const { isMobileOpen, toggleMobileSidebar } = useSidebar();
  const { user, logout } = useAuth();
  const location = useLocation();

  const isActive = (path: string) =>
    path === "/"
      ? location.pathname === "/"
      : location.pathname.startsWith(path);

  const closeOnMobile = () => {
    if (isMobileOpen) toggleMobileSidebar();
  };

  const initials =
    `${user?.firstName?.[0] ?? ""}${user?.lastName?.[0] ?? ""}`.toUpperCase() ||
    "AX";
  const fullName = user
    ? `${user.firstName} ${user.lastName}`.trim()
    : "Account";

  return (
    <>
      <aside
        className={`fixed left-0 top-0 z-50 flex h-screen w-64 flex-col border-r border-outline-variant bg-surface-container-low transition-transform duration-300 ease-in-out lg:translate-x-0 ${
          isMobileOpen ? "translate-x-0" : "-translate-x-full"
        }`}
      >
        {/* Brand — client logo + name, mirrors the admin back office */}
        <div className="px-6 pb-5 pt-6">
          <Link to="/" onClick={closeOnMobile} className="flex items-center gap-3">
            <span className="flex h-10 w-10 shrink-0 items-center justify-center rounded-lg bg-white p-1 ring-1 ring-outline-variant">
              <img
                src={brand.logo}
                alt={brand.name}
                className="h-full w-full object-contain"
              />
            </span>
            <span className="min-w-0 leading-tight">
              <span className="block font-display text-base font-bold tracking-tight text-primary">
                {brand.name}
              </span>
              <span className="block text-xs text-on-surface-variant opacity-70">
                {brand.tagline}
              </span>
            </span>
          </Link>
        </div>

        {/* New Trade */}
        <div className="px-4 pb-4">
          <Link
            to="/orders/new"
            onClick={closeOnMobile}
            className="flex items-center justify-center gap-2 rounded-lg bg-secondary px-4 py-2.5 text-xs font-semibold uppercase tracking-wide text-on-secondary transition-colors hover:bg-on-secondary-fixed-variant"
          >
            <Icon name="add" size={18} />
            New Trade
          </Link>
        </div>

        {/* Primary nav */}
        <nav className="custom-scrollbar flex-1 space-y-1 overflow-y-auto px-3 py-2">
          {navItems.map((item) => {
            const active = isActive(item.path);
            return (
              <Link
                key={item.name}
                to={item.path}
                onClick={closeOnMobile}
                className={`group flex items-center gap-3 rounded-lg px-3 py-2.5 text-sm transition-colors ${
                  active
                    ? "border-r-2 border-secondary bg-surface-container font-bold text-secondary"
                    : "text-on-surface-variant hover:bg-surface-container hover:text-on-surface"
                }`}
              >
                <Icon name={item.icon} filled={active} />
                <span>{item.name}</span>
              </Link>
            );
          })}
        </nav>

        {/* Footer: Settings / Support + user block */}
        <div className="border-t border-outline-variant px-3 py-3">
          <Link
            to="/change-password"
            onClick={closeOnMobile}
            className="flex items-center gap-3 rounded-lg px-3 py-2 text-sm text-on-surface-variant transition-colors hover:bg-surface-container"
          >
            <Icon name="settings" size={20} />
            <span>Settings</span>
          </Link>
          <a
            href={`mailto:${brand.supportEmail}`}
            className="flex items-center gap-3 rounded-lg px-3 py-2 text-sm text-on-surface-variant transition-colors hover:bg-surface-container"
          >
            <Icon name="help" size={20} />
            <span>Support</span>
          </a>

          <div className="mt-2 flex items-center gap-3 border-t border-outline-variant/60 px-3 pt-3">
            <div className="flex h-9 w-9 shrink-0 items-center justify-center rounded-full bg-primary-container text-xs font-bold text-on-primary-fixed">
              {initials}
            </div>
            <div className="min-w-0 flex-1">
              <p className="truncate text-sm font-semibold text-on-surface">
                {fullName}
              </p>
              <p className="truncate text-[10px] uppercase tracking-wide text-on-surface-variant">
                {user?.role || "Client"}
              </p>
            </div>
            <button
              type="button"
              onClick={logout}
              aria-label="Sign out"
              className="rounded-lg p-1.5 text-on-surface-variant transition-colors hover:bg-surface-container hover:text-axis-error"
            >
              <Icon name="logout" size={20} />
            </button>
          </div>
        </div>

        {/* Vendor attribution — Axis by Martens Africa (matches the admin back office) */}
        <a
          href="https://www.martensafrica.com/"
          target="_blank"
          rel="noopener noreferrer"
          aria-label="Axis by Martens Africa"
          className="flex items-center gap-2 border-t border-outline-variant px-6 py-3 transition-colors hover:bg-surface-container"
        >
          <span className="shrink-0 text-[10px] font-medium uppercase tracking-wide text-on-surface-variant opacity-75">
            Axis by
          </span>
          <span className="inline-flex rounded dark:bg-white/95 dark:p-1">
            <img
              src="/images/logo/martens-logo.png"
              alt="Martens Africa"
              className="h-5 w-auto object-contain"
            />
          </span>
        </a>
      </aside>

      {/* Mobile backdrop */}
      {isMobileOpen && (
        <button
          aria-label="Close menu"
          onClick={toggleMobileSidebar}
          className="fixed inset-0 z-40 bg-inverse-surface/40 backdrop-blur-sm lg:hidden"
        />
      )}
    </>
  );
};

export default AppSidebar;
