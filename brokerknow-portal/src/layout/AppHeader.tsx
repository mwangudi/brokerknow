import { useEffect, useRef } from "react";

import { Link } from "react-router";
import { useSidebar } from "../context/SidebarContext";
import { ThemeToggleButton } from "../components/common/ThemeToggleButton";
import NotificationDropdown from "../components/header/NotificationDropdown";
import UserDropdown from "../components/header/UserDropdown";
import Icon from "../components/ui/Icon";

const AppHeader: React.FC = () => {
  const { toggleMobileSidebar } = useSidebar();
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    const handleKeyDown = (event: KeyboardEvent) => {
      if ((event.metaKey || event.ctrlKey) && event.key === "k") {
        event.preventDefault();
        inputRef.current?.focus();
      }
    };
    document.addEventListener("keydown", handleKeyDown);
    return () => document.removeEventListener("keydown", handleKeyDown);
  }, []);

  return (
    <header className="sticky top-0 z-40 flex h-16 w-full items-center justify-between border-b border-outline-variant bg-surface px-4 lg:px-6">
      <div className="flex flex-1 items-center gap-3">
        {/* Mobile sidebar toggle */}
        <button
          className="flex h-10 w-10 items-center justify-center rounded-lg text-on-surface-variant transition-colors hover:bg-surface-container lg:hidden"
          onClick={toggleMobileSidebar}
          aria-label="Toggle menu"
        >
          <Icon name="menu" />
        </button>

        <Link to="/" className="font-display text-xl font-bold tracking-tight text-primary lg:hidden">
          Axis
        </Link>

        {/* Search */}
        <div className="relative hidden w-full max-w-md lg:block">
          <Icon
            name="search"
            size={20}
            className="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2 text-outline"
          />
          <input
            ref={inputRef}
            type="text"
            placeholder="Search instruments, tickers, or markets…"
            className="w-full rounded-full border border-outline-variant bg-surface-container-low py-2 pl-10 pr-4 text-sm text-on-surface placeholder:text-outline focus:border-secondary focus:outline-none focus:ring-1 focus:ring-secondary"
          />
        </div>
      </div>

      <div className="flex items-center gap-2 sm:gap-3">
        <NotificationDropdown />
        <ThemeToggleButton />
        <div className="mx-1 hidden h-8 w-px bg-outline-variant sm:block" />
        <Link
          to="/orders/new"
          className="hidden rounded-lg bg-primary px-4 py-2 text-xs font-semibold uppercase tracking-wide text-on-primary transition-opacity hover:opacity-90 sm:inline-flex"
        >
          Place Order
        </Link>
        <UserDropdown />
      </div>
    </header>
  );
};

export default AppHeader;
