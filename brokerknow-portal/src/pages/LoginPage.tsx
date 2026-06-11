import { useState, type FormEvent } from "react";
import { useNavigate, Link } from "react-router";
import { useAuth } from "../context/AuthContext";
import { brand } from "../lib/brand";

export default function LoginPage() {
  const { login, loading } = useAuth();
  const navigate = useNavigate();
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");

  async function handleSubmit(e: FormEvent) {
    e.preventDefault();
    setError("");
    const result = await login(email, password);
    if (result.error) {
      setError(result.error);
    } else if (result.requiresPasswordChange) {
      // Temp password / monthly rotation — send straight to the change form.
      navigate("/change-password", { replace: true });
    } else {
      navigate("/", { replace: true });
    }
  }

  return (
    <div className="flex min-h-screen bg-background">
      {/* Brand panel — Axis showcase + Martens vendor lockup */}
      <div className="relative hidden w-1/2 flex-col justify-between overflow-hidden bg-primary-container p-12 lg:flex">
        <div className="pointer-events-none absolute -right-20 -top-20 h-72 w-72 rounded-full bg-secondary/10" />
        <div className="pointer-events-none absolute -bottom-24 -left-16 h-64 w-64 rounded-full bg-secondary/5" />

        {/* Product wordmark */}
        <div className="relative">
          <span className="font-display text-2xl font-bold tracking-tight text-white">
            Axis
          </span>
          <p className="mt-1 text-[11px] font-medium uppercase tracking-[0.2em] text-on-primary-container/60">
            Institutional Portal
          </p>
        </div>

        {/* Headline */}
        <div className="relative max-w-md">
          <h1 className="font-display text-4xl font-bold leading-tight text-white">
            Trade with institutional confidence.
          </h1>
          <p className="mt-3 text-on-primary-container/80">
            Monitor your portfolio, place orders, and track settlement — all in
            one secure place.
          </p>
        </div>

        {/* Vendor lockup — Built by Martens Africa */}
        <div className="relative flex items-center gap-3">
          <span className="inline-flex items-center justify-center rounded-lg bg-white px-3 py-2 shadow-sm">
            <img
              src="/images/logo/martens-logo.png"
              alt="Martens Africa"
              className="h-6 w-auto object-contain"
            />
          </span>
          <div className="text-[11px] leading-tight text-on-primary-container/70">
            <p className="font-semibold text-on-primary-container">
              Built by Martens Africa
            </p>
            <p>© {new Date().getFullYear()} · All rights reserved</p>
          </div>
        </div>
      </div>

      {/* Form panel */}
      <div className="flex w-full flex-col items-center justify-center px-4 py-10 lg:w-1/2">
        <div className="w-full max-w-md">
          {/* Client brand, above the Sign in card */}
          <div className="mb-6 flex flex-col items-center text-center">
            <img
              src={brand.logo}
              alt={brand.name}
              className="h-20 w-auto object-contain"
            />
            <p className="mt-2 text-[11px] font-semibold uppercase tracking-[0.15em] text-on-surface-variant">
              Client Portal
            </p>
          </div>

          <form
            onSubmit={handleSubmit}
            className="rounded-xl border border-outline-variant bg-surface-container-lowest p-8 shadow-[0px_4px_12px_rgba(15,23,42,0.03)]"
          >
            <h2 className="font-display text-xl font-semibold text-primary">
              Sign in
            </h2>
            <p className="mb-6 mt-1 text-sm text-on-surface-variant">
              Access your institutional portal.
            </p>

            {error && (
              <div className="mb-4 rounded-lg border border-axis-error/30 bg-axis-error/5 p-3 text-sm text-axis-error">
                {error}
              </div>
            )}

            <div className="mb-4">
              <label className="mb-1.5 block text-[11px] font-semibold uppercase tracking-wide text-on-surface-variant">
                Email
              </label>
              <input
                type="email"
                required
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                className="w-full rounded-lg border border-outline-variant bg-surface-container-lowest px-4 py-2.5 text-sm text-on-surface focus:border-secondary focus:outline-none focus:ring-1 focus:ring-secondary"
                placeholder="you@example.com"
              />
            </div>

            <div className="mb-6">
              <label className="mb-1.5 block text-[11px] font-semibold uppercase tracking-wide text-on-surface-variant">
                Password
              </label>
              <input
                type="password"
                required
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                className="w-full rounded-lg border border-outline-variant bg-surface-container-lowest px-4 py-2.5 text-sm text-on-surface focus:border-secondary focus:outline-none focus:ring-1 focus:ring-secondary"
                placeholder="••••••••"
              />
            </div>

            <button
              type="submit"
              disabled={loading}
              className="w-full rounded-lg bg-primary py-2.5 text-sm font-semibold uppercase tracking-wide text-on-primary transition-opacity hover:opacity-90 disabled:opacity-50"
            >
              {loading ? "Signing in…" : "Sign in"}
            </button>

            <p className="mt-4 text-center text-sm text-on-surface-variant">
              Don't have an account?{" "}
              <Link
                to="/register"
                className="font-semibold text-secondary hover:underline"
              >
                Register here
              </Link>
            </p>
          </form>

          {/* Vendor line — visible on mobile where the brand panel is hidden */}
          <p className="mt-6 flex items-center justify-center gap-1.5 text-[11px] text-on-surface-variant lg:hidden">
            Powered by <span className="font-semibold text-primary">Axis</span> ·
            Built by Martens Africa
          </p>
        </div>
      </div>
    </div>
  );
}
