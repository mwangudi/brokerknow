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
    <div className="flex min-h-screen flex-col bg-background">
      {/* Non-production builds (VITE_ENV_LABEL) show a warning ribbon so a test
          login is unmistakable. Renders nothing in production. */}
      {brand.envLabel && (
        <div
          role="status"
          className="w-full bg-amber-400 px-3 py-1 text-center text-[11px] font-bold uppercase tracking-[0.18em] text-amber-950"
        >
          ⚠ {brand.envLabel} environment — data here is not live
        </div>
      )}
      <div className="flex flex-1">
      {/* Brand panel — TEAL wash differentiates the CLIENT portal from the
          staff back-office (which uses a green wash). Same layout, different
          shade, so users can tell the two sign-ins apart at a glance. */}
      <div className="relative hidden w-1/2 flex-col justify-between overflow-hidden p-12 lg:flex">
        {/* Base teal wash — also the standalone background when no photo is set */}
        <div className="absolute inset-0 bg-gradient-to-br from-[#042f2e] via-[#0a4f48] to-[#0f766e]" />
        {/* Optional hero photo, tinted teal so white text stays legible */}
        {brand.loginPhoto && (
          <>
            <img
              src={brand.loginPhoto}
              alt=""
              className="absolute inset-0 h-full w-full object-cover"
            />
            <div className="absolute inset-0 bg-gradient-to-br from-[#042f2e]/92 via-[#0a4f48]/85 to-[#0f766e]/80" />
          </>
        )}
        <div className="pointer-events-none absolute -right-20 -top-20 h-72 w-72 rounded-full bg-[#5eead4]/15" />
        <div className="pointer-events-none absolute -bottom-24 -left-16 h-64 w-64 rounded-full bg-[#5eead4]/10" />

        {/* Product wordmark */}
        <div className="relative">
          <span className="font-display text-2xl font-bold tracking-tight text-white">
            Axis
          </span>
          <p className="mt-1 text-[11px] font-medium uppercase tracking-[0.2em] text-white/70">
            Client Portal
          </p>
        </div>

        {/* Headline */}
        <div className="relative max-w-md">
          <h1 className="font-display text-4xl font-bold leading-tight text-white">
            {brand.loginHeadline}
          </h1>
          <p className="mt-4 text-sm font-semibold uppercase tracking-[0.12em] text-[#99f6e4]">
            {brand.loginServices}
          </p>
        </div>

        {/* Tagline foot (vendor lockup now lives on the white form side where
            the dark Martens logo reads naturally). */}
        <p className="relative text-[11px] text-white/60">
          © {new Date().getFullYear()} · Secure client access
        </p>
      </div>

      {/* Form panel */}
      <div className="flex w-full flex-col items-center justify-center px-4 py-10 lg:w-1/2">
        <div className="w-full max-w-lg">
          {/* Client brand, above the Sign in card */}
          <div className="mb-8 flex flex-col items-center text-center">
            <img
              src={brand.logo}
              alt={brand.name}
              className="h-24 w-auto object-contain"
            />
            <p className="mt-2 text-xs font-semibold uppercase tracking-[0.15em] text-on-surface-variant">
              Client Portal
            </p>
          </div>

          <form
            onSubmit={handleSubmit}
            className="rounded-2xl border border-outline-variant bg-surface-container-lowest p-10 shadow-[0px_4px_12px_rgba(15,23,42,0.03)]"
          >
            <h2 className="font-display text-2xl font-semibold text-primary">
              Sign in
            </h2>
            <p className="mb-7 mt-1 text-sm text-on-surface-variant">
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
                className="w-full rounded-lg border border-outline-variant bg-surface-container-lowest px-4 py-3 text-base text-on-surface focus:border-secondary focus:outline-none focus:ring-1 focus:ring-secondary"
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
                className="w-full rounded-lg border border-outline-variant bg-surface-container-lowest px-4 py-3 text-base text-on-surface focus:border-secondary focus:outline-none focus:ring-1 focus:ring-secondary"
                placeholder="••••••••"
              />
            </div>

            <button
              type="submit"
              disabled={loading}
              className="w-full rounded-lg bg-primary py-3 text-base font-semibold uppercase tracking-wide text-on-primary transition-opacity hover:opacity-90 disabled:opacity-50"
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

          {/* Vendor lockup — Built by Martens Africa. On the white panel where
              the dark navy/red logo reads naturally (no white chip). */}
          <div className="mt-8 flex items-center justify-center gap-3">
            <img
              src="/images/logo/martens-logo.png"
              alt="Martens Africa"
              className="h-7 w-auto object-contain"
            />
            <div className="border-l border-outline-variant pl-3 text-[11px] leading-tight text-on-surface-variant">
              <p className="font-semibold text-on-surface">Built by Martens Africa</p>
              <p>© {new Date().getFullYear()} · All rights reserved</p>
            </div>
          </div>
        </div>
      </div>
      </div>
    </div>
  );
}
