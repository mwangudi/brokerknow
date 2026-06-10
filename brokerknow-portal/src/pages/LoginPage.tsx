import { useState, type FormEvent } from "react";
import { useNavigate, Link } from "react-router";
import { useAuth } from "../context/AuthContext";

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
      {/* Brand panel */}
      <div className="relative hidden w-1/2 flex-col justify-between overflow-hidden bg-primary-container p-12 lg:flex">
        <div className="pointer-events-none absolute -right-16 -top-16 h-64 w-64 rounded-full bg-secondary/10" />
        <div className="relative">
          <span className="inline-flex rounded bg-white p-2 ring-1 ring-white/20">
            <img
              src="/images/logo/martens-logo.png"
              alt="Martens Africa"
              className="h-7 w-auto object-contain"
            />
          </span>
        </div>
        <div className="relative">
          <h1 className="font-display text-4xl font-bold tracking-tight text-white">
            Axis
          </h1>
          <p className="mt-2 max-w-sm text-on-primary-container">
            The institutional trading portal by Martens Africa. Monitor your
            portfolio, place orders, and track settlement — all in one place.
          </p>
        </div>
        <p className="relative text-xs text-on-primary-container/60">
          © {new Date().getFullYear()} Martens Africa · Powered by Axis
        </p>
      </div>

      {/* Form panel */}
      <div className="flex w-full items-center justify-center px-4 lg:w-1/2">
        <div className="w-full max-w-md">
          <div className="mb-8 lg:hidden">
            <span className="inline-flex rounded bg-white p-1.5 ring-1 ring-outline-variant">
              <img
                src="/images/logo/martens-logo.png"
                alt="Martens Africa"
                className="h-6 w-auto object-contain"
              />
            </span>
            <h1 className="mt-3 font-display text-2xl font-bold text-primary">
              Axis
            </h1>
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
        </div>
      </div>
    </div>
  );
}
