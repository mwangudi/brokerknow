import { useState, type FormEvent } from "react";
import { useNavigate, Link } from "react-router";
import { useAuth } from "../context/AuthContext";
import { brand } from "../lib/brand";

export default function LoginPage() {
  const { login, verifyOtp, forgotPassword, resetPassword, loading } = useAuth();
  const navigate = useNavigate();
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [step, setStep] = useState<"credentials" | "otp" | "forgot-email" | "forgot-reset">("credentials");
  const [otpToken, setOtpToken] = useState("");
  const [otpCode, setOtpCode] = useState("");
  const [otpHint, setOtpHint] = useState("");

  // Forgot-password flow state.
  const [resetToken, setResetToken] = useState("");
  const [resetCode, setResetCode] = useState("");
  const [newPassword, setNewPassword] = useState("");
  const [confirmPassword, setConfirmPassword] = useState("");
  const [notice, setNotice] = useState("");

  async function handleForgotRequest(e: FormEvent) {
    e.preventDefault();
    setError("");
    setNotice("");
    if (!email.trim()) {
      setError("Enter your email address.");
      return;
    }
    const result = await forgotPassword(email.trim());
    if (result.error) {
      setError(result.error);
      return;
    }
    setResetToken(result.token || "");
    setNotice(result.message || "If that email is registered, we've sent a reset code.");
    setResetCode("");
    setNewPassword("");
    setConfirmPassword("");
    setStep("forgot-reset");
  }

  async function handleResetSubmit(e: FormEvent) {
    e.preventDefault();
    setError("");
    if (!resetCode.trim()) {
      setError("Enter the 6-digit code from your email.");
      return;
    }
    if (newPassword !== confirmPassword) {
      setError("The two passwords don't match.");
      return;
    }
    const result = await resetPassword(resetToken, resetCode.trim(), newPassword);
    if (result.error) {
      setError(result.error);
      return;
    }
    setNotice(result.message || "Your password has been reset. Please sign in.");
    setStep("credentials");
    setPassword("");
  }

  async function handleSubmit(e: FormEvent) {
    e.preventDefault();
    setError("");
    const result = await login(email, password);
    if (result.error) {
      setError(result.error);
    } else if (result.otpRequired) {
      setOtpToken(result.otpToken || "");
      setOtpHint(result.otpHint || "your email");
      setOtpCode("");
      setStep("otp");
    } else if (result.requiresPasswordChange) {
      navigate("/change-password", { replace: true });
    } else {
      navigate("/", { replace: true });
    }
  }

  async function handleVerifyOtp(e: FormEvent) {
    e.preventDefault();
    if (!otpCode.trim()) {
      setError("Enter the code we sent you.");
      return;
    }
    setError("");
    const result = await verifyOtp(otpToken, otpCode.trim());
    if (result.error) {
      setError(result.error);
    } else if (result.requiresPasswordChange) {
      navigate("/change-password", { replace: true });
    } else {
      navigate("/", { replace: true });
    }
  }

  return (
    <div className="flex min-h-screen bg-background">
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

        {/* Headline — split on the em-dash so the brand title sits on two
            clean lines (e.g. "Cedar Capital" / "Malawi's Digital Broker"). */}
        <div className="relative max-w-md">
          <h1 className="font-display text-4xl font-bold leading-tight text-white">
            {brand.loginHeadline.split(/\s*\u2014\s*/).map((line, i) => (
              <span key={i} className="block">
                {line}
              </span>
            ))}
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

      {/* Form panel — the TEST ribbon sits at the top of THIS (white) side only */}
      <div className="flex w-full flex-col lg:w-1/2">
        {brand.envLabel && (
          <div
            role="status"
            className="w-full bg-amber-400 px-3 py-1 text-center text-[11px] font-bold uppercase tracking-[0.18em] text-amber-950"
          >
            ⚠ {brand.envLabel} environment — data here is not live
          </div>
        )}
        <div className="flex flex-1 flex-col items-center justify-center px-4 py-10">
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

          {step === "otp" && (
          <form
            onSubmit={handleVerifyOtp}
            className="rounded-2xl border border-outline-variant bg-surface-container-lowest p-10 shadow-[0px_4px_12px_rgba(15,23,42,0.03)]"
          >
            <h2 className="font-display text-2xl font-semibold text-primary">
              Verify it&apos;s you
            </h2>
            <p className="mb-7 mt-1 text-sm text-on-surface-variant">
              We sent a one-time code to{" "}
              <span className="font-medium text-on-surface">{otpHint}</span>. Enter
              it below to finish signing in.
            </p>

            {error && (
              <div className="mb-4 rounded-lg border border-axis-error/30 bg-axis-error/5 p-3 text-sm text-axis-error">
                {error}
              </div>
            )}

            <div className="mb-6">
              <label className="mb-1.5 block text-[11px] font-semibold uppercase tracking-wide text-on-surface-variant">
                Verification code
              </label>
              <input
                type="text"
                inputMode="numeric"
                autoComplete="one-time-code"
                value={otpCode}
                onChange={(e) => setOtpCode(e.target.value.replace(/[^0-9]/g, "").slice(0, 6))}
                className="w-full rounded-lg border border-outline-variant bg-surface-container-lowest px-4 py-3 text-center text-2xl tracking-[0.5em] text-on-surface focus:border-secondary focus:outline-none focus:ring-1 focus:ring-secondary"
                placeholder="••••••"
                autoFocus
              />
            </div>

            <button
              type="submit"
              disabled={loading}
              className="w-full rounded-lg bg-primary py-3 text-base font-semibold uppercase tracking-wide text-on-primary transition-opacity hover:opacity-90 disabled:opacity-50"
            >
              {loading ? "Verifying…" : "Verify & sign in"}
            </button>
            <button
              type="button"
              onClick={() => { setStep("credentials"); setOtpCode(""); setError(""); }}
              className="mt-3 w-full text-center text-sm text-on-surface-variant hover:underline"
            >
              Back to sign in
            </button>
          </form>
          )}

          {step === "forgot-email" && (
          <form
            onSubmit={handleForgotRequest}
            className="rounded-2xl border border-outline-variant bg-surface-container-lowest p-10 shadow-[0px_4px_12px_rgba(15,23,42,0.03)]"
          >
            <h2 className="font-display text-2xl font-semibold text-primary">
              Reset your password
            </h2>
            <p className="mb-7 mt-1 text-sm text-on-surface-variant">
              Enter your email and we&apos;ll send you a 6-digit code to reset your password.
            </p>
            {error && (
              <div className="mb-4 rounded-lg border border-axis-error/30 bg-axis-error/5 p-3 text-sm text-axis-error">
                {error}
              </div>
            )}
            <div className="mb-6">
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
                autoFocus
              />
            </div>
            <button
              type="submit"
              disabled={loading}
              className="w-full rounded-lg bg-primary py-3 text-base font-semibold uppercase tracking-wide text-on-primary transition-opacity hover:opacity-90 disabled:opacity-50"
            >
              {loading ? "Sending…" : "Send reset code"}
            </button>
            <button
              type="button"
              onClick={() => { setStep("credentials"); setError(""); setNotice(""); }}
              className="mt-3 w-full text-center text-sm text-on-surface-variant hover:underline"
            >
              Back to sign in
            </button>
          </form>
          )}

          {step === "forgot-reset" && (
          <form
            onSubmit={handleResetSubmit}
            className="rounded-2xl border border-outline-variant bg-surface-container-lowest p-10 shadow-[0px_4px_12px_rgba(15,23,42,0.03)]"
          >
            <h2 className="font-display text-2xl font-semibold text-primary">
              Enter your reset code
            </h2>
            <p className="mb-5 mt-1 text-sm text-on-surface-variant">
              {notice || "Check your email for a 6-digit code, then choose a new password."}
            </p>
            {error && (
              <div className="mb-4 rounded-lg border border-axis-error/30 bg-axis-error/5 p-3 text-sm text-axis-error">
                {error}
              </div>
            )}
            <div className="mb-4">
              <label className="mb-1.5 block text-[11px] font-semibold uppercase tracking-wide text-on-surface-variant">
                Reset code
              </label>
              <input
                type="text"
                inputMode="numeric"
                autoComplete="one-time-code"
                value={resetCode}
                onChange={(e) => setResetCode(e.target.value.replace(/[^0-9]/g, "").slice(0, 6))}
                className="w-full rounded-lg border border-outline-variant bg-surface-container-lowest px-4 py-3 text-center text-2xl tracking-[0.5em] text-on-surface focus:border-secondary focus:outline-none focus:ring-1 focus:ring-secondary"
                placeholder="••••••"
                autoFocus
              />
            </div>
            <div className="mb-4">
              <label className="mb-1.5 block text-[11px] font-semibold uppercase tracking-wide text-on-surface-variant">
                New password
              </label>
              <input
                type="password"
                required
                value={newPassword}
                onChange={(e) => setNewPassword(e.target.value)}
                className="w-full rounded-lg border border-outline-variant bg-surface-container-lowest px-4 py-3 text-base text-on-surface focus:border-secondary focus:outline-none focus:ring-1 focus:ring-secondary"
                placeholder="••••••••"
              />
            </div>
            <div className="mb-2">
              <label className="mb-1.5 block text-[11px] font-semibold uppercase tracking-wide text-on-surface-variant">
                Confirm new password
              </label>
              <input
                type="password"
                required
                value={confirmPassword}
                onChange={(e) => setConfirmPassword(e.target.value)}
                className="w-full rounded-lg border border-outline-variant bg-surface-container-lowest px-4 py-3 text-base text-on-surface focus:border-secondary focus:outline-none focus:ring-1 focus:ring-secondary"
                placeholder="••••••••"
              />
            </div>
            <p className="mb-6 text-[11px] text-on-surface-variant">
              At least 8 characters with an uppercase letter, a lowercase letter and a digit.
            </p>
            <button
              type="submit"
              disabled={loading}
              className="w-full rounded-lg bg-primary py-3 text-base font-semibold uppercase tracking-wide text-on-primary transition-opacity hover:opacity-90 disabled:opacity-50"
            >
              {loading ? "Resetting…" : "Reset password"}
            </button>
            <button
              type="button"
              onClick={() => { setStep("forgot-email"); setError(""); }}
              className="mt-3 w-full text-center text-sm text-on-surface-variant hover:underline"
            >
              Back
            </button>
          </form>
          )}

          {step === "credentials" && (
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

            {notice && (
              <div className="mb-4 rounded-lg border border-secondary/30 bg-secondary/5 p-3 text-sm text-secondary">
                {notice}
              </div>
            )}

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
              <div className="mt-1.5 text-right">
                <button
                  type="button"
                  onClick={() => { setStep("forgot-email"); setError(""); setNotice(""); }}
                  className="text-sm font-medium text-secondary hover:underline"
                >
                  Forgot password?
                </button>
              </div>
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
          )}

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
