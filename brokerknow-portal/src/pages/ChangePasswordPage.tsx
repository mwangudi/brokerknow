import { useState, type FormEvent } from "react";
import { useNavigate, Navigate } from "react-router";
import api from "../lib/api";
import { useAuth } from "../context/AuthContext";

/**
 * Forced password-change page. Reached either:
 *   - immediately after the first login with a temporary password, or
 *   - after the monthly rotation policy expires the existing one.
 *
 * The user is already authenticated when they get here (they have a valid
 * access token), so we just hit POST /auth/change-password.
 */
export default function ChangePasswordPage() {
  const { user, token, markPasswordChanged, deferPasswordChange } = useAuth();
  const navigate = useNavigate();
  const [current, setCurrent] = useState("");
  const [next, setNext] = useState("");
  const [confirm, setConfirm] = useState("");
  const [error, setError] = useState("");
  const [saving, setSaving] = useState(false);

  // Hard-block access when there is no session at all.
  if (!token || !user) return <Navigate to="/login" replace />;

  const isForced = !!user.mustChangePassword;

  async function handleSubmit(e: FormEvent) {
    e.preventDefault();
    setError("");
    if (next !== confirm) { setError("New passwords do not match."); return; }
    if (next.length < 8) { setError("Password must be at least 8 characters."); return; }
    if (!isForced && !current) { setError("Please enter your current password."); return; }

    setSaving(true);
    try {
      await api.post("/auth/change-password", {
        currentPassword: isForced ? "" : current,
        newPassword: next,
      });
      markPasswordChanged();
      navigate("/", { replace: true });
    } catch (err: any) {
      setError(err.response?.data?.error || "Failed to change password.");
    } finally {
      setSaving(false);
    }
  }

  const skipsRaw = user.passwordChangeSkipsRemaining;
  // A session created before this field existed won't carry it (undefined). Treat
  // "unknown" as skips-available so the button still shows on older sessions — the
  // server still enforces the 3-deferral cap on /auth/defer-password-change; only
  // an explicit 0 (a fresh login that used every skip) hides it.
  const skipsKnown = typeof skipsRaw === "number";
  const canSkip = !skipsKnown || skipsRaw > 0;

  // Escape hatch: defer a forced change ("Not now"). The server counts the
  // deferrals and refuses past the cap (3), so it can't be skipped forever.
  // On success we clear the local flag and continue; reminded again next login.
  async function handleSkip() {
    setError("");
    setSaving(true);
    try {
      await deferPasswordChange();
      navigate("/", { replace: true });
    } catch (err: any) {
      setError(err.response?.data?.error || "You must set a new password now.");
    } finally {
      setSaving(false);
    }
  }

  const reason = isForced
    ? canSkip
      ? "Your password is temporary or has expired. Set a new one to continue."
      : "Your password has expired and must be changed now \u2014 no more deferrals."
    : "Pick a new password.";

  return (
    <div className="flex min-h-screen items-center justify-center bg-gradient-to-br from-gray-50 to-brand-25 px-4 py-12">
      <div className="w-full max-w-md">
        <form onSubmit={handleSubmit} className="rounded-2xl border border-gray-200 bg-white p-8 shadow-sm">
          <h2 className="mb-1 text-xl font-semibold text-gray-900">Change your password</h2>
          <p className="mb-6 text-sm text-gray-500">{reason}</p>

          {error && (
            <div className="mb-4 rounded-lg bg-red-50 p-3 text-sm text-red-700">{error}</div>
          )}

          <Field label="Current password" value={current} onChange={setCurrent} autoComplete="current-password" hidden={isForced} />
          <Field label="New password" value={next} onChange={setNext} autoComplete="new-password" />
          <Field label="Confirm new password" value={confirm} onChange={setConfirm} autoComplete="new-password" />

          <ul className="mb-6 list-inside list-disc space-y-0.5 text-xs text-gray-500">
            <li>At least 8 characters</li>
            <li>Mix of upper- and lower-case letters</li>
            <li>At least one digit</li>
            <li>Different from your current password</li>
          </ul>

          <button
            type="submit"
            disabled={saving}
            className="w-full rounded-lg bg-brand-600 py-2.5 text-sm font-semibold text-white hover:bg-brand-700 disabled:opacity-50"
          >
            {saving ? "Updating..." : "Update password"}
          </button>

          {isForced ? (
            canSkip && (
              <button
                type="button"
                onClick={handleSkip}
                disabled={saving}
                className="mt-3 w-full rounded-lg border border-gray-300 py-2.5 text-sm font-medium text-gray-600 hover:bg-gray-50 disabled:opacity-50"
              >
                Not now — continue{skipsKnown ? ` (${skipsRaw} ${skipsRaw === 1 ? "skip" : "skips"} left)` : ""}
              </button>
            )
          ) : (
            <button
              type="button"
              onClick={() => navigate("/", { replace: true })}
              className="mt-3 w-full rounded-lg border border-gray-300 py-2.5 text-sm font-medium text-gray-600 hover:bg-gray-50"
            >
              Cancel
            </button>
          )}
        </form>
      </div>
    </div>
  );
}

function Field({
  label,
  value,
  onChange,
  autoComplete,
  hidden,
}: {
  label: string;
  value: string;
  onChange: (v: string) => void;
  autoComplete: string;
  hidden?: boolean;
}) {
  const [show, setShow] = useState(false);
  if (hidden) return null;
  return (
    <div className="mb-4">
      <label className="mb-1.5 block text-sm font-medium text-gray-700">{label}</label>
      <div className="relative">
        <input
          type={show ? "text" : "password"}
          required
          value={value}
          autoComplete={autoComplete}
          onChange={(e) => onChange(e.target.value)}
          className="w-full rounded-lg border border-gray-300 px-4 py-2.5 pr-11 text-sm focus:border-brand-500 focus:ring-2 focus:ring-brand-500/20 focus:outline-none"
          placeholder="••••••••"
        />
        <button
          type="button"
          onClick={() => setShow((s) => !s)}
          aria-label={show ? "Hide password" : "Show password"}
          title={show ? "Hide password" : "Show password"}
          className="absolute inset-y-0 right-0 flex items-center pr-3 text-gray-400 hover:text-gray-600 focus:outline-none"
        >
          <PasswordEye show={show} />
        </button>
      </div>
    </div>
  );
}

/** Eye / eye-off icon used by the password fields' show-hide toggle. */
function PasswordEye({ show }: { show: boolean }) {
  return show ? (
    <svg className="h-5 w-5" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor">
      <path strokeLinecap="round" strokeLinejoin="round" d="M3.98 8.223A10.477 10.477 0 0 0 1.934 12C3.226 16.338 7.244 19.5 12 19.5c.993 0 1.953-.138 2.863-.395M6.228 6.228A10.451 10.451 0 0 1 12 4.5c4.756 0 8.774 3.162 10.066 7.498a10.523 10.523 0 0 1-4.293 5.774M6.228 6.228 3 3m3.228 3.228 3.65 3.65m7.894 7.894L21 21m-3.228-3.228-3.65-3.65m0 0a3 3 0 1 0-4.243-4.243m4.242 4.242L9.88 9.88" />
    </svg>
  ) : (
    <svg className="h-5 w-5" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor">
      <path strokeLinecap="round" strokeLinejoin="round" d="M2.036 12.322a1.012 1.012 0 0 1 0-.639C3.423 7.51 7.36 4.5 12 4.5c4.638 0 8.573 3.007 9.963 7.178.07.207.07.431 0 .639C20.577 16.49 16.64 19.5 12 19.5c-4.638 0-8.573-3.007-9.963-7.178Z" />
      <path strokeLinecap="round" strokeLinejoin="round" d="M15 12a3 3 0 1 1-6 0 3 3 0 0 1 6 0Z" />
    </svg>
  );
}
