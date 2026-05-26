import { useMemo, useState, type FormEvent, type ReactNode } from "react";
import { Link } from "react-router";
import { useAuth } from "../context/AuthContext";
import DatePicker from "../components/form/DatePicker";

// ─── Types ──────────────────────────────────────────────────────────
type ClientKind = "existing" | "new" | null;

interface FormState {
  // Branch flag
  clientKind: ClientKind;

  // Personal
  firstName: string;
  lastName: string;
  idNumber: string;
  cdsNumber: string;
  dateOfBirth: string;

  // Contact
  email: string;
  phone: string;
  officePhone: string;
  homePhone: string;

  // Address & next of kin
  postalAddress: string;
  physicalAddress: string;
  contactPerson: string;
}

type Errors = Partial<Record<keyof FormState, string>>;

const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const PHONE_RE = /^[+\d\s()-]{7,20}$/;

const EMPTY: FormState = {
  clientKind: null,
  firstName: "",
  lastName: "",
  idNumber: "",
  cdsNumber: "",
  dateOfBirth: "",
  email: "",
  phone: "",
  officePhone: "",
  homePhone: "",
  postalAddress: "",
  physicalAddress: "",
  contactPerson: "",
};

// ─── Per-step validation ────────────────────────────────────────────
function validateStep(step: number, form: FormState): Errors {
  const e: Errors = {};
  const isExisting = form.clientKind === "existing";

  if (step === 0) {
    // Account-type chooser — handled by button click.
    return e;
  }

  // Step 1 — Personal
  if (step === 1) {
    if (!form.firstName.trim()) e.firstName = "First name is required.";
    else if (form.firstName.trim().length < 2)
      e.firstName = "First name must be at least 2 characters.";

    if (!form.lastName.trim()) e.lastName = "Last name is required.";
    else if (form.lastName.trim().length < 2)
      e.lastName = "Last name must be at least 2 characters.";

    if (isExisting && !form.cdsNumber.trim())
      e.cdsNumber = "Your CDS number is required so we can link your existing account.";

    if (form.dateOfBirth) {
      const dob = new Date(form.dateOfBirth);
      if (isNaN(dob.getTime())) e.dateOfBirth = "Enter a valid date.";
      else {
        const today = new Date();
        const age =
          today.getFullYear() -
          dob.getFullYear() -
          (today.getMonth() < dob.getMonth() ||
          (today.getMonth() === dob.getMonth() &&
            today.getDate() < dob.getDate())
            ? 1
            : 0);
        if (age < 18)
          e.dateOfBirth = "You must be at least 18 years old to register.";
        else if (age > 120)
          e.dateOfBirth = "Please enter a valid date of birth.";
      }
    } else if (!isExisting) {
      e.dateOfBirth = "Date of birth is required.";
    }

    if (!isExisting && !form.idNumber.trim())
      e.idNumber = "ID / Passport number is required.";
  }

  // Step 2 — Contact
  if (step === 2) {
    if (!form.email.trim()) e.email = "Email is required.";
    else if (!EMAIL_RE.test(form.email.trim()))
      e.email = "Enter a valid email address.";

    if (!form.phone.trim()) e.phone = "A cell phone number is required.";
    else if (!PHONE_RE.test(form.phone))
      e.phone = "Enter a valid phone number.";

    if (form.officePhone && !PHONE_RE.test(form.officePhone))
      e.officePhone = "Enter a valid phone number.";
    if (form.homePhone && !PHONE_RE.test(form.homePhone))
      e.homePhone = "Enter a valid phone number.";
  }

  // Step 3 — Addresses (only for new applicants)
  if (step === 3 && !isExisting) {
    if (!form.postalAddress.trim() && !form.physicalAddress.trim())
      e.postalAddress = "Provide at least one address (postal or physical).";
  }

  return e;
}

// ─── Component ──────────────────────────────────────────────────────
export default function RegisterPage() {
  const { register, loading } = useAuth();
  const [form, setForm] = useState<FormState>(EMPTY);
  const [step, setStep] = useState(0);
  const [errors, setErrors] = useState<Errors>({});
  const [submitError, setSubmitError] = useState("");
  const [success, setSuccess] = useState("");

  const isExisting = form.clientKind === "existing";

  // Existing clients have a 3-step flow (chooser → details → review).
  // New clients have a 5-step flow (chooser → personal → contact → address → review).
  const steps = useMemo(
    () =>
      isExisting
        ? ["Account", "Your details", "Review"]
        : ["Account", "Personal", "Contact", "Addresses", "Review"],
    [isExisting],
  );
  const lastStep = steps.length - 1;
  const isReview = step === lastStep;

  function set<K extends keyof FormState>(field: K, value: FormState[K]) {
    setForm((prev) => ({ ...prev, [field]: value }));
    setErrors((prev) => ({ ...prev, [field]: undefined }));
  }

  function chooseKind(kind: ClientKind) {
    setForm({ ...EMPTY, clientKind: kind });
    setErrors({});
    setSubmitError("");
    setStep(1);
  }

  function next() {
    setSubmitError("");
    const stepErrors = validateStep(step, form);
    if (Object.keys(stepErrors).length) {
      setErrors(stepErrors);
      return;
    }
    setStep((s) => Math.min(s + 1, lastStep));
  }

  function back() {
    setSubmitError("");
    setErrors({});
    setStep((s) => Math.max(0, s - 1));
  }

  async function handleSubmit(e: FormEvent) {
    e.preventDefault();
    setSubmitError("");

    // Re-validate all relevant steps before sending.
    const all: Errors = {};
    for (let s = 1; s < lastStep; s++) {
      Object.assign(all, validateStep(s, form));
    }
    if (Object.keys(all).length) {
      setErrors(all);
      setSubmitError(
        "Some required fields are missing. Use Back to fix them and try again.",
      );
      return;
    }

    const result = await register({
      email: form.email,
      firstName: form.firstName,
      lastName: form.lastName,
      phone: form.phone || undefined,
      officePhone: form.officePhone || undefined,
      homePhone: form.homePhone || undefined,
      idNumber: form.idNumber || undefined,
      cdsNumber: form.cdsNumber || undefined,
      dateOfBirth: form.dateOfBirth || undefined,
      physicalAddress: form.physicalAddress || undefined,
      postalAddress: form.postalAddress || undefined,
      contactPerson: form.contactPerson || undefined,
    });

    if (result.error) {
      setSubmitError(result.error);
    } else {
      setSuccess(result.message || "Registration submitted!");
    }
  }

  // ── Success screen ──────────────────────────────────────────────
  if (success) {
    return (
      <Shell>
        <div className="rounded-2xl border border-green-200 bg-white p-8 text-center shadow-sm">
          <div className="mb-4 inline-flex h-16 w-16 items-center justify-center rounded-full bg-green-100">
            <svg className="h-8 w-8 text-green-600" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
            </svg>
          </div>
          <h2 className="mb-2 text-xl font-semibold text-gray-900">Registration Received</h2>
          <p className="mb-6 text-gray-600">{success}</p>
          <p className="mb-6 text-sm text-gray-500">
            {isExisting
              ? "An administrator will verify your CDS number against our records and email you your login credentials."
              : "Our team will review your application and email you your login credentials once approved."}
          </p>
          <Link
            to="/login"
            className="inline-block rounded-lg bg-brand-600 px-6 py-2.5 text-sm font-semibold text-white hover:bg-brand-700"
          >
            Back to Login
          </Link>
        </div>
      </Shell>
    );
  }

  return (
    <Shell>
      <form
        onSubmit={handleSubmit}
        noValidate
        className="rounded-2xl border border-gray-200 bg-white p-6 shadow-sm sm:p-8"
      >
        <Stepper steps={steps} current={step} />

        {submitError && (
          <div className="mb-4 rounded-lg bg-red-50 p-3 text-sm text-red-700">
            {submitError}
          </div>
        )}

        {/* ── Step 0: Account-type chooser ─────────────────────── */}
        {step === 0 && (
          <div>
            <h2 className="mb-1 text-xl font-semibold text-gray-900">
              Welcome — let's get you started
            </h2>
            <p className="mb-6 text-sm text-gray-500">
              Are you an existing CDS account holder, or applying for the first time?
            </p>
            <div className="grid gap-4 sm:grid-cols-2">
              <ChoiceCard
                title="I have an existing CDS account"
                description="You already trade through us and just need an online portal login. We'll verify your CDS number and link this login to your account."
                onClick={() => chooseKind("existing")}
              />
              <ChoiceCard
                title="I'm new — I'd like to open an account"
                description="You don't yet have a CDS account. Complete the full application and we'll onboard you."
                onClick={() => chooseKind("new")}
              />
            </div>
            <p className="mt-6 text-center text-sm text-gray-500">
              Already registered?{" "}
              <Link to="/login" className="font-medium text-brand-600 hover:underline">
                Sign in
              </Link>
            </p>
          </div>
        )}

        {/* ── Step 1: Personal ─────────────────────────────────── */}
        {step === 1 && (
          <div>
            <h2 className="mb-1 text-xl font-semibold text-gray-900">
              {isExisting ? "Tell us who you are" : "Personal information"}
            </h2>
            <p className="mb-6 text-sm text-gray-500">
              {isExisting
                ? "Provide your name and CDS number so our team can locate your account."
                : "Please fill in your personal details."}
            </p>
            <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
              <Field label="First name" required value={form.firstName}
                error={errors.firstName} onChange={(v) => set("firstName", v)}
                placeholder="e.g. John" />
              <Field label="Last name" required value={form.lastName}
                error={errors.lastName} onChange={(v) => set("lastName", v)}
                placeholder="e.g. Banda" />
              <Field
                label="CDS number"
                required={isExisting}
                value={form.cdsNumber}
                error={errors.cdsNumber}
                onChange={(v) => set("cdsNumber", v)}
                placeholder="e.g. CDS-000123"
                hint={isExisting ? "Required so admin can link your portal login." : undefined}
              />
              <Field
                label="ID / Passport number"
                required={!isExisting}
                value={form.idNumber}
                error={errors.idNumber}
                onChange={(v) => set("idNumber", v)}
                placeholder="National ID or Passport no."
              />
              <div>
                <label className="mb-1.5 block text-sm font-medium text-gray-700">
                  Date of birth {!isExisting && <span className="text-red-500">*</span>}
                </label>
                <DatePicker
                  value={form.dateOfBirth}
                  onChange={(v) => set("dateOfBirth", v)}
                  placeholder="Pick date of birth"
                  maxDate={new Date()}
                />
                {errors.dateOfBirth && (
                  <p className="mt-1 text-xs text-red-600">{errors.dateOfBirth}</p>
                )}
              </div>
            </div>
          </div>
        )}

        {/* ── Step 2: Contact ──────────────────────────────────── */}
        {step === 2 && (
          <div>
            <h2 className="mb-1 text-xl font-semibold text-gray-900">Contact details</h2>
            <p className="mb-6 text-sm text-gray-500">
              We'll send your login credentials to the email you provide here.
            </p>
            <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
              <Field label="Email" type="email" required value={form.email}
                error={errors.email} onChange={(v) => set("email", v)}
                placeholder="you@example.com" />
              <Field label="Cell phone" type="tel" required value={form.phone}
                error={errors.phone} onChange={(v) => set("phone", v)}
                placeholder="+265 999 123 456" />
              <Field label="Office phone" type="tel" value={form.officePhone}
                error={errors.officePhone} onChange={(v) => set("officePhone", v)}
                placeholder="+265 1 234 567" />
              <Field label="Home phone" type="tel" value={form.homePhone}
                error={errors.homePhone} onChange={(v) => set("homePhone", v)}
                placeholder="+265 1 765 432" />
            </div>
          </div>
        )}

        {/* ── Step 3 (new only): Addresses & next of kin ──────── */}
        {step === 3 && !isExisting && (
          <div>
            <h2 className="mb-1 text-xl font-semibold text-gray-900">
              Addresses & next of kin
            </h2>
            <p className="mb-6 text-sm text-gray-500">
              Provide at least one address. The contact person will be used in case
              we cannot reach you directly.
            </p>
            <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
              <Field label="Postal address" value={form.postalAddress}
                error={errors.postalAddress} onChange={(v) => set("postalAddress", v)}
                placeholder="P.O. Box 123, Blantyre" />
              <Field label="Physical address" value={form.physicalAddress}
                onChange={(v) => set("physicalAddress", v)}
                placeholder="Street, suburb, city" />
              <Field label="Contact person (next of kin)" value={form.contactPerson}
                onChange={(v) => set("contactPerson", v)}
                placeholder="Full name + phone"
                className="sm:col-span-2" />
            </div>
          </div>
        )}

        {/* ── Final: Review ────────────────────────────────────── */}
        {isReview && (
          <div>
            <h2 className="mb-1 text-xl font-semibold text-gray-900">Review &amp; submit</h2>
            <p className="mb-6 text-sm text-gray-500">
              Please confirm your details below. You can go back to fix anything.
            </p>
            <Summary form={form} />
          </div>
        )}

        {/* ── Footer nav ───────────────────────────────────────── */}
        {step > 0 && (
          <div className="mt-8 flex items-center justify-between gap-3">
            <button
              type="button"
              onClick={back}
              className="rounded-lg border border-gray-300 bg-white px-5 py-2.5 text-sm font-medium text-gray-700 hover:bg-gray-50"
            >
              Back
            </button>
            {isReview ? (
              <button
                type="submit"
                disabled={loading}
                className="rounded-lg bg-brand-600 px-6 py-2.5 text-sm font-semibold text-white hover:bg-brand-700 disabled:opacity-50"
              >
                {loading ? "Submitting…" : "Submit application"}
              </button>
            ) : (
              <button
                type="button"
                onClick={next}
                className="rounded-lg bg-brand-600 px-6 py-2.5 text-sm font-semibold text-white hover:bg-brand-700"
              >
                Continue
              </button>
            )}
          </div>
        )}
      </form>
    </Shell>
  );
}

// ─── Layout ─────────────────────────────────────────────────────────
function Shell({ children }: { children: ReactNode }) {
  return (
    <div className="flex min-h-screen items-center justify-center bg-gradient-to-br from-gray-50 to-brand-25 px-4 py-12">
      <div className="w-full max-w-5xl">
        <div className="mb-6 text-center">
        </div>
        {children}
      </div>
    </div>
  );
}

function Stepper({ steps, current }: { steps: string[]; current: number }) {
  return (
    <ol className="mb-6 flex items-center gap-2 overflow-x-auto pb-1">
      {steps.map((label, i) => {
        const done = i < current;
        const active = i === current;
        return (
          <li key={label} className="flex items-center gap-2">
            <span
              className={`flex h-7 w-7 shrink-0 items-center justify-center rounded-full text-xs font-semibold ${
                done
                  ? "bg-brand-600 text-white"
                  : active
                  ? "bg-brand-100 text-brand-700 ring-2 ring-brand-500"
                  : "bg-gray-100 text-gray-500"
              }`}
            >
              {done ? "✓" : i + 1}
            </span>
            <span
              className={`whitespace-nowrap text-xs font-medium ${
                active ? "text-brand-700" : done ? "text-gray-700" : "text-gray-400"
              }`}
            >
              {label}
            </span>
            {i < steps.length - 1 && (
              <span className="mx-2 h-px w-6 bg-gray-200 sm:w-10" />
            )}
          </li>
        );
      })}
    </ol>
  );
}

function ChoiceCard({
  title,
  description,
  onClick,
}: {
  title: string;
  description: string;
  onClick: () => void;
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      className="rounded-xl border border-gray-200 bg-white p-5 text-left transition hover:border-brand-400 hover:shadow-md focus:outline-none focus:ring-2 focus:ring-brand-500/30"
    >
      <h3 className="mb-1 text-sm font-semibold text-gray-900">{title}</h3>
      <p className="text-xs text-gray-500">{description}</p>
      <span className="mt-3 inline-flex items-center gap-1 text-xs font-medium text-brand-600">
        Continue →
      </span>
    </button>
  );
}

function Summary({ form }: { form: FormState }) {
  const isExisting = form.clientKind === "existing";
  const rows: { label: string; value: string }[] = [
    { label: "Account type", value: isExisting ? "Existing CDS client" : "New applicant" },
    { label: "Name", value: `${form.firstName} ${form.lastName}`.trim() || "—" },
    { label: "CDS number", value: form.cdsNumber || "—" },
    { label: "ID / Passport", value: form.idNumber || "—" },
    { label: "Date of birth", value: form.dateOfBirth || "—" },
    { label: "Email", value: form.email || "—" },
    { label: "Cell phone", value: form.phone || "—" },
    { label: "Office phone", value: form.officePhone || "—" },
    { label: "Home phone", value: form.homePhone || "—" },
  ];
  if (!isExisting) {
    rows.push(
      { label: "Postal address", value: form.postalAddress || "—" },
      { label: "Physical address", value: form.physicalAddress || "—" },
      { label: "Contact person", value: form.contactPerson || "—" },
    );
  }
  return (
    <dl className="divide-y divide-gray-100 rounded-lg border border-gray-200">
      {rows.map((r) => (
        <div key={r.label} className="grid grid-cols-3 gap-3 px-4 py-2.5 text-sm">
          <dt className="text-gray-500">{r.label}</dt>
          <dd className="col-span-2 font-medium text-gray-800">{r.value}</dd>
        </div>
      ))}
    </dl>
  );
}

function Field({
  label,
  value,
  onChange,
  type = "text",
  required,
  placeholder,
  hint,
  error,
  className = "",
}: {
  label: string;
  value: string;
  onChange: (v: string) => void;
  type?: string;
  required?: boolean;
  placeholder?: string;
  hint?: string;
  error?: string;
  className?: string;
}) {
  return (
    <div className={className}>
      <label className="mb-1.5 block text-sm font-medium text-gray-700">
        {label} {required && <span className="text-red-500">*</span>}
      </label>
      <input
        type={type}
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder}
        autoComplete="off"
        aria-invalid={error ? "true" : "false"}
        className={`w-full rounded-lg border px-4 py-2.5 text-sm focus:outline-none focus:ring-2 ${
          error
            ? "border-red-300 focus:border-red-500 focus:ring-red-500/20"
            : "border-gray-300 focus:border-brand-500 focus:ring-brand-500/20"
        }`}
      />
      {error ? (
        <p className="mt-1 text-xs text-red-600">{error}</p>
      ) : hint ? (
        <p className="mt-1 text-xs text-gray-500">{hint}</p>
      ) : null}
    </div>
  );
}
