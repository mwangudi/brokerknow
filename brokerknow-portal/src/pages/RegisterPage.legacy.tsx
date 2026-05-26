import { useState, type FormEvent } from "react";
import { Link } from "react-router";
import { useAuth } from "../context/AuthContext";
import DatePicker from "../components/form/DatePicker";

interface FormState {
  firstName: string;
  lastName: string;
  email: string;
  phone: string;
  officePhone: string;
  homePhone: string;
  idNumber: string;
  cdsNumber: string;
  dateOfBirth: string;
  physicalAddress: string;
  postalAddress: string;
  contactPerson: string;
}

type Errors = Partial<Record<keyof FormState, string>>;

const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const PHONE_RE = /^[+\d\s()-]{7,20}$/;

function validate(form: FormState): Errors {
  const errors: Errors = {};
  if (!form.firstName.trim()) errors.firstName = "First name is required.";
  else if (form.firstName.trim().length < 2) errors.firstName = "First name must be at least 2 characters.";

  if (!form.lastName.trim()) errors.lastName = "Last name is required.";
  else if (form.lastName.trim().length < 2) errors.lastName = "Last name must be at least 2 characters.";

  if (!form.email.trim()) errors.email = "Email is required.";
  else if (!EMAIL_RE.test(form.email.trim())) errors.email = "Enter a valid email address.";

  if (form.phone && !PHONE_RE.test(form.phone)) errors.phone = "Enter a valid phone number.";
  if (form.officePhone && !PHONE_RE.test(form.officePhone)) errors.officePhone = "Enter a valid phone number.";
  if (form.homePhone && !PHONE_RE.test(form.homePhone)) errors.homePhone = "Enter a valid phone number.";

  if (form.dateOfBirth) {
    const dob = new Date(form.dateOfBirth);
    if (isNaN(dob.getTime())) {
      errors.dateOfBirth = "Enter a valid date.";
    } else {
      const today = new Date();
      const age = today.getFullYear() - dob.getFullYear()
        - (today.getMonth() < dob.getMonth() ||
          (today.getMonth() === dob.getMonth() && today.getDate() < dob.getDate()) ? 1 : 0);
      if (age < 18) errors.dateOfBirth = "You must be at least 18 years old to register.";
      else if (age > 120) errors.dateOfBirth = "Please enter a valid date of birth.";
    }
  }

  return errors;
}

export default function RegisterPage() {
  const { register, loading } = useAuth();
  const [form, setForm] = useState<FormState>({
    firstName: "",
    lastName: "",
    email: "",
    phone: "",
    officePhone: "",
    homePhone: "",
    idNumber: "",
    cdsNumber: "",
    dateOfBirth: "",
    physicalAddress: "",
    postalAddress: "",
    contactPerson: "",
  });
  const [errors, setErrors] = useState<Errors>({});
  const [touched, setTouched] = useState<Partial<Record<keyof FormState, boolean>>>({});
  const [submitError, setSubmitError] = useState("");
  const [success, setSuccess] = useState("");

  function set<K extends keyof FormState>(field: K, value: string) {
    setForm((prev) => ({ ...prev, [field]: value }));
    if (touched[field]) {
      // Re-validate this field on each keystroke once the user has touched it
      const next = { ...form, [field]: value };
      const fieldErrors = validate(next);
      setErrors((prev) => ({ ...prev, [field]: fieldErrors[field] }));
    }
  }

  function touch(field: keyof FormState) {
    setTouched((prev) => ({ ...prev, [field]: true }));
    const fieldErrors = validate(form);
    setErrors((prev) => ({ ...prev, [field]: fieldErrors[field] }));
  }

  async function handleSubmit(e: FormEvent) {
    e.preventDefault();
    setSubmitError("");
    setSuccess("");

    const validationErrors = validate(form);
    setErrors(validationErrors);
    setTouched({
      firstName: true, lastName: true, email: true, phone: true,
      officePhone: true, homePhone: true, dateOfBirth: true,
      idNumber: true, cdsNumber: true, physicalAddress: true,
      postalAddress: true, contactPerson: true,
    });

    if (Object.keys(validationErrors).length > 0) {
      setSubmitError("Please fix the errors highlighted below.");
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

  return (
    <div className="flex min-h-screen items-center justify-center bg-gradient-to-br from-blue-50 to-indigo-100 px-4 py-12">
      <div className="w-full max-w-3xl">
        <div className="mb-8 text-center">
          <img
            src="/images/brokerknow-logo.jpg"
            alt="BrokerKnow"
            className="mx-auto h-16 mb-3 mix-blend-multiply"
            onError={(e) => { (e.target as HTMLImageElement).style.display = "none"; }}
          />
        </div>

        {success ? (
          <div className="rounded-2xl border border-green-200 bg-white p-8 text-center shadow-sm">
            <div className="mb-4 inline-flex h-16 w-16 items-center justify-center rounded-full bg-green-100">
              <svg className="h-8 w-8 text-green-600" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
              </svg>
            </div>
            <h2 className="mb-2 text-xl font-semibold text-gray-900">Registration Received</h2>
            <p className="mb-6 text-gray-600">{success}</p>
            <p className="mb-6 text-sm text-gray-500">
              Once your application is reviewed and approved, you will receive your login credentials by email.
            </p>
            <Link
              to="/login"
              className="inline-block rounded-lg bg-blue-600 px-6 py-2.5 text-sm font-semibold text-white hover:bg-blue-700"
            >
              Back to Login
            </Link>
          </div>
        ) : (
          <form onSubmit={handleSubmit} noValidate className="rounded-2xl border border-gray-200 bg-white p-8 shadow-sm">
            <h2 className="mb-2 text-xl font-semibold text-gray-900">Apply for an account</h2>
            <p className="mb-6 text-sm text-gray-500">
              Complete all relevant details below. We'll review your application and email your login credentials once approved.
            </p>

            {submitError && (
              <div className="mb-4 rounded-lg bg-red-50 p-3 text-sm text-red-700">{submitError}</div>
            )}

            {/* Personal Information */}
            <Section title="Personal Information">
              <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
                <Field label="First Name" required value={form.firstName} error={touched.firstName ? errors.firstName : undefined}
                  onChange={(v) => set("firstName", v)} onBlur={() => touch("firstName")} placeholder="First Name" />
                <Field label="Last Name" required value={form.lastName} error={touched.lastName ? errors.lastName : undefined}
                  onChange={(v) => set("lastName", v)} onBlur={() => touch("lastName")} placeholder="Last Name" />
                <Field label="ID / Passport Number" value={form.idNumber}
                  onChange={(v) => set("idNumber", v)} placeholder="National ID or Passport no." />
                <Field label="CDS Number" value={form.cdsNumber}
                  onChange={(v) => set("cdsNumber", v)} placeholder="e.g. CDS-000123" />
                <div>
                  <label className="mb-1.5 block text-sm font-medium text-gray-700">Date of Birth</label>
                  <DatePicker
                    value={form.dateOfBirth}
                    onChange={(v) => { set("dateOfBirth", v); touch("dateOfBirth"); }}
                    placeholder="Pick date of birth"
                    maxDate={new Date()}
                  />
                  {touched.dateOfBirth && errors.dateOfBirth && (
                    <p className="mt-1 text-xs text-red-600">{errors.dateOfBirth}</p>
                  )}
                </div>
                <Field label="Contact Person" value={form.contactPerson}
                  onChange={(v) => set("contactPerson", v)} placeholder="Next of kin or alternate contact" />
              </div>
            </Section>

            {/* Contact Details */}
            <Section title="Contact Details">
              <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
                <Field label="Email" type="email" required value={form.email} error={touched.email ? errors.email : undefined}
                  onChange={(v) => set("email", v)} onBlur={() => touch("email")} placeholder="you@example.com"
                  hint="Your login credentials will be sent here." />
                <Field label="Cell Phone" type="tel" value={form.phone} error={touched.phone ? errors.phone : undefined}
                  onChange={(v) => set("phone", v)} onBlur={() => touch("phone")} placeholder="+265 999 123 456" />
                <Field label="Office Phone" type="tel" value={form.officePhone} error={touched.officePhone ? errors.officePhone : undefined}
                  onChange={(v) => set("officePhone", v)} onBlur={() => touch("officePhone")} placeholder="+265 1 234 567" />
                <Field label="Home Phone" type="tel" value={form.homePhone} error={touched.homePhone ? errors.homePhone : undefined}
                  onChange={(v) => set("homePhone", v)} onBlur={() => touch("homePhone")} placeholder="+265 1 765 432" />
              </div>
            </Section>

            {/* Addresses */}
            <Section title="Addresses">
              <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
                <Field label="Postal Address" value={form.postalAddress}
                  onChange={(v) => set("postalAddress", v)} placeholder="P.O. Box 123, Blantyre" />
                <Field label="Physical Address" value={form.physicalAddress}
                  onChange={(v) => set("physicalAddress", v)} placeholder="Street, suburb, city" />
              </div>
            </Section>

            <button
              type="submit"
              disabled={loading}
              className="mt-6 w-full rounded-lg bg-blue-600 py-3 text-sm font-semibold text-white hover:bg-blue-700 disabled:opacity-50"
            >
              {loading ? "Submitting..." : "Submit Application"}
            </button>

            <p className="mt-4 text-center text-sm text-gray-500">
              Already have an account?{" "}
              <Link to="/login" className="font-medium text-blue-600 hover:underline">
                Sign in
              </Link>
            </p>
          </form>
        )}
      </div>
    </div>
  );
}

function Section({ title, children }: { title: string; children: React.ReactNode }) {
  return (
    <div className="mb-6">
      <h3 className="mb-3 text-xs font-semibold uppercase tracking-wider text-gray-500">{title}</h3>
      {children}
    </div>
  );
}

function Field({
  label, value, onChange, onBlur, type = "text", required, placeholder, hint, error, className = "",
}: {
  label: string;
  value: string;
  onChange: (v: string) => void;
  onBlur?: () => void;
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
        onBlur={onBlur}
        placeholder={placeholder}
        aria-invalid={error ? "true" : "false"}
        className={`w-full rounded-lg border px-4 py-2.5 text-sm focus:outline-none focus:ring-2 ${
          error
            ? "border-red-300 focus:border-red-500 focus:ring-red-500/20"
            : "border-gray-300 focus:border-blue-500 focus:ring-blue-500/20"
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
