import { useMemo, useState, type FormEvent, type ReactNode } from "react";
import { Link } from "react-router";
import { useAuth } from "../context/AuthContext";
import api from "../lib/api";
import DatePicker from "../components/form/DatePicker";
import SearchSelect from "../components/form/SearchSelect";
import { brand } from "../lib/brand";

// ─── Types ──────────────────────────────────────────────────────────
type ClientKind = "existing" | "new" | null;

const ID_TYPE_OPTIONS = [
  { value: "National ID", label: "National ID" },
  { value: "Passport", label: "Passport" },
];

const ACCOUNT_TYPE_OPTIONS = [
  { value: "Individual", label: "Individual" },
  { value: "Joint", label: "Joint (more than one holder)" },
  { value: "ITF", label: "In Trust For (e.g. a minor child)" },
];

interface JointApplicant {
  fullName: string;
  idDocumentType: string;
  idNumber: string;
  relationship: string;
  idDocument: File | null;
}

interface ItfBeneficiary {
  fullName: string;
  dateOfBirth: string;
  relationship: string;
}

interface FormState {
  // Branch flag
  clientKind: ClientKind;

  // Personal
  firstName: string;
  lastName: string;
  idDocumentType: string;
  idNumber: string;
  cdsNumber: string;
  dateOfBirth: string;

  // Account type & joint / in-trust-for intake
  accountType: string;
  jointApplicants: JointApplicant[];
  itfBeneficiary: ItfBeneficiary;

  // Contact
  email: string;
  phone: string;
  officePhone: string;
  homePhone: string;

  // Address & next of kin
  postalAddress: string;
  physicalAddress: string;
  contactPerson: string;

  // FIU / Cedar Capital KYC documents
  idDocument: File | null;
  proofOfAddress: File | null;
  sourceOfFunds: File | null;
  otherDocuments: File[];
}

type Errors = Partial<Record<keyof FormState | "documents", string>>;

const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const PHONE_RE = /^[+\d\s()-]{7,20}$/;
const MAX_FILE_BYTES = 10 * 1024 * 1024;
const ALLOWED_EXT = [
  ".pdf", ".jpg", ".jpeg", ".png", ".gif", ".webp",
  ".doc", ".docx", ".xls", ".xlsx", ".csv", ".txt",
];

const EMPTY: FormState = {
  clientKind: null,
  firstName: "",
  lastName: "",
  idDocumentType: "National ID",
  idNumber: "",
  cdsNumber: "",
  dateOfBirth: "",
  accountType: "Individual",
  jointApplicants: [],
  itfBeneficiary: { fullName: "", dateOfBirth: "", relationship: "" },
  email: "",
  phone: "",
  officePhone: "",
  homePhone: "",
  postalAddress: "",
  physicalAddress: "",
  contactPerson: "",
  idDocument: null,
  proofOfAddress: null,
  sourceOfFunds: null,
  otherDocuments: [],
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

    // Existing clients are already on the CDS — the CSD number is optional
    // (helps us match faster, but staff can also link by name/email).

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

    // Joint / In-Trust-For intake validation.
    if (form.accountType === "Joint") {
      const named = form.jointApplicants.filter((j) => j.fullName.trim());
      if (named.length === 0)
        e.jointApplicants = "Add at least one joint account holder.";
      else if (named.some((j) => !j.idDocument))
        e.jointApplicants = "Attach an ID document for each joint account holder.";
    } else if (form.accountType === "ITF") {
      if (!form.itfBeneficiary.fullName.trim())
        e.itfBeneficiary = "Beneficiary full name is required for an In-Trust-For account.";
      else if (!form.itfBeneficiary.relationship.trim())
        e.itfBeneficiary = "State your relationship to the beneficiary.";
    }
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

  // Documents step — new applicants only (existing clients are already KYC'd
  // so they have no upload step). For new applicants it's step 4. FIU / Cedar
  // Capital require ID + proof of address + proof of source of funds.
  if (!isExisting && step === 4) {
    const missing: string[] = [];
    if (!form.idDocument) missing.push("National ID / Passport");
    if (!form.proofOfAddress) missing.push("Proof of address");
    if (!form.sourceOfFunds) missing.push("Proof of source of funds");
    if (missing.length) e.documents = `Please attach: ${missing.join(", ")}.`;
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
  // True while we ask the API whether the ID/CDS/email is already registered.
  const [checking, setChecking] = useState(false);

  const isExisting = form.clientKind === "existing";

  // Existing clients have a 3-step flow (chooser → details → review).
  // New clients have a 5-step flow (chooser → personal → contact → address → review).
  const steps = useMemo(
    () =>
      isExisting
        ? // Existing clients are already KYC'd and on the CDS, so they skip the
          // document-upload + address steps and CSD is optional — we just need
          // their name + contact email to match them and create the login.
          ["Account", "Personal", "Contact", "Review"]
        : ["Account", "Personal", "Contact", "Addresses", "Documents", "Review"],
    [isExisting],
  );
  const lastStep = steps.length - 1;
  const isReview = step === lastStep;
  // Documents only exist in the new-applicant flow (step 4).
  const isDocuments = !isExisting && step === lastStep - 1;

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

  // Ask the API whether any of the supplied identifiers are already in use.
  // On a network/server error we don't block the applicant — the final submit
  // (and the server-side guard) will still reject a genuine duplicate.
  async function checkDuplicates(fields: {
    email?: string;
    idNumber?: string;
    cdsNumber?: string;
  }): Promise<Errors> {
    try {
      const { data } = await api.post("/auth/check-availability", fields);
      const e: Errors = {};
      if (data.emailTaken)
        e.email = "An account with this email already exists.";
      if (data.idNumberTaken)
        e.idNumber = "This ID / Passport number is already registered.";
      if (data.cdsNumberTaken)
        e.cdsNumber = "This CSD number is already registered.";
      return e;
    } catch {
      return {};
    }
  }

  async function next() {
    setSubmitError("");
    const stepErrors = validateStep(step, form);
    if (Object.keys(stepErrors).length) {
      setErrors(stepErrors);
      return;
    }

    // Duplicate-prevention gate: step 1 carries ID/Passport + CDS, step 2
    // carries the email. Block advancing if any is already registered.
    //
    // For EXISTING clients we skip this gate — their email/ID/CSD already live
    // in the client master (that's the record we're matching them to), so the
    // shared availability check would wrongly flag them. The server's
    // existing-client-aware guard still blocks a duplicate portal *login*.
    let dupFields: { email?: string; idNumber?: string; cdsNumber?: string } | null = null;
    if (!isExisting && step === 1) {
      dupFields = {};
      if (form.idNumber.trim()) dupFields.idNumber = form.idNumber.trim();
      if (form.cdsNumber.trim()) dupFields.cdsNumber = form.cdsNumber.trim();
    } else if (!isExisting && step === 2 && form.email.trim()) {
      dupFields = { email: form.email.trim() };
    }

    if (dupFields && Object.keys(dupFields).length) {
      setChecking(true);
      const dupErrors = await checkDuplicates(dupFields);
      setChecking(false);
      if (Object.keys(dupErrors).length) {
        setErrors(dupErrors);
        return;
      }
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

    // Guard against Enter-key submits from earlier steps: only the final
    // Review step is allowed to actually post the application.
    if (!isReview) {
      void next();
      return;
    }

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
      idDocumentType: form.idDocumentType || undefined,
      cdsNumber: form.cdsNumber || undefined,
      dateOfBirth: form.dateOfBirth || undefined,
      accountType: form.accountType || undefined,
      jointApplicants:
        form.accountType === "Joint" && form.jointApplicants.some((j) => j.fullName.trim())
          ? JSON.stringify(
              form.jointApplicants
                .filter((j) => j.fullName.trim())
                .map(({ idDocument: _idDocument, ...rest }) => rest),
            )
          : undefined,
      jointIdDocuments:
        form.accountType === "Joint"
          ? form.jointApplicants
              .filter((j) => j.fullName.trim())
              .map((j) => j.idDocument)
              .filter((f): f is File => !!f)
          : undefined,
      itfBeneficiary:
        form.accountType === "ITF" && form.itfBeneficiary.fullName.trim()
          ? JSON.stringify(form.itfBeneficiary)
          : undefined,
      physicalAddress: form.physicalAddress || undefined,
      postalAddress: form.postalAddress || undefined,
      contactPerson: form.contactPerson || undefined,
      isExistingClient: isExisting,
      idDocument: form.idDocument ?? undefined,
      proofOfAddress: form.proofOfAddress ?? undefined,
      sourceOfFunds: form.sourceOfFunds ?? undefined,
      otherDocuments: form.otherDocuments.length ? form.otherDocuments : undefined,
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
              ? "An administrator will verify your details against our records and email you your login credentials."
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
              Welcome to {brand.name}, let's get you started.
            </h2>
            <p className="mb-6 text-sm text-gray-500">
              Tell us which best describes you so we can take you down the right path.
            </p>
            <div className="grid gap-4 sm:grid-cols-2">
              <ChoiceCard
                title={`I am already a ${brand.name} Client`}
                description="You already have a CSD account with us and just need an online portal login. We'll verify your CSD number and link this login to your existing client record."
                onClick={() => chooseKind("existing")}
              />
              <ChoiceCard
                title={`I am new, I would like to be a ${brand.name} Client`}
                description="You don't yet have a CSD account. Complete the full application below and our onboarding team will set you up."
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
                ? "Provide your name so our team can locate your existing account. CSD number is optional but helps us match you faster."
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
                label="CSD number"
                value={form.cdsNumber}
                error={errors.cdsNumber}
                onChange={(v) => set("cdsNumber", v)}
                placeholder="e.g. CSD-000123"
                hint={isExisting ? "Optional — helps admin match your existing account." : undefined}
              />
              <SearchSelect
                label="ID document type"
                required={!isExisting}
                value={form.idDocumentType}
                onChange={(v) => set("idDocumentType", v)}
                options={ID_TYPE_OPTIONS}
                searchPlaceholder="Search ID type…"
              />
              <Field
                label={form.idDocumentType === "Passport" ? "Passport number" : "National ID number"}
                required={!isExisting}
                value={form.idNumber}
                error={errors.idNumber}
                onChange={(v) => set("idNumber", v)}
                placeholder={form.idDocumentType === "Passport" ? "Passport no." : "National ID no."}
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

            {!isExisting && (
              <div className="mt-6 border-t border-gray-100 pt-6">
                <h3 className="mb-1 text-base font-semibold text-gray-900">Account type</h3>
                <p className="mb-4 text-sm text-gray-500">
                  Choose how this account will be held. Joint accounts have more than one holder; an
                  In Trust For (ITF) account is opened on behalf of someone else, such as a minor child.
                </p>
                <div className="max-w-md">
                  <SearchSelect
                    label="Account will be held as"
                    value={form.accountType}
                    onChange={(v) =>
                      setForm((prev) => ({
                        ...prev,
                        accountType: v,
                        jointApplicants:
                          v === "Joint" && prev.jointApplicants.length === 0
                            ? [{ fullName: "", idDocumentType: "National ID", idNumber: "", relationship: "", idDocument: null }]
                            : prev.jointApplicants,
                      }))
                    }
                    options={ACCOUNT_TYPE_OPTIONS}
                    searchPlaceholder="Search account type…"
                  />
                </div>

                {form.accountType === "Joint" && (
                  <div className="mt-4">
                    <div className="mb-3 rounded-lg border border-brand-200 bg-brand-50 p-3 text-sm text-brand-800">
                      <span className="font-semibold">{`${form.firstName} ${form.lastName}`.trim() || "You"}</span>{" "}
                      (the person registering) will be the primary contact person and the only login holder for
                      this joint account. Add each additional holder and their KYC below.
                    </div>
                    <div className="mb-2 flex items-center justify-between">
                      <h4 className="text-sm font-semibold text-gray-800">Joint account holders</h4>
                      <button
                        type="button"
                        onClick={() =>
                          set("jointApplicants", [
                            ...form.jointApplicants,
                            { fullName: "", idDocumentType: "National ID", idNumber: "", relationship: "", idDocument: null },
                          ])
                        }
                        className="rounded-lg border border-brand-300 px-3 py-1.5 text-sm font-medium text-brand-700 hover:bg-brand-50"
                      >
                        + Add holder
                      </button>
                    </div>
                    {errors.jointApplicants && (
                      <p className="mb-2 text-xs text-red-600">{errors.jointApplicants}</p>
                    )}
                    <div className="space-y-4">
                      {form.jointApplicants.map((j, idx) => (
                        <div key={idx} className="rounded-xl border border-gray-200 p-4">
                          <div className="mb-2 flex items-center justify-between">
                            <span className="text-xs font-semibold uppercase tracking-wide text-gray-500">
                              Holder {idx + 1}
                            </span>
                            <button
                              type="button"
                              onClick={() =>
                                set(
                                  "jointApplicants",
                                  form.jointApplicants.filter((_, i) => i !== idx),
                                )
                              }
                              className="text-sm font-medium text-red-600 hover:underline"
                            >
                              Remove
                            </button>
                          </div>
                          <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
                            <Field
                              label="Full name"
                              value={j.fullName}
                              onChange={(v) => {
                                const next = [...form.jointApplicants];
                                next[idx] = { ...next[idx], fullName: v };
                                set("jointApplicants", next);
                              }}
                              placeholder="e.g. Mary Banda"
                            />
                            <SearchSelect
                              label="ID document type"
                              value={j.idDocumentType}
                              onChange={(v) => {
                                const next = [...form.jointApplicants];
                                next[idx] = { ...next[idx], idDocumentType: v };
                                set("jointApplicants", next);
                              }}
                              options={ID_TYPE_OPTIONS}
                              searchPlaceholder="Search ID type…"
                            />
                            <Field
                              label={j.idDocumentType === "Passport" ? "Passport number" : "National ID number"}
                              value={j.idNumber}
                              onChange={(v) => {
                                const next = [...form.jointApplicants];
                                next[idx] = { ...next[idx], idNumber: v };
                                set("jointApplicants", next);
                              }}
                              placeholder={j.idDocumentType === "Passport" ? "Passport no." : "National ID no."}
                            />
                            <Field
                              label="Relationship"
                              value={j.relationship}
                              onChange={(v) => {
                                const next = [...form.jointApplicants];
                                next[idx] = { ...next[idx], relationship: v };
                                set("jointApplicants", next);
                              }}
                              placeholder="e.g. Spouse"
                            />
                          </div>
                          <div className="mt-3">
                            <FileSlot
                              label="ID document (National ID or Passport copy)"
                              required
                              file={j.idDocument}
                              onChange={(f) => {
                                const next = [...form.jointApplicants];
                                next[idx] = { ...next[idx], idDocument: f };
                                set("jointApplicants", next);
                              }}
                            />
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                )}

                {form.accountType === "ITF" && (
                  <div className="mt-4 rounded-xl border border-gray-200 p-4">
                    <h4 className="mb-1 text-sm font-semibold text-gray-800">Beneficiary (e.g. minor child)</h4>
                    <p className="mb-3 text-xs text-gray-500">
                      You are opening this account in trust for the person below.
                    </p>
                    {errors.itfBeneficiary && (
                      <p className="mb-2 text-xs text-red-600">{errors.itfBeneficiary}</p>
                    )}
                    <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
                      <Field
                        label="Beneficiary full name"
                        required
                        value={form.itfBeneficiary.fullName}
                        onChange={(v) =>
                          set("itfBeneficiary", { ...form.itfBeneficiary, fullName: v })
                        }
                        placeholder="e.g. Junior Banda"
                      />
                      <div>
                        <label className="mb-1.5 block text-sm font-medium text-gray-700">
                          Beneficiary date of birth
                        </label>
                        <DatePicker
                          value={form.itfBeneficiary.dateOfBirth}
                          onChange={(v) =>
                            set("itfBeneficiary", { ...form.itfBeneficiary, dateOfBirth: v })
                          }
                          placeholder="Pick date of birth"
                          maxDate={new Date()}
                        />
                      </div>
                      <Field
                        label="Your relationship to beneficiary"
                        required
                        value={form.itfBeneficiary.relationship}
                        onChange={(v) =>
                          set("itfBeneficiary", { ...form.itfBeneficiary, relationship: v })
                        }
                        placeholder="e.g. Parent / Guardian"
                      />
                    </div>
                  </div>
                )}
              </div>
            )}
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

        {/* ── Documents step (second-to-last) ──────────────────── */}
        {isDocuments && (
          <div>
            <h2 className="mb-1 text-xl font-semibold text-gray-900">
              KYC documents
            </h2>
            <p className="mb-4 text-sm text-gray-500">
              In line with FIU / {brand.name} onboarding requirements, please
              attach the following. Certified copies dated within the last
              three months are preferred.
            </p>
            {errors.documents && (
              <div className="mb-4 rounded-lg bg-red-50 p-3 text-sm text-red-700">
                {errors.documents}
              </div>
            )}
            <div className="space-y-4">
              <FileSlot
                label="National ID or Passport"
                required
                file={form.idDocument}
                onChange={(f) => set("idDocument", f)}
              />
              <FileSlot
                label="Proof of residential address (utility bill, bank statement, lease)"
                required
                file={form.proofOfAddress}
                onChange={(f) => set("proofOfAddress", f)}
              />
              <FileSlot
                label="Proof of source of funds (payslip, bank statement, business registration)"
                required
                file={form.sourceOfFunds}
                onChange={(f) => set("sourceOfFunds", f)}
              />
              <MultiFileSlot
                label="Other supporting documents (optional)"
                files={form.otherDocuments}
                onChange={(files) => set("otherDocuments", files)}
              />
            </div>
            <p className="mt-4 text-xs text-gray-500">
              Accepted: PDF, images, Word, Excel, CSV, text. Maximum 10&nbsp;MB per file.
            </p>
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
                disabled={checking}
                className="rounded-lg bg-brand-600 px-6 py-2.5 text-sm font-semibold text-white hover:bg-brand-700 disabled:opacity-50"
              >
                {checking ? "Checking…" : "Continue"}
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
    <div className="flex min-h-screen items-start justify-center bg-gradient-to-br from-gray-50 to-brand-25 px-3 py-6 sm:items-center sm:px-4 sm:py-12">
      <div className="w-full max-w-5xl">
        <div className="mb-6 text-center">
          <h1 className="text-2xl font-bold text-gray-900 sm:text-3xl">Open your {brand.name} account</h1>
          <p className="mt-1 text-sm text-gray-500 sm:text-base">
            Complete the application below — it only takes a few minutes.
          </p>
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
      className="rounded-xl border border-gray-200 bg-white p-6 text-left transition hover:border-brand-400 hover:shadow-md focus:outline-none focus:ring-2 focus:ring-brand-500/30"
    >
      <h3 className="mb-1 text-base font-semibold text-gray-900">{title}</h3>
      <p className="text-sm text-gray-500">{description}</p>
      <span className="mt-3 inline-flex items-center gap-1 text-sm font-medium text-brand-600">
        Continue →
      </span>
    </button>
  );
}

function Summary({ form }: { form: FormState }) {
  const isExisting = form.clientKind === "existing";
  const rows: { label: string; value: string }[] = [
    { label: "Applicant", value: isExisting ? "Existing CSD client" : "New applicant" },
    { label: "Name", value: `${form.firstName} ${form.lastName}`.trim() || "—" },
    { label: "CSD number", value: form.cdsNumber || "—" },
    { label: "ID type", value: form.idDocumentType || "—" },
    { label: "ID / Passport", value: form.idNumber || "—" },
    { label: "Date of birth", value: form.dateOfBirth || "—" },
    { label: "Email", value: form.email || "—" },
    { label: "Cell phone", value: form.phone || "—" },
    { label: "Office phone", value: form.officePhone || "—" },
    { label: "Home phone", value: form.homePhone || "—" },
  ];
  if (!isExisting) {
    const accountTypeLabel =
      ACCOUNT_TYPE_OPTIONS.find((o) => o.value === form.accountType)?.label ?? form.accountType;
    rows.push({ label: "Account held as", value: accountTypeLabel || "—" });
    if (form.accountType === "Joint") {
      const holders = form.jointApplicants.filter((j) => j.fullName.trim());
      rows.push({
        label: "Primary contact / login",
        value: `${form.firstName} ${form.lastName}`.trim() || "—",
      });
      rows.push({
        label: "Joint holders",
        value: holders.length
          ? holders
              .map(
                (j) =>
                  `${j.fullName}${j.relationship ? ` (${j.relationship})` : ""}${
                    j.idDocument ? " — ID attached" : " — ID missing"
                  }`,
              )
              .join(", ")
          : "—",
      });
    }
    if (form.accountType === "ITF") {
      rows.push({
        label: "Beneficiary",
        value: form.itfBeneficiary.fullName
          ? `${form.itfBeneficiary.fullName}${
              form.itfBeneficiary.relationship ? ` (${form.itfBeneficiary.relationship})` : ""
            }`
          : "—",
      });
    }
    rows.push(
      { label: "Postal address", value: form.postalAddress || "—" },
      { label: "Physical address", value: form.physicalAddress || "—" },
      { label: "Contact person", value: form.contactPerson || "—" },
    );
  }
  rows.push(
    { label: "ID document", value: form.idDocument?.name ?? "—" },
    { label: "Proof of address", value: form.proofOfAddress?.name ?? "—" },
    { label: "Source of funds", value: form.sourceOfFunds?.name ?? "—" },
    {
      label: "Other documents",
      value: form.otherDocuments.length
        ? form.otherDocuments.map((f) => f.name).join(", ")
        : "—",
    },
  );
  return (
    <dl className="divide-y divide-gray-100 rounded-lg border border-gray-200">
      {rows.map((r) => (
        <div key={r.label} className="grid grid-cols-1 gap-1 px-4 py-2.5 text-sm sm:grid-cols-3 sm:gap-3">
          <dt className="text-gray-500">{r.label}</dt>
          <dd className="font-medium text-gray-800 sm:col-span-2">{r.value}</dd>
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
        className={`w-full rounded-lg border px-4 py-3 text-base focus:outline-none focus:ring-2 ${
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

// ── File upload helpers ────────────────────────────────
function validateFile(f: File): string | null {
  if (f.size > MAX_FILE_BYTES) return `${f.name} exceeds the 10 MB limit.`;
  const lower = f.name.toLowerCase();
  if (!ALLOWED_EXT.some((ext) => lower.endsWith(ext)))
    return `${f.name}: file type not allowed.`;
  return null;
}

function formatSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

function FileSlot({
  label,
  required,
  file,
  onChange,
}: {
  label: string;
  required?: boolean;
  file: File | null;
  onChange: (f: File | null) => void;
}) {
  const [err, setErr] = useState<string | null>(null);
  function pick(f: File | null) {
    if (!f) { onChange(null); setErr(null); return; }
    const e = validateFile(f);
    if (e) { setErr(e); return; }
    setErr(null);
    onChange(f);
  }
  return (
    <div className="rounded-lg border border-gray-200 bg-gray-50 p-3">
      <div className="mb-2 text-sm font-medium text-gray-700">
        {label} {required && <span className="text-red-500">*</span>}
      </div>
      {file ? (
        <div className="flex items-center justify-between gap-3 rounded-md border border-gray-200 bg-white px-3 py-2 text-sm">
          <div className="min-w-0 truncate">
            <span className="font-medium text-gray-800">{file.name}</span>{" "}
            <span className="text-xs text-gray-500">({formatSize(file.size)})</span>
          </div>
          <button
            type="button"
            onClick={() => pick(null)}
            className="text-xs font-medium text-red-600 hover:underline"
          >
            Remove
          </button>
        </div>
      ) : (
        <label className="flex cursor-pointer items-center justify-center rounded-md border border-dashed border-gray-300 bg-white px-3 py-4 text-sm text-gray-500 hover:border-brand-400 hover:text-brand-600">
          <input
            type="file"
            className="sr-only"
            accept={ALLOWED_EXT.join(",")}
            onChange={(e) => pick(e.target.files?.[0] ?? null)}
          />
          Click to choose a file
        </label>
      )}
      {err && <p className="mt-1 text-xs text-red-600">{err}</p>}
    </div>
  );
}

function MultiFileSlot({
  label,
  files,
  onChange,
}: {
  label: string;
  files: File[];
  onChange: (f: File[]) => void;
}) {
  const [err, setErr] = useState<string | null>(null);
  function add(list: FileList | null) {
    if (!list || list.length === 0) return;
    const next = [...files];
    for (const f of Array.from(list)) {
      const e = validateFile(f);
      if (e) { setErr(e); continue; }
      next.push(f);
    }
    setErr(null);
    onChange(next);
  }
  function remove(i: number) {
    const next = [...files];
    next.splice(i, 1);
    onChange(next);
  }
  return (
    <div className="rounded-lg border border-gray-200 bg-gray-50 p-3">
      <div className="mb-2 text-sm font-medium text-gray-700">{label}</div>
      {files.length > 0 && (
        <ul className="mb-2 space-y-1">
          {files.map((f, i) => (
            <li key={`${f.name}-${i}`} className="flex items-center justify-between gap-3 rounded-md border border-gray-200 bg-white px-3 py-2 text-sm">
              <div className="min-w-0 truncate">
                <span className="font-medium text-gray-800">{f.name}</span>{" "}
                <span className="text-xs text-gray-500">({formatSize(f.size)})</span>
              </div>
              <button
                type="button"
                onClick={() => remove(i)}
                className="text-xs font-medium text-red-600 hover:underline"
              >
                Remove
              </button>
            </li>
          ))}
        </ul>
      )}
      <label className="flex cursor-pointer items-center justify-center rounded-md border border-dashed border-gray-300 bg-white px-3 py-3 text-sm text-gray-500 hover:border-brand-400 hover:text-brand-600">
        <input
          type="file"
          multiple
          className="sr-only"
          accept={ALLOWED_EXT.join(",")}
          onChange={(e) => { add(e.target.files); e.currentTarget.value = ""; }}
        />
        Click to add files
      </label>
      {err && <p className="mt-1 text-xs text-red-600">{err}</p>}
    </div>
  );
}
