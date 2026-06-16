import { useEffect, useState } from "react";
import { Link } from "react-router";
import api from "../lib/api";
import { useAuth } from "../context/AuthContext";
import PageMeta from "../components/common/PageMeta";
import Icon from "../components/ui/Icon";

interface Profile {
  clientDpa: number;
  clientName: string;
  clientIdPass?: string;
  clientCdsNo?: string;
  clientEmail?: string;
  clientCellTel?: string;
  clientOfficeTel?: string;
  clientHomeTel?: string;
  clientAddr?: string;
  clientPAddr?: string;
  clientContact?: string;
  clientBDate?: string;
}

export default function ProfilePage() {
  const { user } = useAuth();
  const [profile, setProfile] = useState<Profile | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [photoUrl, setPhotoUrl] = useState<string | null>(null);

  useEffect(() => {
    api
      .get("/portal/profile")
      .then((r) => setProfile(r.data))
      .catch((err) =>
        setError(err.response?.data?.error || "Failed to load profile."),
      )
      .finally(() => setLoading(false));
  }, []);

  // Load the passport photo (if any) as an authenticated blob.
  useEffect(() => {
    let revoke: string | null = null;
    api
      .get("/portal/profile/photo", { responseType: "blob" })
      .then((r) => {
        const url = URL.createObjectURL(r.data as Blob);
        revoke = url;
        setPhotoUrl(url);
      })
      .catch(() => setPhotoUrl(null));
    return () => { if (revoke) URL.revokeObjectURL(revoke); };
  }, []);

  if (loading) {
    return (
      <div className="flex h-64 items-center justify-center text-on-surface-variant">
        Loading profile…
      </div>
    );
  }
  if (error) {
    return (
      <div className="rounded-xl border border-axis-error/30 bg-axis-error/5 p-6 text-axis-error">
        {error}
      </div>
    );
  }
  if (!profile) return null;

  const status = user?.status || "Active";

  return (
    <>
      <PageMeta title="Profile" description="View and manage your profile" />

      <div className="space-y-6">
        {/* Hero */}
        <section className="relative overflow-hidden rounded-xl border border-outline-variant bg-surface-container-lowest p-6 shadow-[0px_4px_12px_rgba(15,23,42,0.03)] md:p-8">
          <div className="pointer-events-none absolute right-6 top-1/2 -translate-y-1/2 opacity-[0.06]">
            <Icon name="verified_user" size={150} />
          </div>
          <div className="relative flex flex-col justify-between gap-6 md:flex-row md:items-center">
            <div className="flex items-center gap-5">
              <div className="flex h-20 w-20 shrink-0 items-center justify-center overflow-hidden rounded-full border-4 border-surface bg-primary-container text-2xl font-bold text-on-primary-fixed">
                {photoUrl ? (
                  <img src={photoUrl} alt={profile.clientName} className="h-full w-full object-cover" />
                ) : (
                  (profile.clientName?.[0] ?? "C").toUpperCase()
                )}
              </div>
              <div className="min-w-0">
                <div className="mb-1 flex flex-wrap items-center gap-2">
                  <h2 className="truncate font-display text-2xl font-semibold text-primary">
                    {profile.clientName}
                  </h2>
                  <span className="inline-flex items-center gap-1 rounded-full bg-secondary-container px-2.5 py-0.5 text-[10px] font-bold uppercase tracking-wider text-on-secondary-container">
                    <Icon name="verified" size={14} filled />
                    {status}
                  </span>
                </div>
                <div className="flex flex-wrap items-center gap-3 text-on-surface-variant">
                  <span className="rounded bg-surface-container px-2 py-0.5 font-mono text-xs">
                    ID: #{profile.clientDpa}
                  </span>
                  <span className="h-1 w-1 rounded-full bg-outline" />
                  <span className="text-sm">{user?.email}</span>
                </div>
              </div>
            </div>
            <div className="flex gap-3">
              <Link
                to="/change-password"
                className="rounded-lg bg-primary px-5 py-2.5 text-xs font-semibold uppercase tracking-wide text-on-primary transition-opacity hover:opacity-90"
              >
                Account Settings
              </Link>
            </div>
          </div>
        </section>

        {/* Bento */}
        <div className="grid grid-cols-12 gap-6">
          <div className="col-span-12 space-y-6 lg:col-span-7">
            <Card title="Personal Information" trailing={<Icon name="info" size={18} className="text-on-surface-variant" />}>
              <Field label="Full Name" value={profile.clientName} />
              <Field label="ID / Passport" value={profile.clientIdPass} />
              <Field label="CSD Number" value={profile.clientCdsNo} mono />
              <Field
                label="Date of Birth"
                value={
                  profile.clientBDate
                    ? new Date(profile.clientBDate).toLocaleDateString("en-GB", {
                        day: "2-digit",
                        month: "long",
                        year: "numeric",
                      })
                    : undefined
                }
              />
              <Field label="Contact Person" value={profile.clientContact} />
            </Card>

            <Card title="Contact Details">
              <Field label="Email Address" value={profile.clientEmail} />
              <Field label="Mobile Number" value={profile.clientCellTel} />
              <Field label="Office Number" value={profile.clientOfficeTel} />
              <Field label="Home Number" value={profile.clientHomeTel} />
              <Field label="Physical Address" value={profile.clientAddr} full />
              <Field label="Postal Address" value={profile.clientPAddr} full />
            </Card>
          </div>

          <div className="col-span-12 lg:col-span-5">
            <DocumentsCard />
          </div>
        </div>
      </div>
    </>
  );
}

interface PortalAttachment {
  id: string;
  name: string;
  size: number;
  uploadedAt: string;
  contentType: string;
}

function formatBytes(n: number): string {
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / 1024 / 1024).toFixed(2)} MB`;
}

function DocumentsCard() {
  const [items, setItems] = useState<PortalAttachment[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    api
      .get<PortalAttachment[]>("/portal/attachments")
      .then((r) => {
        if (!cancelled) setItems(r.data ?? []);
      })
      .catch((err) => {
        if (!cancelled)
          setError(err.response?.data?.error || "Failed to load documents.");
      })
      .finally(() => {
        if (!cancelled) setLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, []);

  async function download(a: PortalAttachment) {
    try {
      const res = await api.get(
        `/portal/attachments/${encodeURIComponent(a.id)}`,
        { responseType: "blob" },
      );
      const url = window.URL.createObjectURL(res.data as Blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = a.name;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);
    } catch (err) {
      const e = err as { message?: string };
      setError(e.message ?? "Download failed.");
    }
  }

  return (
    <div className="flex h-full flex-col overflow-hidden rounded-xl border border-outline-variant bg-surface-container-lowest shadow-[0px_4px_12px_rgba(15,23,42,0.03)]">
      <div className="flex items-center justify-between border-b border-outline-variant bg-surface-container-low px-6 py-4">
        <h3 className="font-display text-lg font-semibold text-primary">
          Document Management (KYC)
        </h3>
        <span className="text-xs font-semibold text-secondary">
          {items.length} on file
        </span>
      </div>

      <div className="flex-1 space-y-3 p-6">
        {error && (
          <div className="rounded-lg border border-axis-error/30 bg-axis-error/5 px-3 py-2 text-xs text-axis-error">
            {error}
          </div>
        )}

        {loading ? (
          <p className="text-sm text-on-surface-variant">Loading documents…</p>
        ) : items.length === 0 ? (
          <div className="rounded-lg border border-dashed border-outline-variant bg-surface-container-low px-4 py-10 text-center text-sm text-on-surface-variant">
            No documents on file yet. Contact your account manager if anything is
            missing.
          </div>
        ) : (
          items.map((a) => (
            <div
              key={a.id}
              className="group flex items-center justify-between rounded-lg border border-outline-variant p-3 transition-colors hover:border-secondary"
            >
              <div className="flex min-w-0 items-center gap-3">
                <div className="flex h-10 w-10 shrink-0 items-center justify-center rounded bg-primary-fixed-dim text-primary">
                  <Icon name="description" size={20} />
                </div>
                <div className="min-w-0">
                  <button
                    type="button"
                    onClick={() => download(a)}
                    className="block truncate text-left text-sm font-bold text-on-surface hover:text-secondary"
                    title={a.name}
                  >
                    {a.name}
                  </button>
                  <p className="font-mono text-[11px] text-outline">
                    {formatBytes(a.size)} ·{" "}
                    {new Date(a.uploadedAt).toLocaleDateString("en-GB", {
                      day: "2-digit",
                      month: "short",
                      year: "numeric",
                    })}
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => download(a)}
                aria-label={`Download ${a.name}`}
                className="rounded-full p-2 text-on-surface-variant transition-colors hover:bg-surface-container"
              >
                <Icon name="download" size={20} />
              </button>
            </div>
          ))
        )}
      </div>
    </div>
  );
}

function Card({
  title,
  trailing,
  children,
}: {
  title: string;
  trailing?: React.ReactNode;
  children: React.ReactNode;
}) {
  return (
    <div className="overflow-hidden rounded-xl border border-outline-variant bg-surface-container-lowest shadow-[0px_4px_12px_rgba(15,23,42,0.03)]">
      <div className="flex items-center justify-between border-b border-outline-variant px-6 py-4">
        <h3 className="font-display text-lg font-semibold text-primary">{title}</h3>
        {trailing}
      </div>
      <dl className="grid grid-cols-1 gap-x-6 gap-y-5 p-6 md:grid-cols-2">
        {children}
      </dl>
    </div>
  );
}

function Field({
  label,
  value,
  mono,
  full,
}: {
  label: string;
  value?: string | null;
  mono?: boolean;
  full?: boolean;
}) {
  return (
    <div className={full ? "md:col-span-2" : undefined}>
      <dt className="mb-1 text-[11px] font-semibold uppercase tracking-wider text-outline">
        {label}
      </dt>
      <dd
        className={`text-on-surface ${mono ? "font-mono text-sm" : "text-base"}`}
      >
        {value || <span className="italic text-outline">Not provided</span>}
      </dd>
    </div>
  );
}
