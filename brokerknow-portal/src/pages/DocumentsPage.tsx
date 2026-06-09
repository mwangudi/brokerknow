import { useEffect, useState } from "react";
import api from "../lib/api";

interface DocRow {
  id: string;
  name: string;
  size: number;
  uploadedAt: string;
  contentType: string;
}

const fmtDate = (iso: string) =>
  new Date(iso).toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });

const fmtSize = (bytes: number) => {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
};

/** Strip the leading category prefix ("id-", "address-", "funds-", "other-") for a friendly label. */
const prettyName = (name: string) =>
  name.replace(/^(id|address|funds|other)-/i, "").trim() || name;

export default function DocumentsPage() {
  const [docs, setDocs] = useState<DocRow[]>([]);
  const [loading, setLoading] = useState(true);
  const [downloadError, setDownloadError] = useState<string | null>(null);

  useEffect(() => {
    setLoading(true);
    api
      .get<DocRow[]>("/portal/attachments")
      .then((r) => setDocs(r.data))
      .catch(() => {/* leave empty */})
      .finally(() => setLoading(false));
  }, []);

  function download(doc: DocRow) {
    setDownloadError(null);
    api
      .get(`/portal/attachments/${encodeURIComponent(doc.id)}`, { responseType: "blob" })
      .then((r) => {
        const url = window.URL.createObjectURL(new Blob([r.data], { type: doc.contentType }));
        const link = document.createElement("a");
        link.href = url;
        link.setAttribute("download", prettyName(doc.name));
        document.body.appendChild(link);
        link.click();
        link.remove();
        window.URL.revokeObjectURL(url);
      })
      .catch(() => setDownloadError("Could not download that document. Please try again."));
  }

  return (
    <div>
      <div className="mb-6">
        <h2 className="text-2xl font-bold text-gray-900">My Documents</h2>
        <p className="mt-1 text-sm text-gray-500">
          The KYC documents on file for your account, including those you submitted at registration.
        </p>
      </div>

      {downloadError && (
        <div className="mb-4 rounded-lg border border-rose-200 bg-rose-50 px-4 py-3 text-sm text-rose-700">
          {downloadError}
        </div>
      )}

      <div className="rounded-xl border border-gray-200 bg-white">
        <div className="overflow-x-auto">
          <table className="min-w-full text-sm">
            <thead>
              <tr className="border-b border-gray-100 text-left text-xs font-medium uppercase text-gray-500">
                <th className="px-4 py-3">Document</th>
                <th className="px-4 py-3">Uploaded</th>
                <th className="px-4 py-3">Size</th>
                <th className="px-4 py-3"></th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {loading ? (
                <tr><td colSpan={4} className="px-4 py-8 text-center text-gray-500">Loading...</td></tr>
              ) : docs.length === 0 ? (
                <tr><td colSpan={4} className="px-4 py-8 text-center text-gray-500">No documents on file yet.</td></tr>
              ) : docs.map((d, i) => (
                <tr key={d.id} className={i % 2 === 0 ? "bg-blue-50/50" : ""}>
                  <td className="px-4 py-2.5 font-medium text-gray-900">{prettyName(d.name)}</td>
                  <td className="px-4 py-2.5 text-gray-700">{fmtDate(d.uploadedAt)}</td>
                  <td className="px-4 py-2.5 text-gray-500">{fmtSize(d.size)}</td>
                  <td className="px-4 py-2.5 text-right">
                    <button
                      onClick={() => download(d)}
                      className="text-sm font-medium text-blue-600 hover:underline"
                    >
                      Download
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}
