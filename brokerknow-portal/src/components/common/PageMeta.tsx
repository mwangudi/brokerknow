import { HelmetProvider, Helmet } from "react-helmet-async";

// Non-production builds (VITE_ENV_LABEL set, e.g. "TEST") prefix the browser
// tab so staff can tell a test tab from a live one at a glance.
const ENV_LABEL = (import.meta.env.VITE_ENV_LABEL as string) || "";
const ENV_PREFIX = ENV_LABEL ? `[${ENV_LABEL}] ` : "";

const PageMeta = ({
  title,
  description,
}: {
  title: string;
  description: string;
}) => (
  <Helmet>
    <title>{ENV_PREFIX + (title ? `${title} · Axis` : "Axis")}</title>
    <meta name="description" content={description} />
  </Helmet>
);

export const AppWrapper = ({ children }: { children: React.ReactNode }) => (
  <HelmetProvider>{children}</HelmetProvider>
);

export default PageMeta;
