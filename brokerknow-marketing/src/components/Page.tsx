import { Helmet } from "react-helmet-async";

export default function Page({
  title,
  description,
  children,
}: {
  title: string;
  description?: string;
  children: React.ReactNode;
}) {
  return (
    <>
      <Helmet>
        <title>{`${title} — Cedar Capital`}</title>
        {description && <meta name="description" content={description} />}
      </Helmet>
      {children}
    </>
  );
}
