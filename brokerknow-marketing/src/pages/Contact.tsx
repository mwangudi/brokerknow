import Page from "../components/Page";
import { SITE } from "../data/site";

export default function Contact() {
  return (
    <Page
      title="Contact"
      description="Get in touch with Cedar Capital — Livingstone Towers, Blantyre, Malawi."
    >
      <section className="bg-brand-50 py-12 lg:py-16">
        <div className="mx-auto max-w-6xl px-4 lg:px-6">
          <h1 className="text-4xl font-bold text-brand-800 lg:text-5xl">Contact us</h1>
          <p className="mt-3 max-w-2xl text-brand-700">
            Drop in, give us a call or send an email — we&rsquo;d love to hear from you.
          </p>
        </div>
      </section>

      <section className="mx-auto max-w-6xl px-4 py-12 lg:px-6">
        <div className="grid gap-10 lg:grid-cols-2">
          <div>
            <h2 className="text-xl font-semibold text-brand-800">Office</h2>
            <address className="mt-3 space-y-1 not-italic text-slate-700">
              <div>{SITE.address.line1}</div>
              <div>{SITE.address.line2}</div>
              <div>{SITE.address.city}, {SITE.address.country}</div>
            </address>

            <h2 className="mt-8 text-xl font-semibold text-brand-800">Phone</h2>
            <p className="mt-3">
              <a className="text-brand-700 hover:text-brand-900" href={`tel:${SITE.phoneTel}`}>
                {SITE.phone}
              </a>
            </p>

            <h2 className="mt-8 text-xl font-semibold text-brand-800">Email</h2>
            <p className="mt-3">
              <a className="text-brand-700 hover:text-brand-900" href={`mailto:${SITE.email}`}>
                {SITE.email}
              </a>
            </p>

            <h2 className="mt-8 text-xl font-semibold text-brand-800">Hours</h2>
            <p className="mt-3 text-slate-700">Monday &ndash; Friday, 08:00 &ndash; 17:00 CAT</p>
          </div>

          <div className="overflow-hidden rounded-lg border border-slate-200 shadow-sm">
            {/* Embedded Google map; replace 'q' with the exact coords if needed */}
            <iframe
              title="Cedar Capital location"
              src="https://www.google.com/maps?q=Livingstone+Towers,+Glyn+Jones+Road,+Blantyre,+Malawi&output=embed"
              className="aspect-square w-full"
              loading="lazy"
              referrerPolicy="no-referrer-when-downgrade"
            />
          </div>
        </div>
      </section>
    </Page>
  );
}
