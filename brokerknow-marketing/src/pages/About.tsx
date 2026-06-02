import Page from "../components/Page";

export default function About() {
  return (
    <Page
      title="About Us"
      description="Cedar Capital was formed following the divestment of African Alliance from Malawi. Led by Armstrong Kamphoni, the firm is the most sought-after broker on the Malawi Stock Exchange."
    >
      <section className="bg-brand-50 py-12 lg:py-16">
        <div className="mx-auto max-w-6xl px-4 lg:px-6">
          <h1 className="text-4xl font-bold text-brand-800 lg:text-5xl">About Cedar Capital</h1>
          <p className="mt-3 max-w-3xl text-brand-700">
            Two decades of stockbroking experience, deep market relationships and an
            unwavering focus on client outcomes.
          </p>
        </div>
      </section>

      <section className="mx-auto max-w-6xl px-4 py-12 lg:px-6 lg:py-16">
        <div className="grid gap-10 lg:grid-cols-3">
          <article className="prose prose-slate max-w-none lg:col-span-2">
            <p>
              Cedar Capital Ltd was formed following the divestment from Malawi of
              African Alliance. The company is managed by <strong>Armstrong Kamphoni</strong>{" "}
              &mdash; a seasoned investment banker and stockbroker in the Malawi
              financial market &mdash; supported by a team of staff well versed in
              international standards of stockbroking.
            </p>
            <p>
              Cedar Capital&rsquo;s board ensures best international practice is
              adhered to, with a diverse group of directors from within Malawi and
              the region.
            </p>
            <p>
              In terms of market share and innovation, Cedar Capital is the leader in
              the Malawi market. Our service offering includes stockbroking,
              corporate advisory, research and personal investment advice. The
              company enjoys strong links with international brokerage houses making
              it the most sought-after broker in matching and executing high-value
              trades in the market.
            </p>
            <p>
              Cedar Capital seeks to forge long-term client and counter-party
              relationships based on trust and service delivery. Our processes and
              procedures, and the combined experience from our energetic and
              client-centric dealing staff, are integral to delivering world-class
              stockbroking services.
            </p>
            <p>
              We aim to provide the most effective and efficient service available,
              tailor-made to meet individual needs of our institutional and
              high-net-worth clients. Our objective is always to create an
              environment in which our clients get the best execution and achieve
              strong performance from their investments. We execute trades in the
              best interest of our clients, in a price-sensitive and confidential
              manner.
            </p>
            <p>
              Cedar Capital is the sole provider of regular research on listed
              companies on the Malawi Stock Exchange and issues a weekly market
              commentary together with an economic and political overview.
            </p>
            <p className="text-sm text-slate-600">
              Cedar Capital Limited is registered in Malawi under the Companies Act
              2013, is a licensed Broker/Dealer by the Registrar of Financial
              Institutions under the Securities Act 2010, and a registered member of
              the Malawi Stock Exchange.
            </p>
          </article>

          <aside className="rounded-lg border border-slate-200 bg-white p-6 shadow-sm">
            {/* TODO replace with high-res portrait of Armstrong Kamphoni */}
            <img
              src="/images/armstrong-kamphoni.jpg"
              alt="Armstrong Kamphoni"
              className="mb-4 aspect-[4/5] w-full rounded-md object-cover"
            />
            <h3 className="text-lg font-semibold text-brand-800">Armstrong Kamphoni</h3>
            <p className="text-sm font-medium text-slate-600">Managing Director</p>
            <p className="mt-3 text-sm text-slate-700">
              A seasoned investment banker and stockbroker who has led Cedar Capital
              since its founding, with extensive experience across the Malawi and
              regional capital markets.
            </p>
          </aside>
        </div>
      </section>
    </Page>
  );
}
