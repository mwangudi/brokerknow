import { Link } from "react-router";
import { NAV, SITE } from "../data/site";

export default function Footer() {
  return (
    <footer className="mt-16 bg-brand-900 text-brand-100">
      <div className="mx-auto grid max-w-6xl gap-10 px-4 py-12 sm:grid-cols-2 lg:grid-cols-4 lg:px-6">
        <div>
          <img
            src="/images/logo.png"
            alt="Cedar Capital"
            className="mb-4 h-12 w-auto brightness-0 invert"
          />
          <p className="text-sm text-brand-200">
            Malawi&rsquo;s leading private financial services firm — stockbroking, research and corporate advisory.
          </p>
        </div>

        <div>
          <h3 className="mb-3 text-sm font-semibold uppercase tracking-wide text-white">Navigate</h3>
          <ul className="space-y-2 text-sm">
            {NAV.map((n) => (
              <li key={n.to}>
                <Link className="text-brand-200 hover:text-white" to={n.to}>
                  {n.label}
                </Link>
              </li>
            ))}
          </ul>
        </div>

        <div>
          <h3 className="mb-3 text-sm font-semibold uppercase tracking-wide text-white">Clients</h3>
          <ul className="space-y-2 text-sm">
            <li>
              <a className="text-brand-200 hover:text-white" href={SITE.portalUrl}>
                Client Portal
              </a>
            </li>
            <li>
              <Link className="text-brand-200 hover:text-white" to="/forms">
                Account Opening Forms
              </Link>
            </li>
            <li>
              <Link className="text-brand-200 hover:text-white" to="/weekly-reports">
                Weekly Market Reports
              </Link>
            </li>
          </ul>
        </div>

        <div>
          <h3 className="mb-3 text-sm font-semibold uppercase tracking-wide text-white">Contact</h3>
          <address className="space-y-1 text-sm not-italic text-brand-200">
            <div>{SITE.address.line1}</div>
            <div>{SITE.address.line2}</div>
            <div>{SITE.address.city}, {SITE.address.country}</div>
            <div className="pt-2">
              <a className="hover:text-white" href={`mailto:${SITE.email}`}>{SITE.email}</a>
            </div>
            <div>
              <a className="hover:text-white" href={`tel:${SITE.phoneTel}`}>{SITE.phone}</a>
            </div>
          </address>
        </div>
      </div>

      <div className="border-t border-brand-800">
        <div className="mx-auto flex max-w-6xl flex-col items-center justify-between gap-2 px-4 py-4 text-xs text-brand-300 sm:flex-row lg:px-6">
          <div>&copy; {new Date().getFullYear()} Cedar Capital Limited. All rights reserved.</div>
          <div>Licensed Broker/Dealer · Member of the Malawi Stock Exchange</div>
        </div>
      </div>
    </footer>
  );
}
