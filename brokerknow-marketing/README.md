# Cedar Capital — Marketing Site

Static React + Vite + Tailwind 4 site, sibling of `brokerknow-portal` and
`brokerknow-web`. Deploys to `/var/www/marketing/` on the BrokerKnow droplet
and is served by nginx on `cedarcapital.mw` / `www.cedarcapital.mw`.

## Local dev

```powershell
cd brokerknow-marketing
npm install
npm run dev    # http://localhost:5173
```

## Build & deploy

```powershell
npm run build  # outputs ./dist
tar -czf marketing.tgz -C dist .
scp -O marketing.tgz root@46.101.6.131:/tmp/
ssh root@46.101.6.131 'rm -rf /var/www/marketing/*; tar -xzf /tmp/marketing.tgz -C /var/www/marketing/; chown -R www-data:www-data /var/www/marketing'
```

Then add two `server` blocks to `/etc/nginx/sites-available/brokerknow`:

```nginx
server {
    listen 80;
    server_name cedarcapital.mw www.cedarcapital.mw;
    root /var/www/marketing;
    index index.html;
    location / { try_files $uri $uri/ /index.html; }
}
```

(Certbot will upgrade these to HTTPS once DNS resolves.)

## Content sources

- `src/data/site.ts` — contact info, nav, services.
- `src/data/forms.ts` — account-opening / KYC PDFs.
- `src/data/research.ts` — quarterly / thematic research PDFs.
- `src/data/weekly.ts` — weekly market commentary archive (2018-2025).

All PDFs currently point at the legacy WordPress origin
(`cedarcapital.mw/content/uploads/...`) so the new site can ship without
migrating files. Move PDFs onto a self-hosted `/var/www/marketing/uploads/`
folder when retiring the WP origin.

## Images required from client

Place under `public/images/` with these exact filenames:

| File | Use | Recommended |
|---|---|---|
| `logo.png`             | Header (dark text on white) | SVG/PNG, transparent background, 120-200 px tall |
| `logo-light.png`       | Footer (white/light variant) | SVG/PNG, transparent background |
| `favicon.png`          | Browser tab icon            | 64×64 PNG (or .ico) |
| `hero.jpg`             | Home hero background        | 2000×1100, Blantyre skyline / financial district vibe |
| `about-hero.jpg`       | Home about-section image    | 1200×900, office or Livingstone Towers exterior |
| `armstrong-kamphoni.jpg` | About page portrait        | Square or 4:5 portrait, min 800 px on short edge |

The site falls back gracefully (broken-image alt text) until these exist.

## Pages

| Route | Source |
|---|---|
| `/`               | `src/pages/Home.tsx` |
| `/about`          | `src/pages/About.tsx` |
| `/weekly-reports` | `src/pages/WeeklyReports.tsx` |
| `/research`       | `src/pages/Research.tsx` |
| `/forms`          | `src/pages/Forms.tsx` |
| `/contact`        | `src/pages/Contact.tsx` |
