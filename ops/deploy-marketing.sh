#!/usr/bin/env bash
set -euo pipefail

WEBROOT=/var/www/marketing
SITES_AVAILABLE=/etc/nginx/sites-available/marketing
SITES_ENABLED=/etc/nginx/sites-enabled/marketing
BUNDLE=${1:-/tmp/marketing.tgz}

echo "=== 1. Deploy bundle ==="
mkdir -p "$WEBROOT"
rm -rf "$WEBROOT"/*
tar -xzf "$BUNDLE" -C "$WEBROOT"
chown -R www-data:www-data "$WEBROOT"
ls -la "$WEBROOT" | head

echo "=== 2. Install nginx vhost ==="
cat > "$SITES_AVAILABLE" <<'EOF'
# Marketing site (cedarcapital.mw + www)
server {
    listen 80;
    server_name cedarcapital.mw www.cedarcapital.mw marketing.test;

    root /var/www/marketing;
    index index.html;

    client_max_body_size 10M;

    # static SPA
    location / {
        try_files $uri $uri/ /index.html;
    }

    # long-cache hashed assets
    location /assets/ {
        expires 30d;
        add_header Cache-Control "public, immutable";
        try_files $uri =404;
    }
}
EOF
ln -sf "$SITES_AVAILABLE" "$SITES_ENABLED"

echo "=== 3. Reload nginx ==="
nginx -t
systemctl reload nginx
echo NGINX-OK

echo "=== 4. Smoke ==="
curl -sS -o /dev/null -w "marketing Host=%{http_code}\n" -H "Host: cedarcapital.mw"     http://127.0.0.1/
curl -sS -o /dev/null -w "marketing www =%{http_code}\n" -H "Host: www.cedarcapital.mw" http://127.0.0.1/
curl -sS -o /dev/null -w "marketing asset=%{http_code}\n" -H "Host: cedarcapital.mw"    http://127.0.0.1/images/logo.png
