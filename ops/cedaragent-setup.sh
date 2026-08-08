#!/usr/bin/env bash
# One-shot setup for cedaragent.martensafrica.com (dedicated agent portal host).
#
# BLOCKED until the DNS A record exists. Run this only after adding, in cPanel
# -> Zone Editor -> martensafrica.com:   A   cedaragent   46.101.6.131   TTL 300
#
# It refuses to touch certbot until ALL FOUR authoritative nameservers agree, so
# a half-replicated record cannot burn Let's Encrypt's 5-failures-per-hour limit.
#
# Usage: bash cedaragent-setup.sh
set -euo pipefail

HOST=cedaragent.martensafrica.com
IP=46.101.6.131
ROOT=/var/www/cedaragent
CONF=/etc/nginx/sites-enabled/brokerknow
EMAIL=mutinyud@martensafrica.com

echo "== 1. DNS guard: all four authoritative NS must return $IP =="
ok=0
for ns in ns1 ns2 ns3 ns4; do
  got=$(dig +short @${ns}.supercp.com "$HOST" A 2>/dev/null | head -1)
  printf "   %-16s -> %s\n" "${ns}.supercp.com" "${got:-(none)}"
  [ "$got" = "$IP" ] && ok=$((ok+1))
done
if [ "$ok" -ne 4 ]; then
  echo
  echo "   STOP: only $ok/4 nameservers resolve $HOST to $IP."
  echo "   Add the A record (or wait for supercp replication) and re-run."
  exit 1
fi
echo "   all 4 agree"

echo
echo "== 2. document root =="
if [ ! -f "$ROOT/index.html" ]; then
  echo "   STOP: $ROOT/index.html missing. Upload the root-based agent build first:"
  echo "     brokerknow-web: VITE_AUDIENCE=agent VITE_BASENAME=/ npm run build -- --base=/ --outDir dist-agent-root"
  exit 1
fi
echo "   $ROOT ok"

echo
echo "== 3. nginx vhost =="
if grep -q "server_name $HOST;" "$CONF"; then
  echo "   already present, skipping"
else
  cp -a "$CONF" "/root/nginx-backups/brokerknow.$(date +%Y%m%d%H%M%S).pre-cedaragent"
  cat >> "$CONF" <<EOF

# =============== PROD: $HOST (agent portal, host-rooted) ===============
# Agent SPA at /, API -> :5260 (same backend the agents use today at /agent/).
server {
    listen 80;
    server_name $HOST;
    client_max_body_size 50M;

    location /api/ {
        proxy_pass         http://127.0.0.1:5260;
        proxy_http_version 1.1;
        proxy_set_header   Host              \$host;
        proxy_set_header   X-Real-IP         \$remote_addr;
        proxy_set_header   X-Forwarded-For   \$proxy_add_x_forwarded_for;
        proxy_set_header   X-Forwarded-Proto \$scheme;
        proxy_set_header   Upgrade           \$http_upgrade;
        proxy_set_header   Connection        keep-alive;
        proxy_read_timeout 300s;
    }

    location /assets/ {
        root $ROOT;
        expires 1y;
        add_header Cache-Control "public, immutable";
        access_log off;
        try_files \$uri =404;
    }

    location / {
        root $ROOT;
        try_files \$uri \$uri/ /index.html;
        add_header Cache-Control "no-cache" always;
    }
}
EOF
  if ! nginx -t; then
    echo "   nginx -t failed, rolling back"
    cp -a "$(ls -t /root/nginx-backups/brokerknow.*.pre-cedaragent | head -1)" "$CONF"
    exit 1
  fi
  systemctl reload nginx
  echo "   vhost added (HTTP)"
fi

echo
echo "== 4. TLS =="
if [ -d "/etc/letsencrypt/live/$HOST" ]; then
  echo "   certificate already issued, skipping"
else
  certbot --nginx -d "$HOST" --non-interactive --agree-tos -m "$EMAIL" --redirect
fi

echo
echo "== 5. verify =="
R="--resolve $HOST:443:$IP"
printf "   https://%s/          -> %s\n" "$HOST" "$(curl -s $R -o /dev/null -w '%{http_code}' https://$HOST/)"
printf "   https://%s/api/ping  -> %s\n" "$HOST" "$(curl -s $R -o /dev/null -w '%{http_code}' https://$HOST/api/ping)"
printf "   http -> https        -> %s\n" "$(curl -s -o /dev/null -w '%{http_code}' --resolve $HOST:80:$IP http://$HOST/)"
echo
echo "DONE. Agents sign in at https://$HOST/"
