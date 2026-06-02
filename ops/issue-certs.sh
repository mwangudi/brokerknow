#!/usr/bin/env bash
# Issue Let's Encrypt certs for all BrokerKnow hostnames.
# Pre-req: DNS A records for the hosts below must resolve to this droplet.
# Re-runnable: certbot --nginx is idempotent and handles renewals via systemd timer.
set -euo pipefail

EMAIL="${1:?usage: issue-certs.sh <admin-email> [--staging]}"
STAGING_FLAG=""
[ "${2:-}" = "--staging" ] && STAGING_FLAG="--staging"

HOSTS=(
  portal.cedarcapital.mw
  admin.cedarcapital.mw
  test-portal.cedarcapital.mw
  test-admin.cedarcapital.mw
)

echo "=== 1. Verify DNS resolves to this droplet ==="
THIS_IP=$(curl -s4 https://ifconfig.me || hostname -I | awk '{print $1}')
echo "Droplet IP: $THIS_IP"
TO_ISSUE=()
for H in "${HOSTS[@]}"; do
  RESOLVED=$(dig +short A "$H" | tail -n1)
  if [ "$RESOLVED" = "$THIS_IP" ]; then
    echo "  OK  $H -> $RESOLVED"
    TO_ISSUE+=("$H")
  else
    echo "  SKIP $H -> '$RESOLVED' (expected $THIS_IP)"
  fi
done
if [ "${#TO_ISSUE[@]}" -eq 0 ]; then
  echo "No hosts ready. Aborting."
  exit 1
fi

echo "=== 2. Ensure certbot installed ==="
if ! command -v certbot >/dev/null; then
  apt-get update -qq
  apt-get install -y certbot python3-certbot-nginx
fi

echo "=== 3. Issue certs (one per hostname via --nginx) ==="
DOMAIN_ARGS=()
for H in "${TO_ISSUE[@]}"; do DOMAIN_ARGS+=(-d "$H"); done

certbot --nginx --non-interactive --agree-tos --email "$EMAIL" \
  --redirect $STAGING_FLAG \
  "${DOMAIN_ARGS[@]}"

echo "=== 4. Verify renewal timer ==="
systemctl status certbot.timer --no-pager | head -5 || true
certbot renew --dry-run | tail -10

echo "=== 5. Final smoke ==="
for H in "${TO_ISSUE[@]}"; do
  CODE=$(curl -sS -o /dev/null -w "%{http_code}" "https://$H/")
  echo "  $H  https=$CODE"
done
