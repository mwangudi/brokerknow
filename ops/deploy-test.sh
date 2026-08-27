#!/usr/bin/env bash
# Deploy API + Portal to TEST environments only (no prod touch).
set -euo pipefail
STAMP="${1:?stamp required}"

echo "== TEST-ONLY deploy, STAMP=$STAMP"

# ---- API (test) ----
ROOT=/opt/brokerknow-test; SVC=brokerknow-api-test; PORT=5261
echo "== API: test ($SVC on :$PORT)"
cp "$ROOT/api/appsettings.json" "/tmp/test-appsettings.json"
cp -ra "$ROOT/api" "$ROOT/api.bak-${STAMP}"
# these are ~120 MB each and were never pruned; 104 had piled up by Aug 2026
ls -dt "$ROOT"/api.bak-* 2>/dev/null | tail -n +4 | xargs -r rm -rf
rm -rf "$ROOT/api"/*
tar -xzf "/tmp/api-publish-${STAMP}.tgz" -C "$ROOT/api"
cp "/tmp/test-appsettings.json" "$ROOT/api/appsettings.json"
systemctl restart "$SVC"
sleep 4
echo "   test health=$(curl -s -o /dev/null -w '%{http_code}' http://127.0.0.1:${PORT}/health) status=$(systemctl is-active $SVC)"

# ---- Portal (test only) ----
D=/var/www/test-portal
echo "== Portal: $D"
cp -ra "$D" "${D}.bak-${STAMP}"
ls -dt "${D}".bak-* 2>/dev/null | tail -n +4 | xargs -r rm -rf
rm -rf "$D"/assets "$D"/index.html "$D"/favicon.svg "$D"/favicon.png
mkdir -p "$D"
tar -xzf "/tmp/portal-${STAMP}.tgz" -C "$D"
chown -R www-data:www-data "$D"

nginx -t >/dev/null 2>&1 && echo "nginx OK" || { echo "nginx FAILED"; exit 3; }
echo "TEST DONE"
