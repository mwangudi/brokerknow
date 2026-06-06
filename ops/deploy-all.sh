#!/usr/bin/env bash
# Deploy everything: API (test+prod), Admin (test+prod+legacy), Portal (prod+test).
set -euo pipefail
STAMP="${1:?stamp required}"

echo "== STAMP=$STAMP"

# ---- API ----
for E in test prod; do
  if [ "$E" = "test" ]; then
    ROOT=/opt/brokerknow-test; SVC=brokerknow-api-test; PORT=5261
  else
    ROOT=/opt/brokerknow; SVC=brokerknow-api; PORT=5260
  fi
  echo "== API: $E ($SVC on :$PORT)"
  cp "$ROOT/api/appsettings.json" "/tmp/${E}-appsettings.json"
  cp -ra "$ROOT/api" "$ROOT/api.bak-${STAMP}"
  rm -rf "$ROOT/api"/*
  tar -xzf "/tmp/api-publish-${STAMP}.tgz" -C "$ROOT/api"
  cp "/tmp/${E}-appsettings.json" "$ROOT/api/appsettings.json"
  systemctl restart "$SVC"
  sleep 4
  H=$(curl -s -o /dev/null -w '%{http_code}' "http://127.0.0.1:${PORT}/health")
  echo "   $E health=$H status=$(systemctl is-active $SVC)"
done

# ---- Admin web ----
bash /tmp/deploy-admin.sh "$STAMP" both

# ---- Portal SPA ----
for D in /var/www/portal /var/www/test-portal; do
  echo "== Portal: $D"
  if [ -d "$D" ]; then
    cp -ra "$D" "${D}.bak-${STAMP}"
  fi
  rm -rf "$D"/assets "$D"/index.html "$D"/favicon.svg "$D"/favicon.png
  mkdir -p "$D"
  tar -xzf "/tmp/portal-${STAMP}.tgz" -C "$D"
  chown -R www-data:www-data "$D"
done

nginx -t >/dev/null 2>&1 && echo "nginx OK" || { echo "nginx FAILED"; exit 3; }
echo "ALL DONE"
