#!/usr/bin/env bash
# Throwaway: stand up the BrokerKnow API against BrokerKnow_Clean on :5263.
# Clones the test deployment, repoints DB + port, runs as a transient systemd unit.
set -euo pipefail

SRC=/opt/brokerknow-test/api
DST=/opt/brokerknow-clean/api

# stop any prior clean instance first
systemctl stop brokerknow-api-clean 2>/dev/null || true
systemctl reset-failed brokerknow-api-clean 2>/dev/null || true

# fresh copy of the test deployment
rm -rf /opt/brokerknow-clean
mkdir -p /opt/brokerknow-clean
cp -a "$SRC" "$DST"

# repoint: clean DB + port 5263 (both the conn string and the appsettings Urls)
sed -i 's/Database=BrokerKnow_Test/Database=BrokerKnow_Clean/' "$DST/appsettings.json"
sed -i 's#http://127.0.0.1:5261#http://127.0.0.1:5263#' "$DST/appsettings.json"

echo "==== effective conn + urls ===="
grep -E '"BrokerKnow":|"Urls":' "$DST/appsettings.json"

# launch transient unit (Restart=no so a crash surfaces; auto-removed on stop)
systemd-run --unit=brokerknow-api-clean \
  --property=WorkingDirectory="$DST" \
  --property=Restart=no \
  --setenv=ASPNETCORE_ENVIRONMENT=Production \
  --setenv=ASPNETCORE_URLS=http://127.0.0.1:5263 \
  --setenv=DOTNET_PRINT_TELEMETRY_MESSAGE=false \
  /usr/bin/dotnet "$DST/BrokerKnow.Api.dll" >/dev/null

echo "==== waiting for health on :5263 ===="
code=000
for i in $(seq 1 40); do
  code=$(curl -s -o /dev/null -w '%{http_code}' http://127.0.0.1:5263/api/ping || true)
  if [ "$code" = "200" ]; then echo "HEALTH_OK after ${i}s (ping=200)"; break; fi
  sleep 1
done
echo "final ping code: $code"
echo "==== last 25 log lines ===="
journalctl -u brokerknow-api-clean --no-pager -n 25 | sed 's/Password=[^;"]*/Password=***/g'
