#!/usr/bin/env bash
# deploy_api.sh <stamp> — swap in /tmp/api-publish-<stamp>.tgz for test then prod.
# Backs up the current install and PRESERVES each environment's appsettings.json:
# the published tarball carries the developer's localdb config, which kills the
# service on Linux ("LocalDB is not supported on this platform").
set -euo pipefail

STAMP="${1:?usage: deploy_api.sh <stamp>}"
TARBALL="/tmp/api-publish-${STAMP}.tgz"
[ -f "$TARBALL" ] || { echo "ERROR: missing $TARBALL" >&2; exit 1; }

deploy_one() {
  root="$1"; unit="$2"; api="${root}/api"
  echo "== ${unit}  (${api})"
  if [ ! -d "$api" ]; then echo "   no install, skipping"; return 0; fi

  keep="$(mktemp -d)"
  for f in appsettings.json appsettings.Production.json; do
    [ -f "$api/$f" ] && cp -p "$api/$f" "$keep/$f"
  done
  # Refuse to continue if we could not capture the live config.
  if [ ! -s "$keep/appsettings.json" ]; then
    echo "   ERROR: no appsettings.json captured from $api — aborting" >&2
    rm -rf "$keep"; return 1
  fi

  cp -a "$api" "${root}/api.bak-${STAMP}"
  rm -rf "${api:?}/"*
  tar -xzf "$TARBALL" -C "$api"
  for f in appsettings.json appsettings.Production.json; do
    [ -f "$keep/$f" ] && cp -p "$keep/$f" "$api/$f"
  done
  rm -rf "$keep"

  # Sanity: the restored config must not be the developer's localdb one.
  if grep -qi 'localdb' "$api/appsettings.json"; then
    echo "   ERROR: appsettings.json still points at LocalDB — rolling back" >&2
    rm -rf "${api:?}/"*; cp -a "${root}/api.bak-${STAMP}/." "$api/"
    systemctl restart "$unit" || true
    return 1
  fi

  systemctl restart "$unit"
  sleep 4
  echo "   is-active=$(systemctl is-active "$unit")   backup=${root}/api.bak-${STAMP}"
}

deploy_one /opt/brokerknow-test brokerknow-api-test
deploy_one /opt/brokerknow      brokerknow-api
echo "DEPLOY_DONE"
