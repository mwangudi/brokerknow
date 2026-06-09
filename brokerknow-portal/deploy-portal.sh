#!/usr/bin/env bash
# Deploy the client portal SPA to TEST and/or PROD web roots.
# Usage on droplet: bash /tmp/deploy-portal.sh <STAMP> test|prod|both
set -euo pipefail

STAMP="${1:?stamp required}"
TARGET="${2:-both}"
TGZ="/tmp/portal-${STAMP}.tgz"

[[ -f "$TGZ" ]] || { echo "missing $TGZ"; exit 1; }

snapshot_and_deploy() {
  local dest="$1"
  echo "== ${dest} <- $(basename "$TGZ")"
  if [[ -d "$dest" ]]; then
    cp -a "$dest" "${dest}.bak-${STAMP}"
    echo "   snapshot -> ${dest}.bak-${STAMP}"
  fi
  rm -rf "${dest:?}"/assets
  rm -f  "${dest:?}"/index.html "${dest:?}"/favicon.svg "${dest:?}"/favicon.png "${dest:?}"/icons.svg
  mkdir -p "$dest"
  tar -xzf "$TGZ" -C "$dest"
  chown -R www-data:www-data "$dest"
  echo "   index.html: $(test -f "$dest/index.html" && echo yes || echo NO), martens-logo: $(test -f "$dest/images/logo/martens-logo.png" && echo yes || echo NO)"
}

case "$TARGET" in
  test) snapshot_and_deploy /var/www/test-portal ;;
  prod) snapshot_and_deploy /var/www/portal ;;
  both)
    snapshot_and_deploy /var/www/test-portal
    snapshot_and_deploy /var/www/portal
    ;;
  *) echo "Unknown target: $TARGET (use test|prod|both)"; exit 2 ;;
esac

nginx -t >/dev/null 2>&1 && echo "nginx config OK" || { echo "nginx config FAILED"; exit 3; }
echo "DONE"
