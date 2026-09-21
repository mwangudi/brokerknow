#!/usr/bin/env bash
# Extract the admin bundle into a host-based web root (basename=/ build).
# Existing hashed assets are left in place so the previous bundle stays loadable.
set -euo pipefail
STAMP="${1:?usage: deploy_web.sh <stamp> <root>...}"
shift
TARBALL="/tmp/admin-dist-${STAMP}.tgz"
[ -f "$TARBALL" ] || { echo "ERROR: missing $TARBALL" >&2; exit 1; }

for root in "$@"; do
  [ -d "$root" ] || { echo "skip ${root} (missing)"; continue; }
  tar -xzf "$TARBALL" -C "$root"
  printf '%-26s now serving %s\n' "$root" \
    "$(grep -oE 'index-[A-Za-z0-9_-]+\.js' "$root/index.html" | head -1)"
done
