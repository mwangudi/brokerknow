#!/usr/bin/env bash
# Snapshot the admin web roots before a bundle deploy.
set -euo pipefail
STAMP="${1:?usage: backup_web.sh <stamp>}"
for d in admin-host test-admin admin; do
  src="/var/www/${d}"
  if [ -d "$src" ]; then
    cp -a "$src" "/var/www/${d}.bak-${STAMP}"
    echo "backed up ${src} -> /var/www/${d}.bak-${STAMP}"
  fi
done
ls -1d /var/www/*.bak-"${STAMP}"
