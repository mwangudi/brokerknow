#!/usr/bin/env bash
# Deploy the remaining SPA roots that were stale. Each root gets the bundle
# built for its own base path and tenant mode - they are NOT interchangeable.
set -uo pipefail
S="${1:?stamp}"

deploy() {
  local tgz="/tmp/$1-$S.tgz" dir="$2" expect="$3"
  if [ ! -f "$tgz" ]; then echo "  SKIP $dir (missing $tgz)"; return; fi
  if [ ! -d "$dir" ]; then echo "  SKIP $dir (no such root)"; return; fi
  cp -a "$dir" "${dir}.bak-${S}"
  ls -dt "${dir}".bak-* 2>/dev/null | tail -n +4 | xargs -r rm -rf
  rm -rf "$dir/assets" "$dir/index.html"
  tar xzf "$tgz" -C "$dir"
  chown -R www-data:www-data "$dir"
  local got
  got=$(grep -oE 'src="[^"]*/assets/' "$dir/index.html" | head -1 | sed 's|src="||; s|assets/||')
  [ -z "$got" ] && got="/"
  if [ "$got" = "$expect" ]; then
    echo "  OK   $(basename "$dir")  base=$got"
  else
    echo "  WRONG BASE $(basename "$dir")  expected=$expect got=$got  -- rolling back"
    rm -rf "$dir/assets" "$dir/index.html"
    cp -a "${dir}.bak-${S}/." "$dir/"
    chown -R www-data:www-data "$dir"
  fi
}

echo "== agent portal (live surface, was 2 months stale) =="
deploy agent      /var/www/agent-host       /agent/
deploy agent      /var/www/test-agent       /agent/
deploy cedaragent /var/www/cedaragent       /

echo
echo "== demo tenants: admin =="
deploy keadmin    /var/www/kenya-admin      /ke/admin/
deploy rwadmin    /var/www/rwanda-admin     /rw/admin/
deploy rwtadmin   /var/www/rwandatest-admin /admin/

echo
echo "== demo tenants: client portal =="
deploy keportal   /var/www/kenya-portal     /ke/
deploy rwportal   /var/www/rwanda-portal    /rw/
deploy rwtportal  /var/www/rwandatest-portal /

echo
nginx -t 2>&1 | tail -2
