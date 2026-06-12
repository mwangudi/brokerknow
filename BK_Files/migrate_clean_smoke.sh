#!/usr/bin/env bash
# Read-only smoke against the clean-DB API on :5263.
# Validates: login (PortalUsers migration + bcrypt + JWT), auth guard, and that
# migrated data is queryable through EF + the app.* views across core endpoints.
B=http://127.0.0.1:5263/api
pass=0; fail=0
note() { printf '  %-6s %s\n' "$1" "$2"; }

jval() { python3 -c "import sys,json
try:
    d=json.load(sys.stdin)
    print(d.get('$1','') if isinstance(d,dict) else '')
except Exception: print('')"; }

jlen() { python3 -c "import sys,json
try:
    d=json.load(sys.stdin)
    if isinstance(d,list): print(len(d))
    elif isinstance(d,dict):
        for k in ('items','data','results','rows'):
            if isinstance(d.get(k),list): print(len(d[k])); break
        else: print('obj')
    else: print('')
except Exception: print('')"; }

echo "=== AUTH ==="
printf '{\"email\":\"admin@cedar.mw\",\"password\":\"Passw0rd\",\"audience\":\"admin\"}' > /tmp/login_clean.json
LOGIN=$(curl -s -w '\n%{http_code}' -X POST "$B/auth/login" -H 'Content-Type: application/json' --data-binary @/tmp/login_clean.json)
CODE=$(printf '%s' "$LOGIN" | tail -1); BODY=$(printf '%s' "$LOGIN" | sed '$d')
TOKEN=$(printf '%s' "$BODY" | jval accessToken)
if [ "$CODE" = "200" ] && [ -n "$TOKEN" ]; then note PASS "POST /auth/login (admin@cedar.mw) -> 200, token len=${#TOKEN}"; pass=$((pass+1));
else note FAIL "POST /auth/login -> $CODE"; echo "       $(printf '%s' "$BODY" | head -c 200)"; fail=$((fail+1)); fi

AUTH="Authorization: Bearer $TOKEN"

C=$(curl -s -o /dev/null -w '%{http_code}' "$B/users/online" -H "$AUTH")
if [ "$C" = "200" ]; then note PASS "GET /users/online (with token) -> 200"; pass=$((pass+1)); else note FAIL "GET /users/online (with token) -> $C"; fail=$((fail+1)); fi

C=$(curl -s -o /dev/null -w '%{http_code}' "$B/users/online")
if [ "$C" = "401" ]; then note PASS "GET /users/online (no token) -> 401 (guard works)"; pass=$((pass+1)); else note FAIL "GET /users/online (no token) -> $C (expected 401)"; fail=$((fail+1)); fi

echo "=== READ-THROUGH (migrated data via app endpoints) ==="
for ep in brokers agents securities clients orders payments banks account-managers holidays levies commissions contracts; do
  R=$(curl -s -w '\n%{http_code}' "$B/$ep" -H "$AUTH")
  CODE=$(printf '%s' "$R" | tail -1); BODY=$(printf '%s' "$R" | sed '$d')
  LEN=$(printf '%s' "$BODY" | jlen)
  if [ "$CODE" = "200" ]; then note PASS "GET /$ep -> 200 (len=$LEN)"; pass=$((pass+1));
  else note FAIL "GET /$ep -> $CODE"; fail=$((fail+1)); fi
done

echo "=== DETAIL FETCH (validates joins / app.* views) ==="
FIRST_CLIENT=$(curl -s "$B/clients" -H "$AUTH" | python3 -c "import sys,json
try:
    d=json.load(sys.stdin); a=d if isinstance(d,list) else d.get('items',[])
    print(a[0].get('clientDpa') or a[0].get('id') or a[0].get('clientId') or '')
except Exception: print('')")
if [ -n "$FIRST_CLIENT" ]; then
  C=$(curl -s -o /dev/null -w '%{http_code}' "$B/clients/$FIRST_CLIENT" -H "$AUTH")
  if [ "$C" = "200" ]; then note PASS "GET /clients/$FIRST_CLIENT -> 200"; pass=$((pass+1)); else note FAIL "GET /clients/$FIRST_CLIENT -> $C"; fail=$((fail+1)); fi
else note INFO "no client id parsed; skipping client detail"; fi

FIRST_ORDER=$(curl -s "$B/orders" -H "$AUTH" | python3 -c "import sys,json
try:
    d=json.load(sys.stdin); a=d if isinstance(d,list) else d.get('items',[])
    print(a[0].get('orderDpa') or a[0].get('id') or a[0].get('orderId') or '')
except Exception: print('')")
if [ -n "$FIRST_ORDER" ]; then
  C=$(curl -s -o /dev/null -w '%{http_code}' "$B/orders/$FIRST_ORDER" -H "$AUTH")
  if [ "$C" = "200" ]; then note PASS "GET /orders/$FIRST_ORDER -> 200"; pass=$((pass+1)); else note FAIL "GET /orders/$FIRST_ORDER -> $C"; fail=$((fail+1)); fi
else note INFO "no order id parsed; skipping order detail"; fi

echo "=== RESULT: pass=$pass fail=$fail ==="
