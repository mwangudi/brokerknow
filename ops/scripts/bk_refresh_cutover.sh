#!/usr/bin/env bash
# bk_refresh_cutover.sh <BuildDb> <Tag> CONFIRM
#   e.g. bk_refresh_cutover.sh BrokerKnow_Malawi0918 0918 CONFIRM
#
# Promotes <BuildDb> to axis_db_prod by RENAME (never a restore over live) and keeps
# the outgoing database as axis_db_prod_pre<Tag> for rollback. Aborts before touching
# anything if a guard fails; rolls back automatically if the API does not come back.
set -euo pipefail

BUILD="${1:?usage: bk_refresh_cutover.sh <BuildDb> <Tag> CONFIRM}"
TAG="${2:?missing tag, e.g. 0918}"
[ "${3:-}" = "CONFIRM" ] || { echo "refusing: pass CONFIRM as the third argument" >&2; exit 2; }

LIVE=axis_db_prod
PRE="axis_db_prod_pre${TAG}"
UNIT=brokerknow-api
SQLCMD=/opt/mssql-tools18/bin/sqlcmd
CFG=/opt/brokerknow/api/appsettings.json
SA_PWD="$(grep -oiE 'password=[^;\"]+' "$CFG" | head -1 | cut -d= -f2-)"
run() { "$SQLCMD" -S localhost -U sa -P "$SA_PWD" -C -I -b -W "$@"; }

echo "== guards =="
run -h -1 -Q "
SET NOCOUNT ON;
IF DB_ID('${BUILD}') IS NULL RAISERROR('build db missing',16,1);
IF DB_ID('${PRE}')   IS NOT NULL RAISERROR('${PRE} already exists - pick another tag',16,1);
DECLARE @lu int = (SELECT COUNT(*) FROM ${LIVE}.dbo.PortalUsers);
DECLARE @bu int = (SELECT COUNT(*) FROM ${BUILD}.dbo.PortalUsers);
DECLARE @lc int = (SELECT COUNT(*) FROM ${LIVE}.dbo.Client);
DECLARE @bc int = (SELECT COUNT(*) FROM ${BUILD}.dbo.Client);
DECLARE @lcds int = (SELECT COUNT(*) FROM ${LIVE}.dbo.Client WHERE ISNULL(ClientCDSNo,'') <> '');
DECLARE @bcds int = (SELECT COUNT(*) FROM ${BUILD}.dbo.Client WHERE ISNULL(ClientCDSNo,'') <> '');
IF @bu <> @lu   RAISERROR('portal users %d != live %d',16,1,@bu,@lu);
IF @bcds < @lcds RAISERROR('CDS regression: build %d < live %d',16,1,@bcds,@lcds);
IF @bc  < @lc   RAISERROR('client regression: build %d < live %d',16,1,@bc,@lc);
SELECT 'guards ok  users=' + CONVERT(varchar(10),@bu)
     + '  clients=' + CONVERT(varchar(10),@bc)
     + '  cds=' + CONVERT(varchar(10),@bcds);"

echo "== identity gate =="
if [ -f /tmp/bk/refresh_identity_check.sql ]; then
  run -d "$LIVE" -i /tmp/bk/refresh_identity_check.sql -v BuildDb="$BUILD"
else
  echo "   WARN: /tmp/bk/refresh_identity_check.sql missing — gate skipped"
fi

echo "== stopping ${UNIT} =="
systemctl stop "$UNIT"

echo "== rename ${LIVE} -> ${PRE}, ${BUILD} -> ${LIVE} =="
run -h -1 -Q "
ALTER DATABASE [${LIVE}]  SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
ALTER DATABASE [${LIVE}]  MODIFY NAME = [${PRE}];
ALTER DATABASE [${PRE}]   SET MULTI_USER;
ALTER DATABASE [${BUILD}] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
ALTER DATABASE [${BUILD}] MODIFY NAME = [${LIVE}];
ALTER DATABASE [${LIVE}]  SET MULTI_USER;
SELECT 'renamed';"

echo "== starting ${UNIT} =="
systemctl start "$UNIT"
sleep 6
code="$(curl -s -o /dev/null -w '%{http_code}' -X POST http://127.0.0.1:5260/api/auth/forgot-password \
        -H 'Content-Type: application/json' --data '{"identifier":"zz-healthprobe-zz"}' || echo 000)"
echo "   is-active=$(systemctl is-active "$UNIT")  probe=${code}"

if [ "$code" != "200" ]; then
  echo "!! API unhealthy — ROLLING BACK" >&2
  systemctl stop "$UNIT" || true
  run -h -1 -Q "
  ALTER DATABASE [${LIVE}] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
  ALTER DATABASE [${LIVE}] MODIFY NAME = [${BUILD}];
  ALTER DATABASE [${BUILD}] SET MULTI_USER;
  ALTER DATABASE [${PRE}]  SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
  ALTER DATABASE [${PRE}]  MODIFY NAME = [${LIVE}];
  ALTER DATABASE [${LIVE}] SET MULTI_USER;
  SELECT 'rolled back';"
  systemctl start "$UNIT"
  exit 1
fi

echo "CUTOVER_COMPLETE — rollback available as ${PRE}"
