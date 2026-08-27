#!/usr/bin/env bash
# Bootstrap test environment on the droplet.
# Idempotent: safe to re-run.
set -euo pipefail

TS="${1:?usage: bootstrap-test.sh <timestamp>}"
# Credentials are not kept in git. See Ops_Runbook.md section 2.
DB_CONF=${BROKERKNOW_DB_CONF:-/etc/brokerknow/db.conf}
[ -r "$DB_CONF" ] && . "$DB_CONF"
SA_PWD="${SA_PWD:?SA_PWD not set - create $DB_CONF containing SA_PWD='...' (chmod 600, root only)}"
PROD_DB='BrokerKnow'
TEST_DB='BrokerKnow_Test'

echo '=== 1. Clone prod DB to test DB ==='
mkdir -p /var/opt/mssql/backup
chown mssql:mssql /var/opt/mssql/backup
/opt/mssql-tools18/bin/sqlcmd -S localhost -U sa -P "$SA_PWD" -C -b -Q "
BACKUP DATABASE [$PROD_DB] TO DISK = N'/var/opt/mssql/backup/bk-clone.bak'
WITH INIT, COMPRESSION, CHECKSUM, FORMAT;
"

# Capture logical file names
mapfile -t FILES < <(/opt/mssql-tools18/bin/sqlcmd -S localhost -U sa -P "$SA_PWD" -C -h -1 -W -s'|' -Q "
SET NOCOUNT ON;
RESTORE FILELISTONLY FROM DISK = N'/var/opt/mssql/backup/bk-clone.bak';
" | awk -F'|' 'NF>2 && $3 ~ /^(D|L)$/ {print $1"|"$3}')

DATA_FILE=$(printf '%s\n' "${FILES[@]}" | awk -F'|' '$2=="D"{print $1; exit}')
LOG_FILE=$(printf '%s\n' "${FILES[@]}" | awk -F'|' '$2=="L"{print $1; exit}')
echo "Data logical: $DATA_FILE"
echo "Log  logical: $LOG_FILE"

/opt/mssql-tools18/bin/sqlcmd -S localhost -U sa -P "$SA_PWD" -C -b -Q "
IF DB_ID('$TEST_DB') IS NOT NULL
BEGIN
  ALTER DATABASE [$TEST_DB] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
END
RESTORE DATABASE [$TEST_DB] FROM DISK = N'/var/opt/mssql/backup/bk-clone.bak'
WITH REPLACE, RECOVERY,
  MOVE N'$DATA_FILE' TO N'/var/opt/mssql/data/${TEST_DB}.mdf',
  MOVE N'$LOG_FILE'  TO N'/var/opt/mssql/data/${TEST_DB}_log.ldf';
ALTER DATABASE [$TEST_DB] SET MULTI_USER;
"
echo "Test DB ready."

echo '=== 2. Mirror /opt/brokerknow/api -> /opt/brokerknow-test/api ==='
systemctl stop brokerknow-api-test 2>/dev/null || true
mkdir -p /opt/brokerknow-test/uploads/{client-attachments,portal-applications}
rm -rf /opt/brokerknow-test/api
mkdir -p /opt/brokerknow-test/api
cp -a /opt/brokerknow/api/. /opt/brokerknow-test/api/

# Patch appsettings: swap DB to test, override URLs and uploads paths
python3 <<PY
import json, pathlib
p = pathlib.Path('/opt/brokerknow-test/api/appsettings.json')
d = json.loads(p.read_text())
cs = d.setdefault('ConnectionStrings', {})
cs['BrokerKnow'] = cs.get('BrokerKnow','').replace('Database=$PROD_DB', 'Database=$TEST_DB')
d['Urls'] = 'http://127.0.0.1:5261'
d.setdefault('Uploads', {})
d['Uploads']['ClientAttachmentsRoot'] = '/opt/brokerknow-test/uploads/client-attachments'
d['Uploads']['PortalApplicationsRoot'] = '/opt/brokerknow-test/uploads/portal-applications'
p.write_text(json.dumps(d, indent=2))
print('Patched test appsettings.')
PY

chown -R deploy:deploy /opt/brokerknow-test

echo '=== 3. Install systemd unit for test API ==='
cat >/etc/systemd/system/brokerknow-api-test.service <<'UNIT'
[Unit]
Description=BrokerKnow API (test)
After=network.target mssql-server.service
Requires=mssql-server.service

[Service]
WorkingDirectory=/opt/brokerknow-test/api
ExecStart=/usr/bin/dotnet /opt/brokerknow-test/api/BrokerKnow.Api.dll
Restart=always
RestartSec=10
User=deploy
Environment=ASPNETCORE_ENVIRONMENT=Production
Environment=ASPNETCORE_URLS=http://127.0.0.1:5261
Environment=DOTNET_PRINT_TELEMETRY_MESSAGE=false

[Install]
WantedBy=multi-user.target
UNIT
systemctl daemon-reload
systemctl enable brokerknow-api-test
systemctl start brokerknow-api-test
sleep 5
systemctl is-active brokerknow-api-test
curl -sS -o /dev/null -w 'test-api health=%{http_code}\n' http://127.0.0.1:5261/health

echo '=== 4. Deploy SPA bundles (root-base admin + test copies) ==='
# admin-host (base=/): used by admin.cedarcapital.mw + test-admin.cedarcapital.mw
rm -rf /var/www/admin-host /var/www/test-admin /var/www/test-portal
mkdir -p /var/www/admin-host /var/www/test-admin /var/www/test-portal
tar -xzf "/tmp/admin-host-${TS}.tgz" -C /var/www/admin-host/
tar -xzf "/tmp/admin-host-${TS}.tgz" -C /var/www/test-admin/
tar -xzf "/tmp/portal-${TS}.tgz"     -C /var/www/test-portal/
chown -R www-data:www-data /var/www/admin-host /var/www/test-admin /var/www/test-portal
echo 'SPA bundles deployed.'

echo '=== 5. Done. Next: install nginx vhosts, then reload. ==='
