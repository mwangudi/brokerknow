#!/usr/bin/env bash
# Install backups on the droplet:
#   - Switches user DBs to FULL recovery model
#   - Takes an initial FULL so log backups have a base
#   - Drops scripts into /usr/local/sbin
#   - Installs cron entries
#   - Installs logrotate
set -euo pipefail

# Credentials are not kept in git. See Ops_Runbook.md section 2.
DB_CONF=${BROKERKNOW_DB_CONF:-/etc/brokerknow/db.conf}
[ -r "$DB_CONF" ] && . "$DB_CONF"
SA_PWD="${SA_PWD:?SA_PWD not set - create $DB_CONF containing SA_PWD='...' (chmod 600, root only)}"

echo '=== Install backup scripts ==='
install -m 0750 -o root -g root /tmp/backup-nightly.sh /usr/local/sbin/backup-nightly.sh
install -m 0750 -o root -g root /tmp/backup-log.sh     /usr/local/sbin/backup-log.sh
mkdir -p /var/backups/brokerknow/sql /var/backups/brokerknow/fs
chown mssql:mssql /var/backups/brokerknow/sql
touch /var/log/brokerknow-backup.log
chmod 0640 /var/log/brokerknow-backup.log

echo '=== Switch user DBs to FULL recovery + initial FULL backup ==='
DBS=$(/opt/mssql-tools18/bin/sqlcmd -S localhost -U sa -P "$SA_PWD" -C -h -1 -W -Q \
  "SET NOCOUNT ON; SELECT name FROM sys.databases WHERE database_id > 4 AND state_desc = 'ONLINE';" \
  | grep -v '^$' | grep -v 'rows affected')
for DB in $DBS; do
  echo "  $DB -> FULL recovery"
  /opt/mssql-tools18/bin/sqlcmd -S localhost -U sa -P "$SA_PWD" -C -b -Q \
    "ALTER DATABASE [$DB] SET RECOVERY FULL;"
done

echo '=== First FULL backup (seeds the log chain) ==='
/usr/local/sbin/backup-nightly.sh

echo '=== Install cron ==='
cat >/etc/cron.d/brokerknow-backup <<'CRON'
# BrokerKnow backups
SHELL=/bin/bash
PATH=/usr/local/sbin:/usr/local/bin:/usr/sbin:/usr/bin:/sbin:/bin

# Full DB + filesystem nightly at 02:00
0 2 * * * root /usr/local/sbin/backup-nightly.sh

# Transaction log every 15 min
*/15 * * * * root /usr/local/sbin/backup-log.sh
CRON
chmod 0644 /etc/cron.d/brokerknow-backup

echo '=== Install logrotate ==='
cat >/etc/logrotate.d/brokerknow-backup <<'LR'
/var/log/brokerknow-backup.log {
    weekly
    rotate 8
    compress
    delaycompress
    missingok
    notifempty
    create 0640 root root
}
LR

echo '=== Verify ==='
ls -la /var/backups/brokerknow/sql/ | tail -10
echo '---'
cat /etc/cron.d/brokerknow-backup
echo '---'
tail -20 /var/log/brokerknow-backup.log
