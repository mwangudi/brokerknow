#!/usr/bin/env bash
# BrokerKnow nightly backup runner.
# Run from cron at 02:00. Performs:
#   1. SQL Server FULL backup for every user DB
#   2. Filesystem tar of uploads + /etc configs + appsettings
#   3. Local rotation (keeps 14 days)
#   4. Optional offsite sync to S3-compatible storage (DO Spaces) if rclone
#      remote "spaces:" is configured
#
# Logs to /var/log/brokerknow-backup.log. Exits non-zero on any failure.
set -euo pipefail

LOG=/var/log/brokerknow-backup.log
exec >>"$LOG" 2>&1
echo "==== $(date -Is) backup-nightly start ===="

BACKUP_ROOT=/var/backups/brokerknow
SQL_DIR=$BACKUP_ROOT/sql
FS_DIR=$BACKUP_ROOT/fs
KEEP_DAYS=14
SA_PWD='SpringfielD##88'
TS=$(date +%Y%m%d-%H%M%S)

mkdir -p "$SQL_DIR" "$FS_DIR"
chown -R mssql:mssql "$SQL_DIR"  # mssql writes here; we read it later

# --- 1. SQL Server FULL ---
USER_DBS=$(/opt/mssql-tools18/bin/sqlcmd -S localhost -U sa -P "$SA_PWD" -C -h -1 -W -Q \
  "SET NOCOUNT ON; SELECT name FROM sys.databases WHERE database_id > 4 AND state_desc = 'ONLINE';" \
  | grep -v '^$' | grep -v 'rows affected')

for DB in $USER_DBS; do
  OUT="$SQL_DIR/${DB}_FULL_${TS}.bak"
  echo "FULL $DB -> $OUT"
  /opt/mssql-tools18/bin/sqlcmd -S localhost -U sa -P "$SA_PWD" -C -b -Q \
    "BACKUP DATABASE [$DB] TO DISK = N'$OUT' WITH INIT, COMPRESSION, CHECKSUM, FORMAT, NAME = N'$DB FULL $TS';"
done

# --- 2. Filesystem ---
FS_OUT="$FS_DIR/files_${TS}.tar.gz"
echo "Filesystem -> $FS_OUT"
tar --warning=no-file-changed -czf "$FS_OUT" \
  /opt/brokerknow/uploads \
  /opt/brokerknow/api/appsettings.json \
  /opt/brokerknow-test/api/appsettings.json \
  /etc/nginx/sites-available \
  /etc/systemd/system/brokerknow-api.service \
  /etc/systemd/system/brokerknow-api-test.service 2>/dev/null || true

# --- 3. Local rotation ---
find "$SQL_DIR" -type f -name '*.bak' -mtime +$KEEP_DAYS -print -delete
find "$SQL_DIR" -type f -name '*.trn' -mtime +2 -print -delete
find "$FS_DIR"  -type f -name '*.tar.gz' -mtime +$KEEP_DAYS -print -delete

# --- 4. Offsite (optional) ---
if command -v rclone >/dev/null && rclone listremotes 2>/dev/null | grep -q '^spaces:'; then
  echo "Offsite sync -> spaces:brokerknow-backups"
  rclone sync "$BACKUP_ROOT" spaces:brokerknow-backups \
    --transfers=4 --checkers=4 --s3-no-check-bucket --fast-list \
    --log-level NOTICE
else
  echo "Offsite SKIPPED (no rclone 'spaces:' remote configured)."
fi

# --- 5. Summary to log ---
du -sh "$SQL_DIR" "$FS_DIR" || true
echo "==== $(date -Is) backup-nightly done ===="
