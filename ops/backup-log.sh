#!/usr/bin/env bash
# Transaction log backup runner. Cron every 15 min.
# Only runs against DBs in FULL recovery model with at least one prior full backup.
set -euo pipefail

LOG=/var/log/brokerknow-backup.log
exec >>"$LOG" 2>&1

SQL_DIR=/var/backups/brokerknow/sql
SA_PWD='SpringfielD##88'
TS=$(date +%Y%m%d-%H%M%S)

mkdir -p "$SQL_DIR"

DBS=$(/opt/mssql-tools18/bin/sqlcmd -S localhost -U sa -P "$SA_PWD" -C -h -1 -W -Q \
  "SET NOCOUNT ON;
   SELECT d.name
   FROM sys.databases d
   WHERE d.database_id > 4
     AND d.state_desc = 'ONLINE'
     AND d.recovery_model_desc = 'FULL'
     AND EXISTS (SELECT 1 FROM msdb.dbo.backupset bs
                 WHERE bs.database_name = d.name AND bs.type = 'D');" \
  | grep -v '^$' | grep -v 'rows affected')

for DB in $DBS; do
  OUT="$SQL_DIR/${DB}_LOG_${TS}.trn"
  /opt/mssql-tools18/bin/sqlcmd -S localhost -U sa -P "$SA_PWD" -C -b -Q \
    "BACKUP LOG [$DB] TO DISK = N'$OUT' WITH COMPRESSION, CHECKSUM, NAME = N'$DB LOG $TS';" \
    || echo "[$(date -Is)] LOG backup failed for $DB"
done
