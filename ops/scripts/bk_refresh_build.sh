#!/usr/bin/env bash
# bk_refresh_build.sh <BuildDb>  — prepare a freshly restored legacy dump for cutover.
#
# Reads from axis_db_prod but never writes to it. All writes land in <BuildDb>.
# Replaces the June-era hardcoded graft: the app-table list is derived at run time
# from whatever exists on live, so anything added since (IPO, SupportTicket,
# AppSettings, ClientCDS_backup_*, ...) is carried across instead of silently lost.
set -euo pipefail

BUILD="${1:?usage: bk_refresh_build.sh <BuildDb>}"
LIVE=axis_db_prod
SQLCMD=/opt/mssql-tools18/bin/sqlcmd
CFG=/opt/brokerknow/api/appsettings.json
SA_PWD="$(grep -oiE 'password=[^;\"]+' "$CFG" | head -1 | cut -d= -f2-)"
[ -n "$SA_PWD" ] || { echo "ERROR: no DB credential" >&2; exit 1; }
run() { "$SQLCMD" -S localhost -U sa -P "$SA_PWD" -C -I -b -W "$@"; }

run -h -1 -Q "IF DB_ID('${BUILD}') IS NULL RAISERROR('build db ${BUILD} not found',16,1);"

echo "== 1/5 graft app-layer tables (dynamic) =="
run -d "$BUILD" -Q "
SET NOCOUNT ON;
DECLARE @sql nvarchar(max) = N'', @n int = 0;
SELECT @sql = @sql + N'SELECT * INTO dbo.' + QUOTENAME(l.name)
                   + N' FROM ${LIVE}.dbo.' + QUOTENAME(l.name) + N';' + CHAR(10),
       @n = @n + 1
FROM ${LIVE}.sys.tables l
WHERE l.schema_id = SCHEMA_ID('dbo')
  AND NOT EXISTS (SELECT 1 FROM sys.tables d
                  WHERE d.name = l.name AND d.schema_id = SCHEMA_ID('dbo'));
IF @sql <> N'' EXEC sp_executesql @sql;
SELECT CONVERT(varchar(10), @n) + ' table(s) grafted' AS grafted;"

echo "== 2/5 restore keys SELECT INTO drops (PK on Id) =="
run -d "$BUILD" -Q "
SET NOCOUNT ON;
DECLARE @t sysname, @sql nvarchar(max), @done int = 0, @skipped int = 0;
-- Only tables where LIVE itself has a single-column PK on Id, and the grafted
-- copy has no NULLs there. Anything else is left alone.
DECLARE c CURSOR LOCAL FAST_FORWARD FOR
  SELECT t.name
  FROM sys.tables t
  WHERE t.schema_id = SCHEMA_ID('dbo')
    AND NOT EXISTS (SELECT 1 FROM sys.key_constraints k
                    WHERE k.parent_object_id = t.object_id AND k.type = 'PK')
    AND EXISTS (
      SELECT 1
      FROM ${LIVE}.sys.key_constraints lk
      JOIN ${LIVE}.sys.tables lt ON lt.object_id = lk.parent_object_id
      JOIN ${LIVE}.sys.index_columns lic ON lic.object_id = lt.object_id AND lic.index_id = lk.unique_index_id
      JOIN ${LIVE}.sys.columns lc ON lc.object_id = lt.object_id AND lc.column_id = lic.column_id
      WHERE lk.type = 'PK' AND lt.name = t.name AND lc.name = 'Id'
      GROUP BY lt.name HAVING COUNT(*) = 1);
OPEN c; FETCH NEXT FROM c INTO @t;
WHILE @@FETCH_STATUS = 0
BEGIN
  BEGIN TRY
    SET @sql = N'IF NOT EXISTS (SELECT 1 FROM dbo.' + QUOTENAME(@t) + N' WHERE [Id] IS NULL) BEGIN '
             + N'ALTER TABLE dbo.' + QUOTENAME(@t) + N' ALTER COLUMN [Id] int NOT NULL; '
             + N'ALTER TABLE dbo.' + QUOTENAME(@t) + N' ADD CONSTRAINT ' + QUOTENAME('PK_' + @t)
             + N' PRIMARY KEY ([Id]); END';
    EXEC sp_executesql @sql;
    SET @done = @done + 1;
  END TRY
  BEGIN CATCH
    SET @skipped = @skipped + 1;
  END CATCH
  FETCH NEXT FROM c INTO @t;
END
CLOSE c; DEALLOCATE c;
SELECT CONVERT(varchar(10),@done) + ' PK(s) restored, ' + CONVERT(varchar(10),@skipped) + ' skipped' AS keys;"

echo "== 3/5 carry CDS numbers across from live (by Client_DPA_) =="
# The legacy schema still declares ClientCDSNo nvarchar(20); Axis widened it and
# real values reach 23 chars, so widen to match live before copying or the
# UPDATE fails with a truncation error.
run -d "$BUILD" -Q "
SET NOCOUNT ON;
DECLARE @live_len int = (
  SELECT c.max_length/2 FROM ${LIVE}.sys.columns c
  WHERE c.object_id = OBJECT_ID('${LIVE}.dbo.Client') AND c.name = 'ClientCDSNo');
DECLARE @build_len int = (
  SELECT c.max_length/2 FROM sys.columns c
  WHERE c.object_id = OBJECT_ID('dbo.Client') AND c.name = 'ClientCDSNo');
IF @live_len > @build_len
BEGIN
  DECLARE @sql nvarchar(max) =
    N'ALTER TABLE dbo.Client ALTER COLUMN ClientCDSNo nvarchar(' + CONVERT(varchar(10), @live_len) + N') NULL;';
  EXEC sp_executesql @sql;
  SELECT 'widened ClientCDSNo ' + CONVERT(varchar(10),@build_len) + ' -> ' + CONVERT(varchar(10),@live_len) AS widen;
END
ELSE SELECT 'ClientCDSNo width already adequate (' + CONVERT(varchar(10),@build_len) + ')' AS widen;"

run -d "$BUILD" -Q "
SET NOCOUNT ON;
SET XACT_ABORT ON;
DECLARE @expected int = (SELECT COUNT(*) FROM ${LIVE}.dbo.Client l
                         JOIN dbo.Client b ON b.Client_DPA_ = l.Client_DPA_
                         WHERE ISNULL(l.ClientCDSNo,'') <> '');
BEGIN TRAN;
UPDATE b SET b.ClientCDSNo = l.ClientCDSNo
FROM dbo.Client b
JOIN ${LIVE}.dbo.Client l ON l.Client_DPA_ = b.Client_DPA_
WHERE ISNULL(l.ClientCDSNo,'') <> '';
DECLARE @updated int = @@ROWCOUNT;
IF @updated <> @expected
BEGIN
    ROLLBACK TRAN;
    RAISERROR('CDS carry-across mismatch: updated %d, expected %d', 16, 1, @updated, @expected);
END
ELSE
BEGIN
    COMMIT TRAN;
    SELECT CONVERT(varchar(10), @updated) + ' CDS numbers carried across' AS cds;
END"

echo "== 4/5 re-apply the BankAcc name/number swap fix =="
if [ -f /tmp/bk/fix_swapped_bankacc_name_number.sql ]; then
  run -d "$BUILD" -i /tmp/bk/fix_swapped_bankacc_name_number.sql
else
  echo "   WARN: /tmp/bk/fix_swapped_bankacc_name_number.sql missing — skipped"
fi

echo "== 5/5 verification =="
run -d "$BUILD" -Q "
SELECT
 (SELECT COUNT(*) FROM dbo.Client)                                          AS clients,
 (SELECT COUNT(*) FROM dbo.Client WHERE ISNULL(ClientCDSNo,'') <> '')       AS with_cds,
 (SELECT COUNT(*) FROM dbo.PortalUsers)                                     AS portal_users,
 (SELECT COUNT(*) FROM sys.tables WHERE schema_id = SCHEMA_ID('dbo'))       AS tables_total;"

run -h -1 -Q "
SELECT 'STILL MISSING: ' + l.name
FROM ${LIVE}.sys.tables l
WHERE l.schema_id = SCHEMA_ID('dbo')
  AND NOT EXISTS (SELECT 1 FROM ${BUILD}.sys.tables d
                  WHERE d.name = l.name AND d.schema_id = SCHEMA_ID('dbo'));"

echo "BUILD_COMPLETE for ${BUILD} — ${LIVE} untouched."
