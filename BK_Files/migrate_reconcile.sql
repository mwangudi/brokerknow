/* =====================================================================
   MIGRATION RECONCILIATION — compares BrokerKnow_Clean vs prod BrokerKnow.
   Per-table COUNT(*) diff + money control totals. READ-ONLY on both.
   AuditTrail/AuditTrailItem are expected mismatches (intentionally not
   migrated) — flagged separately, not counted as failures.
   ===================================================================== */
SET NOCOUNT ON;
USE BrokerKnow_Clean;
GO

IF OBJECT_ID('tempdb..#rec') IS NOT NULL DROP TABLE #rec;
CREATE TABLE #rec (tbl sysname, clean_rows int, src_rows int);

DECLARE @t sysname, @sql nvarchar(max);
DECLARE c CURSOR LOCAL FAST_FORWARD FOR
  SELECT name FROM sys.tables WHERE schema_id=SCHEMA_ID('dbo') ORDER BY name;
OPEN c; FETCH NEXT FROM c INTO @t;
WHILE @@FETCH_STATUS=0
BEGIN
  SET @sql = 'INSERT INTO #rec SELECT ''' + @t + ''','
    + '(SELECT COUNT(*) FROM dbo.' + QUOTENAME(@t) + '),'
    + '(SELECT COUNT(*) FROM BrokerKnow.dbo.' + QUOTENAME(@t) + ')';
  BEGIN TRY EXEC sp_executesql @sql; END TRY
  BEGIN CATCH INSERT INTO #rec VALUES (@t, -1, -1); END CATCH
  FETCH NEXT FROM c INTO @t;
END
CLOSE c; DEALLOCATE c;

PRINT '===== MISMATCHES (excluding intentionally-skipped audit logs) =====';
SELECT tbl, clean_rows, src_rows, src_rows - clean_rows AS missing
FROM #rec
WHERE clean_rows <> src_rows AND tbl NOT IN ('AuditTrail','AuditTrailItem')
ORDER BY tbl;

PRINT '===== summary =====';
SELECT COUNT(*) AS tables_compared,
       SUM(CASE WHEN clean_rows = src_rows THEN 1 ELSE 0 END) AS matching,
       SUM(CASE WHEN clean_rows <> src_rows AND tbl NOT IN ('AuditTrail','AuditTrailItem') THEN 1 ELSE 0 END) AS real_mismatches,
       SUM(CASE WHEN tbl IN ('AuditTrail','AuditTrailItem') THEN 1 ELSE 0 END) AS intentional_skips
FROM #rec;

PRINT '===== money control totals: clean vs source =====';
SELECT 'Payment_sum' AS metric,
       (SELECT CAST(SUM(PaymentAmount) AS decimal(38,2)) FROM dbo.Payment WHERE Deleted=0 OR Deleted IS NULL) AS clean_val,
       (SELECT CAST(SUM(PaymentAmount) AS decimal(38,2)) FROM BrokerKnow.dbo.Payment WHERE Deleted=0 OR Deleted IS NULL) AS src_val
UNION ALL SELECT 'ClientBalances_sum',
       (SELECT CAST(SUM(CurrentBal) AS decimal(38,2)) FROM dbo.ClientBalances),
       (SELECT CAST(SUM(CurrentBal) AS decimal(38,2)) FROM BrokerKnow.dbo.ClientBalances)
UNION ALL SELECT 'Lot_gross_sum',
       (SELECT CAST(SUM(LotGrossAmount) AS decimal(38,2)) FROM dbo.Lot WHERE Deleted=0 OR Deleted IS NULL),
       (SELECT CAST(SUM(LotGrossAmount) AS decimal(38,2)) FROM BrokerKnow.dbo.Lot WHERE Deleted=0 OR Deleted IS NULL)
UNION ALL SELECT 'LevyContract_sum',
       (SELECT CAST(SUM(LevyAmount) AS decimal(38,2)) FROM dbo.LevyContract),
       (SELECT CAST(SUM(LevyAmount) AS decimal(38,2)) FROM BrokerKnow.dbo.LevyContract);
GO
