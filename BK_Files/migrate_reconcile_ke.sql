/* KE migration reconciliation — BrokerKnow_KE_Clean vs BrokerKnow_KE_Legacy.
   Per-table COUNT(*) diff (only tables present in BOTH) + money control totals.
   AuditTrail/AuditTrailItem + app-layer tables are expected differences. */
SET NOCOUNT ON;
USE BrokerKnow_KE_Clean;
GO

IF OBJECT_ID('tempdb..#rec') IS NOT NULL DROP TABLE #rec;
CREATE TABLE #rec (tbl sysname, clean_rows int, src_rows int);

DECLARE @t sysname, @sql nvarchar(max);
DECLARE c CURSOR LOCAL FAST_FORWARD FOR
  SELECT name FROM sys.tables WHERE schema_id=SCHEMA_ID('dbo')
    AND OBJECT_ID('BrokerKnow_KE_Legacy.dbo.' + QUOTENAME(name)) IS NOT NULL
  ORDER BY name;
OPEN c; FETCH NEXT FROM c INTO @t;
WHILE @@FETCH_STATUS=0
BEGIN
  SET @sql = 'INSERT INTO #rec SELECT ''' + @t + ''','
    + '(SELECT COUNT(*) FROM dbo.' + QUOTENAME(@t) + '),'
    + '(SELECT COUNT(*) FROM BrokerKnow_KE_Legacy.dbo.' + QUOTENAME(@t) + ')';
  BEGIN TRY EXEC sp_executesql @sql; END TRY
  BEGIN CATCH INSERT INTO #rec VALUES (@t, -1, -1); END CATCH
  FETCH NEXT FROM c INTO @t;
END
CLOSE c; DEALLOCATE c;

PRINT '===== MISMATCHES (excluding intentionally-skipped audit logs) =====';
SELECT tbl, clean_rows, src_rows, src_rows - clean_rows AS missing
FROM #rec
WHERE clean_rows <> src_rows AND tbl NOT IN ('AuditTrail','AuditTrailItem')
ORDER BY ABS(src_rows - clean_rows) DESC;

PRINT '===== summary =====';
SELECT COUNT(*) AS tables_compared,
       SUM(CASE WHEN clean_rows = src_rows THEN 1 ELSE 0 END) AS matching,
       SUM(CASE WHEN clean_rows <> src_rows AND tbl NOT IN ('AuditTrail','AuditTrailItem') THEN 1 ELSE 0 END) AS real_mismatches,
       SUM(CASE WHEN tbl IN ('AuditTrail','AuditTrailItem') THEN 1 ELSE 0 END) AS intentional_skips
FROM #rec;

PRINT '===== money control totals: clean vs source =====';
SELECT 'Payment_sum' AS metric,
       (SELECT CAST(SUM(PaymentAmount) AS decimal(38,2)) FROM dbo.Payment WHERE Deleted=0 OR Deleted IS NULL) AS clean_val,
       (SELECT CAST(SUM(PaymentAmount) AS decimal(38,2)) FROM BrokerKnow_KE_Legacy.dbo.Payment WHERE Deleted=0 OR Deleted IS NULL) AS src_val
UNION ALL SELECT 'Lot_gross_sum',
       (SELECT CAST(SUM(LotGrossAmount) AS decimal(38,2)) FROM dbo.Lot WHERE Deleted=0 OR Deleted IS NULL),
       (SELECT CAST(SUM(LotGrossAmount) AS decimal(38,2)) FROM BrokerKnow_KE_Legacy.dbo.Lot WHERE Deleted=0 OR Deleted IS NULL)
UNION ALL SELECT 'LevyContract_sum',
       (SELECT CAST(SUM(LevyAmount) AS decimal(38,2)) FROM dbo.LevyContract),
       (SELECT CAST(SUM(LevyAmount) AS decimal(38,2)) FROM BrokerKnow_KE_Legacy.dbo.LevyContract);
GO
