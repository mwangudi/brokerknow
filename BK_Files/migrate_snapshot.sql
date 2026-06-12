/* READ-ONLY pre-migration snapshot. Source = prod BrokerKnow. Captures the
   per-table row counts we will reconcile the clean DB against, plus the
   money-path control totals (the numbers that MUST match exactly). */
SET NOCOUNT ON;
USE BrokerKnow;
GO
PRINT '===== row counts: all dbo tables (source of truth) =====';
SELECT t.name AS table_name, SUM(p.rows) AS row_count
FROM sys.tables t
JOIN sys.partitions p ON p.object_id=t.object_id AND p.index_id IN (0,1)
WHERE t.schema_id = SCHEMA_ID('dbo')
GROUP BY t.name
ORDER BY t.name;

PRINT '===== money-path control totals =====';
SELECT 'Payment_sum' AS metric, CAST(SUM(PaymentAmount) AS decimal(38,2)) AS val FROM dbo.Payment WHERE Deleted=0 OR Deleted IS NULL
UNION ALL SELECT 'Payment_rows', COUNT(*) FROM dbo.Payment
UNION ALL SELECT 'ClientBalances_sum', CAST(SUM(CurrentBal) AS decimal(38,2)) FROM dbo.ClientBalances
UNION ALL SELECT 'Lot_gross_sum', CAST(SUM(LotGrossAmount) AS decimal(38,2)) FROM dbo.Lot WHERE Deleted=0 OR Deleted IS NULL
UNION ALL SELECT 'LevyContract_sum', CAST(SUM(LevyAmount) AS decimal(38,2)) FROM dbo.LevyContract
UNION ALL SELECT 'Client_rows', COUNT(*) FROM dbo.Client
UNION ALL SELECT 'Contract_rows', COUNT(*) FROM dbo.Contract;
GO
