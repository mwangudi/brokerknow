/* PROD pre-check for PK_Payment / PK_Security. READ-ONLY. */
SET NOCOUNT ON;
USE BrokerKnow;
GO
PRINT '===== PAYMENT =====';
SELECT
  (SELECT COUNT(*) FROM dbo.Payment) AS rows_total,
  (SELECT COUNT(*) FROM dbo.Payment WHERE Payment_DPA_ IS NULL) AS null_keys,
  (SELECT COUNT(*) FROM (SELECT Payment_DPA_ FROM dbo.Payment
       WHERE Payment_DPA_ IS NOT NULL GROUP BY Payment_DPA_ HAVING COUNT(*)>1) d) AS dup_keys,
  (SELECT CASE WHEN EXISTS (SELECT 1 FROM sys.indexes WHERE object_id=OBJECT_ID('dbo.Payment') AND is_primary_key=1) THEN 1 ELSE 0 END) AS has_pk;

PRINT '===== PAYMENT duplicate detail (if any) =====';
;WITH d AS (SELECT Payment_DPA_ FROM dbo.Payment WHERE Payment_DPA_ IS NOT NULL GROUP BY Payment_DPA_ HAVING COUNT(*)>1)
SELECT p.Payment_DPA_, p.Entity_DPA_, p.PaymentAmount, p.PaymentPDate, p.PaymentReference, p.Deleted
FROM dbo.Payment p JOIN d ON d.Payment_DPA_=p.Payment_DPA_ ORDER BY p.Payment_DPA_;

PRINT '===== SECURITY =====';
SELECT
  (SELECT COUNT(*) FROM dbo.Security) AS rows_total,
  (SELECT COUNT(*) FROM dbo.Security WHERE Security_DPA_ IS NULL) AS null_keys,
  (SELECT COUNT(*) FROM (SELECT Security_DPA_ FROM dbo.Security
       WHERE Security_DPA_ IS NOT NULL GROUP BY Security_DPA_ HAVING COUNT(*)>1) d) AS dup_keys,
  (SELECT CASE WHEN EXISTS (SELECT 1 FROM sys.indexes WHERE object_id=OBJECT_ID('dbo.Security') AND is_primary_key=1) THEN 1 ELSE 0 END) AS has_pk;
GO
