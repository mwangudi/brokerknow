/* Phase-2 investigation — READ-ONLY. Why Payment has 2 dup keys, and the
   true state of Security_DPA_. No writes. Target: BrokerKnow_Test. */
SET NOCOUNT ON;
USE BrokerKnow_Test;
GO

/* ---- PAYMENT: the duplicate key values ----------------------------- */
PRINT '===== PAYMENT duplicate Payment_DPA_ values =====';
SELECT Payment_DPA_, COUNT(*) AS copies
FROM dbo.Payment
WHERE Payment_DPA_ IS NOT NULL
GROUP BY Payment_DPA_ HAVING COUNT(*) > 1;

/* ---- PAYMENT: full rows for those dup keys (to judge if true dup or
        distinct payments that collided on an id) ---------------------- */
PRINT '===== PAYMENT rows behind the duplicate keys =====';
SELECT p.Payment_DPA_, p.EntityType_DPA_, p.Entity_DPA_, p.PayType_DPA_,
       p.PaymentAmount, p.PaymentPDate, p.PaymentReference, p.PaymentNarrative,
       p.Deleted, p.Voucher_DPA_, p.Contract_DPA_, p.Order_DPA_
FROM dbo.Payment p
JOIN (SELECT Payment_DPA_ FROM dbo.Payment
      WHERE Payment_DPA_ IS NOT NULL
      GROUP BY Payment_DPA_ HAVING COUNT(*) > 1) d
  ON d.Payment_DPA_ = p.Payment_DPA_
ORDER BY p.Payment_DPA_;

/* ---- SECURITY: column metadata + actual null/dup state ------------- */
PRINT '===== SECURITY column metadata =====';
SELECT c.name AS column_name, ty.name AS type, c.is_nullable, c.is_identity
FROM sys.columns c
JOIN sys.types ty ON ty.user_type_id = c.user_type_id
WHERE c.object_id = OBJECT_ID('dbo.Security') AND c.name = 'Security_DPA_';

PRINT '===== SECURITY null / duplicate counts =====';
SELECT
  (SELECT COUNT(*) FROM dbo.Security) AS rows_total,
  (SELECT COUNT(*) FROM dbo.Security WHERE Security_DPA_ IS NULL) AS null_keys,
  (SELECT COUNT(*) FROM (SELECT Security_DPA_ FROM dbo.Security
       WHERE Security_DPA_ IS NOT NULL
       GROUP BY Security_DPA_ HAVING COUNT(*) > 1) d) AS dup_keys;

PRINT '===== SECURITY all rows (only 21) =====';
SELECT Security_DPA_, SecurityCode, SecurityName, OrderSecType_DPA_
FROM dbo.Security ORDER BY Security_DPA_;
GO
