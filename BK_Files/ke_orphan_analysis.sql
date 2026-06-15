SET NOCOUNT ON;
USE BrokerKnow_KE_Clean;

PRINT '=== FK_tbOrder_Client: orders whose Client_DPA_ has no Client ===';
SELECT COUNT(*) AS orphan_orders,
       SUM(CASE WHEN o.Client_DPA_ = 0 THEN 1 ELSE 0 END) AS zero_sentinel,
       SUM(CASE WHEN o.Client_DPA_ IS NULL THEN 1 ELSE 0 END) AS null_refs,
       SUM(CASE WHEN o.Client_DPA_ <> 0 AND o.Client_DPA_ IS NOT NULL THEN 1 ELSE 0 END) AS real_dangling
FROM dbo.tbOrder o
WHERE NOT EXISTS (SELECT 1 FROM dbo.Client c WHERE c.Client_DPA_ = o.Client_DPA_);

PRINT '=== FK_Lot_Contract: lots whose Contract_DPA_ has no Contract ===';
SELECT COUNT(*) AS orphan_lots,
       SUM(CASE WHEN l.Contract_DPA_ = 0 THEN 1 ELSE 0 END) AS zero_sentinel,
       SUM(CASE WHEN l.Contract_DPA_ IS NULL THEN 1 ELSE 0 END) AS null_refs,
       SUM(CASE WHEN l.Contract_DPA_ <> 0 AND l.Contract_DPA_ IS NOT NULL THEN 1 ELSE 0 END) AS real_dangling
FROM dbo.Lot l
WHERE NOT EXISTS (SELECT 1 FROM dbo.Contract c WHERE c.Contract_DPA_ = l.Contract_DPA_);

PRINT '=== FK_Payment_PayType: payments whose PayType_DPA_ has no PayType ===';
SELECT COUNT(*) AS orphan_payments,
       SUM(CASE WHEN p.PayType_DPA_ = 0 THEN 1 ELSE 0 END) AS zero_sentinel,
       SUM(CASE WHEN p.PayType_DPA_ IS NULL THEN 1 ELSE 0 END) AS null_refs,
       SUM(CASE WHEN p.PayType_DPA_ <> 0 AND p.PayType_DPA_ IS NOT NULL THEN 1 ELSE 0 END) AS real_dangling
FROM dbo.Payment p
WHERE NOT EXISTS (SELECT 1 FROM dbo.PayType pt WHERE pt.PayType_DPA_ = p.PayType_DPA_);

PRINT '=== distinct missing PayType ids referenced ===';
SELECT DISTINCT p.PayType_DPA_ AS missing_paytype FROM dbo.Payment p
WHERE NOT EXISTS (SELECT 1 FROM dbo.PayType pt WHERE pt.PayType_DPA_ = p.PayType_DPA_)
ORDER BY p.PayType_DPA_;

PRINT '=== row totals for context ===';
SELECT (SELECT COUNT(*) FROM dbo.tbOrder) AS orders, (SELECT COUNT(*) FROM dbo.Lot) AS lots, (SELECT COUNT(*) FROM dbo.Payment) AS payments;
