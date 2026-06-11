/* Final PROD verification: PKs present, 0 Payment dups, index count. READ-ONLY. */
SET NOCOUNT ON;
USE BrokerKnow;
GO
SELECT 'PKs' AS check_name, t.name AS tbl, kc.name AS pk
FROM sys.key_constraints kc JOIN sys.tables t ON t.object_id=kc.parent_object_id
WHERE kc.type='PK' AND t.name IN ('Payment','Security') ORDER BY t.name;

SELECT 'PaymentDups' AS check_name,
  (SELECT COUNT(*) FROM (SELECT Payment_DPA_ FROM dbo.Payment WHERE Payment_DPA_ IS NOT NULL GROUP BY Payment_DPA_ HAVING COUNT(*)>1) d) AS remaining_dups,
  (SELECT COUNT(*) FROM dbo.Payment) AS payment_rows;

SELECT 'Indexes' AS check_name, COUNT(*) AS ix_count
FROM sys.indexes i JOIN sys.tables t ON t.object_id=i.object_id
WHERE i.name LIKE 'IX[_]%'
  AND t.name IN ('LevyContract','Payment','Lot','OrdDetail','JournalEntry','tbOrder','Holdings','Contract','Client');
GO
