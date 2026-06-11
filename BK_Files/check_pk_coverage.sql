/* READ-ONLY: PK coverage + key health for (a) tables behind the remaining
   MAX+1 code sites, and (b) real business tables still lacking a PK.
   Decides whether more PKs (additive, low-risk) are the right next step. */
SET NOCOUNT ON;
USE BrokerKnow_Test;
GO

PRINT '===== PK presence on MAX+1 target + candidate tables =====';
SELECT t.name AS tbl, p.rows AS row_count,
       CASE WHEN pk.name IS NULL THEN 'NO PK' ELSE pk.name END AS pk,
       i.name AS clustered_index
FROM sys.tables t
JOIN sys.partitions p ON p.object_id=t.object_id AND p.index_id IN (0,1)
LEFT JOIN sys.key_constraints pk ON pk.parent_object_id=t.object_id AND pk.type='PK'
LEFT JOIN sys.indexes i ON i.object_id=t.object_id AND i.index_id=1
WHERE t.name IN ('Owner','Agent','Broker','Bank','BnkBranch','BankAcc',
                 'Holdings','ClientBalances','ClientTotal','tbOrder','OrdDetail','Contract')
ORDER BY pk, t.name;

PRINT '===== key health for the no-PK real tables (nulls / dups in the natural key) =====';
-- BankAcc.BankAccount_DPA_
SELECT 'BankAcc' AS tbl, 'BankAccount_DPA_' AS keycol,
  (SELECT COUNT(*) FROM dbo.BankAcc) AS rows_total,
  (SELECT COUNT(*) FROM dbo.BankAcc WHERE BankAccount_DPA_ IS NULL) AS null_keys,
  (SELECT COUNT(*) FROM (SELECT BankAccount_DPA_ FROM dbo.BankAcc WHERE BankAccount_DPA_ IS NOT NULL GROUP BY BankAccount_DPA_ HAVING COUNT(*)>1) d) AS dup_keys;
-- ClientBalances.client_DPA_
SELECT 'ClientBalances','client_DPA_',
  (SELECT COUNT(*) FROM dbo.ClientBalances),
  (SELECT COUNT(*) FROM dbo.ClientBalances WHERE client_DPA_ IS NULL),
  (SELECT COUNT(*) FROM (SELECT client_DPA_ FROM dbo.ClientBalances WHERE client_DPA_ IS NOT NULL GROUP BY client_DPA_ HAVING COUNT(*)>1) d);
-- ClientTotal.Client_DPA_  (col name may differ — guard with COL_LENGTH)
GO
IF COL_LENGTH('dbo.ClientTotal','Client_DPA_') IS NOT NULL
  SELECT 'ClientTotal' AS tbl, 'Client_DPA_' AS keycol,
    (SELECT COUNT(*) FROM dbo.ClientTotal) AS rows_total,
    (SELECT COUNT(*) FROM dbo.ClientTotal WHERE Client_DPA_ IS NULL) AS null_keys,
    (SELECT COUNT(*) FROM (SELECT Client_DPA_ FROM dbo.ClientTotal WHERE Client_DPA_ IS NOT NULL GROUP BY Client_DPA_ HAVING COUNT(*)>1) d) AS dup_keys;
ELSE PRINT 'ClientTotal: no Client_DPA_ column — inspect separately';
GO
-- Holdings.HoldingID (identity?)
SELECT 'Holdings' AS tbl, c.name AS keycol, c.is_identity, c.is_nullable
FROM sys.columns c WHERE c.object_id=OBJECT_ID('dbo.Holdings') AND c.name='HoldingID';
GO
