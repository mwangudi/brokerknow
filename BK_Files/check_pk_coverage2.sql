SET NOCOUNT ON;
USE BrokerKnow_Test;
GO
PRINT '===== PK presence =====';
SELECT t.name AS tbl, p.rows AS row_count,
       ISNULL(pk.name,'** NO PK **') AS pk
FROM sys.tables t
JOIN sys.partitions p ON p.object_id=t.object_id AND p.index_id IN (0,1)
LEFT JOIN sys.key_constraints pk ON pk.parent_object_id=t.object_id AND pk.type='PK'
WHERE t.name IN ('Owner','Agent','Broker','Bank','BnkBranch','BankAcc',
                 'Holdings','ClientBalances','ClientTotal','tbOrder','OrdDetail','Contract','Lot')
ORDER BY pk, t.name;
GO
PRINT '===== BankAcc columns =====';
SELECT c.column_id, c.name, ty.name AS typ, c.is_identity, c.is_nullable
FROM sys.columns c JOIN sys.types ty ON ty.user_type_id=c.user_type_id
WHERE c.object_id=OBJECT_ID('dbo.BankAcc') ORDER BY c.column_id;
GO
