/* =====================================================================
   PK_BankAcc — declare the primary key on dbo.BankAcc.
   Target: BrokerKnow_Test (:5261) then prod (change USE).
   BankAcc_DPA_ is already an IDENTITY NOT NULL column with 0 nulls / 0 dups,
   so this is a purely declarative, additive PRIMARY KEY (NONCLUSTERED to
   avoid rebuilding the heap). Idempotent. Reversible: DROP CONSTRAINT.
   ===================================================================== */
SET NOCOUNT ON;
USE BrokerKnow_Test;
GO

IF EXISTS (SELECT 1 FROM sys.key_constraints WHERE parent_object_id=OBJECT_ID('dbo.BankAcc') AND type='PK')
    PRINT 'SKIP: BankAcc already has a primary key.';
ELSE IF EXISTS (SELECT 1 FROM dbo.BankAcc WHERE BankAcc_DPA_ IS NULL)
    PRINT 'SKIP: NULL BankAcc_DPA_ present.';
ELSE IF EXISTS (SELECT BankAcc_DPA_ FROM dbo.BankAcc GROUP BY BankAcc_DPA_ HAVING COUNT(*)>1)
    PRINT 'SKIP: duplicate BankAcc_DPA_ present.';
ELSE
BEGIN
    ALTER TABLE dbo.BankAcc ADD CONSTRAINT PK_BankAcc PRIMARY KEY NONCLUSTERED (BankAcc_DPA_);
    PRINT 'OK: PK_BankAcc created.';
END
GO

PRINT '===== verify =====';
SELECT t.name AS tbl, kc.name AS pk, i.type_desc
FROM sys.key_constraints kc
JOIN sys.tables t ON t.object_id=kc.parent_object_id
JOIN sys.indexes i ON i.object_id=kc.parent_object_id AND i.index_id=kc.unique_index_id
WHERE kc.type='PK' AND t.name='BankAcc';
GO
