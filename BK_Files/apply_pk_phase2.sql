/* =====================================================================
   BrokerKnow Phase-2 — PRIMARY KEYS on Payment & Security.
   Target: BrokerKnow_Test (:5261). TEST ONLY — NOT prod.

   Why NONCLUSTERED PK: both tables are heaps today. A clustered PK would
   rebuild the whole table (long Sch-M lock) — risky for a table the legacy
   app shares. A NONCLUSTERED PK enforces NOT NULL + uniqueness (the integrity
   goal) while leaving the heap physically intact. Reversible:
     ALTER TABLE dbo.Payment  DROP CONSTRAINT PK_Payment;
     ALTER TABLE dbo.Payment  ALTER COLUMN Payment_DPA_  int NULL;   -- if needed
     (same for Security)

   Safe: each table is guarded — the PK is added ONLY when there are 0 NULLs
   and 0 duplicate key values. Otherwise it prints why and skips (no error,
   no DDL). Idempotent: skips if the PK already exists. Pure read in the
   pre-check; the only writes are the guarded ALTERs.
   ===================================================================== */
SET NOCOUNT ON;
USE BrokerKnow_Test;
GO

/* ===================== PAYMENT ======================================= */
DECLARE @rows int, @nulls int, @dups int, @hasPk bit, @clustered sysname;

SELECT @rows  = COUNT(*),
       @nulls = SUM(CASE WHEN Payment_DPA_ IS NULL THEN 1 ELSE 0 END)
FROM dbo.Payment;

SELECT @dups = COUNT(*) FROM (
    SELECT Payment_DPA_ FROM dbo.Payment
    WHERE Payment_DPA_ IS NOT NULL
    GROUP BY Payment_DPA_ HAVING COUNT(*) > 1) d;

SELECT @hasPk = CASE WHEN EXISTS (SELECT 1 FROM sys.indexes
    WHERE object_id = OBJECT_ID('dbo.Payment') AND is_primary_key = 1) THEN 1 ELSE 0 END;

SELECT @clustered = i.name FROM sys.indexes i
WHERE i.object_id = OBJECT_ID('dbo.Payment') AND i.index_id = 1;

PRINT '===== PAYMENT pre-check =====';
PRINT '  rows           = ' + CAST(@rows AS varchar(20));
PRINT '  NULL keys      = ' + CAST(@nulls AS varchar(20));
PRINT '  duplicate keys = ' + CAST(@dups AS varchar(20));
PRINT '  existing PK?   = ' + CAST(@hasPk AS varchar(1));
PRINT '  clustered idx  = ' + ISNULL(@clustered, '(heap)');

IF @hasPk = 1
    PRINT '  -> SKIP: Payment already has a primary key.';
ELSE IF @nulls > 0 OR @dups > 0
    PRINT '  -> SKIP: cannot add PK (NULL or duplicate Payment_DPA_ values present). Investigate first.';
ELSE
BEGIN
    PRINT '  -> APPLYING: NOT NULL + NONCLUSTERED PK_Payment ...';
    ALTER TABLE dbo.Payment ALTER COLUMN Payment_DPA_ int NOT NULL;
    ALTER TABLE dbo.Payment ADD CONSTRAINT PK_Payment PRIMARY KEY NONCLUSTERED (Payment_DPA_);
    PRINT '  -> DONE: PK_Payment created.';
END
GO

/* ===================== SECURITY ===================================== */
DECLARE @rows int, @nulls int, @dups int, @hasPk bit, @clustered sysname;

SELECT @rows  = COUNT(*),
       @nulls = SUM(CASE WHEN Security_DPA_ IS NULL THEN 1 ELSE 0 END)
FROM dbo.Security;

SELECT @dups = COUNT(*) FROM (
    SELECT Security_DPA_ FROM dbo.Security
    WHERE Security_DPA_ IS NOT NULL
    GROUP BY Security_DPA_ HAVING COUNT(*) > 1) d;

SELECT @hasPk = CASE WHEN EXISTS (SELECT 1 FROM sys.indexes
    WHERE object_id = OBJECT_ID('dbo.Security') AND is_primary_key = 1) THEN 1 ELSE 0 END;

SELECT @clustered = i.name FROM sys.indexes i
WHERE i.object_id = OBJECT_ID('dbo.Security') AND i.index_id = 1;

PRINT '===== SECURITY pre-check =====';
PRINT '  rows           = ' + CAST(@rows AS varchar(20));
PRINT '  NULL keys      = ' + CAST(@nulls AS varchar(20));
PRINT '  duplicate keys = ' + CAST(@dups AS varchar(20));
PRINT '  existing PK?   = ' + CAST(@hasPk AS varchar(1));
PRINT '  clustered idx  = ' + ISNULL(@clustered, '(heap)');

IF @hasPk = 1
    PRINT '  -> SKIP: Security already has a primary key.';
ELSE IF @nulls > 0 OR @dups > 0
    PRINT '  -> SKIP: cannot add PK (NULL or duplicate Security_DPA_ values present). Investigate first.';
ELSE
BEGIN
    PRINT '  -> APPLYING: NOT NULL + NONCLUSTERED PK_Security ...';
    ALTER TABLE dbo.Security ALTER COLUMN Security_DPA_ int NOT NULL;
    ALTER TABLE dbo.Security ADD CONSTRAINT PK_Security PRIMARY KEY NONCLUSTERED (Security_DPA_);
    PRINT '  -> DONE: PK_Security created.';
END
GO

/* ===================== VERIFY ======================================= */
PRINT '===== VERIFY: primary keys now present =====';
SELECT t.name AS table_name, kc.name AS pk_name, i.type_desc AS pk_type
FROM sys.key_constraints kc
JOIN sys.tables  t ON t.object_id = kc.parent_object_id
JOIN sys.indexes i ON i.object_id = kc.parent_object_id AND i.index_id = kc.unique_index_id
WHERE kc.type = 'PK' AND t.name IN ('Payment','Security')
ORDER BY t.name;
GO
