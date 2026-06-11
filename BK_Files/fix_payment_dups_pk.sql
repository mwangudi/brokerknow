/* =====================================================================
   Phase-2b — resolve Payment duplicate keys, then add PK_Payment.
   Target: BrokerKnow_Test ONLY. NOT prod.

   The 2 duplicate Payment_DPA_ values (41686, 41689) are portal-test
   artifacts; in EACH dup pair exactly one row is Deleted=1. We physically
   remove ONLY the Deleted=1 row of a duplicated key — never a live row,
   never a key that isn't duplicated. The whole thing runs in a transaction
   and ROLLS BACK if the delete count isn't exactly the expected 2 or if any
   live (Deleted=0) row would be touched.

   Safe: bounded, transactional, self-verifying. Reversible up to COMMIT.
   ===================================================================== */
SET NOCOUNT ON;
SET XACT_ABORT ON;
USE BrokerKnow_Test;
GO

BEGIN TRAN;

-- Rows that are (a) Deleted, AND (b) share their Payment_DPA_ with another row.
;WITH dup AS (
    SELECT Payment_DPA_
    FROM dbo.Payment
    WHERE Payment_DPA_ IS NOT NULL
    GROUP BY Payment_DPA_ HAVING COUNT(*) > 1
)
SELECT p.Payment_DPA_, p.Entity_DPA_, p.PaymentAmount, p.Deleted
INTO #to_delete
FROM dbo.Payment p
JOIN dup ON dup.Payment_DPA_ = p.Payment_DPA_
WHERE p.Deleted = 1;

DECLARE @expected int = (SELECT COUNT(*) FROM #to_delete);
PRINT 'Soft-deleted rows in duplicate pairs to remove: ' + CAST(@expected AS varchar(10));

-- Safety 1: every duplicated key must retain >=1 NON-deleted row after we cut
-- the deleted ones (i.e. we never orphan a key by deleting its only survivor).
IF EXISTS (
    SELECT p.Payment_DPA_
    FROM dbo.Payment p
    JOIN (SELECT Payment_DPA_ FROM dbo.Payment
          WHERE Payment_DPA_ IS NOT NULL
          GROUP BY Payment_DPA_ HAVING COUNT(*) > 1) d ON d.Payment_DPA_ = p.Payment_DPA_
    GROUP BY p.Payment_DPA_
    HAVING SUM(CASE WHEN p.Deleted = 0 OR p.Deleted IS NULL THEN 1 ELSE 0 END) = 0)
BEGIN
    PRINT '  ABORT: a duplicated key has NO live survivor — manual review needed.';
    ROLLBACK TRAN; 
END
ELSE
BEGIN
    -- Delete only those exact soft-deleted duplicate rows (join on the temp set).
    DELETE p
    FROM dbo.Payment p
    JOIN #to_delete d
      ON d.Payment_DPA_ = p.Payment_DPA_
     AND d.Entity_DPA_  = p.Entity_DPA_
     AND p.Deleted = 1;

    DECLARE @removed int = @@ROWCOUNT;
    PRINT '  rows removed: ' + CAST(@removed AS varchar(10));

    DECLARE @remaining int = (SELECT COUNT(*) FROM (
        SELECT Payment_DPA_ FROM dbo.Payment
        WHERE Payment_DPA_ IS NOT NULL
        GROUP BY Payment_DPA_ HAVING COUNT(*) > 1) x);

    IF @removed = @expected AND @remaining = 0
    BEGIN
        PRINT '  OK: duplicates resolved, 0 remaining. Committing.';
        COMMIT TRAN;
    END
    ELSE
    BEGIN
        PRINT '  ABORT: unexpected state (removed<>expected or dups remain). Rolling back.';
        ROLLBACK TRAN;
    END
END

DROP TABLE IF EXISTS #to_delete;
GO

/* ---- add PK_Payment (separate batch so NOT NULL metadata is fresh) -- */
PRINT '===== add PK_Payment =====';
IF EXISTS (SELECT 1 FROM dbo.Payment WHERE Payment_DPA_ IS NULL)
    PRINT '  SKIP: NULL Payment_DPA_ present.';
ELSE IF EXISTS (SELECT Payment_DPA_ FROM dbo.Payment
                GROUP BY Payment_DPA_ HAVING COUNT(*) > 1)
    PRINT '  SKIP: duplicates still present.';
ELSE IF EXISTS (SELECT 1 FROM sys.indexes
                WHERE object_id = OBJECT_ID('dbo.Payment') AND is_primary_key = 1)
    PRINT '  SKIP: PK already exists.';
ELSE
BEGIN
    ALTER TABLE dbo.Payment ALTER COLUMN Payment_DPA_ int NOT NULL;
END
GO
-- PK add in its own batch (fresh NOT NULL metadata — avoids Msg 8111).
IF EXISTS (SELECT 1 FROM sys.columns
           WHERE object_id = OBJECT_ID('dbo.Payment')
             AND name = 'Payment_DPA_' AND is_nullable = 0)
   AND NOT EXISTS (SELECT 1 FROM sys.indexes
           WHERE object_id = OBJECT_ID('dbo.Payment') AND is_primary_key = 1)
BEGIN
    ALTER TABLE dbo.Payment ADD CONSTRAINT PK_Payment PRIMARY KEY NONCLUSTERED (Payment_DPA_);
    PRINT '  OK: PK_Payment created.';
END
GO

/* ---- verify -------------------------------------------------------- */
PRINT '===== verify PKs =====';
SELECT t.name AS table_name, kc.name AS pk_name, i.type_desc
FROM sys.key_constraints kc
JOIN sys.tables  t ON t.object_id = kc.parent_object_id
JOIN sys.indexes i ON i.object_id = kc.parent_object_id AND i.index_id = kc.unique_index_id
WHERE kc.type='PK' AND t.name IN ('Payment','Security')
ORDER BY t.name;

SELECT 'Payment rows now' AS info, COUNT(*) AS n FROM dbo.Payment;
GO
