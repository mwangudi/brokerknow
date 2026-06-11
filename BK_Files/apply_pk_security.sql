/* Phase-2b — Security PK: diagnose the NOT NULL failure, then apply safely.
   Target: BrokerKnow_Test. Security data is clean (21 rows, 0 null, 0 dup).
   Integrity-only change, no data rows touched. Reversible:
     ALTER TABLE dbo.Security DROP CONSTRAINT PK_Security;
     ALTER TABLE dbo.Security ALTER COLUMN Security_DPA_ int NULL;   -- if needed
*/
SET NOCOUNT ON;
USE BrokerKnow_Test;
GO

/* ---- 1. schema-bound views/functions referencing dbo.Security ------ */
PRINT '===== schema-bound dependencies on dbo.Security =====';
SELECT OBJECT_NAME(d.referencing_id) AS referencing_object,
       o.type_desc
FROM sys.sql_expression_dependencies d
JOIN sys.objects o ON o.object_id = d.referencing_id
WHERE d.referenced_id = OBJECT_ID('dbo.Security')
  AND d.is_schema_bound_reference = 1;

/* ---- 2. all indexes currently on dbo.Security ---------------------- */
PRINT '===== indexes on dbo.Security =====';
SELECT i.name AS index_name, i.type_desc, i.is_primary_key, i.is_unique
FROM sys.indexes i
WHERE i.object_id = OBJECT_ID('dbo.Security') AND i.index_id > 0;

/* ---- 3. attempt NOT NULL, capturing the REAL error ----------------- */
PRINT '===== attempt ALTER COLUMN Security_DPA_ NOT NULL =====';
BEGIN TRY
    ALTER TABLE dbo.Security ALTER COLUMN Security_DPA_ int NOT NULL;
    PRINT '  OK: column is now NOT NULL.';
END TRY
BEGIN CATCH
    PRINT '  FAILED: ' + ERROR_MESSAGE();
END CATCH
GO

/* ---- 4. add the PK only if the column is now NOT NULL -------------- */
PRINT '===== add PK_Security (only if column is NOT NULL) =====';
IF EXISTS (SELECT 1 FROM sys.columns
           WHERE object_id = OBJECT_ID('dbo.Security')
             AND name = 'Security_DPA_' AND is_nullable = 0)
   AND NOT EXISTS (SELECT 1 FROM sys.indexes
           WHERE object_id = OBJECT_ID('dbo.Security') AND is_primary_key = 1)
BEGIN
    BEGIN TRY
        ALTER TABLE dbo.Security ADD CONSTRAINT PK_Security PRIMARY KEY NONCLUSTERED (Security_DPA_);
        PRINT '  OK: PK_Security created.';
    END TRY
    BEGIN CATCH
        PRINT '  FAILED: ' + ERROR_MESSAGE();
    END CATCH
END
ELSE
    PRINT '  SKIP: column still nullable or PK already exists.';
GO

/* ---- 5. verify ----------------------------------------------------- */
PRINT '===== verify =====';
SELECT c.is_nullable AS security_dpa_nullable,
       (SELECT name FROM sys.key_constraints
        WHERE parent_object_id = OBJECT_ID('dbo.Security') AND type='PK') AS pk_name
FROM sys.columns c
WHERE c.object_id = OBJECT_ID('dbo.Security') AND c.name = 'Security_DPA_';
GO
