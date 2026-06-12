/* migrate_fk_check.sql — run IN BrokerKnow_Clean.
   Re-enables every dbo FK WITH CHECK to validate migrated referential integrity.
   Any FK that fails has orphan rows in the migrated data. Read-only report of results. */
SET NOCOUNT ON;
USE BrokerKnow_Clean;
GO
CREATE TABLE #r (fk sysname, tbl sysname, status varchar(20), err nvarchar(2000));
DECLARE @fk sysname, @tbl sysname, @sql nvarchar(max);
DECLARE c CURSOR LOCAL FAST_FORWARD FOR
    SELECT fk.name, OBJECT_NAME(fk.parent_object_id)
    FROM sys.foreign_keys fk
    JOIN sys.tables t ON t.object_id = fk.parent_object_id
    WHERE t.schema_id = SCHEMA_ID('dbo');
OPEN c; FETCH NEXT FROM c INTO @fk, @tbl;
WHILE @@FETCH_STATUS = 0
BEGIN
    SET @sql = 'ALTER TABLE dbo.' + QUOTENAME(@tbl) + ' WITH CHECK CHECK CONSTRAINT ' + QUOTENAME(@fk) + ';';
    BEGIN TRY
        EXEC (@sql);
        INSERT #r VALUES (@fk, @tbl, 'OK', NULL);
    END TRY
    BEGIN CATCH
        INSERT #r VALUES (@fk, @tbl, 'VIOLATION', ERROR_MESSAGE());
    END CATCH;
    FETCH NEXT FROM c INTO @fk, @tbl;
END
CLOSE c; DEALLOCATE c;

PRINT '===== FK VIOLATIONS (orphan rows in migrated data) =====';
SELECT fk, tbl, err FROM #r WHERE status = 'VIOLATION' ORDER BY tbl, fk;

PRINT '===== summary =====';
SELECT COUNT(*) AS total_fks,
       SUM(CASE WHEN status = 'OK' THEN 1 ELSE 0 END) AS trusted,
       SUM(CASE WHEN status = 'VIOLATION' THEN 1 ELSE 0 END) AS violations
FROM #r;

-- Confirm all FKs are now trusted (is_not_trusted = 0)
PRINT '===== untrusted FKs remaining (should be none) =====';
SELECT fk.name, OBJECT_NAME(fk.parent_object_id) AS tbl
FROM sys.foreign_keys fk
JOIN sys.tables t ON t.object_id = fk.parent_object_id
WHERE t.schema_id = SCHEMA_ID('dbo') AND fk.is_not_trusted = 1
ORDER BY tbl;
GO
