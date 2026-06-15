/* KE migration FK check — re-enable every FK WITH CHECK in BrokerKnow_KE_Clean
   and report any that fail (i.e. rows that violate referential integrity).
   A FK that re-enables WITH CHECK and stays trusted = that relationship is clean
   across the entire migrated KE history. READ-ONLY-ish (only toggles constraints).*/
SET NOCOUNT ON;
USE BrokerKnow_KE_Clean;
GO

DECLARE @fk sysname, @tbl sysname, @sql nvarchar(max), @ok int = 0, @bad int = 0;
IF OBJECT_ID('tempdb..#fk') IS NOT NULL DROP TABLE #fk;
CREATE TABLE #fk (fk sysname, tbl sysname, result nvarchar(400));

DECLARE c CURSOR LOCAL FAST_FORWARD FOR
  SELECT fk.name, OBJECT_NAME(fk.parent_object_id)
  FROM sys.foreign_keys fk ORDER BY fk.name;
OPEN c; FETCH NEXT FROM c INTO @fk, @tbl;
WHILE @@FETCH_STATUS = 0
BEGIN
    SET @sql = N'ALTER TABLE dbo.' + QUOTENAME(@tbl) + ' WITH CHECK CHECK CONSTRAINT ' + QUOTENAME(@fk) + ';';
    BEGIN TRY
        EXEC sp_executesql @sql;
        INSERT INTO #fk VALUES (@fk, @tbl, 'OK (trusted)');
        SET @ok = @ok + 1;
    END TRY
    BEGIN CATCH
        INSERT INTO #fk VALUES (@fk, @tbl, 'VIOLATION: ' + ERROR_MESSAGE());
        SET @bad = @bad + 1;
    END CATCH
    FETCH NEXT FROM c INTO @fk, @tbl;
END
CLOSE c; DEALLOCATE c;

PRINT '===== FK results =====';
SELECT fk, tbl, result FROM #fk WHERE result <> 'OK (trusted)' ORDER BY fk;
PRINT '===== summary =====';
SELECT @ok AS fks_trusted, @bad AS fks_violated;

-- Confirm none are left untrusted (is_not_trusted=1 means it didn't validate).
PRINT '===== untrusted FKs remaining (should be 0) =====';
SELECT name, OBJECT_NAME(parent_object_id) AS tbl
FROM sys.foreign_keys WHERE is_not_trusted = 1 ORDER BY name;
GO
