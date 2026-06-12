/* =====================================================================
   MIGRATION LOAD — copy data from prod BrokerKnow -> BrokerKnow_Clean.
   Runs IN BrokerKnow_Clean. Source (BrokerKnow) is READ-ONLY.
   FKs are disabled during load (order-independent), identity preserved,
   computed/rowversion columns skipped. Audit logs excluded by design.
   Re-running: clean DB should be freshly built (empty) before this.
   ===================================================================== */
SET NOCOUNT ON;
USE BrokerKnow_Clean;
GO

-- 1. disable all FK constraints so load order doesn't matter
DECLARE @off nvarchar(max) = N'';
SELECT @off = @off + 'ALTER TABLE ' + QUOTENAME(SCHEMA_NAME(t.schema_id)) + '.' + QUOTENAME(t.name) + ' NOCHECK CONSTRAINT ALL;' + CHAR(10)
FROM sys.tables t WHERE t.schema_id = SCHEMA_ID('dbo');
EXEC sp_executesql @off;
PRINT '-- FKs disabled.';

-- 2. copy each table (except audit logs)
DECLARE @t sysname, @cols nvarchar(max), @hasIdent bit, @sql nvarchar(max), @rows int, @total bigint = 0, @failed int = 0;
DECLARE c CURSOR LOCAL FAST_FORWARD FOR
  SELECT t.name FROM sys.tables t
  WHERE t.schema_id = SCHEMA_ID('dbo')
    AND t.name NOT IN ('AuditTrail','AuditTrailItem')   -- intentional: historical logging excluded
  ORDER BY t.name;
OPEN c; FETCH NEXT FROM c INTO @t;
WHILE @@FETCH_STATUS = 0
BEGIN
    SELECT @cols = STRING_AGG(QUOTENAME(c.name), ',') WITHIN GROUP (ORDER BY c.column_id)
    FROM sys.columns c JOIN sys.types ty ON ty.user_type_id = c.user_type_id
    WHERE c.object_id = OBJECT_ID('dbo.' + QUOTENAME(@t))
      AND c.is_computed = 0 AND ty.name NOT IN ('timestamp','rowversion');

    SET @hasIdent = CASE WHEN EXISTS (SELECT 1 FROM sys.identity_columns WHERE object_id = OBJECT_ID('dbo.' + QUOTENAME(@t))) THEN 1 ELSE 0 END;

    SET @sql =
        CASE WHEN @hasIdent = 1 THEN 'SET IDENTITY_INSERT dbo.' + QUOTENAME(@t) + ' ON; ' ELSE '' END
      + 'INSERT INTO dbo.' + QUOTENAME(@t) + ' (' + @cols + ') SELECT ' + @cols + ' FROM BrokerKnow.dbo.' + QUOTENAME(@t) + ';'
      + CASE WHEN @hasIdent = 1 THEN ' SET IDENTITY_INSERT dbo.' + QUOTENAME(@t) + ' OFF;' ELSE '' END;

    BEGIN TRY
        EXEC sp_executesql @sql;
        SET @rows = @@ROWCOUNT; SET @total = @total + @rows;
        PRINT @t + ': ' + CAST(@rows AS varchar(20));
    END TRY
    BEGIN CATCH
        SET @failed = @failed + 1;
        PRINT @t + ' !! FAILED: ' + ERROR_MESSAGE();
    END CATCH
    FETCH NEXT FROM c INTO @t;
END
CLOSE c; DEALLOCATE c;
PRINT '-- TOTAL rows loaded: ' + CAST(@total AS varchar(30)) + ' ; tables failed: ' + CAST(@failed AS varchar(10));
GO
