/* =====================================================================
   MIGRATION LOAD (KENYA) — copy data from BrokerKnow_KE_Legacy ->
   BrokerKnow_KE_Clean. Runs IN BrokerKnow_KE_Clean (the hardened target).
   Source (BrokerKnow_KE_Legacy) is READ-ONLY. FKs disabled during load,
   identity preserved, computed/rowversion columns skipped. Audit logs
   excluded by design. App-layer tables (PortalUsers etc.) are NOT in the
   legacy source — they fail-and-skip here and get grafted from KE_Demo after.
   Only loads columns that exist in BOTH target and source (KE legacy schema
   is older, so some clean columns may be absent there).
   ===================================================================== */
SET NOCOUNT ON;
USE BrokerKnow_KE_Clean;
GO

-- 1. disable all FK constraints so load order doesn't matter
DECLARE @off nvarchar(max) = N'';
SELECT @off = @off + 'ALTER TABLE ' + QUOTENAME(SCHEMA_NAME(t.schema_id)) + '.' + QUOTENAME(t.name) + ' NOCHECK CONSTRAINT ALL;' + CHAR(10)
FROM sys.tables t WHERE t.schema_id = SCHEMA_ID('dbo');
EXEC sp_executesql @off;
PRINT '-- FKs disabled.';

-- 2. copy each table (except audit logs), intersecting columns target∩source
DECLARE @t sysname, @cols nvarchar(max), @hasIdent bit, @sql nvarchar(max),
        @rows int, @total bigint = 0, @failed int = 0, @skipped int = 0;
DECLARE c CURSOR LOCAL FAST_FORWARD FOR
  SELECT t.name FROM sys.tables t
  WHERE t.schema_id = SCHEMA_ID('dbo')
    AND t.name NOT IN ('AuditTrail','AuditTrailItem')   -- intentional: historical logging excluded
  ORDER BY t.name;
OPEN c; FETCH NEXT FROM c INTO @t;
WHILE @@FETCH_STATUS = 0
BEGIN
    -- Skip target tables that don't exist in the legacy source (e.g. app-layer).
    IF OBJECT_ID('BrokerKnow_KE_Legacy.dbo.' + QUOTENAME(@t)) IS NULL
    BEGIN
        SET @skipped = @skipped + 1;
        PRINT @t + ' -- skipped (not in legacy source)';
        FETCH NEXT FROM c INTO @t; CONTINUE;
    END

    -- Column list = columns present in BOTH target and source (non-computed, non-rowversion).
    SELECT @cols = STRING_AGG(QUOTENAME(tc.name), ',') WITHIN GROUP (ORDER BY tc.column_id)
    FROM sys.columns tc
    JOIN sys.types ty ON ty.user_type_id = tc.user_type_id
    WHERE tc.object_id = OBJECT_ID('dbo.' + QUOTENAME(@t))
      AND tc.is_computed = 0 AND ty.name NOT IN ('timestamp','rowversion')
      AND EXISTS (SELECT 1 FROM BrokerKnow_KE_Legacy.sys.columns sc
                  WHERE sc.object_id = OBJECT_ID('BrokerKnow_KE_Legacy.dbo.' + QUOTENAME(@t))
                    AND sc.name = tc.name);

    IF @cols IS NULL
    BEGIN
        SET @skipped = @skipped + 1;
        PRINT @t + ' -- skipped (no common columns)';
        FETCH NEXT FROM c INTO @t; CONTINUE;
    END

    SET @hasIdent = CASE WHEN EXISTS (SELECT 1 FROM sys.identity_columns WHERE object_id = OBJECT_ID('dbo.' + QUOTENAME(@t))) THEN 1 ELSE 0 END;

    SET @sql =
        CASE WHEN @hasIdent = 1 THEN 'SET IDENTITY_INSERT dbo.' + QUOTENAME(@t) + ' ON; ' ELSE '' END
      + 'INSERT INTO dbo.' + QUOTENAME(@t) + ' (' + @cols + ') SELECT ' + @cols + ' FROM BrokerKnow_KE_Legacy.dbo.' + QUOTENAME(@t) + ';'
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
PRINT '-- TOTAL rows: ' + CAST(@total AS varchar(30)) + ' ; failed: ' + CAST(@failed AS varchar(10)) + ' ; skipped: ' + CAST(@skipped AS varchar(10));
GO
