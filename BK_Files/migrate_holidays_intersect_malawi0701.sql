SET NOCOUNT ON;
USE BrokerKnow_Clean;
-- Load Holidays using ONLY the columns present in BOTH schemas (the clean
-- schema added 'Recurring'; the fresh Malawi legacy dump lacks it → let it
-- default). Same intersect approach used for the KE migration.
DECLARE @cols nvarchar(max), @hasId bit, @sql nvarchar(max);
SELECT @cols = STRING_AGG(QUOTENAME(c.name), ',') WITHIN GROUP (ORDER BY c.column_id)
FROM sys.columns c JOIN sys.types ty ON ty.user_type_id = c.user_type_id
WHERE c.object_id = OBJECT_ID('dbo.Holidays') AND c.is_computed = 0
  AND ty.name NOT IN ('timestamp','rowversion')
  AND EXISTS (SELECT 1 FROM BrokerKnow_Malawi0701.sys.columns sc
              WHERE sc.object_id = OBJECT_ID('BrokerKnow_Malawi0701.dbo.Holidays') AND sc.name = c.name);
SET @hasId = CASE WHEN EXISTS (SELECT 1 FROM sys.identity_columns WHERE object_id = OBJECT_ID('dbo.Holidays')) THEN 1 ELSE 0 END;
SET @sql = CASE WHEN @hasId = 1 THEN 'SET IDENTITY_INSERT dbo.Holidays ON; ' ELSE '' END
    + 'INSERT INTO dbo.Holidays (' + @cols + ') SELECT ' + @cols + ' FROM BrokerKnow_Malawi0701.dbo.Holidays;'
    + CASE WHEN @hasId = 1 THEN ' SET IDENTITY_INSERT dbo.Holidays OFF;' ELSE '' END;
EXEC sp_executesql @sql;
SELECT 'Holidays loaded' AS info, @@ROWCOUNT AS rows_loaded;
