/* KE migration fix pass 2 — GENERALIZED. For the 3 still-failing tables
   (Account, Client, PaymentRequests), find EVERY not-null column in the clean
   schema that actually contains NULLs in the KE legacy source, relax them all to
   nullable in one pass, then reload. Avoids fixing one column at a time. */
SET NOCOUNT ON;
USE BrokerKnow_KE_Clean;
GO

DECLARE @tables TABLE (name sysname);
INSERT INTO @tables VALUES ('Account'),('Client'),('PaymentRequests');

-- 1. Discover not-null clean columns that exist in legacy AND have >=1 NULL there.
IF OBJECT_ID('tempdb..#fix') IS NOT NULL DROP TABLE #fix;
CREATE TABLE #fix (tbl sysname, col sysname, typedef nvarchar(200));

DECLARE @t sysname, @c sysname, @typedef nvarchar(200), @sql nvarchar(max), @null_cnt int;
DECLARE cur CURSOR LOCAL FAST_FORWARD FOR
  SELECT t.name, c.name,
         ty.name +
         CASE
           WHEN ty.name IN ('decimal','numeric') THEN '(' + CAST(c.precision AS varchar(5)) + ',' + CAST(c.scale AS varchar(5)) + ')'
           WHEN ty.name IN ('varchar','char','varbinary','binary') THEN '(' + CASE WHEN c.max_length = -1 THEN 'max' ELSE CAST(c.max_length AS varchar(10)) END + ')'
           WHEN ty.name IN ('nvarchar','nchar') THEN '(' + CASE WHEN c.max_length = -1 THEN 'max' ELSE CAST(c.max_length/2 AS varchar(10)) END + ')'
           WHEN ty.name IN ('datetime2','time','datetimeoffset') THEN '(' + CAST(c.scale AS varchar(5)) + ')'
           ELSE ''
         END AS typedef
  FROM sys.columns c
  JOIN sys.tables t ON t.object_id = c.object_id
  JOIN sys.types ty ON ty.user_type_id = c.user_type_id
  WHERE t.name IN (SELECT name FROM @tables)
    AND c.is_nullable = 0
    AND c.is_computed = 0
    AND c.is_identity = 0
    AND EXISTS (SELECT 1 FROM BrokerKnow_KE_Legacy.sys.columns sc
                WHERE sc.object_id = OBJECT_ID('BrokerKnow_KE_Legacy.dbo.' + QUOTENAME(t.name))
                  AND sc.name = c.name);
OPEN cur; FETCH NEXT FROM cur INTO @t, @c, @typedef;
WHILE @@FETCH_STATUS = 0
BEGIN
    SET @sql = N'SELECT @n = COUNT(*) FROM BrokerKnow_KE_Legacy.dbo.' + QUOTENAME(@t) + ' WHERE ' + QUOTENAME(@c) + ' IS NULL;';
    SET @null_cnt = 0;
    EXEC sp_executesql @sql, N'@n int OUTPUT', @n = @null_cnt OUTPUT;
    IF @null_cnt > 0 INSERT INTO #fix VALUES (@t, @c, @typedef);
    FETCH NEXT FROM cur INTO @t, @c, @typedef;
END
CLOSE cur; DEALLOCATE cur;

PRINT '=== columns to relax (have NULLs in KE legacy) ===';
SELECT tbl, col, typedef FROM #fix ORDER BY tbl, col;

-- 2. Relax them all.
DECLARE cur2 CURSOR LOCAL FAST_FORWARD FOR SELECT tbl, col, typedef FROM #fix;
OPEN cur2; FETCH NEXT FROM cur2 INTO @t, @c, @typedef;
WHILE @@FETCH_STATUS = 0
BEGIN
    SET @sql = N'ALTER TABLE dbo.' + QUOTENAME(@t) + ' ALTER COLUMN ' + QUOTENAME(@c) + ' ' + @typedef + ' NULL;';
    EXEC sp_executesql @sql;
    FETCH NEXT FROM cur2 INTO @t, @c, @typedef;
END
CLOSE cur2; DEALLOCATE cur2;
DECLARE @fixn int = (SELECT COUNT(*) FROM #fix);
PRINT '-- relaxed ' + CAST(@fixn AS varchar(10)) + ' column(s).';

-- 3. Reload the 3 tables (intersect columns, identity-safe).
DECLARE @cols nvarchar(max), @hasIdent bit, @rows int;
DECLARE cur3 CURSOR LOCAL FAST_FORWARD FOR SELECT name FROM @tables;
OPEN cur3; FETCH NEXT FROM cur3 INTO @t;
WHILE @@FETCH_STATUS = 0
BEGIN
    EXEC('ALTER TABLE dbo.' + @t + ' NOCHECK CONSTRAINT ALL;');
    -- empty first in case a partial earlier attempt left rows (failed inserts are atomic, but be safe)
    EXEC('DELETE FROM dbo.' + @t + ';');

    SELECT @cols = STRING_AGG(QUOTENAME(tc.name), ',') WITHIN GROUP (ORDER BY tc.column_id)
    FROM sys.columns tc JOIN sys.types ty ON ty.user_type_id = tc.user_type_id
    WHERE tc.object_id = OBJECT_ID('dbo.' + QUOTENAME(@t))
      AND tc.is_computed = 0 AND ty.name NOT IN ('timestamp','rowversion')
      AND EXISTS (SELECT 1 FROM BrokerKnow_KE_Legacy.sys.columns sc
                  WHERE sc.object_id = OBJECT_ID('BrokerKnow_KE_Legacy.dbo.' + QUOTENAME(@t)) AND sc.name = tc.name);

    SET @hasIdent = CASE WHEN EXISTS (SELECT 1 FROM sys.identity_columns WHERE object_id = OBJECT_ID('dbo.' + QUOTENAME(@t))) THEN 1 ELSE 0 END;
    SET @sql =
        CASE WHEN @hasIdent = 1 THEN 'SET IDENTITY_INSERT dbo.' + QUOTENAME(@t) + ' ON; ' ELSE '' END
      + 'INSERT INTO dbo.' + QUOTENAME(@t) + ' (' + @cols + ') SELECT ' + @cols + ' FROM BrokerKnow_KE_Legacy.dbo.' + QUOTENAME(@t) + ';'
      + CASE WHEN @hasIdent = 1 THEN ' SET IDENTITY_INSERT dbo.' + QUOTENAME(@t) + ' OFF;' ELSE '' END;
    BEGIN TRY
        EXEC sp_executesql @sql; SET @rows = @@ROWCOUNT;
        PRINT @t + ': ' + CAST(@rows AS varchar(20));
    END TRY
    BEGIN CATCH PRINT @t + ' !! FAILED: ' + ERROR_MESSAGE(); END CATCH
    FETCH NEXT FROM cur3 INTO @t;
END
CLOSE cur3; DEALLOCATE cur3;
GO
