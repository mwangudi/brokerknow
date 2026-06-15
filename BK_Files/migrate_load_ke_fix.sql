/* KE migration fix: the hardened (Malawi-derived) clean schema made 5 columns
   NOT NULL that contain NULLs in KE's older legacy data. Relax just those 5 to
   nullable so the real source data loads FAITHFULLY (no invented values), then
   reload only the 5 tables that failed atomically (they hold 0 rows now).
   DIVERGENCE NOTE (for real cutover): decide per-column whether to keep nullable
   or backfill a default (e.g. Client.CreditLimit NULL -> 0). NULL counts in KE:
   Client.CreditLimit=2, Account.ReconStartDate=17, PaymentRequests.FirstApproval=3772,
   Share.SharePDate=10, WebtbOrder.Action=60. */
SET NOCOUNT ON;
USE BrokerKnow_KE_Clean;
GO

ALTER TABLE dbo.Account          ALTER COLUMN ReconStartDate datetime      NULL;
ALTER TABLE dbo.Client           ALTER COLUMN CreditLimit    money         NULL;
ALTER TABLE dbo.PaymentRequests  ALTER COLUMN FirstApproval  bit           NULL;
ALTER TABLE dbo.Share            ALTER COLUMN SharePDate     smalldatetime NULL;
ALTER TABLE dbo.WebtbOrder       ALTER COLUMN [Action]       int           NULL;
GO
PRINT '-- 5 columns relaxed to NULL.';
GO

-- Reload ONLY the 5 previously-failed tables (intersect columns, identity-safe).
DECLARE @list TABLE (name sysname);
INSERT INTO @list VALUES ('Account'),('Client'),('PaymentRequests'),('Share'),('WebtbOrder');

DECLARE @t sysname, @cols nvarchar(max), @hasIdent bit, @sql nvarchar(max), @rows int;
DECLARE c CURSOR LOCAL FAST_FORWARD FOR SELECT name FROM @list;
OPEN c; FETCH NEXT FROM c INTO @t;
WHILE @@FETCH_STATUS = 0
BEGIN
    EXEC('ALTER TABLE dbo.' + @t + ' NOCHECK CONSTRAINT ALL;');

    SELECT @cols = STRING_AGG(QUOTENAME(tc.name), ',') WITHIN GROUP (ORDER BY tc.column_id)
    FROM sys.columns tc
    JOIN sys.types ty ON ty.user_type_id = tc.user_type_id
    WHERE tc.object_id = OBJECT_ID('dbo.' + QUOTENAME(@t))
      AND tc.is_computed = 0 AND ty.name NOT IN ('timestamp','rowversion')
      AND EXISTS (SELECT 1 FROM BrokerKnow_KE_Legacy.sys.columns sc
                  WHERE sc.object_id = OBJECT_ID('BrokerKnow_KE_Legacy.dbo.' + QUOTENAME(@t))
                    AND sc.name = tc.name);

    SET @hasIdent = CASE WHEN EXISTS (SELECT 1 FROM sys.identity_columns WHERE object_id = OBJECT_ID('dbo.' + QUOTENAME(@t))) THEN 1 ELSE 0 END;

    SET @sql =
        CASE WHEN @hasIdent = 1 THEN 'SET IDENTITY_INSERT dbo.' + QUOTENAME(@t) + ' ON; ' ELSE '' END
      + 'INSERT INTO dbo.' + QUOTENAME(@t) + ' (' + @cols + ') SELECT ' + @cols + ' FROM BrokerKnow_KE_Legacy.dbo.' + QUOTENAME(@t) + ';'
      + CASE WHEN @hasIdent = 1 THEN ' SET IDENTITY_INSERT dbo.' + QUOTENAME(@t) + ' OFF;' ELSE '' END;

    BEGIN TRY
        EXEC sp_executesql @sql;
        SET @rows = @@ROWCOUNT;
        PRINT @t + ': ' + CAST(@rows AS varchar(20));
    END TRY
    BEGIN CATCH
        PRINT @t + ' !! FAILED: ' + ERROR_MESSAGE();
    END CATCH
    FETCH NEXT FROM c INTO @t;
END
CLOSE c; DEALLOCATE c;
GO
