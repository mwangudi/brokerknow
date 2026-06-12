/* =====================================================================
   Tier-1 QUARANTINE — move 11 zero-reference backup/temp tables out of dbo
   into a [trash] schema. Target: BrokerKnow_Test first, then prod.

   WHY quarantine (not DROP): the legacy desktop app shares this DB and we
   have a side-by-side run (Aug) ahead. Moving to [trash] makes dbo clean NOW
   but is INSTANTLY REVERSIBLE if the side-by-side reveals a need:
       ALTER SCHEMA dbo TRANSFER trash.<table>;
   After the side-by-side proves nothing references them, final cleanup is a
   plain DROP TABLE trash.<table>.

   ONLY the 11 tables proven to have: 0 inbound FKs AND 0 view/proc/function
   references (see quarantine_mapping.sql). The 5 referenced ones
   (datastream_Market, clientBalancesTemp, excep_SummaryHoldings,
   PrimaryIssues, temp_MarchContractSchedule) are intentionally LEFT in dbo.

   Idempotent: skips a table already in trash / not in dbo.
   ===================================================================== */
SET NOCOUNT ON;
SET XACT_ABORT ON;
USE BrokerKnow_Test;
GO

IF NOT EXISTS (SELECT 1 FROM sys.schemas WHERE name = 'trash')
    EXEC('CREATE SCHEMA trash');
GO

DECLARE @safe TABLE (name sysname PRIMARY KEY);
INSERT INTO @safe (name) VALUES
 ('LevyContract_20081120'),('LevyContract_20081120a'),('LevyContract20081126'),
 ('LevyContract2009Feb9'),('MenuGroups--'),('MenuGroupsBKP'),('Menus--'),
 ('OrdDetail2009Feb5'),('temp01'),('TempInstitutionMapping'),('Users20081124');

DECLARE @n sysname, @sql nvarchar(max), @moved int = 0;
DECLARE c CURSOR LOCAL FAST_FORWARD FOR SELECT name FROM @safe;
OPEN c; FETCH NEXT FROM c INTO @n;
WHILE @@FETCH_STATUS = 0
BEGIN
    IF EXISTS (SELECT 1 FROM sys.tables t JOIN sys.schemas s ON s.schema_id=t.schema_id
               WHERE s.name='dbo' AND t.name=@n)
    BEGIN
        -- final safety: re-confirm no inbound FK right now
        IF EXISTS (SELECT 1 FROM sys.foreign_keys fk WHERE OBJECT_NAME(fk.referenced_object_id)=@n)
            PRINT 'SKIP (has inbound FK): ' + @n;
        ELSE
        BEGIN
            SET @sql = N'ALTER SCHEMA trash TRANSFER dbo.' + QUOTENAME(@n) + N';';
            EXEC sp_executesql @sql;
            SET @moved = @moved + 1;
            PRINT 'moved -> trash: ' + @n;
        END
    END
    ELSE PRINT 'skip (not in dbo): ' + @n;
    FETCH NEXT FROM c INTO @n;
END
CLOSE c; DEALLOCATE c;
PRINT 'TOTAL moved: ' + CAST(@moved AS varchar(10));
GO

PRINT '===== trash schema contents =====';
SELECT s.name + '.' + t.name AS quarantined, p.rows AS row_count
FROM sys.tables t JOIN sys.schemas s ON s.schema_id=t.schema_id
JOIN sys.partitions p ON p.object_id=t.object_id AND p.index_id IN (0,1)
WHERE s.name='trash' ORDER BY t.name;
GO
