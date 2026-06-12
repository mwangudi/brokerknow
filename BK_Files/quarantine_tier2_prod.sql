/* =====================================================================
   Tier-2 QUARANTINE ? move 30 EMPTY (0-row) unmapped tables with 0 inbound
   FKs and 0 view/proc/function references into [trash]. Target: test then prod.
   Same reversible approach as Tier-1 (ALTER SCHEMA dbo TRANSFER trash.<t>).
   The 28 referenced empty tables (InterTransfer, CPortfolio, BondProposals,
   ForwardRate, etc.) are LEFT in dbo ? their legacy views/procs must be
   retired first. Idempotent + a live re-check of refs/FKs/row-count per table.
   ===================================================================== */
SET NOCOUNT ON;
SET XACT_ABORT ON;
USE BrokerKnow;
GO

IF NOT EXISTS (SELECT 1 FROM sys.schemas WHERE name='trash') EXEC('CREATE SCHEMA trash');
GO

DECLARE @safe TABLE (name sysname PRIMARY KEY);
INSERT INTO @safe (name) VALUES
 ('_ATS_Downloaded_Files_'),('_ImportBondsEquitiesTurnover'),('_ImportPriceList'),
 ('BatchLog'),('BondsInput'),('BondsInputOutput'),('BondsInputs'),('BondsOutput'),
 ('BondsOutputs'),('DataStream_Securities_'),('EmailConfigurations'),('EmailedDocs'),
 ('FineTradingScheduleFinal'),('MailMerge'),('MailMergeDetails'),('PathConfigurations'),
 ('Results'),('SendBatchItems'),('SendClientReports'),('SendDaily'),('SendMonthly'),
 ('SendQuarterly'),('SendWeekly'),('tblAppender'),('tblIniSetup'),('tblIPOsetup'),
 ('tblschedule'),('tblUsers'),('TotalKnown'),('Views');

DECLARE @n sysname, @sql nvarchar(max), @moved int = 0;
DECLARE c CURSOR LOCAL FAST_FORWARD FOR SELECT name FROM @safe;
OPEN c; FETCH NEXT FROM c INTO @n;
WHILE @@FETCH_STATUS = 0
BEGIN
  IF NOT EXISTS (SELECT 1 FROM sys.tables t JOIN sys.schemas s ON s.schema_id=t.schema_id WHERE s.name='dbo' AND t.name=@n)
      PRINT 'skip (not in dbo): ' + @n;
  ELSE IF EXISTS (SELECT 1 FROM sys.foreign_keys fk WHERE OBJECT_NAME(fk.referenced_object_id)=@n)
      PRINT 'SKIP (inbound FK): ' + @n;
  ELSE IF EXISTS (SELECT 1 FROM sys.sql_expression_dependencies d WHERE d.referenced_entity_name=@n)
      PRINT 'SKIP (referenced): ' + @n;
  ELSE IF EXISTS (SELECT 1 FROM sys.partitions p JOIN sys.tables t ON t.object_id=p.object_id
                  WHERE t.name=@n AND t.schema_id=SCHEMA_ID('dbo') AND p.index_id IN (0,1) AND p.rows > 0)
      PRINT 'SKIP (not empty anymore): ' + @n;
  ELSE
  BEGIN
      SET @sql = N'ALTER SCHEMA trash TRANSFER dbo.' + QUOTENAME(@n) + N';';
      EXEC sp_executesql @sql;
      SET @moved = @moved + 1;
      PRINT 'moved -> trash: ' + @n;
  END
  FETCH NEXT FROM c INTO @n;
END
CLOSE c; DEALLOCATE c;
PRINT 'TOTAL moved: ' + CAST(@moved AS varchar(10));
GO

PRINT '===== trash schema total =====';
SELECT COUNT(*) AS trash_tables FROM sys.tables t JOIN sys.schemas s ON s.schema_id=t.schema_id WHERE s.name='trash';
SELECT COUNT(*) AS dbo_tables FROM sys.tables t JOIN sys.schemas s ON s.schema_id=t.schema_id WHERE s.name='dbo';
GO
