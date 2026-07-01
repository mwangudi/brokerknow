/* Runs IN BrokerKnow_Clean. (1) widen the two columns the fresh .bak reverted,
   (2) add the legacy tables Malawi0701 has that clean lacks (skipping the client-
   clear snapshots) so the desktop's legacy views can rebuild. */
SET NOCOUNT ON;

-- (1) Column widenings (fix the pre-existing prod regression, keep nullability).
DECLARE @n nvarchar(10);
SET @n = (SELECT CASE WHEN is_nullable=1 THEN 'NULL' ELSE 'NOT NULL' END FROM sys.columns WHERE object_id=OBJECT_ID('dbo.Broker') AND name='BrokerCode');
IF @n IS NOT NULL AND (SELECT max_length FROM sys.columns WHERE object_id=OBJECT_ID('dbo.Broker') AND name='BrokerCode') < 40
    EXEC('ALTER TABLE dbo.Broker ALTER COLUMN BrokerCode nvarchar(20) ' + @n);
SET @n = (SELECT CASE WHEN is_nullable=1 THEN 'NULL' ELSE 'NOT NULL' END FROM sys.columns WHERE object_id=OBJECT_ID('dbo.Client') AND name='ClientCDSNo');
IF @n IS NOT NULL AND (SELECT max_length FROM sys.columns WHERE object_id=OBJECT_ID('dbo.Client') AND name='ClientCDSNo') < 100
    EXEC('ALTER TABLE dbo.Client ALTER COLUMN ClientCDSNo nvarchar(50) ' + @n);
PRINT 'widenings applied';

-- (2) Add legacy tables present in Malawi0701 but missing from clean (so views resolve).
DECLARE @t sysname, @sql nvarchar(max), @added int = 0;
DECLARE c CURSOR LOCAL FAST_FORWARD FOR
    SELECT s.name FROM BrokerKnow_Malawi0701.sys.tables s
    WHERE s.schema_id = SCHEMA_ID('dbo')
      AND s.name NOT LIKE '%_removed_clients_%'
      AND NOT EXISTS (SELECT 1 FROM sys.tables c WHERE c.name = s.name AND c.schema_id = SCHEMA_ID('dbo'));
OPEN c; FETCH NEXT FROM c INTO @t;
WHILE @@FETCH_STATUS = 0
BEGIN
    SET @sql = 'SELECT * INTO dbo.' + QUOTENAME(@t) + ' FROM BrokerKnow_Malawi0701.dbo.' + QUOTENAME(@t) + ';';
    BEGIN TRY EXEC sp_executesql @sql; SET @added += 1; END TRY
    BEGIN CATCH PRINT 'FAIL ' + @t + ': ' + ERROR_MESSAGE(); END CATCH
    FETCH NEXT FROM c INTO @t;
END
CLOSE c; DEALLOCATE c;
PRINT CONCAT('legacy tables added: ', @added);
