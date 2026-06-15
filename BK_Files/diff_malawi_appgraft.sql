SET NOCOUNT ON;
-- New-system tables present in LIVE BrokerKnow.dbo but ABSENT in the Malawi dump.
-- These must be grafted (schema + data) to make the dump app-ready.
SELECT l.name AS missing_in_malawi
FROM BrokerKnow.sys.tables l
WHERE l.schema_id = SCHEMA_ID('dbo')
  AND NOT EXISTS (
    SELECT 1 FROM BrokerKnow_Malawi0612.sys.tables m
    WHERE m.name = l.name AND m.schema_id = SCHEMA_ID('dbo'))
ORDER BY l.name;
GO
PRINT '===== row counts of the key app tables in LIVE (what we would copy) =====';
SELECT 'PortalUsers' AS tbl, COUNT(*) AS n FROM BrokerKnow.dbo.PortalUsers
UNION ALL SELECT 'PortalRefreshTokens', COUNT(*) FROM BrokerKnow.dbo.PortalRefreshTokens
UNION ALL SELECT 'Groups', COUNT(*) FROM BrokerKnow.dbo.Groups
UNION ALL SELECT 'app schema views', (SELECT COUNT(*) FROM BrokerKnow.sys.views WHERE SCHEMA_NAME(schema_id)='app');
GO
