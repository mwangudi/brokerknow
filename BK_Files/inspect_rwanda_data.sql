SET NOCOUNT ON;
USE BrokerKnow_RW_0612;

PRINT '=== Core business data counts (Rwanda) ===';
SELECT
  (SELECT COUNT(*) FROM dbo.Client)        AS Clients,
  (SELECT COUNT(*) FROM dbo.tbOrder)       AS Orders,
  (SELECT COUNT(*) FROM dbo.Contract)      AS Contracts,
  (SELECT COUNT(*) FROM dbo.Lot)           AS Lots,
  (SELECT COUNT(*) FROM dbo.Payment)       AS Payments,
  (SELECT COUNT(*) FROM dbo.Security)      AS Securities,
  (SELECT COUNT(*) FROM dbo.Agent)         AS Agents,
  (SELECT COUNT(*) FROM dbo.Broker)        AS Brokers;

PRINT '=== Securities on the Rwanda book (the RSE listings) ===';
SELECT TOP 40 Security_DPA_ AS Id, SecurityCode, SecurityName, SecurityMktPrice
FROM dbo.Security
ORDER BY Security_DPA_;

PRINT '=== Newest order date (snapshot freshness) ===';
SELECT MAX(OrderDate) AS NewestOrder, MIN(OrderDate) AS OldestOrder FROM dbo.tbOrder;

PRINT '=== App-layer (new-system) tables present? 0 = MISSING, needs graft ===';
SELECT
  OBJECT_ID('dbo.PortalUsers')         AS PortalUsers,
  OBJECT_ID('dbo.PortalRefreshTokens') AS PortalRefreshTokens,
  OBJECT_ID('dbo.UserPageAccess')      AS UserPageAccess,
  OBJECT_ID('dbo.AppSettings')         AS AppSettings,
  OBJECT_ID('dbo.LoginOtps')           AS LoginOtps,
  OBJECT_ID('dbo.MarketQuotes')        AS MarketQuotes,
  SCHEMA_ID('app')                     AS app_schema;

PRINT '=== Total dbo table count (schema-size sanity vs Malawi 119/160) ===';
SELECT COUNT(*) AS dbo_tables FROM sys.tables WHERE schema_id = SCHEMA_ID('dbo');
