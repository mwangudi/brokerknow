SET NOCOUNT ON;
USE BrokerKnow_KE_Demo;
SELECT
  (SELECT COUNT(*) FROM sys.tables WHERE schema_id = SCHEMA_ID('dbo')) AS dbo_tables,
  (SELECT COUNT(*) FROM sys.foreign_keys)                              AS fks,
  (SELECT COUNT(*) FROM sys.views WHERE schema_id = SCHEMA_ID('app'))  AS app_views,
  (SELECT COUNT(*) FROM dbo.Client)                                    AS clients,
  (SELECT COUNT(*) FROM dbo.Security)                                  AS securities,
  (SELECT COUNT(*) FROM dbo.PortalUsers)                               AS portal_users;
SELECT Id, Email, Role, Status, ClientDpa FROM dbo.PortalUsers ORDER BY Id;
