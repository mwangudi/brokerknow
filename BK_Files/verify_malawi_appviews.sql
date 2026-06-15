SET NOCOUNT ON;
USE BrokerKnow_Malawi0612;
SELECT 'app views' AS k, COUNT(*) AS n
FROM sys.views v JOIN sys.schemas s ON s.schema_id = v.schema_id
WHERE s.name = 'app';
-- sanity: app.Clients + a couple of core endpoints' tables resolve
SELECT TOP 1 'app.Clients ok' AS probe FROM app.Clients;
SELECT 'OrderHoldType' AS tbl, COUNT(*) AS n FROM dbo.OrderHoldType;
GO
