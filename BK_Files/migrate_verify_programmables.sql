SET NOCOUNT ON;
USE BrokerKnow_Clean;
GO
PRINT '===== programmables in BrokerKnow_Clean by type =====';
SELECT type_desc, COUNT(*) AS cnt
FROM sys.objects
WHERE type IN ('V','P','FN','IF','TF') AND SCHEMA_NAME(schema_id) IN ('dbo','app')
GROUP BY type_desc ORDER BY type_desc;

PRINT '===== app.* friendly views =====';
SELECT s.name AS sch, o.name AS view_name
FROM sys.objects o JOIN sys.schemas s ON s.schema_id = o.schema_id
WHERE s.name = 'app' AND o.type = 'V'
ORDER BY o.name;

PRINT '===== sanity: query an app view + a recovered QI-OFF view =====';
SELECT TOP 0 * FROM app.Clients;
SELECT TOP 0 * FROM dbo.CompanyInfoList;
GO
