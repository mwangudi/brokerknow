SET NOCOUNT ON;
USE BrokerKnow_KE_Legacy;

PRINT '=== Core business data counts (Kenya legacy) ===';
SELECT
  (SELECT COUNT(*) FROM dbo.Client)   AS Clients,
  (SELECT COUNT(*) FROM dbo.tbOrder)  AS Orders,
  (SELECT COUNT(*) FROM dbo.Contract) AS Contracts,
  (SELECT COUNT(*) FROM dbo.Lot)      AS Lots,
  (SELECT COUNT(*) FROM dbo.Payment)  AS Payments,
  (SELECT COUNT(*) FROM dbo.Security) AS Securities,
  (SELECT COUNT(*) FROM dbo.Agent)    AS Agents,
  (SELECT COUNT(*) FROM dbo.Broker)   AS Brokers;

PRINT '=== Snapshot freshness ===';
SELECT MAX(OrderDate) AS NewestOrder, MIN(OrderDate) AS OldestOrder FROM dbo.tbOrder;

PRINT '=== Biggest tables (where the 4.4GB lives) ===';
SELECT TOP 15 t.name AS tbl, p.rows,
       CAST(SUM(a.total_pages)*8.0/1024 AS DECIMAL(10,1)) AS MB
FROM sys.tables t
JOIN sys.partitions p ON p.object_id=t.object_id AND p.index_id IN (0,1)
JOIN sys.allocation_units a ON a.container_id=p.partition_id
WHERE t.schema_id=SCHEMA_ID('dbo')
GROUP BY t.name, p.rows
ORDER BY SUM(a.total_pages) DESC;

PRINT '=== dbo table count (vs Malawi clean 119, raw 143-160) ===';
SELECT COUNT(*) AS dbo_tables FROM sys.tables WHERE schema_id=SCHEMA_ID('dbo');

PRINT '=== Schema diff: tables in clean baseline MISSING from KE legacy ===';
-- compare against KE_Demo (the hardened clean schema) to see how far KE legacy diverges
SELECT d.name AS in_clean_not_in_ke_legacy
FROM BrokerKnow_KE_Demo.sys.tables d
WHERE SCHEMA_NAME(d.schema_id)='dbo'
  AND d.name NOT IN (SELECT name FROM BrokerKnow_KE_Legacy.sys.tables WHERE SCHEMA_NAME(schema_id)='dbo')
ORDER BY d.name;
