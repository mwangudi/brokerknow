-- verify_cedar_csd.sql — post-apply checks for the Cedar client CSD import.
-- Run after apply_cedar_csd.sql on the freshly-migrated DB.
SET NOCOUNT ON;

PRINT '== 1) CSD coverage ==';
SELECT COUNT(*) AS TotalClients,
       SUM(CASE WHEN LEN(ISNULL(ClientCDSNo,'')) > 0 THEN 1 ELSE 0 END) AS WithCds,
       SUM(CASE WHEN LEN(ISNULL(ClientCDSNo,'')) > 0 AND LEN(ClientCDSNo) < 23 THEN 1 ELSE 0 END) AS Truncated
FROM dbo.Client WHERE ISNULL(Deleted,0)=0;

PRINT '';
PRINT '== 2) Pending CDS-trade import batches: client codes that now resolve ==';
PRINT '   (re-run reconciliation on these batches in the app, then Commit)';
SELECT t.BatchId,
       COUNT(DISTINCT t.ClientCode) AS DistinctCodes,
       COUNT(DISTINCT CASE WHEN c.Client_DPA_ IS NOT NULL THEN t.ClientCode END) AS Resolved,
       COUNT(DISTINCT CASE WHEN c.Client_DPA_ IS NULL     THEN t.ClientCode END) AS StillUnknown
FROM dbo.CdsImportedTrades t
LEFT JOIN dbo.Client c ON c.ClientCDSNo = t.ClientCode AND ISNULL(c.Deleted,0)=0
WHERE t.MatchStatus NOT IN ('Matched','Ignored')
GROUP BY t.BatchId
ORDER BY t.BatchId;

PRINT '';
PRINT '== 3) Import client codes STILL unknown (clients not in the CSD listing) ==';
SELECT TOP 30 t.ClientCode, COUNT(*) AS Rows
FROM dbo.CdsImportedTrades t
WHERE t.MatchStatus NOT IN ('Matched','Ignored')
  AND NOT EXISTS (SELECT 1 FROM dbo.Client c WHERE c.ClientCDSNo = t.ClientCode AND ISNULL(c.Deleted,0)=0)
GROUP BY t.ClientCode
ORDER BY COUNT(*) DESC;

PRINT '';
PRINT '== 4) CSD numbers shared by >1 client (reconcile is dup-tolerant; first wins) ==';
SELECT ClientCDSNo, COUNT(*) AS Clients
FROM dbo.Client
WHERE ISNULL(Deleted,0)=0 AND LEN(ISNULL(ClientCDSNo,'')) > 0
GROUP BY ClientCDSNo
HAVING COUNT(*) > 1
ORDER BY COUNT(*) DESC;
