/* Spot-check: friendly join reads cleanly, and the views agree with the
   base tables (live-row counts match the same filter applied directly). */
SET NOCOUNT ON;
USE BrokerKnow_Test;
GO

PRINT '===== friendly join: top 5 clients by live order count =====';
SELECT TOP 5
       c.ClientId, c.Name, c.CdsNumber,
       COUNT(o.OrderId) AS Orders
FROM app.Clients c
JOIN app.Orders  o ON o.ClientId = c.ClientId
GROUP BY c.ClientId, c.Name, c.CdsNumber
ORDER BY COUNT(o.OrderId) DESC;

PRINT '===== friendly trade row: lot -> security -> client (sample 3) =====';
SELECT TOP 3
       l.ContractNumber, s.Code AS Security, l.Quantity, l.Price, l.GrossAmount,
       cl.Name AS Client
FROM app.Lots l
JOIN app.OrderDetails od ON od.OrderDetailId = l.OrderDetailId
JOIN app.Securities s    ON s.SecurityId    = od.SecurityId
JOIN app.Orders o        ON o.OrderId       = od.OrderId
JOIN app.Clients cl      ON cl.ClientId     = o.ClientId
ORDER BY l.TradeDate DESC;

PRINT '===== integrity: view count == base live-row count (expect 0 diffs) =====';
SELECT 'Clients' AS entity,
       (SELECT COUNT(*) FROM app.Clients) AS view_rows,
       (SELECT COUNT(*) FROM dbo.Client WHERE Deleted = 0 OR Deleted IS NULL) AS base_live,
       (SELECT COUNT(*) FROM app.Clients) - (SELECT COUNT(*) FROM dbo.Client WHERE Deleted = 0 OR Deleted IS NULL) AS diff
UNION ALL SELECT 'Payments',
       (SELECT COUNT(*) FROM app.Payments),
       (SELECT COUNT(*) FROM dbo.Payment WHERE Deleted = 0 OR Deleted IS NULL),
       (SELECT COUNT(*) FROM app.Payments) - (SELECT COUNT(*) FROM dbo.Payment WHERE Deleted = 0 OR Deleted IS NULL)
UNION ALL SELECT 'Securities',
       (SELECT COUNT(*) FROM app.Securities),
       (SELECT COUNT(*) FROM dbo.Security),
       (SELECT COUNT(*) FROM app.Securities) - (SELECT COUNT(*) FROM dbo.Security);
GO
