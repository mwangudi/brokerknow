/* READ-ONLY: exact columns for the core tables we'll wrap in friendly views. */
SET NOCOUNT ON;
USE BrokerKnow_Test;
GO
SELECT t.name AS tbl, c.column_id, c.name AS col, ty.name AS typ
FROM sys.tables t
JOIN sys.columns c ON c.object_id = t.object_id
JOIN sys.types ty ON ty.user_type_id = c.user_type_id
WHERE t.name IN ('Client','tbOrder','OrdDetail','Lot','Contract','Payment','Security','Agent','Broker')
ORDER BY t.name, c.column_id;
GO
