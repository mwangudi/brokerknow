SET NOCOUNT ON;
PRINT '===== OrderHoldType (singular) columns in prod =====';
USE BrokerKnow;
SELECT c.column_id, c.name, ty.name AS typ, c.is_nullable, c.is_identity
FROM sys.columns c JOIN sys.types ty ON ty.user_type_id=c.user_type_id
WHERE c.object_id=OBJECT_ID('dbo.OrderHoldType') ORDER BY c.column_id;
GO
PRINT '===== RW demo state =====';
USE BrokerKnow_RW_Demo;
SELECT 'OrderHoldOptions' AS tbl, COUNT(*) AS n FROM OrderHoldOptions
UNION ALL SELECT 'OrderHoldType', COUNT(*) FROM OrderHoldType;
SELECT fk.name AS rw_tborder_fk FROM sys.foreign_keys fk WHERE fk.parent_object_id = OBJECT_ID('dbo.tbOrder');
GO
PRINT '===== KE demo state =====';
USE BrokerKnow_KE_Demo;
SELECT 'OrderHoldOptions' AS tbl, COUNT(*) AS n FROM OrderHoldOptions
UNION ALL SELECT 'OrderHoldType', COUNT(*) FROM OrderHoldType;
SELECT fk.name AS ke_tborder_fk FROM sys.foreign_keys fk WHERE fk.parent_object_id = OBJECT_ID('dbo.tbOrder');
GO
