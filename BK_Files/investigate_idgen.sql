/* READ-ONLY: understand the legacy ID-generation mechanism + concurrency
   config, to harden Payment/Order key generation correctly. Target: test. */
SET NOCOUNT ON;
USE BrokerKnow_Test;
GO

/* ---- does the legacy sequence table exist? ------------------------- */
PRINT '===== tables matching %Initial%Table%ID% or %_ID_% =====';
SELECT name FROM sys.tables
WHERE name LIKE '%Initial%' OR name LIKE '%[_]ID[_]%' OR name LIKE '%TableID%' OR name LIKE '%Sequence%';

/* ---- if _Initial_Table_ID_ exists, show its shape + the Payment/Order rows */
IF OBJECT_ID('dbo._Initial_Table_ID_') IS NOT NULL
BEGIN
    PRINT '===== _Initial_Table_ID_ columns =====';
    SELECT c.name, ty.name AS type FROM sys.columns c
    JOIN sys.types ty ON ty.user_type_id=c.user_type_id
    WHERE c.object_id = OBJECT_ID('dbo._Initial_Table_ID_') ORDER BY c.column_id;

    PRINT '===== _Initial_Table_ID_ contents (all) =====';
    SELECT * FROM dbo._Initial_Table_ID_;
END
ELSE PRINT '  _Initial_Table_ID_ NOT present in this DB.';

/* ---- compare: legacy seq value vs actual MAX for Payment & tbOrder -- */
PRINT '===== actual MAX keys =====';
SELECT 'Payment'  AS tbl, MAX(Payment_DPA_)  AS max_key FROM dbo.Payment
UNION ALL SELECT 'tbOrder', MAX(Order_DPA_)   FROM dbo.tbOrder
UNION ALL SELECT 'OrdDetail', MAX(OrdDetail_DPA_) FROM dbo.OrdDetail
UNION ALL SELECT 'Contract', MAX(Contract_DPA_) FROM dbo.Contract;

/* ---- is there a trigger on Payment that assigns the id? ------------ */
PRINT '===== triggers on Payment / tbOrder =====';
SELECT t.name AS trigger_name, OBJECT_NAME(t.parent_id) AS on_table, t.is_disabled
FROM sys.triggers t
WHERE OBJECT_NAME(t.parent_id) IN ('Payment','tbOrder','OrdDetail','Contract');
GO
