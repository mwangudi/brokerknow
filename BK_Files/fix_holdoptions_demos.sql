/* fix_holdoptions_demos.sql — seed the order-hold lookups in BOTH demo DBs.
   The portal Place Order flow needs a default OrderHoldOptions row, and the
   FK-hardened clean clone (KE) needs a matching OrderHoldType row for the
   FK_tbOrder_OrderHoldType constraint. Values mirror prod BrokerKnow exactly.
   Idempotent. */
SET QUOTED_IDENTIFIER ON;
SET ANSI_NULLS ON;
SET NOCOUNT ON;
GO

/* ============================ RWANDA ============================ */
USE BrokerKnow_RW_Demo;
GO
IF NOT EXISTS (SELECT 1 FROM OrderHoldOptions WHERE OrderHoldOptionID = 2)
    INSERT INTO OrderHoldOptions (OrderHoldOptionID, Description, RequiresDate, DefaultSelection)
    VALUES (2, 'Awaiting Manual Release', 0, 1);
IF NOT EXISTS (SELECT 1 FROM OrderHoldOptions WHERE OrderHoldOptionID = 3)
    INSERT INTO OrderHoldOptions (OrderHoldOptionID, Description, RequiresDate, DefaultSelection)
    VALUES (3, 'Until Given Date', 1, 0);

IF NOT EXISTS (SELECT 1 FROM OrderHoldType)
BEGIN
    SET IDENTITY_INSERT OrderHoldType ON;
    INSERT INTO OrderHoldType (OrderHoldType_DPA_, OrderHoldTypeName, DefaultSelection) VALUES
     (1, 'Released', 0),
     (2, 'Awaiting manual release', 1),
     (3, 'Awaiting payment', 0),
     (4, 'Until given date', 0);
    SET IDENTITY_INSERT OrderHoldType OFF;
END
SELECT 'RW OrderHoldOptions' AS tbl, COUNT(*) AS n FROM OrderHoldOptions
UNION ALL SELECT 'RW OrderHoldType', COUNT(*) FROM OrderHoldType;
GO

/* ============================ KENYA ============================= */
USE BrokerKnow_KE_Demo;
GO
IF NOT EXISTS (SELECT 1 FROM OrderHoldOptions WHERE OrderHoldOptionID = 2)
    INSERT INTO OrderHoldOptions (OrderHoldOptionID, Description, RequiresDate, DefaultSelection)
    VALUES (2, 'Awaiting Manual Release', 0, 1);
IF NOT EXISTS (SELECT 1 FROM OrderHoldOptions WHERE OrderHoldOptionID = 3)
    INSERT INTO OrderHoldOptions (OrderHoldOptionID, Description, RequiresDate, DefaultSelection)
    VALUES (3, 'Until Given Date', 1, 0);

IF NOT EXISTS (SELECT 1 FROM OrderHoldType)
BEGIN
    SET IDENTITY_INSERT OrderHoldType ON;
    INSERT INTO OrderHoldType (OrderHoldType_DPA_, OrderHoldTypeName, DefaultSelection) VALUES
     (1, 'Released', 0),
     (2, 'Awaiting manual release', 1),
     (3, 'Awaiting payment', 0),
     (4, 'Until given date', 0);
    SET IDENTITY_INSERT OrderHoldType OFF;
END
SELECT 'KE OrderHoldOptions' AS tbl, COUNT(*) AS n FROM OrderHoldOptions
UNION ALL SELECT 'KE OrderHoldType', COUNT(*) FROM OrderHoldType;
GO
