/* Phase-3 FK behaviour smoke — TEST. Proves the new FKs (a) accept a normal
   order insert with valid refs and (b) reject an order with a bogus client.
   Everything runs inside transactions that ROLL BACK — no data is changed. */
SET NOCOUNT ON;
SET XACT_ABORT OFF;
USE BrokerKnow_Test;
GO

-- Pick real, valid FK values from existing data.
DECLARE @client int   = (SELECT TOP 1 Client_DPA_   FROM dbo.Client      ORDER BY Client_DPA_);
DECLARE @branch int   = (SELECT TOP 1 Branch_DPA_   FROM dbo.Branch       ORDER BY Branch_DPA_);
DECLARE @otype  int   = (SELECT TOP 1 OrderType_DPA_    FROM dbo.OrderType     ORDER BY OrderType_DPA_);
DECLARE @ostype int   = (SELECT TOP 1 OrderSecType_DPA_ FROM dbo.OrderSecType  ORDER BY OrderSecType_DPA_);
DECLARE @ohold  int   = (SELECT TOP 1 OrderHoldType_DPA_ FROM dbo.OrderHoldType ORDER BY OrderHoldType_DPA_);
DECLARE @sec    int   = (SELECT TOP 1 Security_DPA_  FROM dbo.Security     ORDER BY Security_DPA_);
DECLARE @oid    int   = (SELECT ISNULL(MAX(Order_DPA_),0)+1 FROM dbo.tbOrder);
DECLARE @odid   int   = (SELECT ISNULL(MAX(OrdDetail_DPA_),0)+1 FROM dbo.OrdDetail);

PRINT '=== TEST 1: VALID order insert (expect SUCCESS, then rollback) ===';
BEGIN TRAN;
BEGIN TRY
    INSERT INTO dbo.tbOrder (Order_DPA_, Client_DPA_, Branch_DPA_, OrderDate, OrderRef,
        OrderSecType_DPA_, OrderType_DPA_, OrderHold, OrderCanceled, OrderHoldType_DPA_, OrderCompounded)
      VALUES (@oid, @client, @branch, GETDATE(), 'FKSMOKE',
        @ostype, @otype, 1, 0, @ohold, 0);
    INSERT INTO dbo.OrdDetail (OrdDetail_DPA_, Order_DPA_, Security_DPA_, OrdDetailPrice, OrdDetailQty, OrdDetailCompound)
      VALUES (@odid, @oid, @sec, '10.00', 100, 0);
    PRINT '  OK: valid order+detail inserted (FKs accepted real refs).';
END TRY
BEGIN CATCH
    PRINT '  !! UNEXPECTED FAILURE: ' + ERROR_MESSAGE();
END CATCH
ROLLBACK TRAN;

PRINT '=== TEST 2: INVALID client (expect FK REJECTION, then rollback) ===';
BEGIN TRAN;
BEGIN TRY
    INSERT INTO dbo.tbOrder (Order_DPA_, Client_DPA_, Branch_DPA_, OrderDate, OrderRef,
        OrderSecType_DPA_, OrderType_DPA_, OrderHold, OrderCanceled, OrderHoldType_DPA_, OrderCompounded)
      VALUES (@oid, 99999999, @branch, GETDATE(), 'FKSMOKE-BAD',
        @ostype, @otype, 1, 0, @ohold, 0);
    PRINT '  !! PROBLEM: bogus Client_DPA_=99999999 was ACCEPTED (FK not working).';
END TRY
BEGIN CATCH
    PRINT '  OK: FK correctly REJECTED bogus client -> ' + ERROR_MESSAGE();
END CATCH
IF @@TRANCOUNT > 0 ROLLBACK TRAN;

PRINT '=== TEST 3: INVALID security on detail (expect FK REJECTION) ===';
BEGIN TRAN;
BEGIN TRY
    INSERT INTO dbo.tbOrder (Order_DPA_, Client_DPA_, Branch_DPA_, OrderDate, OrderRef,
        OrderSecType_DPA_, OrderType_DPA_, OrderHold, OrderCanceled, OrderHoldType_DPA_, OrderCompounded)
      VALUES (@oid, @client, @branch, GETDATE(), 'FKSMOKE',
        @ostype, @otype, 1, 0, @ohold, 0);
    INSERT INTO dbo.OrdDetail (OrdDetail_DPA_, Order_DPA_, Security_DPA_, OrdDetailPrice, OrdDetailQty, OrdDetailCompound)
      VALUES (@odid, @oid, 99999999, '10.00', 100, 0);
    PRINT '  !! PROBLEM: bogus Security_DPA_ was ACCEPTED.';
END TRY
BEGIN CATCH
    PRINT '  OK: FK correctly REJECTED bogus security -> ' + ERROR_MESSAGE();
END CATCH
IF @@TRANCOUNT > 0 ROLLBACK TRAN;

PRINT '=== verify NO rows leaked (FKSMOKE refs should be gone) ===';
SELECT COUNT(*) AS leftover_smoke_orders FROM dbo.tbOrder WHERE OrderRef LIKE 'FKSMOKE%';
GO
