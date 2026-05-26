-- ============================================================================
-- BrokerKnow CDS Trade Imports — demo seed
-- ----------------------------------------------------------------------------
-- Pairs with: BK_Files/Samples/MSE_Transaction_Statement_2026-05-12.csv
--
-- After running this script, upload the CSV via Operations → CDS Trade Imports.
-- The reconciliation pass will classify the rows below into all six unmatch
-- reasons plus several commit-eligible candidates so you can exercise both
-- the "Commit selected" and "Commit all eligible" flows.
--
-- Trade date in the file: 2026-05-12.  Participant: CEDAMWMW.
-- This script only INSERTs / ALTERs when the target rows are missing — it
-- is safe to run multiple times.
-- ============================================================================

SET NOCOUNT ON;

DECLARE @TradeDate datetime = '2026-05-12';

-- ── 0. Make BrokerCode + ClientCDSNo wide enough for SWIFT/BIC participant
--      codes and 23-char demo CDS numbers (legacy widths are too small). ────
IF EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('dbo.Broker')
      AND name = 'BrokerCode'
      AND max_length < 40   -- nvarchar(20) = 40 bytes
)
BEGIN
    ALTER TABLE dbo.Broker ALTER COLUMN BrokerCode nvarchar(20) NOT NULL;
    PRINT 'Broker.BrokerCode widened to nvarchar(20).';
END

IF EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('dbo.Client')
      AND name = 'ClientCDSNo'
      AND max_length < 80   -- nvarchar(40) = 80 bytes
)
BEGIN
    ALTER TABLE dbo.Client ALTER COLUMN ClientCDSNo nvarchar(40) NULL;
    PRINT 'Client.ClientCDSNo widened to nvarchar(40).';
END

-- ── 1. Broker (the file's "Participant: CEDAMWMW") ─────────────────────────
IF NOT EXISTS (SELECT 1 FROM dbo.Broker WHERE BrokerCode = 'CEDAMWMW')
BEGIN
    DECLARE @nextBrokerDpa int = ISNULL((SELECT MAX(Broker_DPA_) FROM dbo.Broker), 0) + 1;
    INSERT INTO dbo.Broker (Broker_DPA_, Broker_EIT_, BrokerName, BrokerCode,
                            EntityType_DPA_, BrokerOpeningBal)
    VALUES (@nextBrokerDpa, NEWID(), 'CDS Demo Broker (CEDAMWMW)', 'CEDAMWMW',
            3, 0);
    PRINT 'Inserted broker CEDAMWMW.';
END

DECLARE @BrokerDpa int = (SELECT Broker_DPA_ FROM dbo.Broker WHERE BrokerCode = 'CEDAMWMW');

-- ── 2. Resolve OrderType (Purchase/Sale) and a default OrderSecType / Branch /
--      Class / Commission / Residency / OrderHoldType — every NOT NULL FK on
--      tbOrder / Client points at one of these lookups. We just grab whatever
--      the first existing row is so the demo works on any restored DB. ──────
DECLARE @PurchaseTypeDpa int = (SELECT TOP 1 OrderType_DPA_ FROM dbo.OrderType WHERE OrderTypeSale = 0 ORDER BY OrderType_DPA_);
DECLARE @SaleTypeDpa     int = (SELECT TOP 1 OrderType_DPA_ FROM dbo.OrderType WHERE OrderTypeSale = 1 ORDER BY OrderType_DPA_);
DECLARE @SecTypeDpa      int = (SELECT TOP 1 OrderSecType_DPA_ FROM dbo.OrderSecType ORDER BY OrderSecType_DPA_);
DECLARE @BranchDpa       int = (SELECT TOP 1 Branch_DPA_ FROM dbo.Branch ORDER BY Branch_DPA_);
DECLARE @ClassDpa        int = (SELECT TOP 1 Class_DPA_ FROM dbo.Class ORDER BY Class_DPA_);
DECLARE @CommDpa         int = (SELECT TOP 1 Commission_DPA_ FROM dbo.Commission ORDER BY Commission_DPA_);
DECLARE @ResDpa          int = (SELECT TOP 1 Residency_DPA_ FROM dbo.Residency ORDER BY Residency_DPA_);
DECLARE @HoldTypeDpa     int = ISNULL((SELECT TOP 1 OrderHoldType_DPA_ FROM dbo.OrderHoldType ORDER BY OrderHoldType_DPA_), 0);

IF @PurchaseTypeDpa IS NULL OR @SaleTypeDpa IS NULL OR @SecTypeDpa IS NULL
   OR @BranchDpa IS NULL OR @ClassDpa IS NULL OR @CommDpa IS NULL OR @ResDpa IS NULL
BEGIN
    RAISERROR('One or more lookup tables (OrderType / OrderSecType / Branch / Class / Commission / Residency) are empty. Cannot seed demo data.', 16, 1);
    RETURN;
END

-- ── 3. Demo clients — five of the CDS numbers that appear in the file.
--      We deliberately leave CEDAXXX0000000000005648 OUT so it falls through
--      to the "UnknownClient" reason. ────────────────────────────────────────
DECLARE @ClientsToSeed TABLE (Cds nvarchar(40), Name nvarchar(100));
INSERT INTO @ClientsToSeed VALUES
    ('CEDAXXX0000000000000281', 'Demo Client 0281 (TNM seller)'),
    ('CEDAXXX0000000000007243', 'Demo Client 7243 (TNM/FDHB buyer)'),
    ('CEDAXXX0000000000004342', 'Demo Client 4342 (FDHB - held)'),
    ('CEDAXXX0000000000007456', 'Demo Client 7456 (FDHB - over-allocated)'),
    ('CEDAXXX0000000000007731', 'Demo Client 7731 (SUNBIRD buyer)');

DECLARE @cds nvarchar(40), @name nvarchar(100), @nextClientDpa int;
DECLARE c CURSOR LOCAL FAST_FORWARD FOR SELECT Cds, Name FROM @ClientsToSeed;
OPEN c;
FETCH NEXT FROM c INTO @cds, @name;
WHILE @@FETCH_STATUS = 0
BEGIN
    IF NOT EXISTS (SELECT 1 FROM dbo.Client WHERE ClientCDSNo = @cds)
    BEGIN
        SET @nextClientDpa = ISNULL((SELECT MAX(Client_DPA_) FROM dbo.Client), 0) + 1;
        INSERT INTO dbo.Client (
            Client_DPA_, Client_EIT_, ClientName, ClientCDSNo,
            Branch_DPA_, Class_DPA_, Commission_DPA_, Residency_DPA_,
            EntityType_DPA_, ClientVIP, ClientOpeningBal, ClientRegDate,
            OnlineRegistration, CreditLimit, IsCustodian, TimeCreated)
        VALUES (
            @nextClientDpa, NEWID(), @name, @cds,
            @BranchDpa, @ClassDpa, @CommDpa, @ResDpa,
            1, 0, 0, GETDATE(),
            0, 0, 0, GETDATE());
        PRINT 'Inserted client ' + @cds + '.';
    END
    FETCH NEXT FROM c INTO @cds, @name;
END
CLOSE c; DEALLOCATE c;

-- ── 4. Open released orders that match six specific CSV rows ───────────────
--      Each (CDS, Symbol, Side) must already have a corresponding Security row
--      in dbo.Security with the exact SecurityCode (TNM / FDHB / SUNBIRD).
--      The classification each row will receive is in the comment.
DECLARE @SeedOrders TABLE (
    Cds      nvarchar(40),
    Symbol   nvarchar(50),
    Side     char(1),       -- 'B' or 'S'
    Qty      int,           -- order quantity (set < CSV qty for OverAllocated)
    Price    decimal(19,4),
    Hold     bit,
    Note     nvarchar(200)
);
INSERT INTO @SeedOrders VALUES
    -- 1. eligible BUY: order qty exactly matches CSV.
    ('CEDAXXX0000000000007243', 'TNM',     'B', 33001, 29.94,  0, 'Eligible: TNM Buy 33001 → Ready · Order #'),
    -- 2. eligible SELL: order qty exactly matches CSV.
    ('CEDAXXX0000000000000281', 'TNM',     'S', 33001, 29.94,  0, 'Eligible: TNM Sell 33001 → Ready · Order #'),
    -- 3. eligible BUY: SUNBIRD 190 (matches one of the SUNBIRD buy rows).
    ('CEDAXXX0000000000007731', 'SUNBIRD', 'B', 190,   2585.26, 0, 'Eligible: SUNBIRD Buy 190 → Ready · Order #'),
    -- 4. OrderHeld: matching qty but order still on hold.
    ('CEDAXXX0000000000004342', 'FDHB',    'B', 693,   568.92,  1, 'OrderHeld: order is on hold; release it to commit.'),
    -- 5. OverAllocated: order only allows 1000 but CSV trade is 2171.
    ('CEDAXXX0000000000007456', 'FDHB',    'B', 1000,  568.92,  0, 'OverAllocated: trade qty 2171 > remaining 1000.');
    -- (no row inserted for CEDAXXX0000000000004037 → NoPendingOrder)
    -- (no client inserted for CEDAXXX0000000000005648 → UnknownClient)

DECLARE @SymCode nvarchar(50), @Side char(1), @Qty int, @Price decimal(19,4), @Hold bit, @Note nvarchar(200);
DECLARE @ClientDpa int, @SecurityDpa int, @OrderTypeDpa int, @nextOrderDpa int, @nextDetailDpa int;

DECLARE o CURSOR LOCAL FAST_FORWARD FOR
    SELECT Cds, Symbol, Side, Qty, Price, Hold, Note FROM @SeedOrders;
OPEN o;
FETCH NEXT FROM o INTO @cds, @SymCode, @Side, @Qty, @Price, @Hold, @Note;
WHILE @@FETCH_STATUS = 0
BEGIN
    SET @ClientDpa = (SELECT TOP 1 Client_DPA_ FROM dbo.Client WHERE ClientCDSNo = @cds);
    SET @SecurityDpa = (SELECT TOP 1 Security_DPA_ FROM dbo.Security WHERE SecurityCode = @SymCode);
    SET @OrderTypeDpa = CASE WHEN @Side = 'B' THEN @PurchaseTypeDpa ELSE @SaleTypeDpa END;

    IF @ClientDpa IS NULL OR @SecurityDpa IS NULL
    BEGIN
        PRINT 'Skipped (missing Client or Security): ' + @cds + ' / ' + @SymCode;
    END
    ELSE IF EXISTS (
        -- Skip if a matching open order detail already exists for this client/symbol/side.
        SELECT 1
        FROM dbo.OrdDetail od
        JOIN dbo.tbOrder o ON o.Order_DPA_ = od.Order_DPA_
        WHERE od.Security_DPA_ = @SecurityDpa
          AND o.Client_DPA_ = @ClientDpa
          AND o.OrderType_DPA_ = @OrderTypeDpa
          AND ISNULL(o.OrderCanceled, 0) = 0
          AND ISNULL(od.Deleted, 0) = 0
    )
    BEGIN
        PRINT 'Order already present for ' + @cds + ' / ' + @SymCode + ' / ' + @Side + ' — skipping.';
    END
    ELSE
    BEGIN
        SET @nextOrderDpa  = ISNULL((SELECT MAX(Order_DPA_) FROM dbo.tbOrder), 0) + 1;
        SET @nextDetailDpa = ISNULL((SELECT MAX(OrdDetail_DPA_) FROM dbo.OrdDetail), 0) + 1;

        INSERT INTO dbo.tbOrder (
            Order_DPA_, Order_EIT_, OrderDate, OrderHold, OrderRef,
            Branch_DPA_, Client_DPA_, OrderType_DPA_, OrderSecType_DPA_,
            OrderCanceled, OrderHoldType_DPA_, OrderCompounded,
            IsCustodian, InterBank, ToPay, AgentReturnable,
            CreatedBy, TimeCreated, ChangedBy, TimeChanged, Deleted)
        VALUES (
            @nextOrderDpa, NEWID(), @TradeDate, @Hold, 'CDS-DEMO-' + CONVERT(varchar(20), @nextOrderDpa),
            @BranchDpa, @ClientDpa, @OrderTypeDpa, @SecTypeDpa,
            0, @HoldTypeDpa, 0,
            0, 0, 0, 0,
            1, GETDATE(), 1, GETDATE(), 0);

        INSERT INTO dbo.OrdDetail (
            OrdDetail_DPA_, OrdDetail_EIT_, Order_DPA_, Security_DPA_,
            OrdDetailPrice, OrdDetailQty, OrdDetailCompound, Best, Deleted)
        VALUES (
            @nextDetailDpa, NEWID(), @nextOrderDpa, @SecurityDpa,
            CONVERT(varchar(20), @Price), @Qty, 0, 0, 0);

        PRINT 'Created Order #' + CONVERT(varchar(20), @nextOrderDpa)
              + ' / OrdDetail #' + CONVERT(varchar(20), @nextDetailDpa)
              + ' — ' + @cds + ' / ' + @SymCode + ' / ' + @Side
              + ' qty=' + CONVERT(varchar(20), @Qty)
              + ' hold=' + CONVERT(varchar(1), @Hold)
              + ' (' + @Note + ')';
    END

    FETCH NEXT FROM o INTO @cds, @SymCode, @Side, @Qty, @Price, @Hold, @Note;
END
CLOSE o; DEALLOCATE o;

-- ── 5. Sanity report ───────────────────────────────────────────────────────
PRINT '';
PRINT 'Demo seed complete. Open: dbo.Client / dbo.tbOrder / dbo.OrdDetail to verify.';
PRINT 'Now upload BK_Files/Samples/MSE_Transaction_Statement_2026-05-12.csv';
PRINT 'Expected outcomes for the seeded scenarios:';
PRINT '  CSV row → CEDAXXX...7243 / TNM / Buy 33001    : Ready · Order # (commit-eligible)';
PRINT '  CSV row → CEDAXXX...0281 / TNM / Sell 33001   : Ready · Order # (commit-eligible)';
PRINT '  CSV row → CEDAXXX...7731 / SUNBIRD / Buy 190  : Ready · Order # (commit-eligible)';
PRINT '  CSV row → CEDAXXX...4342 / FDHB / Buy 693     : OrderHeld';
PRINT '  CSV row → CEDAXXX...7456 / FDHB / Buy 2171    : OverAllocated';
PRINT '  CSV row → CEDAXXX...4037 / FDHB / Buy 90      : NoPendingOrder';
PRINT '  CSV row → CEDAXXX...5648 / SUNBIRD / Sell ... : UnknownClient';
PRINT '  All other rows                               : UnknownClient (no client seeded)';
