/* =====================================================================
   Friendly read-only VIEWS layer ? schema [app].
   Target: BrokerKnow_Test (:5261) first, then prod.

   Purpose: make the legacy _DPA_ schema pleasant to read & query for
   reporting / BI / debugging, WITHOUT touching base tables or the legacy
   app. Pure metadata ? zero data change, instantly reversible (DROP VIEW).

   Conventions:
     * <Entity>_DPA_  -> <Entity>Id ;  other _DPA_ FKs -> <Name>Id
     * business columns renamed to clean names; audit/rare columns omitted
     * soft-deleted rows are EXCLUDED (WHERE Deleted = 0/NULL) so these views
       show LIVE data only ? the common reporting need. (Security/Broker have
       no Deleted column, so those show all rows.)
     * NOT schemabound -> base tables remain freely alterable by legacy.
   Idempotent: CREATE OR ALTER. Each view is its own batch (T-SQL rule).
   ===================================================================== */
SET NOCOUNT ON;
USE BrokerKnow;
GO

IF NOT EXISTS (SELECT 1 FROM sys.schemas WHERE name = 'app')
    EXEC('CREATE SCHEMA app');
GO

CREATE OR ALTER VIEW app.Clients AS
SELECT  Client_DPA_        AS ClientId,
        ClientName         AS Name,
        ClientCDSNo        AS CdsNumber,
        ClientIDPass       AS IdNumber,
        ClientEmail        AS Email,
        ClientCellTel      AS Mobile,
        Agent_DPA_         AS AgentId,
        Branch_DPA_        AS BranchId,
        Class_DPA_         AS ClassId,
        Commission_DPA_    AS CommissionId,
        Residency_DPA_     AS ResidencyId,
        ClientOpeningBal   AS OpeningBalance,
        CreditLimit        AS CreditLimit,
        ClientVIP          AS IsVip,
        IsCustodian        AS IsCustodian,
        IsNominee          AS IsNominee,
        IsFrozen           AS IsFrozen,
        ClientRegDate      AS RegisteredOn
FROM dbo.Client
WHERE Deleted = 0 OR Deleted IS NULL;
GO

CREATE OR ALTER VIEW app.Agents AS
SELECT  Agent_DPA_         AS AgentId,
        AgentName          AS Name,
        AgentEmail         AS Email,
        AgentCellTel       AS Mobile,
        Branch_DPA_        AS BranchId,
        Commission_DPA_    AS CommissionId,
        AgentOpeningBal    AS OpeningBalance,
        AgentRegDate       AS RegisteredOn
FROM dbo.Agent
WHERE Deleted = 0 OR Deleted IS NULL;
GO

CREATE OR ALTER VIEW app.Brokers AS
SELECT  Broker_DPA_        AS BrokerId,
        BrokerCode         AS Code,
        BrokerName         AS Name,
        BrokerOpeningBal   AS OpeningBalance,
        BrokerRegDate      AS RegisteredOn
FROM dbo.Broker;     -- no Deleted column
GO

CREATE OR ALTER VIEW app.Securities AS
SELECT  Security_DPA_      AS SecurityId,
        SecurityCode       AS Code,
        SecurityName       AS Name,
        SecurityMktPrice   AS MarketPrice,
        OrderSecType_DPA_  AS SecTypeId,
        Sector_DPA_        AS SectorId,
        CanTrade           AS CanTrade,
        Immobilised        AS IsImmobilised
FROM dbo.Security;   -- no Deleted column
GO

CREATE OR ALTER VIEW app.Orders AS
SELECT  Order_DPA_         AS OrderId,
        Client_DPA_        AS ClientId,
        Branch_DPA_        AS BranchId,
        Agent_DPA_         AS AgentId,
        OrderType_DPA_     AS OrderTypeId,
        OrderSecType_DPA_  AS SecTypeId,
        OrderHoldType_DPA_ AS HoldTypeId,
        OrderDate          AS OrderDate,
        OrderRef           AS Reference,
        OrderHold          AS IsOnHold,
        OrderCanceled      AS IsCancelled,
        OrderCompounded    AS IsCompounded,
        Remarks            AS Remarks,
        TimeCreated        AS CreatedOn
FROM dbo.tbOrder
WHERE Deleted = 0 OR Deleted IS NULL;
GO

CREATE OR ALTER VIEW app.OrderDetails AS
SELECT  OrdDetail_DPA_     AS OrderDetailId,
        Order_DPA_         AS OrderId,
        Security_DPA_      AS SecurityId,
        OrdDetailPrice     AS Price,
        OrdDetailQty       AS Quantity,
        Amount             AS Amount,
        Best               AS IsBest
FROM dbo.OrdDetail
WHERE Deleted = 0 OR Deleted IS NULL;
GO

CREATE OR ALTER VIEW app.Contracts AS
SELECT  Contract_DPA_           AS ContractId,
        Status_DPA_             AS StatusId,
        ContractSettlementDate  AS SettlementDate,
        IsInterBank             AS IsInterBank,
        Voucher_DPA_            AS VoucherId,
        ClientVoucher_DPA_      AS ClientVoucherId
FROM dbo.Contract
WHERE Deleted = 0 OR Deleted IS NULL;
GO

CREATE OR ALTER VIEW app.Lots AS
SELECT  Lot_DPA_                AS LotId,
        Contract_DPA_           AS ContractId,
        OrdDetail_DPA_          AS OrderDetailId,
        Broker_DPA_             AS BrokerId,
        ContractNumber          AS ContractNumber,
        LotPrice                AS Price,
        LotQty                  AS Quantity,
        LotGrossAmount          AS GrossAmount,
        LotTDate                AS TradeDate,
        ContractSettlementDate  AS SettlementDate
FROM dbo.Lot
WHERE Deleted = 0 OR Deleted IS NULL;
GO

CREATE OR ALTER VIEW app.Payments AS
SELECT  Payment_DPA_       AS PaymentId,
        EntityType_DPA_    AS EntityTypeId,
        Entity_DPA_        AS EntityId,
        PayType_DPA_       AS PayTypeId,
        PaymentAmount      AS Amount,
        PaymentReceiptNo   AS ReceiptNo,
        PaymentPDate       AS PaymentDate,
        PaymentReference   AS Reference,
        PaymentNarrative   AS Narrative,
        BankAccount_DPA_   AS BankAccountId,
        Contract_DPA_      AS ContractId,
        Order_DPA_         AS OrderId
FROM dbo.Payment
WHERE Deleted = 0 OR Deleted IS NULL;
GO

/* ---- verify: list app views + a row count from each ---------------- */
PRINT '===== app.* views created =====';
SELECT s.name + '.' + v.name AS view_name
FROM sys.views v JOIN sys.schemas s ON s.schema_id = v.schema_id
WHERE s.name = 'app' ORDER BY v.name;

SELECT 'app.Clients' AS view_name, COUNT(*) AS rows FROM app.Clients
UNION ALL SELECT 'app.Agents', COUNT(*) FROM app.Agents
UNION ALL SELECT 'app.Brokers', COUNT(*) FROM app.Brokers
UNION ALL SELECT 'app.Securities', COUNT(*) FROM app.Securities
UNION ALL SELECT 'app.Orders', COUNT(*) FROM app.Orders
UNION ALL SELECT 'app.OrderDetails', COUNT(*) FROM app.OrderDetails
UNION ALL SELECT 'app.Contracts', COUNT(*) FROM app.Contracts
UNION ALL SELECT 'app.Lots', COUNT(*) FROM app.Lots
UNION ALL SELECT 'app.Payments', COUNT(*) FROM app.Payments
ORDER BY view_name;
GO
