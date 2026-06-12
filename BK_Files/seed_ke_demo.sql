/* =====================================================================
   BrokerKnow_KE_Demo — synthetic seed data
   Axis Kenya placeholder white-label demo (isolated DB, KES / NSE).
   NO real Cedar/Malawi data. Safe to re-run (idempotent deletes first).
   Demo login:  demo@axis-kenya.demo  /  Demo@2026   (BCrypt wf=11)
   Admin login: admin@axis-kenya.demo /  Admin@2026

   NOTE: BrokerKnow_KE_Demo is cloned from BrokerKnow_Clean, which carries the
   22 hardened FKs (FK_Client_Branch/Class/Commission/Residency/EntityType,
   FK_Payment_EntityType/PayType, ...). So unlike the older Rwanda seed we must
   seed the PARENT lookup rows first or the Client/Payment inserts are rejected.
   ===================================================================== */
SET QUOTED_IDENTIFIER ON;   -- required: PortalUsers has filtered indexes
SET ANSI_NULLS ON;
SET NOCOUNT ON;
USE BrokerKnow_KE_Demo;
GO

/* ---- Idempotency: clear prior demo rows (children -> parents) ------ */
DELETE FROM ClientBalances WHERE client_DPA_ = 1001;
DELETE FROM Payment        WHERE Payment_DPA_ IN (5001, 5002, 5003);
DELETE FROM PortalUsers    WHERE Email IN ('demo@axis-kenya.demo', 'admin@axis-kenya.demo');
DELETE FROM MarketQuotes   WHERE SecurityDpa BETWEEN 1 AND 6;
DELETE FROM Client         WHERE Client_DPA_ = 1001;
DELETE FROM Security       WHERE Security_DPA_ BETWEEN 1 AND 6;
DELETE FROM OrderSecType   WHERE OrderSecType_DPA_ IN (1, 2);
DELETE FROM OrderType      WHERE OrderType_DPA_ IN (1, 2);
DELETE FROM PayType        WHERE PayType_DPA_ IN (1, 2);
DELETE FROM OrderHoldType  WHERE OrderHoldType_DPA_ IN (1, 2, 3, 4);
DELETE FROM OrderHoldOptions WHERE OrderHoldOptionID IN (2, 3);
DELETE FROM EntityType     WHERE EntityType_DPA_ = 1;
DELETE FROM Residency      WHERE Residency_DPA_ = 1;
DELETE FROM Commission     WHERE Commission_DPA_ = 1;
DELETE FROM Class          WHERE Class_DPA_ = 1;
DELETE FROM Branch         WHERE Branch_DPA_ = 1;
GO

/* ===== PARENT LOOKUPS (required by the hardened FKs) =============== */

/* ---- Branch -------------------------------------------------------- */
INSERT INTO Branch (Branch_DPA_, BranchDescription, DefaultSelection) VALUES
 (1, 'Nairobi Office', 1);
GO

/* ---- Class (client classification) -------------------------------- */
INSERT INTO Class (Class_DPA_, ClassDescription, DefaultSelection, IsCda, AgentStatus) VALUES
 (1, 'Individual', 1, 0, 0);
GO

/* ---- Commission (identity PK -> force id 1) ----------------------- */
SET IDENTITY_INSERT Commission ON;
INSERT INTO Commission
 (Commission_DPA_, CommissionRate, DefaultSelection, BondCommission,
  SecurityBoundary, BondBoundary, SecondBondBoundary,
  UpperBondCommission, UpperSecurityCommission,
  MedianBondCommission, MedianSecurityCommission,
  MinimumBondCommission, MinimumSecurityCommission,
  CMARegulated, Immobilised, SystemMaintained, Vatable)
VALUES
 (1, 2.1, 1, 0.0, 0, 0, 0, 0.0, 2.1, 0, 0, 0, 0, 0, 0, 0, 0);
SET IDENTITY_INSERT Commission OFF;
GO

/* ---- Residency ----------------------------------------------------- */
INSERT INTO Residency (Residency_DPA_, ResidencyDescription, DefaultSelection) VALUES
 (1, 'Resident', 1);
GO

/* ---- EntityType (1 = Client) -------------------------------------- */
INSERT INTO EntityType
 (AccountType_DPA_, EntityType_DPA_, EntityTypeName, DefaultSelection,
  SystemMaintained, Nominal, EntityTypeCode)
VALUES
 (1, 1, 'Client', 1, 1, 0, 'CLI');
GO

/* ===== TRANSACTIONAL LOOKUPS ====================================== */

/* ---- Payment types ------------------------------------------------- */
INSERT INTO PayType (PayType_DPA_, PayTypeDescription, PayTypeIn) VALUES
 (1, 'Receipt', 1),   -- incoming  -> statement Credit
 (2, 'Payment', 0);   -- outgoing  -> statement Debit
GO

/* ---- Order types (Buy/Sell) --------------------------------------- */
INSERT INTO OrderType (OrderType_DPA_, OrderTypeDescription, OrderTypeSale, DefaultSelection, RequireCertificate) VALUES
 (1, 'Purchase', 0, 1, 0),
 (2, 'Sale',     1, 0, 0);
GO

/* ---- Security types ----------------------------------------------- */
INSERT INTO OrderSecType (OrderSecType_DPA_, OrderSecTypeDescription, OrderSecTypeDisplayName, DefaultSelection) VALUES
 (1, 'Equity', 'Equity', 1),
 (2, 'Bond',   'Bond',   0);
GO

/* ---- Order-hold lookups (REQUIRED for placing orders) -------------
   The portal Place Order flow picks a default OrderHoldOptions row, and the
   FK-hardened clone enforces FK_tbOrder_OrderHoldType, so OrderHoldType must
   carry the matching ids. Values mirror prod BrokerKnow exactly. */
INSERT INTO OrderHoldOptions (OrderHoldOptionID, Description, RequiresDate, DefaultSelection) VALUES
 (2, 'Awaiting Manual Release', 0, 1),
 (3, 'Until Given Date',        1, 0);
GO
SET IDENTITY_INSERT OrderHoldType ON;
INSERT INTO OrderHoldType (OrderHoldType_DPA_, OrderHoldTypeName, DefaultSelection) VALUES
 (1, 'Released',                0),
 (2, 'Awaiting manual release', 1),
 (3, 'Awaiting payment',        0),
 (4, 'Until given date',        0);
SET IDENTITY_INSERT OrderHoldType OFF;
GO

/* ---- Securities: Nairobi Securities Exchange listings (KES) ------- */
INSERT INTO Security (Security_DPA_, SecurityCode, SecurityName, SecurityMktPrice, OrderSecType_DPA_, Immobilised, CanTrade) VALUES
 (1, 'SCOM', 'Safaricom PLC',                 15.00, 1, 0, 1),
 (2, 'EQTY', 'Equity Group Holdings',         45.00, 1, 0, 1),
 (3, 'KCB',  'KCB Group',                     38.00, 1, 0, 1),
 (4, 'EABL', 'East African Breweries',       180.00, 1, 0, 1),
 (5, 'COOP', 'Co-operative Bank of Kenya',    13.00, 1, 0, 1),
 (6, 'ABSA', 'Absa Bank Kenya',               14.50, 1, 0, 1);
GO

/* ---- Market quotes: latest day (NSE) ------------------------------ */
DECLARE @d date = CAST(GETDATE() AS date);
INSERT INTO MarketQuotes (SecurityDpa, QuoteDate, [Open], High, Low, [Close], PreviousClose, Volume, Exchange) VALUES
 (1, @d,  14.80,  15.20,  14.70,  15.00,  14.85, 1850000, 'NSE'),
 (2, @d,  44.50,  45.50,  44.25,  45.00,  44.60,  420000, 'NSE'),
 (3, @d,  37.50,  38.40,  37.40,  38.00,  37.60,  310000, 'NSE'),
 (4, @d, 178.00, 182.00, 177.50, 180.00, 178.50,   95000, 'NSE'),
 (5, @d,  12.90,  13.20,  12.85,  13.00,  12.95,  540000, 'NSE'),
 (6, @d,  14.30,  14.70,  14.20,  14.50,  14.35,  280000, 'NSE');
GO

/* ---- Demo client (sample investor of Axis Kenya) ------------------ */
INSERT INTO Client
 (Client_DPA_, ClientName, Branch_DPA_, Class_DPA_, Commission_DPA_, Residency_DPA_,
  EntityType_DPA_, ClientVIP, OnlineRegistration, ClientOpeningBal, CreditLimit,
  ClientCDSNo, ClientEmail, ClientCellTel, ClientAddr, ClientRegDate, TimeCreated, Deleted)
VALUES
 (1001, 'Wanjiku Kamau', 1, 1, 1, 1,
  1, 0, 0, 500000, 1000000,
  'KE-CDS-100123', 'wkamau@example.ke', '+254 712 345 678',
  'Kimathi Street, Nairobi', '2025-09-01', SYSUTCDATETIME(), 0);
GO

/* ---- Portal login (Role=Client, Approved, Active) ----------------- */
INSERT INTO PortalUsers
 (Email, PasswordHash, FirstName, LastName, Role, Status, Active, CreatedAt,
  ApprovedAt, ClientDpa, MustChangePassword, PasswordChangedAt, Phone, CdsNumber)
VALUES
 ('demo@axis-kenya.demo',
  '$2a$11$vldeisofY8uqVfUkn7vjDuXqWQpYiU/Hn.TqjpY897cfLWBAP3W5O',  -- Demo@2026
  'Wanjiku', 'Kamau', 'Client', 'Approved', 1, SYSUTCDATETIME(),
  SYSUTCDATETIME(), 1001, 0, SYSUTCDATETIME(), '+254 712 345 678', 'KE-CDS-100123');
GO

/* ---- Back-office admin login (Role=Administrators -> full access) -- */
INSERT INTO PortalUsers
 (Email, PasswordHash, FirstName, LastName, Role, Status, Active, CreatedAt,
  ApprovedAt, ClientDpa, MustChangePassword, PasswordChangedAt, Phone)
VALUES
 ('admin@axis-kenya.demo',
  '$2a$11$ZvOoeM2CRMdhffyo6.A35uiZofi4oeXfk34tCU1vUab3X.YkzfULe',  -- Admin@2026
  'David', 'Otieno', 'Administrators', 'Approved', 1, SYSUTCDATETIME(),
  SYSUTCDATETIME(), NULL, 0, SYSUTCDATETIME(), '+254 733 654 321');
GO

/* ---- Cash movements (statement + balance) ------------------------- */
/* EntityType 1 = Client; Entity_DPA_ = client; BankAccount_DPA_ = 1     */
INSERT INTO Payment
 (Payment_DPA_, EntityType_DPA_, Entity_DPA_, PayType_DPA_, BankAccount_DPA_,
  PaymentAmount, PaymentPDate, PaymentReference, PaymentNarrative, Deleted)
VALUES
 (5001, 1, 1001, 1, 1, 1000000, '2026-01-15', 'DEP-0001', 'Initial cash deposit', 0),
 (5002, 1, 1001, 1, 1,   75000, '2026-02-10', 'DEP-0002', 'Top-up deposit',        0),
 (5003, 1, 1001, 2, 1,  100000, '2026-03-05', 'WTH-0001', 'Cash withdrawal',       0);
GO

/* ---- Materialized balance -----------------------------------------
   CurrentBal = Opening 500,000 + receipts 1,075,000 - payments 100,000
              = 1,475,000                                                 */
INSERT INTO ClientBalances (client_DPA_, CurrentBal) VALUES (1001, 1475000);
GO

/* ---- Verify -------------------------------------------------------- */
SELECT 'Branch'        AS tbl, COUNT(*) AS n FROM Branch
UNION ALL SELECT 'Class',         COUNT(*) FROM Class
UNION ALL SELECT 'Commission',    COUNT(*) FROM Commission
UNION ALL SELECT 'Residency',     COUNT(*) FROM Residency
UNION ALL SELECT 'EntityType',    COUNT(*) FROM EntityType
UNION ALL SELECT 'PayType',       COUNT(*) FROM PayType
UNION ALL SELECT 'OrderType',     COUNT(*) FROM OrderType
UNION ALL SELECT 'OrderSecType',  COUNT(*) FROM OrderSecType
UNION ALL SELECT 'OrderHoldOptions', COUNT(*) FROM OrderHoldOptions
UNION ALL SELECT 'OrderHoldType', COUNT(*) FROM OrderHoldType
UNION ALL SELECT 'Security',      COUNT(*) FROM Security
UNION ALL SELECT 'MarketQuotes',  COUNT(*) FROM MarketQuotes
UNION ALL SELECT 'Client',        COUNT(*) FROM Client
UNION ALL SELECT 'PortalUsers',   COUNT(*) FROM PortalUsers
UNION ALL SELECT 'Payment',       COUNT(*) FROM Payment
UNION ALL SELECT 'ClientBalances',COUNT(*) FROM ClientBalances;
GO
