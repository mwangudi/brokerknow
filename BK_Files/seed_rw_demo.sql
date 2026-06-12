/* =====================================================================
   BrokerKnow_RW_Demo — synthetic seed data
   African Alliance Rwanda Limited white-label demo (isolated DB).
   NO real Cedar/Malawi data. Safe to re-run (idempotent deletes first).
   Demo login:  demo@aar.bsp.rw  /  Demo@2026   (BCrypt wf=11)
   ===================================================================== */
SET QUOTED_IDENTIFIER ON;   -- required: PortalUsers has filtered/computed indexes
SET ANSI_NULLS ON;
SET NOCOUNT ON;
USE BrokerKnow_RW_Demo;
GO

/* ---- Idempotency: clear prior demo rows (fixed demo keys) ---------- */
DELETE FROM ClientBalances WHERE client_DPA_ = 1001;
DELETE FROM Payment        WHERE Payment_DPA_ IN (5001, 5002, 5003);
DELETE FROM PortalUsers    WHERE Email IN ('demo@aar.bsp.rw', 'admin@aar.bsp.rw');
DELETE FROM MarketQuotes   WHERE SecurityDpa BETWEEN 1 AND 6;
DELETE FROM Client         WHERE Client_DPA_ = 1001;
DELETE FROM Security       WHERE Security_DPA_ BETWEEN 1 AND 6;
DELETE FROM OrderSecType   WHERE OrderSecType_DPA_ IN (1, 2);
DELETE FROM OrderType      WHERE OrderType_DPA_ IN (1, 2);
DELETE FROM PayType        WHERE PayType_DPA_ IN (1, 2);
DELETE FROM OrderHoldType  WHERE OrderHoldType_DPA_ IN (1, 2, 3, 4);
DELETE FROM OrderHoldOptions WHERE OrderHoldOptionID IN (2, 3);
GO

/* ---- Lookups: payment types --------------------------------------- */
INSERT INTO PayType (PayType_DPA_, PayTypeDescription, PayTypeIn) VALUES
 (1, 'Receipt', 1),   -- incoming  -> statement Credit
 (2, 'Payment', 0);   -- outgoing  -> statement Debit
GO

/* ---- Lookups: order types (Buy/Sell) ------------------------------ */
INSERT INTO OrderType (OrderType_DPA_, OrderTypeDescription, OrderTypeSale, DefaultSelection, RequireCertificate) VALUES
 (1, 'Purchase', 0, 1, 0),
 (2, 'Sale',     1, 0, 0);
GO

/* ---- Lookups: security types -------------------------------------- */
INSERT INTO OrderSecType (OrderSecType_DPA_, OrderSecTypeDescription, OrderSecTypeDisplayName, DefaultSelection) VALUES
 (1, 'Equity', 'Equity', 1),
 (2, 'Bond',   'Bond',   0);
GO

/* ---- Order-hold lookups (REQUIRED for placing orders) -------------
   The portal Place Order flow picks a default OrderHoldOptions row; the
   OrderHoldType ids are the FK target of tbOrder.OrderHoldType_DPA_ on the
   FK-hardened clone. Values mirror prod BrokerKnow exactly. */
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

/* ---- Securities: Rwanda Stock Exchange listings ------------------- */
INSERT INTO Security (Security_DPA_, SecurityCode, SecurityName, SecurityMktPrice, OrderSecType_DPA_, Immobilised, CanTrade) VALUES
 (1, 'BK',   'Bank of Kigali Holdings',        290.00, 1, 0, 1),
 (2, 'BLR',  'Bralirwa',                       158.00, 1, 0, 1),
 (3, 'CTL',  'Crystal Telecom',                 80.00, 1, 0, 1),
 (4, 'IMR',  'I&M Bank (Rwanda)',               38.00, 1, 0, 1),
 (5, 'KCB',  'KCB Group',                      378.00, 1, 0, 1),
 (6, 'EQTY', 'Equity Group Holdings',          460.00, 1, 0, 1);
GO

/* ---- Market quotes: latest day (RSE) ------------------------------ */
DECLARE @d date = CAST(GETDATE() AS date);
INSERT INTO MarketQuotes (SecurityDpa, QuoteDate, [Open], High, Low, [Close], PreviousClose, Volume, Exchange) VALUES
 (1, @d, 288.00, 292.00, 287.00, 290.00, 288.00, 12400, 'RSE'),
 (2, @d, 159.00, 161.00, 156.00, 158.00, 160.00,  8300, 'RSE'),
 (3, @d,  79.00,  81.00,  78.50,  80.00,  79.00, 15200, 'RSE'),
 (4, @d,  38.50,  38.50,  37.80,  38.00,  38.50, 22100, 'RSE'),
 (5, @d, 375.00, 380.00, 374.00, 378.00, 375.00,  5400, 'RSE'),
 (6, @d, 455.00, 462.00, 454.00, 460.00, 455.00,  6700, 'RSE');
GO

/* ---- Demo client (sample investor of African Alliance Rwanda) ----- */
INSERT INTO Client
 (Client_DPA_, ClientName, Branch_DPA_, Class_DPA_, Commission_DPA_, Residency_DPA_,
  ClientVIP, ClientOpeningBal, CreditLimit, ClientCDSNo, ClientEmail, ClientCellTel,
  ClientAddr, ClientRegDate, Deleted)
VALUES
 (1001, 'Jean-Bosco Habimana', 1, 1, 1, 1,
  0, 5000000, 10000000, 'RW-CDS-100123', 'jbhabimana@example.rw', '+250 788 123 456',
  'KN 41st Ave, Kiyovu, Kigali', '2025-09-01', 0);
GO

/* ---- Portal login (Role=Client, Approved, Active) ----------------- */
INSERT INTO PortalUsers
 (Email, PasswordHash, FirstName, LastName, Role, Status, Active, CreatedAt,
  ApprovedAt, ClientDpa, MustChangePassword, PasswordChangedAt, Phone, CdsNumber)
VALUES
 ('demo@aar.bsp.rw',
  '$2a$11$vldeisofY8uqVfUkn7vjDuXqWQpYiU/Hn.TqjpY897cfLWBAP3W5O',  -- Demo@2026
  'Jean-Bosco', 'Habimana', 'Client', 'Approved', 1, SYSUTCDATETIME(),
  SYSUTCDATETIME(), 1001, 0, SYSUTCDATETIME(), '+250 788 123 456', 'RW-CDS-100123');
GO

/* ---- Back-office admin login (Role=Administrators -> full access) -- */
INSERT INTO PortalUsers
 (Email, PasswordHash, FirstName, LastName, Role, Status, Active, CreatedAt,
  ApprovedAt, ClientDpa, MustChangePassword, PasswordChangedAt, Phone)
VALUES
 ('admin@aar.bsp.rw',
  '$2a$11$ZvOoeM2CRMdhffyo6.A35uiZofi4oeXfk34tCU1vUab3X.YkzfULe',  -- Admin@2026
  'Aline', 'Uwase', 'Administrators', 'Approved', 1, SYSUTCDATETIME(),
  SYSUTCDATETIME(), NULL, 0, SYSUTCDATETIME(), '+250 788 654 321');
GO

/* ---- Cash movements (statement + balance) ------------------------- */
/* EntityType 1 = Client; Entity_DPA_ = client; BankAccount_DPA_ = 1     */
INSERT INTO Payment
 (Payment_DPA_, EntityType_DPA_, Entity_DPA_, PayType_DPA_, BankAccount_DPA_,
  PaymentAmount, PaymentPDate, PaymentReference, PaymentNarrative, Deleted)
VALUES
 (5001, 1, 1001, 1, 1, 10000000, '2026-01-15', 'DEP-0001', 'Initial cash deposit', 0),
 (5002, 1, 1001, 1, 1,   750000, '2026-02-10', 'DEP-0002', 'Top-up deposit',        0),
 (5003, 1, 1001, 2, 1,  1000000, '2026-03-05', 'WTH-0001', 'Cash withdrawal',       0);
GO

/* ---- Materialized balance -----------------------------------------
   CurrentBal = Opening 5,000,000 + receipts 10,750,000 - payments 1,000,000
              = 14,750,000                                                */
INSERT INTO ClientBalances (client_DPA_, CurrentBal) VALUES (1001, 14750000);
GO

/* ---- Verify -------------------------------------------------------- */
SELECT 'PayType'       AS tbl, COUNT(*) AS n FROM PayType
UNION ALL SELECT 'OrderType',     COUNT(*) FROM OrderType
UNION ALL SELECT 'OrderSecType',  COUNT(*) FROM OrderSecType
UNION ALL SELECT 'Security',      COUNT(*) FROM Security
UNION ALL SELECT 'MarketQuotes',  COUNT(*) FROM MarketQuotes
UNION ALL SELECT 'Client',        COUNT(*) FROM Client
UNION ALL SELECT 'PortalUsers',   COUNT(*) FROM PortalUsers
UNION ALL SELECT 'Payment',       COUNT(*) FROM Payment
UNION ALL SELECT 'ClientBalances',COUNT(*) FROM ClientBalances;
GO
