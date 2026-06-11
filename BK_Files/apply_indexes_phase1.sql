/* =====================================================================
   BrokerKnow Phase-1 indexes — core _DPA_ join columns.
   Target: BrokerKnow_Test (:5261). TEST ONLY for now — NOT prod.
   Safe: purely ADDITIVE nonclustered indexes. No data change, no DDL on
   columns, legacy app writes unaffected (small insert cost only).
   Idempotent: each CREATE is guarded by IF NOT EXISTS, so re-running is a
   no-op. Reversible: DROP INDEX <name> ON <table>.

   Measurement: SET STATISTICS IO/TIME ON around the two heaviest hot-path
   queries, run BEFORE and AFTER the creates. We use LOGICAL READS (pages
   touched) — independent of the buffer cache, so we do NOT flush buffers
   (this instance is shared with prod; DBCC DROPCLEANBUFFERS would hurt it).
   ===================================================================== */
SET NOCOUNT ON;
USE BrokerKnow_Test;
GO

/* ---------- MEASURE: BEFORE -------------------------------------------- */
PRINT '################ BEFORE INDEXES ################';
GO
SET STATISTICS IO ON;
SET STATISTICS TIME ON;
GO
-- (1) Levies for the busiest contract — full scan of 225K rows today.
DECLARE @c int = (SELECT TOP 1 Contract_DPA_ FROM dbo.LevyContract
                  WHERE Contract_DPA_ IS NOT NULL
                  GROUP BY Contract_DPA_ ORDER BY COUNT(*) DESC);
PRINT '--- BEFORE Q1: LevyContract WHERE Contract_DPA_ = (busiest)';
SELECT COUNT(*) AS levy_rows, SUM(LevyAmount) AS levy_total
FROM dbo.LevyContract WHERE Contract_DPA_ = @c;

-- (2) Client cash receipts — the balance/statement hot path (full scan of Payment today).
DECLARE @cl int = (SELECT TOP 1 Entity_DPA_ FROM dbo.Payment
                   WHERE EntityType_DPA_ = 1
                   GROUP BY Entity_DPA_ ORDER BY COUNT(*) DESC);
PRINT '--- BEFORE Q2: Payment receipts for busiest client';
SELECT SUM(PaymentAmount) AS receipts
FROM dbo.Payment
WHERE EntityType_DPA_ = 1 AND Entity_DPA_ = @cl AND PayType_DPA_ = 1
  AND (Deleted = 0 OR Deleted IS NULL);
GO
SET STATISTICS IO OFF;
SET STATISTICS TIME OFF;
GO

/* ---------- CREATE INDEXES (idempotent) ------------------------------- */
PRINT '################ CREATING INDEXES ################';
GO

-- LevyContract: 225K rows — biggest single win (every levy/contract report).
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_LevyContract_Contract' AND object_id=OBJECT_ID('dbo.LevyContract'))
BEGIN CREATE NONCLUSTERED INDEX IX_LevyContract_Contract ON dbo.LevyContract(Contract_DPA_); PRINT 'created IX_LevyContract_Contract'; END
ELSE PRINT 'exists  IX_LevyContract_Contract';
GO

-- Payment: covering index for balance + statement (hottest money query).
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_Payment_Entity' AND object_id=OBJECT_ID('dbo.Payment'))
BEGIN
    CREATE NONCLUSTERED INDEX IX_Payment_Entity ON dbo.Payment(EntityType_DPA_, Entity_DPA_, PayType_DPA_)
        INCLUDE (PaymentAmount, PaymentPDate, Deleted);
    PRINT 'created IX_Payment_Entity';
END ELSE PRINT 'exists  IX_Payment_Entity';
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_Payment_Contract' AND object_id=OBJECT_ID('dbo.Payment'))
BEGIN CREATE NONCLUSTERED INDEX IX_Payment_Contract ON dbo.Payment(Contract_DPA_); PRINT 'created IX_Payment_Contract'; END
ELSE PRINT 'exists  IX_Payment_Contract';
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_Payment_Order' AND object_id=OBJECT_ID('dbo.Payment'))
BEGIN CREATE NONCLUSTERED INDEX IX_Payment_Order ON dbo.Payment(Order_DPA_); PRINT 'created IX_Payment_Order'; END
ELSE PRINT 'exists  IX_Payment_Order';
GO

-- Lot: trade-match joins (OrdDetail, Contract, Broker).
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_Lot_OrdDetail' AND object_id=OBJECT_ID('dbo.Lot'))
BEGIN CREATE NONCLUSTERED INDEX IX_Lot_OrdDetail ON dbo.Lot(OrdDetail_DPA_); PRINT 'created IX_Lot_OrdDetail'; END
ELSE PRINT 'exists  IX_Lot_OrdDetail';
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_Lot_Contract' AND object_id=OBJECT_ID('dbo.Lot'))
BEGIN CREATE NONCLUSTERED INDEX IX_Lot_Contract ON dbo.Lot(Contract_DPA_); PRINT 'created IX_Lot_Contract'; END
ELSE PRINT 'exists  IX_Lot_Contract';
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_Lot_Broker' AND object_id=OBJECT_ID('dbo.Lot'))
BEGIN CREATE NONCLUSTERED INDEX IX_Lot_Broker ON dbo.Lot(Broker_DPA_); PRINT 'created IX_Lot_Broker'; END
ELSE PRINT 'exists  IX_Lot_Broker';
GO

-- OrdDetail: order → detail and detail → security joins.
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_OrdDetail_Order' AND object_id=OBJECT_ID('dbo.OrdDetail'))
BEGIN CREATE NONCLUSTERED INDEX IX_OrdDetail_Order ON dbo.OrdDetail(Order_DPA_); PRINT 'created IX_OrdDetail_Order'; END
ELSE PRINT 'exists  IX_OrdDetail_Order';
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_OrdDetail_Security' AND object_id=OBJECT_ID('dbo.OrdDetail'))
BEGIN CREATE NONCLUSTERED INDEX IX_OrdDetail_Security ON dbo.OrdDetail(Security_DPA_); PRINT 'created IX_OrdDetail_Security'; END
ELSE PRINT 'exists  IX_OrdDetail_Security';
GO

-- JournalEntry: GL by journal and by entity (statement/balance).
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_JournalEntry_Journal' AND object_id=OBJECT_ID('dbo.JournalEntry'))
BEGIN CREATE NONCLUSTERED INDEX IX_JournalEntry_Journal ON dbo.JournalEntry(Journal_DPA_); PRINT 'created IX_JournalEntry_Journal'; END
ELSE PRINT 'exists  IX_JournalEntry_Journal';
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_JournalEntry_Entity' AND object_id=OBJECT_ID('dbo.JournalEntry'))
BEGIN CREATE NONCLUSTERED INDEX IX_JournalEntry_Entity ON dbo.JournalEntry(EntityType_DPA_, Entity_DPA_)
        INCLUDE (JournalEntryDebit, JournalEntryCredit); PRINT 'created IX_JournalEntry_Entity'; END
ELSE PRINT 'exists  IX_JournalEntry_Entity';
GO

-- tbOrder: orders by client and by agent (agent portal).
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_tbOrder_Client' AND object_id=OBJECT_ID('dbo.tbOrder'))
BEGIN CREATE NONCLUSTERED INDEX IX_tbOrder_Client ON dbo.tbOrder(Client_DPA_); PRINT 'created IX_tbOrder_Client'; END
ELSE PRINT 'exists  IX_tbOrder_Client';
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_tbOrder_Agent' AND object_id=OBJECT_ID('dbo.tbOrder'))
BEGIN CREATE NONCLUSTERED INDEX IX_tbOrder_Agent ON dbo.tbOrder(Agent_DPA_); PRINT 'created IX_tbOrder_Agent'; END
ELSE PRINT 'exists  IX_tbOrder_Agent';
GO

-- Holdings: portfolio by client and by security.
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_Holdings_Client' AND object_id=OBJECT_ID('dbo.Holdings'))
BEGIN CREATE NONCLUSTERED INDEX IX_Holdings_Client ON dbo.Holdings(Client_DPA_); PRINT 'created IX_Holdings_Client'; END
ELSE PRINT 'exists  IX_Holdings_Client';
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_Holdings_Security' AND object_id=OBJECT_ID('dbo.Holdings'))
BEGIN CREATE NONCLUSTERED INDEX IX_Holdings_Security ON dbo.Holdings(Security_DPA_); PRINT 'created IX_Holdings_Security'; END
ELSE PRINT 'exists  IX_Holdings_Security';
GO

-- Contract status (contract listings) and Client→Agent (agent portal lists).
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_Contract_Status' AND object_id=OBJECT_ID('dbo.Contract'))
BEGIN CREATE NONCLUSTERED INDEX IX_Contract_Status ON dbo.Contract(Status_DPA_); PRINT 'created IX_Contract_Status'; END
ELSE PRINT 'exists  IX_Contract_Status';
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_Client_Agent' AND object_id=OBJECT_ID('dbo.Client'))
BEGIN CREATE NONCLUSTERED INDEX IX_Client_Agent ON dbo.Client(Agent_DPA_); PRINT 'created IX_Client_Agent'; END
ELSE PRINT 'exists  IX_Client_Agent';
GO

/* ---------- MEASURE: AFTER -------------------------------------------- */
PRINT '################ AFTER INDEXES ################';
GO
SET STATISTICS IO ON;
SET STATISTICS TIME ON;
GO
DECLARE @c int = (SELECT TOP 1 Contract_DPA_ FROM dbo.LevyContract
                  WHERE Contract_DPA_ IS NOT NULL
                  GROUP BY Contract_DPA_ ORDER BY COUNT(*) DESC);
PRINT '--- AFTER Q1: LevyContract WHERE Contract_DPA_ = (busiest)';
SELECT COUNT(*) AS levy_rows, SUM(LevyAmount) AS levy_total
FROM dbo.LevyContract WHERE Contract_DPA_ = @c;

DECLARE @cl int = (SELECT TOP 1 Entity_DPA_ FROM dbo.Payment
                   WHERE EntityType_DPA_ = 1
                   GROUP BY Entity_DPA_ ORDER BY COUNT(*) DESC);
PRINT '--- AFTER Q2: Payment receipts for busiest client';
SELECT SUM(PaymentAmount) AS receipts
FROM dbo.Payment
WHERE EntityType_DPA_ = 1 AND Entity_DPA_ = @cl AND PayType_DPA_ = 1
  AND (Deleted = 0 OR Deleted IS NULL);
GO
SET STATISTICS IO OFF;
SET STATISTICS TIME OFF;
GO

/* ---------- SUMMARY: indexes now on the core tables ------------------- */
PRINT '################ INDEX SUMMARY ################';
SELECT t.name AS table_name, i.name AS index_name,
       STUFF((SELECT ', ' + c.name
              FROM sys.index_columns ic
              JOIN sys.columns c ON c.object_id=ic.object_id AND c.column_id=ic.column_id
              WHERE ic.object_id=i.object_id AND ic.index_id=i.index_id AND ic.is_included_column=0
              ORDER BY ic.key_ordinal FOR XML PATH('')),1,2,'') AS key_cols
FROM sys.indexes i
JOIN sys.tables t ON t.object_id=i.object_id
WHERE i.name LIKE 'IX[_]%'
  AND t.name IN ('LevyContract','Payment','Lot','OrdDetail','JournalEntry',
                 'tbOrder','Holdings','Contract','Client')
ORDER BY t.name, i.name;
GO
