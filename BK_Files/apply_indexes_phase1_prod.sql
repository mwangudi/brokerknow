/* Phase-1 indexes for PROD (BrokerKnow). Same as test, but USE BrokerKnow.
   Purely additive nonclustered indexes, idempotent (IF NOT EXISTS),
   reversible (DROP INDEX). No before/after timing here — just create + list.
   The CREATE INDEX takes a brief schema-modification lock per table; core
   tables are small (<=225K rows) so each completes in well under a second. */
SET NOCOUNT ON;
USE BrokerKnow;
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_LevyContract_Contract' AND object_id=OBJECT_ID('dbo.LevyContract'))
BEGIN CREATE NONCLUSTERED INDEX IX_LevyContract_Contract ON dbo.LevyContract(Contract_DPA_); PRINT 'created IX_LevyContract_Contract'; END ELSE PRINT 'exists IX_LevyContract_Contract';
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_Payment_Entity' AND object_id=OBJECT_ID('dbo.Payment'))
BEGIN CREATE NONCLUSTERED INDEX IX_Payment_Entity ON dbo.Payment(EntityType_DPA_, Entity_DPA_, PayType_DPA_) INCLUDE (PaymentAmount, PaymentPDate, Deleted); PRINT 'created IX_Payment_Entity'; END ELSE PRINT 'exists IX_Payment_Entity';
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_Payment_Contract' AND object_id=OBJECT_ID('dbo.Payment'))
BEGIN CREATE NONCLUSTERED INDEX IX_Payment_Contract ON dbo.Payment(Contract_DPA_); PRINT 'created IX_Payment_Contract'; END ELSE PRINT 'exists IX_Payment_Contract';
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_Payment_Order' AND object_id=OBJECT_ID('dbo.Payment'))
BEGIN CREATE NONCLUSTERED INDEX IX_Payment_Order ON dbo.Payment(Order_DPA_); PRINT 'created IX_Payment_Order'; END ELSE PRINT 'exists IX_Payment_Order';
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_Lot_OrdDetail' AND object_id=OBJECT_ID('dbo.Lot'))
BEGIN CREATE NONCLUSTERED INDEX IX_Lot_OrdDetail ON dbo.Lot(OrdDetail_DPA_); PRINT 'created IX_Lot_OrdDetail'; END ELSE PRINT 'exists IX_Lot_OrdDetail';
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_Lot_Contract' AND object_id=OBJECT_ID('dbo.Lot'))
BEGIN CREATE NONCLUSTERED INDEX IX_Lot_Contract ON dbo.Lot(Contract_DPA_); PRINT 'created IX_Lot_Contract'; END ELSE PRINT 'exists IX_Lot_Contract';
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_Lot_Broker' AND object_id=OBJECT_ID('dbo.Lot'))
BEGIN CREATE NONCLUSTERED INDEX IX_Lot_Broker ON dbo.Lot(Broker_DPA_); PRINT 'created IX_Lot_Broker'; END ELSE PRINT 'exists IX_Lot_Broker';
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_OrdDetail_Order' AND object_id=OBJECT_ID('dbo.OrdDetail'))
BEGIN CREATE NONCLUSTERED INDEX IX_OrdDetail_Order ON dbo.OrdDetail(Order_DPA_); PRINT 'created IX_OrdDetail_Order'; END ELSE PRINT 'exists IX_OrdDetail_Order';
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_OrdDetail_Security' AND object_id=OBJECT_ID('dbo.OrdDetail'))
BEGIN CREATE NONCLUSTERED INDEX IX_OrdDetail_Security ON dbo.OrdDetail(Security_DPA_); PRINT 'created IX_OrdDetail_Security'; END ELSE PRINT 'exists IX_OrdDetail_Security';
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_JournalEntry_Journal' AND object_id=OBJECT_ID('dbo.JournalEntry'))
BEGIN CREATE NONCLUSTERED INDEX IX_JournalEntry_Journal ON dbo.JournalEntry(Journal_DPA_); PRINT 'created IX_JournalEntry_Journal'; END ELSE PRINT 'exists IX_JournalEntry_Journal';
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_JournalEntry_Entity' AND object_id=OBJECT_ID('dbo.JournalEntry'))
BEGIN CREATE NONCLUSTERED INDEX IX_JournalEntry_Entity ON dbo.JournalEntry(EntityType_DPA_, Entity_DPA_) INCLUDE (JournalEntryDebit, JournalEntryCredit); PRINT 'created IX_JournalEntry_Entity'; END ELSE PRINT 'exists IX_JournalEntry_Entity';
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_tbOrder_Client' AND object_id=OBJECT_ID('dbo.tbOrder'))
BEGIN CREATE NONCLUSTERED INDEX IX_tbOrder_Client ON dbo.tbOrder(Client_DPA_); PRINT 'created IX_tbOrder_Client'; END ELSE PRINT 'exists IX_tbOrder_Client';
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_tbOrder_Agent' AND object_id=OBJECT_ID('dbo.tbOrder'))
BEGIN CREATE NONCLUSTERED INDEX IX_tbOrder_Agent ON dbo.tbOrder(Agent_DPA_); PRINT 'created IX_tbOrder_Agent'; END ELSE PRINT 'exists IX_tbOrder_Agent';
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_Holdings_Client' AND object_id=OBJECT_ID('dbo.Holdings'))
BEGIN CREATE NONCLUSTERED INDEX IX_Holdings_Client ON dbo.Holdings(Client_DPA_); PRINT 'created IX_Holdings_Client'; END ELSE PRINT 'exists IX_Holdings_Client';
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_Holdings_Security' AND object_id=OBJECT_ID('dbo.Holdings'))
BEGIN CREATE NONCLUSTERED INDEX IX_Holdings_Security ON dbo.Holdings(Security_DPA_); PRINT 'created IX_Holdings_Security'; END ELSE PRINT 'exists IX_Holdings_Security';
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_Contract_Status' AND object_id=OBJECT_ID('dbo.Contract'))
BEGIN CREATE NONCLUSTERED INDEX IX_Contract_Status ON dbo.Contract(Status_DPA_); PRINT 'created IX_Contract_Status'; END ELSE PRINT 'exists IX_Contract_Status';
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_Client_Agent' AND object_id=OBJECT_ID('dbo.Client'))
BEGIN CREATE NONCLUSTERED INDEX IX_Client_Agent ON dbo.Client(Agent_DPA_); PRINT 'created IX_Client_Agent'; END ELSE PRINT 'exists IX_Client_Agent';
GO

PRINT '===== PROD index summary =====';
SELECT t.name AS table_name, COUNT(*) AS ix_added
FROM sys.indexes i JOIN sys.tables t ON t.object_id=i.object_id
WHERE i.name LIKE 'IX[_]%'
  AND t.name IN ('LevyContract','Payment','Lot','OrdDetail','JournalEntry','tbOrder','Holdings','Contract','Client')
GROUP BY t.name ORDER BY t.name;
GO
