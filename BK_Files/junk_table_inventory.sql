/* =====================================================================
   JUNK-TABLE CANDIDATE INVENTORY — READ-ONLY. Target: BrokerKnow_Test.
   Lists every base table NOT used by the app, with evidence, plus a
   heuristic category. NOTHING is dropped. For human review before any DROP.
   The app-used set below = EF ToTable() names + raw-SQL refs + legacy infra
   the desktop app needs (AuditTrail/_Initial_Table_ID_/etc are KEEP).
   ===================================================================== */
SET NOCOUNT ON;
USE BrokerKnow_Test;
GO

DECLARE @appTables TABLE (name sysname PRIMARY KEY);
INSERT INTO @appTables (name) VALUES
-- EF-mapped (ToTable)
 ('LevyContract'),('Levy'),('Journal'),('JournalEntry'),('Holdings'),('Client'),
 ('Contract'),('CdsImportedTrades'),('ContractApprovals'),('Agent'),
 ('CdsImportedHoldings'),('OrderHoldOptions'),('Commission'),('LevySecurity'),
 ('OrdDetail'),('Broker'),('Lot'),('PortalPaymentRequests'),('tbOrder'),
 ('PaymentApprovals'),('PriceImportBatches'),('Payment'),('PortalUsers'),
 ('UserPageAccess'),('PortalRefreshTokens'),('PriceImportRows'),('MarketQuotes'),
 ('Branch'),('Gender'),('Residency'),('Institution'),('Owner'),('Class'),
 ('MarketSector'),('OrderType'),('OrderSecType'),('PaymentTypes'),('PayType'),
 ('EntityType'),('Entity'),('Bank'),('BnkBranch'),('BankAcc'),('SecTransFee'),
 ('ClientVoucher'),('Voucher'),('BrokerReceiptVoucher'),('Holidays'),('Bond'),
 ('Status'),('GenericSetting'),('ClientBalances'),('ClientTotal'),('Security'),
 ('Users'),('Groups'),('UserGroups'),('MenuGroups'),('Menus'),
-- legacy infra the DESKTOP app relies on — KEEP even though our API doesn't map them
 ('_Initial_Table_ID_'),('_Parent_Child_Links_'),('_Record_Locks_'),
 ('_ReportsParameters_'),('AuditTrail'),('AuditTrailItem'),('dtproperties'),
 ('_CDS_Imported_Trades_'),('_CDS_Imported_Holdings_'),('_CDS_Imported_Files_'),
 ('MenuTypes'),('Sector'),('OfferType'),('Bonds');

SELECT
  t.name AS table_name,
  p.rows AS row_count,
  CAST(t.create_date AS date) AS created,
  CAST(t.modify_date AS date) AS modified,
  CASE
    WHEN t.name LIKE '%[-][-]'                              THEN 'JUNK: -- suffix'
    WHEN t.name LIKE '%BKP' OR t.name LIKE '%[_]bak%' OR t.name LIKE '%backup%' THEN 'JUNK: backup name'
    WHEN t.name LIKE '%2008%' OR t.name LIKE '%2009%'
      OR t.name LIKE '%Jan%' OR t.name LIKE '%Feb%' OR t.name LIKE '%Mar%'
      OR t.name LIKE '%[0-9][0-9][0-9][0-9][0-9][0-9]%'     THEN 'JUNK: date-stamped backup'
    WHEN t.name LIKE 'temp%' OR t.name LIKE 'tmp%' OR t.name LIKE '%Temp%' THEN 'JUNK: temp'
    WHEN t.name = 'Test' OR t.name LIKE '%Test'             THEN 'JUNK: test'
    WHEN t.name LIKE '%2'                                   THEN 'REVIEW: numbered dup?'
    ELSE 'REVIEW: unmapped'
  END AS category
FROM sys.tables t
JOIN sys.partitions p ON p.object_id = t.object_id AND p.index_id IN (0,1)
WHERE t.name NOT IN (SELECT name FROM @appTables)
ORDER BY category, p.rows DESC, t.name;
GO
