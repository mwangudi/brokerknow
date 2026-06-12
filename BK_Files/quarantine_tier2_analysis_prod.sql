SET NOCOUNT ON;
USE BrokerKnow;
GO
DECLARE @keep TABLE (name sysname PRIMARY KEY);
INSERT INTO @keep (name) VALUES
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
 ('_Initial_Table_ID_'),('_Parent_Child_Links_'),('_Record_Locks_'),
 ('_ReportsParameters_'),('AuditTrail'),('AuditTrailItem'),('dtproperties'),
 ('_CDS_Imported_Trades_'),('_CDS_Imported_Holdings_'),('_CDS_Imported_Files_'),
 ('MenuTypes'),('Sector'),('OfferType'),('Bonds'),('OrderHoldType');
IF OBJECT_ID('tempdb..#t2') IS NOT NULL DROP TABLE #t2;
SELECT t.object_id, t.name, CAST(0 AS int) AS inbound_fks, CAST(0 AS int) AS refs
INTO #t2
FROM sys.tables t
JOIN sys.partitions p ON p.object_id=t.object_id AND p.index_id IN (0,1)
WHERE t.schema_id = SCHEMA_ID('dbo') AND p.rows = 0 AND t.name NOT IN (SELECT name FROM @keep);
UPDATE x SET inbound_fks = (SELECT COUNT(*) FROM sys.foreign_keys fk WHERE fk.referenced_object_id=x.object_id),
             refs = (SELECT COUNT(DISTINCT d.referencing_id) FROM sys.sql_expression_dependencies d WHERE d.referenced_entity_name=x.name)
FROM #t2 x;
SELECT name AS empty_table, inbound_fks, refs,
       CASE WHEN inbound_fks=0 AND refs=0 THEN 'SAFE' ELSE 'REVIEW' END AS verdict
FROM #t2 ORDER BY verdict, name;
PRINT '===== summary =====';
SELECT COUNT(*) AS total, SUM(CASE WHEN inbound_fks=0 AND refs=0 THEN 1 ELSE 0 END) AS safe,
       SUM(CASE WHEN inbound_fks>0 OR refs>0 THEN 1 ELSE 0 END) AS review FROM #t2;
PRINT '===== referenced ones detail =====';
SELECT x.name AS empty_table, OBJECT_NAME(d.referencing_id) AS referenced_by, o.type_desc
FROM #t2 x JOIN sys.sql_expression_dependencies d ON d.referenced_entity_name=x.name
JOIN sys.objects o ON o.object_id=d.referencing_id WHERE x.refs>0 ORDER BY x.name;
DROP TABLE #t2;
GO
