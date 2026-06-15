SET NOCOUNT ON;

PRINT '=== App-layer tables present in BrokerKnow_RW_Demo (graft source) ===';
SELECT t.name
FROM BrokerKnow_RW_Demo.sys.tables t
WHERE t.name IN ('PortalUsers','PortalRefreshTokens','UserPageAccess','AppSettings',
                 'LoginOtps','MarketQuotes','PortalPaymentRequests','ContractApprovals',
                 'PaymentApprovals','PriceImportBatches','PriceImportRows',
                 'CdsImportedTrades','CdsImportedHoldings')
ORDER BY t.name;

PRINT '=== app.* views in RW_Demo ===';
SELECT COUNT(*) AS app_views FROM BrokerKnow_RW_Demo.sys.views WHERE schema_id = SCHEMA_ID('app');

PRINT '=== Does synthetic demo client 1001 collide with a REAL Rwanda client? ===';
SELECT Client_DPA_, ClientName
FROM BrokerKnow_RW_0612.dbo.Client
WHERE Client_DPA_ = 1001;

PRINT '=== PortalUsers identity check in RW_Demo (does Id stay identity after graft?) ===';
SELECT name, is_identity FROM BrokerKnow_RW_Demo.sys.columns
WHERE object_id = OBJECT_ID('BrokerKnow_RW_Demo.dbo.PortalUsers') AND name = 'Id';
