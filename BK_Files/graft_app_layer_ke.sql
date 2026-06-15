/* graft_app_layer_ke.sql — make BrokerKnow_KE_Clean app-ready by copying the
   NEW-system tables from BrokerKnow_KE_Demo (the clean optimised :5264 DB), so the
   Green Margin demo ADMIN login works against the migrated REAL Kenya data.
   BrokerKnow_KE_Demo is only READ from. Re-runnable (drops grafted copies).

   PRIVACY: the demo CLIENT login (demo@greenmargin.demo, ClientDpa=1001) is REMOVED
   because real KE client 1001 = "Sammy Onyancha Maina" — keeping it would expose a
   real client's portfolio. Only admin@greenmargin.demo is kept.

   NOTE: KE_Clean already HAS the app-layer tables (they came from the baseline schema)
   but EMPTY. We DROP + re-SELECT INTO from KE_Demo to get the seeded logins, matching
   the Malawi/Rwanda graft approach. */
SET QUOTED_IDENTIFIER ON;
SET ANSI_NULLS ON;
SET NOCOUNT ON;
USE [BrokerKnow_KE_Clean];
GO

-- Drop the (empty, baseline-created) app-layer tables so we can SELECT INTO copies.
DECLARE @t sysname, @sql nvarchar(max);
DECLARE c CURSOR FOR SELECT name FROM sys.tables WHERE name IN (
    'PortalUsers','PortalRefreshTokens','UserPageAccess','PortalPaymentRequests',
    'PriceImportBatches','PriceImportRows','CdsImportedTrades','CdsImportedHoldings',
    'ContractApprovals','PaymentApprovals','MarketQuotes','AppSettings','LoginOtps');
OPEN c; FETCH NEXT FROM c INTO @t;
WHILE @@FETCH_STATUS = 0
BEGIN
    -- drop FKs that reference/originate on the table first (baseline may add some), then the table
    SET @sql = N'IF OBJECT_ID(''dbo.' + @t + ''') IS NOT NULL DROP TABLE dbo.' + QUOTENAME(@t); 
    BEGIN TRY EXEC sp_executesql @sql; END TRY BEGIN CATCH PRINT 'skip drop ' + @t + ': ' + ERROR_MESSAGE(); END CATCH
    FETCH NEXT FROM c INTO @t;
END
CLOSE c; DEALLOCATE c;
GO

SELECT * INTO dbo.PortalUsers           FROM [BrokerKnow_KE_Demo].dbo.PortalUsers;
SELECT * INTO dbo.PortalRefreshTokens   FROM [BrokerKnow_KE_Demo].dbo.PortalRefreshTokens;
SELECT * INTO dbo.UserPageAccess        FROM [BrokerKnow_KE_Demo].dbo.UserPageAccess;
SELECT * INTO dbo.PortalPaymentRequests FROM [BrokerKnow_KE_Demo].dbo.PortalPaymentRequests;
SELECT * INTO dbo.PriceImportBatches    FROM [BrokerKnow_KE_Demo].dbo.PriceImportBatches;
SELECT * INTO dbo.PriceImportRows        FROM [BrokerKnow_KE_Demo].dbo.PriceImportRows;
SELECT * INTO dbo.CdsImportedTrades      FROM [BrokerKnow_KE_Demo].dbo.CdsImportedTrades;
SELECT * INTO dbo.CdsImportedHoldings    FROM [BrokerKnow_KE_Demo].dbo.CdsImportedHoldings;
SELECT * INTO dbo.ContractApprovals      FROM [BrokerKnow_KE_Demo].dbo.ContractApprovals;
SELECT * INTO dbo.PaymentApprovals       FROM [BrokerKnow_KE_Demo].dbo.PaymentApprovals;
SELECT * INTO dbo.MarketQuotes           FROM [BrokerKnow_KE_Demo].dbo.MarketQuotes;
SELECT * INTO dbo.AppSettings            FROM [BrokerKnow_KE_Demo].dbo.AppSettings;
SELECT * INTO dbo.LoginOtps              FROM [BrokerKnow_KE_Demo].dbo.LoginOtps;
GO

-- Remove the synthetic demo CLIENT login (linked to real client 1001) + its children.
DELETE FROM dbo.PortalRefreshTokens WHERE PortalUserId IN (SELECT Id FROM dbo.PortalUsers WHERE Email = 'demo@greenmargin.demo');
DELETE FROM dbo.UserPageAccess      WHERE PortalUserId IN (SELECT Id FROM dbo.PortalUsers WHERE Email = 'demo@greenmargin.demo');
DELETE FROM dbo.PortalUsers WHERE Email = 'demo@greenmargin.demo';
GO

-- PortalUsers: PK + the two filtered unique indexes.
ALTER TABLE dbo.PortalUsers ALTER COLUMN Id int NOT NULL;
ALTER TABLE dbo.PortalUsers ADD CONSTRAINT PK_PortalUsers PRIMARY KEY (Id);
CREATE UNIQUE INDEX IX_PortalUsers_Email ON dbo.PortalUsers (Email) WHERE Email IS NOT NULL AND Email <> '';
CREATE UNIQUE INDEX IX_PortalUsers_Username ON dbo.PortalUsers (Username) WHERE Username IS NOT NULL;
GO

SELECT 'PortalUsers' AS tbl,
       (SELECT COUNT(*) FROM dbo.PortalUsers) AS rows,
       (SELECT is_identity FROM sys.columns WHERE object_id=OBJECT_ID('dbo.PortalUsers') AND name='Id') AS id_is_identity;
SELECT Id, Email, Role, Status, ClientDpa FROM dbo.PortalUsers ORDER BY Id;
GO
