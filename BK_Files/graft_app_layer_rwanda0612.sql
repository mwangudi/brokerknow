/* graft_app_layer_rwanda0612.sql — make BrokerKnow_RW_0612 app-ready by copying the
   NEW-system tables from the existing BrokerKnow_RW_Demo (the current rwanda :5262 DB),
   so the African Alliance Rwanda demo ADMIN login works against the REAL Rwanda data.
   BrokerKnow_RW_Demo is only READ from — never written. Re-runnable (drops grafted copies).

   IMPORTANT: the synthetic demo CLIENT login (demo@aar.bsp.rw, ClientDpa=1001) is REMOVED
   here because real Rwanda client 1001 = "BK - Mukabagwira Josephine" — keeping that login
   would expose a real client's portfolio in the demo. Only admin@aar.bsp.rw is kept.

   Single-table SELECT ... INTO preserves the IDENTITY property (does NOT carry keys/indexes/
   defaults). We re-add the PK + the two filtered unique indexes. */
SET QUOTED_IDENTIFIER ON;
SET ANSI_NULLS ON;
SET NOCOUNT ON;
USE [BrokerKnow_RW_0612];
GO

-- Drop any prior grafted copies so this is re-runnable.
DECLARE @t sysname, @sql nvarchar(max);
DECLARE c CURSOR FOR SELECT name FROM sys.tables WHERE name IN (
    'PortalUsers','PortalRefreshTokens','UserPageAccess','PortalPaymentRequests',
    'PriceImportBatches','PriceImportRows','CdsImportedTrades','CdsImportedHoldings',
    'ContractApprovals','PaymentApprovals','MarketQuotes','AppSettings','LoginOtps');
OPEN c; FETCH NEXT FROM c INTO @t;
WHILE @@FETCH_STATUS = 0
BEGIN
    SET @sql = N'DROP TABLE dbo.' + QUOTENAME(@t); EXEC sp_executesql @sql;
    FETCH NEXT FROM c INTO @t;
END
CLOSE c; DEALLOCATE c;
GO

-- Copy structure+data from the demo DB (single-table SELECT INTO keeps the identity property).
SELECT * INTO dbo.PortalUsers           FROM [BrokerKnow_RW_Demo].dbo.PortalUsers;
SELECT * INTO dbo.PortalRefreshTokens   FROM [BrokerKnow_RW_Demo].dbo.PortalRefreshTokens;
SELECT * INTO dbo.UserPageAccess        FROM [BrokerKnow_RW_Demo].dbo.UserPageAccess;
SELECT * INTO dbo.PortalPaymentRequests FROM [BrokerKnow_RW_Demo].dbo.PortalPaymentRequests;
SELECT * INTO dbo.PriceImportBatches    FROM [BrokerKnow_RW_Demo].dbo.PriceImportBatches;
SELECT * INTO dbo.PriceImportRows        FROM [BrokerKnow_RW_Demo].dbo.PriceImportRows;
SELECT * INTO dbo.CdsImportedTrades      FROM [BrokerKnow_RW_Demo].dbo.CdsImportedTrades;
SELECT * INTO dbo.CdsImportedHoldings    FROM [BrokerKnow_RW_Demo].dbo.CdsImportedHoldings;
SELECT * INTO dbo.ContractApprovals      FROM [BrokerKnow_RW_Demo].dbo.ContractApprovals;
SELECT * INTO dbo.PaymentApprovals       FROM [BrokerKnow_RW_Demo].dbo.PaymentApprovals;
SELECT * INTO dbo.MarketQuotes           FROM [BrokerKnow_RW_Demo].dbo.MarketQuotes;
SELECT * INTO dbo.AppSettings            FROM [BrokerKnow_RW_Demo].dbo.AppSettings;
SELECT * INTO dbo.LoginOtps              FROM [BrokerKnow_RW_Demo].dbo.LoginOtps;
GO

-- Remove the synthetic demo CLIENT login (linked to real client 1001) + its children.
-- Keep ONLY admin@aar.bsp.rw so the demo browses real data without exposing a real client.
DELETE FROM dbo.PortalRefreshTokens
WHERE PortalUserId IN (SELECT Id FROM dbo.PortalUsers WHERE Email = 'demo@aar.bsp.rw');
DELETE FROM dbo.UserPageAccess
WHERE PortalUserId IN (SELECT Id FROM dbo.PortalUsers WHERE Email = 'demo@aar.bsp.rw');
DELETE FROM dbo.PortalUsers WHERE Email = 'demo@aar.bsp.rw';
GO

-- PortalUsers: PK + the two filtered unique indexes (login + registration dedup).
ALTER TABLE dbo.PortalUsers ALTER COLUMN Id int NOT NULL;
ALTER TABLE dbo.PortalUsers ADD CONSTRAINT PK_PortalUsers PRIMARY KEY (Id);
CREATE UNIQUE INDEX IX_PortalUsers_Email ON dbo.PortalUsers (Email)
    WHERE Email IS NOT NULL AND Email <> '';
CREATE UNIQUE INDEX IX_PortalUsers_Username ON dbo.PortalUsers (Username)
    WHERE Username IS NOT NULL;
GO

-- Verify identity survived the copy; report logins kept.
SELECT 'PortalUsers' AS tbl,
       (SELECT COUNT(*) FROM dbo.PortalUsers) AS rows,
       (SELECT is_identity FROM sys.columns WHERE object_id=OBJECT_ID('dbo.PortalUsers') AND name='Id') AS id_is_identity,
       (SELECT MAX(Id) FROM dbo.PortalUsers) AS max_id;
SELECT Id, Email, Role, Status, ClientDpa FROM dbo.PortalUsers ORDER BY Id;
GO
