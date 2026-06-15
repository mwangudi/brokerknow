/* graft_app_layer_malawi0615.sql — make BrokerKnow_Malawi0615 app-ready by copying
   the NEW-system tables (incl. the 59 PortalUsers logins) from the CURRENT live
   BrokerKnow, so the SAME credentials work against the fresh legacy data.
   LIVE BrokerKnow is only READ from — never written. Re-runnable (drops grafted copies).

   Single-table SELECT ... INTO preserves the IDENTITY property of a column (it does
   NOT carry keys/indexes/defaults). We re-add the PK + the two filtered unique indexes
   and reseed identity, then verify is_identity below. */
SET QUOTED_IDENTIFIER ON;
SET ANSI_NULLS ON;
SET NOCOUNT ON;
USE [BrokerKnow_Malawi0615];
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

-- Copy structure+data from live (single-table SELECT INTO keeps the identity property).
SELECT * INTO dbo.PortalUsers           FROM [BrokerKnow].dbo.PortalUsers;
SELECT * INTO dbo.PortalRefreshTokens   FROM [BrokerKnow].dbo.PortalRefreshTokens;
SELECT * INTO dbo.UserPageAccess        FROM [BrokerKnow].dbo.UserPageAccess;
SELECT * INTO dbo.PortalPaymentRequests FROM [BrokerKnow].dbo.PortalPaymentRequests;
SELECT * INTO dbo.PriceImportBatches    FROM [BrokerKnow].dbo.PriceImportBatches;
SELECT * INTO dbo.PriceImportRows        FROM [BrokerKnow].dbo.PriceImportRows;
SELECT * INTO dbo.CdsImportedTrades      FROM [BrokerKnow].dbo.CdsImportedTrades;
SELECT * INTO dbo.CdsImportedHoldings    FROM [BrokerKnow].dbo.CdsImportedHoldings;
SELECT * INTO dbo.ContractApprovals      FROM [BrokerKnow].dbo.ContractApprovals;
SELECT * INTO dbo.PaymentApprovals       FROM [BrokerKnow].dbo.PaymentApprovals;
SELECT * INTO dbo.MarketQuotes           FROM [BrokerKnow].dbo.MarketQuotes;
SELECT * INTO dbo.AppSettings            FROM [BrokerKnow].dbo.AppSettings;
SELECT * INTO dbo.LoginOtps              FROM [BrokerKnow].dbo.LoginOtps;
GO

-- PortalUsers: PK + the two filtered unique indexes (login + registration dedup).
ALTER TABLE dbo.PortalUsers ALTER COLUMN Id int NOT NULL;
ALTER TABLE dbo.PortalUsers ADD CONSTRAINT PK_PortalUsers PRIMARY KEY (Id);
CREATE UNIQUE INDEX IX_PortalUsers_Email ON dbo.PortalUsers (Email)
    WHERE Email IS NOT NULL AND Email <> '';
CREATE UNIQUE INDEX IX_PortalUsers_Username ON dbo.PortalUsers (Username)
    WHERE Username IS NOT NULL;
GO

-- Verify identity survived the copy; report.
SELECT 'PortalUsers' AS tbl,
       (SELECT COUNT(*) FROM dbo.PortalUsers) AS rows,
       (SELECT is_identity FROM sys.columns WHERE object_id=OBJECT_ID('dbo.PortalUsers') AND name='Id') AS id_is_identity,
       (SELECT MAX(Id) FROM dbo.PortalUsers) AS max_id;
SELECT TOP 3 Email, Role FROM dbo.PortalUsers WHERE Email IS NOT NULL ORDER BY Id;
GO
