/* graft_app_layer_malawi.sql — make BrokerKnow_Malawi0612 app-ready by copying
   the NEW-system tables (incl. the 59 PortalUsers logins) from the current live
   BrokerKnow, so the SAME credentials work against the fresh legacy data.
   The API startup migration also self-creates these, but we seed the DATA here.
   LIVE BrokerKnow is only READ from. Idempotent-ish (drops the grafted copies first). */
SET QUOTED_IDENTIFIER ON;
SET ANSI_NULLS ON;
SET NOCOUNT ON;
USE BrokerKnow_Malawi0612;
GO

-- Copy each new-system table structure+data from live IF not already present.
-- SELECT ... INTO clones columns+data (no constraints/identity — fine for a test
-- graft; the app re-adds what it needs and these are app-managed tables).
IF OBJECT_ID('dbo.PortalUsers') IS NULL
    SELECT * INTO dbo.PortalUsers          FROM BrokerKnow.dbo.PortalUsers;
IF OBJECT_ID('dbo.PortalRefreshTokens') IS NULL
    SELECT * INTO dbo.PortalRefreshTokens  FROM BrokerKnow.dbo.PortalRefreshTokens;
IF OBJECT_ID('dbo.UserPageAccess') IS NULL
    SELECT * INTO dbo.UserPageAccess       FROM BrokerKnow.dbo.UserPageAccess;
IF OBJECT_ID('dbo.PortalPaymentRequests') IS NULL
    SELECT * INTO dbo.PortalPaymentRequests FROM BrokerKnow.dbo.PortalPaymentRequests;
IF OBJECT_ID('dbo.PriceImportBatches') IS NULL
    SELECT * INTO dbo.PriceImportBatches   FROM BrokerKnow.dbo.PriceImportBatches;
IF OBJECT_ID('dbo.PriceImportRows') IS NULL
    SELECT * INTO dbo.PriceImportRows       FROM BrokerKnow.dbo.PriceImportRows;
IF OBJECT_ID('dbo.CdsImportedTrades') IS NULL
    SELECT * INTO dbo.CdsImportedTrades     FROM BrokerKnow.dbo.CdsImportedTrades;
IF OBJECT_ID('dbo.CdsImportedHoldings') IS NULL
    SELECT * INTO dbo.CdsImportedHoldings   FROM BrokerKnow.dbo.CdsImportedHoldings;
IF OBJECT_ID('dbo.ContractApprovals') IS NULL
    SELECT * INTO dbo.ContractApprovals     FROM BrokerKnow.dbo.ContractApprovals;
IF OBJECT_ID('dbo.PaymentApprovals') IS NULL
    SELECT * INTO dbo.PaymentApprovals      FROM BrokerKnow.dbo.PaymentApprovals;
IF OBJECT_ID('dbo.MarketQuotes') IS NULL
    SELECT * INTO dbo.MarketQuotes          FROM BrokerKnow.dbo.MarketQuotes;
GO

-- A unique index on PortalUsers.Id is needed for EF identity + login lookups.
-- (SELECT INTO doesn't carry keys.) Add a PK on Id if the column exists.
IF OBJECT_ID('dbo.PortalUsers') IS NOT NULL
   AND COL_LENGTH('dbo.PortalUsers','Id') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM sys.key_constraints WHERE parent_object_id=OBJECT_ID('dbo.PortalUsers'))
BEGIN
    ALTER TABLE dbo.PortalUsers ALTER COLUMN Id int NOT NULL;
    ALTER TABLE dbo.PortalUsers ADD CONSTRAINT PK_PortalUsers_graft PRIMARY KEY (Id);
END
GO

SELECT 'PortalUsers grafted' AS info, COUNT(*) AS n FROM dbo.PortalUsers;
SELECT TOP 3 Email, Role FROM dbo.PortalUsers WHERE Email IS NOT NULL ORDER BY Id;
GO
