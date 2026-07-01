/* graft_app_layer_malawi0701.sql — copy NEW-system tables (incl. PortalUsers
   logins) from the CURRENT live BrokerKnow_Malawi0629 into Malawi0701 so the
   SAME credentials work against the fresh legacy data. Live 0629 READ-only. */
SET QUOTED_IDENTIFIER ON;
SET ANSI_NULLS ON;
SET NOCOUNT ON;
USE [BrokerKnow_Malawi0701];
GO

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

SELECT * INTO dbo.PortalUsers           FROM [BrokerKnow_Malawi0629].dbo.PortalUsers;
SELECT * INTO dbo.PortalRefreshTokens   FROM [BrokerKnow_Malawi0629].dbo.PortalRefreshTokens;
SELECT * INTO dbo.UserPageAccess        FROM [BrokerKnow_Malawi0629].dbo.UserPageAccess;
SELECT * INTO dbo.PortalPaymentRequests FROM [BrokerKnow_Malawi0629].dbo.PortalPaymentRequests;
SELECT * INTO dbo.PriceImportBatches    FROM [BrokerKnow_Malawi0629].dbo.PriceImportBatches;
SELECT * INTO dbo.PriceImportRows       FROM [BrokerKnow_Malawi0629].dbo.PriceImportRows;
SELECT * INTO dbo.CdsImportedTrades     FROM [BrokerKnow_Malawi0629].dbo.CdsImportedTrades;
SELECT * INTO dbo.CdsImportedHoldings   FROM [BrokerKnow_Malawi0629].dbo.CdsImportedHoldings;
SELECT * INTO dbo.ContractApprovals     FROM [BrokerKnow_Malawi0629].dbo.ContractApprovals;
SELECT * INTO dbo.PaymentApprovals      FROM [BrokerKnow_Malawi0629].dbo.PaymentApprovals;
SELECT * INTO dbo.MarketQuotes          FROM [BrokerKnow_Malawi0629].dbo.MarketQuotes;
SELECT * INTO dbo.AppSettings           FROM [BrokerKnow_Malawi0629].dbo.AppSettings;
SELECT * INTO dbo.LoginOtps             FROM [BrokerKnow_Malawi0629].dbo.LoginOtps;
GO

ALTER TABLE dbo.PortalUsers ALTER COLUMN Id int NOT NULL;
ALTER TABLE dbo.PortalUsers ADD CONSTRAINT PK_PortalUsers PRIMARY KEY (Id);
CREATE UNIQUE INDEX IX_PortalUsers_Email ON dbo.PortalUsers (Email)
    WHERE Email IS NOT NULL AND Email <> '';
CREATE UNIQUE INDEX IX_PortalUsers_Username ON dbo.PortalUsers (Username)
    WHERE Username IS NOT NULL;
GO

SELECT 'PortalUsers' AS tbl,
       (SELECT COUNT(*) FROM dbo.PortalUsers) AS rows,
       (SELECT is_identity FROM sys.columns WHERE object_id=OBJECT_ID('dbo.PortalUsers') AND name='Id') AS id_is_identity,
       (SELECT MAX(Id) FROM dbo.PortalUsers) AS max_id,
       (SELECT COUNT(*) FROM dbo.PortalUsers WHERE Role='Client') AS clients,
       (SELECT COUNT(*) FROM dbo.PortalUsers WHERE Role<>'Client') AS backoffice;
GO
