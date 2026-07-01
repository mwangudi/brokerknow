-- Clear online-client PORTAL LOGINS (Role='Client') from BrokerKnow_Malawi0701,
-- keep ALL back-office users. tbClient NOT touched. Snapshotted + transactional:
-- rolls back unless 0 clients remain AND back-office count unchanged.
SET QUOTED_IDENTIFIER ON;
SET ANSI_NULLS ON;
SET NOCOUNT ON;
SET XACT_ABORT ON;
USE [BrokerKnow_Malawi0701];

DECLARE @backofficeBefore int = (SELECT COUNT(*) FROM dbo.PortalUsers WHERE Role<>'Client');
DECLARE @clientsBefore    int = (SELECT COUNT(*) FROM dbo.PortalUsers WHERE Role='Client');

BEGIN TRAN;

IF OBJECT_ID('dbo.PortalUsers_removed_clients_20260701') IS NOT NULL DROP TABLE dbo.PortalUsers_removed_clients_20260701;
SELECT * INTO dbo.PortalUsers_removed_clients_20260701 FROM dbo.PortalUsers WHERE Role='Client';

IF OBJECT_ID('dbo.PortalRefreshTokens_removed_clients_20260701') IS NOT NULL DROP TABLE dbo.PortalRefreshTokens_removed_clients_20260701;
SELECT rt.* INTO dbo.PortalRefreshTokens_removed_clients_20260701
FROM dbo.PortalRefreshTokens rt JOIN dbo.PortalUsers u ON rt.PortalUserId=u.Id WHERE u.Role='Client';

DELETE rt FROM dbo.PortalRefreshTokens rt JOIN dbo.PortalUsers u ON rt.PortalUserId=u.Id WHERE u.Role='Client';
DECLARE @delTokens int = @@ROWCOUNT;
DELETE FROM dbo.PortalUsers WHERE Role='Client';
DECLARE @delUsers int = @@ROWCOUNT;

DECLARE @clientsAfter int = (SELECT COUNT(*) FROM dbo.PortalUsers WHERE Role='Client');
DECLARE @backofficeAfter int = (SELECT COUNT(*) FROM dbo.PortalUsers WHERE Role<>'Client');

PRINT CONCAT('clients before=', @clientsBefore, ' removed=', @delUsers, ' tokens=', @delTokens);
PRINT CONCAT('backoffice before=', @backofficeBefore, ' after=', @backofficeAfter, ' clients remaining=', @clientsAfter);

IF (@clientsAfter=0 AND @backofficeAfter=@backofficeBefore)
BEGIN COMMIT; PRINT 'COMMITTED: client logins cleared, back-office kept.'; END
ELSE BEGIN ROLLBACK; PRINT 'ROLLED BACK: unexpected counts.'; END
