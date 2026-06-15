-- Remove the online-client PORTAL LOGIN accounts (Role = 'Client') from the
-- live BrokerKnow_Malawi0615 DB. These 19 rows were grafted from the old prep
-- DB; no real online clients have been registered against the fresh data yet.
--
-- SAFE BY DESIGN:
--   * Only dbo.PortalUsers (portal logins) + their refresh tokens are touched.
--   * tbClient (the brokerage's real client trading records) is NOT touched.
--   * Rows are snapshotted to dated backup tables BEFORE deletion (instant restore).
--   * Whole thing runs in ONE transaction and ROLLS BACK unless exactly 19 client
--     logins are removed and zero remain.
--   * old `BrokerKnow` DB still holds the originals too (second safety net).
--
-- NOTE: PortalUsers has filtered unique indexes (Email/Username) so any DML
-- requires QUOTED_IDENTIFIER ON + ANSI_NULLS ON (else Msg 1934).
SET QUOTED_IDENTIFIER ON;
SET ANSI_NULLS ON;
SET NOCOUNT ON;
SET XACT_ABORT ON;

BEGIN TRAN;

-- 1) Snapshot (instant restore) -------------------------------------------------
IF OBJECT_ID('dbo.PortalUsers_removed_clients_20260615') IS NOT NULL
    DROP TABLE dbo.PortalUsers_removed_clients_20260615;
SELECT * INTO dbo.PortalUsers_removed_clients_20260615
FROM dbo.PortalUsers WHERE Role = 'Client';

IF OBJECT_ID('dbo.PortalRefreshTokens_removed_clients_20260615') IS NOT NULL
    DROP TABLE dbo.PortalRefreshTokens_removed_clients_20260615;
SELECT rt.* INTO dbo.PortalRefreshTokens_removed_clients_20260615
FROM dbo.PortalRefreshTokens rt
JOIN dbo.PortalUsers u ON rt.PortalUserId = u.Id
WHERE u.Role = 'Client';

DECLARE @snapUsers  int = (SELECT COUNT(*) FROM dbo.PortalUsers_removed_clients_20260615);
DECLARE @snapTokens int = (SELECT COUNT(*) FROM dbo.PortalRefreshTokens_removed_clients_20260615);

-- 2) Delete children first, then the client logins -----------------------------
DELETE rt
FROM dbo.PortalRefreshTokens rt
JOIN dbo.PortalUsers u ON rt.PortalUserId = u.Id
WHERE u.Role = 'Client';
DECLARE @delTokens int = @@ROWCOUNT;

DELETE FROM dbo.PortalUsers WHERE Role = 'Client';
DECLARE @delUsers int = @@ROWCOUNT;

-- 3) Verify ---------------------------------------------------------------------
DECLARE @remain   int = (SELECT COUNT(*) FROM dbo.PortalUsers WHERE Role = 'Client');
DECLARE @totalNow int = (SELECT COUNT(*) FROM dbo.PortalUsers);

PRINT CONCAT('snapshot:  users=', @snapUsers,  ' tokens=', @snapTokens);
PRINT CONCAT('deleted:   users=', @delUsers,   ' tokens=', @delTokens);
PRINT CONCAT('remaining clients=', @remain, '   total PortalUsers now=', @totalNow);

IF (@remain = 0 AND @delUsers = @snapUsers AND @snapUsers = 19)
BEGIN
    COMMIT;
    PRINT 'COMMITTED: 19 client portal logins removed (snapshotted for restore).';
END
ELSE
BEGIN
    ROLLBACK;
    PRINT 'ROLLED BACK: counts did not match expectations - NO changes made.';
END
