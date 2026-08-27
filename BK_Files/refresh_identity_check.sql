/* Refresh guard: does any client id change WHO it refers to?

   The refresh replaces dbo.Client wholesale from the legacy dump, but the
   app-layer tables (PortalUsers, IpoApplications, ...) are grafted across and
   still hold the OLD Client_DPA_ values. Clients created in BrokerKnow take
   their key from MAX(Client_DPA_)+1 in the app database, while Axis allocates
   the same numbers to different people - so after a refresh a portal login can
   silently point at a stranger. That happened to 4 logins before 2026-08-18.

   The existing delta check compares clients by Client_DPA_ only, so a
   collision looks like a match. This compares IDENTITY.

   Usage (before cutover):
     :setvar BuildDb BrokerKnow_Clean
     sqlcmd -C -I -d axis_db_prod -i refresh_identity_check.sql -v BuildDb="BrokerKnow_Clean"
*/
SET NOCOUNT ON;

DECLARE @build sysname = N'$(BuildDb)';
DECLARE @sql nvarchar(max) = N'
SELECT
    live.Client_DPA_                              AS ClientId,
    LEFT(ISNULL(live.ClientName, ''''), 40)       AS NameNowOnLive,
    LEFT(ISNULL(bld.ClientName, ''''), 40)        AS NameInIncomingDump,
    LEFT(ISNULL(live.ClientEmail, ''''), 30)      AS EmailNowOnLive,
    (SELECT COUNT(*) FROM dbo.PortalUsers p
      WHERE p.ClientDpa = live.Client_DPA_)       AS PortalLoginsAffected
FROM dbo.Client live
JOIN ' + QUOTENAME(@build) + N'.dbo.Client bld
  ON bld.Client_DPA_ = live.Client_DPA_
WHERE ISNULL(live.ClientName, '''') <> ISNULL(bld.ClientName, '''')
ORDER BY PortalLoginsAffected DESC, live.Client_DPA_;';

PRINT '--- client ids whose identity changes at cutover ---';
EXEC sp_executesql @sql;

DECLARE @count int, @withLogins int;
-- the HasLogin flag must be computed in a derived table: SQL Server rejects an
-- EXISTS subquery inside SUM(), and a failure here would print a false "safe".
SET @sql = N'
SELECT @c = COUNT(*), @w = ISNULL(SUM(x.HasLogin), 0)
FROM (
    SELECT CASE WHEN EXISTS (SELECT 1 FROM dbo.PortalUsers p
                             WHERE p.ClientDpa = live.Client_DPA_)
                THEN 1 ELSE 0 END AS HasLogin
    FROM dbo.Client live
    JOIN ' + QUOTENAME(@build) + N'.dbo.Client bld ON bld.Client_DPA_ = live.Client_DPA_
    WHERE ISNULL(live.ClientName, '''') <> ISNULL(bld.ClientName, '''')
) x;';
EXEC sp_executesql @sql, N'@c int OUTPUT, @w int OUTPUT', @c = @count OUTPUT, @w = @withLogins OUTPUT;

PRINT '';
PRINT 'ids changing identity      : ' + CAST(@count AS varchar);
PRINT 'of those, with a portal login: ' + CAST(ISNULL(@withLogins, 0) AS varchar);

IF ISNULL(@withLogins, 0) > 0
BEGIN
    RAISERROR('A client id that a portal login points at will refer to a DIFFERENT person after cutover. Re-point or remove those logins first.', 16, 1);
END
ELSE IF @count > 0
BEGIN
    PRINT 'No portal login is affected. Review the list above, then proceed.';
END
ELSE
BEGIN
    PRINT 'No client id changes identity. Safe to cut over.';
END
