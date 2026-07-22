/* Preserve web-created back-office legacy identities across a fresh desktop import.
   Back-office logins are created in the web app referencing dbo.Users rows; those
   rows (and their dbo.UserGroups memberships) live only in the web database and are
   dropped by every fresh desktop dump (whose UserID range stops earlier). Without
   this, those logins resolve to a MISSING identity/role after JWT login.

   Run against the freshly-built clean DB:  sqlcmd -C -d BrokerKnow_Clean -b -i
   Source of the surviving rows = the current live axis_db_prod. Idempotent. */
SET NOCOUNT ON;
SET XACT_ABORT ON;
BEGIN TRAN;

-- UserIDs referenced by an ACTIVE back-office login, present in live axis_db_prod,
-- but missing from this fresh build.
SELECT DISTINCT p.LegacyUserId AS UserID
INTO #u
FROM dbo.PortalUsers p
WHERE p.Active = 1 AND p.LegacyUserId IS NOT NULL
  AND NOT EXISTS (SELECT 1 FROM dbo.Users u          WHERE u.UserID = p.LegacyUserId)
  AND EXISTS     (SELECT 1 FROM axis_db_prod.dbo.Users s WHERE s.UserID = p.LegacyUserId);

SET IDENTITY_INSERT dbo.Users ON;
INSERT dbo.Users (UserID, UserName, Password, Surname, OtherNames, StaffID, Removed,
                  FirstTime, Description, Expires, Enabled, RemoteUser, Client_DPA_,
                  SecretQuestion, SecretAnswer, RequiresSecretQuestion, email, Accepted)
SELECT s.UserID, s.UserName, s.Password, s.Surname, s.OtherNames, s.StaffID, s.Removed,
       s.FirstTime, s.Description, s.Expires, s.Enabled, s.RemoteUser, s.Client_DPA_,
       s.SecretQuestion, s.SecretAnswer, s.RequiresSecretQuestion, s.email, s.Accepted
FROM axis_db_prod.dbo.Users s
WHERE s.UserID IN (SELECT UserID FROM #u);
SET IDENTITY_INSERT dbo.Users OFF;

-- Their group memberships (MemberID is a surrogate identity -> let it regenerate).
INSERT dbo.UserGroups (GroupID, UserID)
SELECT ug.GroupID, ug.UserID
FROM axis_db_prod.dbo.UserGroups ug
WHERE ug.UserID IN (SELECT UserID FROM #u)
  AND NOT EXISTS (SELECT 1 FROM dbo.UserGroups t WHERE t.UserID = ug.UserID AND t.GroupID = ug.GroupID);

DECLARE @n int = (SELECT COUNT(*) FROM #u);
PRINT CONCAT('preserved web-created Users: ', @n);
COMMIT;

PRINT '=== remaining orphaned active back-office logins (missing in prod too -> need manual decision) ===';
SELECT p.Id, p.Email, p.Username, p.Role, p.LegacyUserId AS orphanUserId
FROM dbo.PortalUsers p
WHERE p.Active = 1 AND p.LegacyUserId IS NOT NULL
  AND NOT EXISTS (SELECT 1 FROM dbo.Users u WHERE u.UserID = p.LegacyUserId)
ORDER BY p.Id;
GO
