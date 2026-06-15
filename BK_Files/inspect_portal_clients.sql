SET NOCOUNT ON;

PRINT '=== Client portal rows (Role = Client) ===';
SELECT Id, Email, Status, ClientDpa, LastSeenAt, CreatedAt
FROM dbo.PortalUsers
WHERE Role = 'Client'
ORDER BY Id;

PRINT '=== Foreign keys that REFERENCE dbo.PortalUsers (children to clean first) ===';
SELECT fk.name AS fk_name,
       OBJECT_SCHEMA_NAME(fk.parent_object_id) + '.' + OBJECT_NAME(fk.parent_object_id) AS child_table,
       COL_NAME(fkc.parent_object_id, fkc.parent_column_id) AS child_col
FROM sys.foreign_keys fk
JOIN sys.foreign_key_columns fkc ON fkc.constraint_object_id = fk.object_id
WHERE fk.referenced_object_id = OBJECT_ID('dbo.PortalUsers')
ORDER BY child_table;

PRINT '=== Activity / child-row counts for client-role users ===';
SELECT
  (SELECT COUNT(*) FROM dbo.PortalUsers WHERE Role = 'Client' AND LastSeenAt IS NOT NULL) AS clients_ever_seen,
  (SELECT COUNT(*) FROM dbo.PortalRefreshTokens rt JOIN dbo.PortalUsers u ON rt.PortalUserId = u.Id WHERE u.Role = 'Client') AS client_refresh_tokens,
  (SELECT COUNT(*) FROM dbo.UserPageAccess pa JOIN dbo.PortalUsers u ON pa.PortalUserId = u.Id WHERE u.Role = 'Client') AS client_page_access;
