SET QUOTED_IDENTIFIER ON;
SET ANSI_NULLS ON;
SET NOCOUNT ON;
USE BrokerKnow_Clean;
GO
IF EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_PortalUsers_Username' AND object_id=OBJECT_ID('dbo.PortalUsers'))
    DROP INDEX IX_PortalUsers_Username ON dbo.PortalUsers;
IF EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_PortalUsers_Email' AND object_id=OBJECT_ID('dbo.PortalUsers'))
    DROP INDEX IX_PortalUsers_Email ON dbo.PortalUsers;
GO
CREATE UNIQUE NONCLUSTERED INDEX IX_PortalUsers_Username ON dbo.PortalUsers(Username) WHERE Username IS NOT NULL;
CREATE UNIQUE NONCLUSTERED INDEX IX_PortalUsers_Email ON dbo.PortalUsers(Email) WHERE Email IS NOT NULL AND Email <> '';
GO
ALTER TABLE dbo.PortalUsers NOCHECK CONSTRAINT ALL;
SET IDENTITY_INSERT dbo.PortalUsers ON;
DECLARE @cols nvarchar(max);
SELECT @cols = STRING_AGG(QUOTENAME(c.name), ',') WITHIN GROUP (ORDER BY c.column_id)
FROM sys.columns c JOIN sys.types ty ON ty.user_type_id=c.user_type_id
WHERE c.object_id=OBJECT_ID('dbo.PortalUsers') AND c.is_computed=0 AND ty.name NOT IN ('timestamp','rowversion');
DECLARE @sql nvarchar(max) = 'INSERT INTO dbo.PortalUsers (' + @cols + ') SELECT ' + @cols + ' FROM BrokerKnow.dbo.PortalUsers';
EXEC sp_executesql @sql;
SET IDENTITY_INSERT dbo.PortalUsers OFF;
GO
SELECT 'clean' AS db, COUNT(*) AS portalusers FROM dbo.PortalUsers
UNION ALL SELECT 'src', COUNT(*) FROM BrokerKnow.dbo.PortalUsers;
GO
