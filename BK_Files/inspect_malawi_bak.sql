SET NOCOUNT ON;
-- Read-only inspection of the Malawi dump BEFORE any restore.
RESTORE HEADERONLY  FROM DISK = N'/var/opt/mssql/restore/Malawi120626_02.bak';
RESTORE FILELISTONLY FROM DISK = N'/var/opt/mssql/restore/Malawi120626_02.bak';
GO
