SET NOCOUNT ON;
-- Cap SQL Server so it can't balloon and starve the 4 dotnet APIs / OS on the
-- 3.8 GB droplet. 1800 MB leaves ~2 GB for APIs + OS. Online, no restart.
EXEC sys.sp_configure 'show advanced options', 1; RECONFIGURE;
EXEC sys.sp_configure 'max server memory (MB)', 1800; RECONFIGURE;
EXEC sys.sp_configure 'show advanced options', 0; RECONFIGURE;
-- Confirm
SELECT name AS setting, CAST(value_in_use AS bigint) AS value_in_use
FROM sys.configurations
WHERE name IN ('max server memory (MB)', 'min server memory (MB)');
GO
