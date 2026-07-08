/* Restore the 08 Jul 2026 Rwanda dump as BrokerKnow_RW_0708 (raw legacy).
   Live BrokerKnow_RW_0612 (:5262) is NOT touched — kept as rollback. */
SET NOCOUNT ON;
IF DB_ID('BrokerKnow_RW_0708') IS NOT NULL
BEGIN
    ALTER DATABASE [BrokerKnow_RW_0708] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
    DROP DATABASE [BrokerKnow_RW_0708];
END
GO
RESTORE DATABASE [BrokerKnow_RW_0708]
FROM DISK = N'/tmp/imports0708/Rwanda080726.bak'
WITH MOVE N'BrokerKnow_dat' TO N'/var/opt/mssql/data/BrokerKnow_RW_0708.mdf',
     MOVE N'BrokerKnow_log' TO N'/var/opt/mssql/data/BrokerKnow_RW_0708_log.ldf',
     REPLACE, RECOVERY, STATS = 10;
GO
ALTER DATABASE [BrokerKnow_RW_0708] SET COMPATIBILITY_LEVEL = 160;
GO
USE [BrokerKnow_RW_0708];
SELECT 'RW_0708' AS info,
    (SELECT COUNT(*) FROM dbo.Client)   AS Clients,
    (SELECT COUNT(*) FROM dbo.tbOrder)  AS Orders,
    (SELECT COUNT(*) FROM dbo.Contract) AS Contracts,
    (SELECT COUNT(*) FROM dbo.Payment)  AS Payments,
    (SELECT COUNT(*) FROM dbo.Security) AS Securities,
    (SELECT MAX(OrderDate) FROM dbo.tbOrder) AS NewestOrder;
GO
