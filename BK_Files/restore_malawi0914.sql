/* Restore the 14 Sep 2026 Malawi dump as BrokerKnow_Malawi0914 (raw legacy).
   Refresh #11. Live axis_db_prod is NOT touched by this script — this is the
   side-by-side snapshot used for reconciliation before any cutover. */
SET NOCOUNT ON;
IF DB_ID('BrokerKnow_Malawi0914') IS NOT NULL
BEGIN
    ALTER DATABASE [BrokerKnow_Malawi0914] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
    DROP DATABASE [BrokerKnow_Malawi0914];
END
GO
RESTORE DATABASE [BrokerKnow_Malawi0914]
FROM DISK = N'/tmp/ml0914/Malawi140926.bak'
WITH MOVE N'BrokerKnow_dat' TO N'/var/opt/mssql/data/BrokerKnow_Malawi0914.mdf',
     MOVE N'BrokerKnow_log' TO N'/var/opt/mssql/data/BrokerKnow_Malawi0914_log.ldf',
     REPLACE, RECOVERY, STATS = 10;
GO
ALTER DATABASE [BrokerKnow_Malawi0914] SET COMPATIBILITY_LEVEL = 160;
GO
USE [BrokerKnow_Malawi0914];
SELECT 'Malawi0914' AS info,
    (SELECT COUNT(*) FROM dbo.Client)   AS Clients,
    (SELECT COUNT(*) FROM dbo.tbOrder)  AS Orders,
    (SELECT COUNT(*) FROM dbo.Contract) AS Contracts,
    (SELECT COUNT(*) FROM dbo.Payment)  AS Payments,
    (SELECT COUNT(*) FROM dbo.Security) AS Securities,
    (SELECT MAX(OrderDate) FROM dbo.tbOrder) AS NewestOrder;
GO
