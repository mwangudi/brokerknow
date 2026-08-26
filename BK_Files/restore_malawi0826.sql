/* Restore the 26 Aug 2026 Malawi dump as BrokerKnow_Malawi0826 (raw legacy).
   Live axis_db_prod (:5260) is NOT touched. */
SET NOCOUNT ON;
IF DB_ID('BrokerKnow_Malawi0826') IS NOT NULL
BEGIN
    ALTER DATABASE [BrokerKnow_Malawi0826] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
    DROP DATABASE [BrokerKnow_Malawi0826];
END
GO
RESTORE DATABASE [BrokerKnow_Malawi0826]
FROM DISK = N'/tmp/ml0826/Malawi260826.bak'
WITH MOVE N'BrokerKnow_dat' TO N'/var/opt/mssql/data/BrokerKnow_Malawi0826.mdf',
     MOVE N'BrokerKnow_log' TO N'/var/opt/mssql/data/BrokerKnow_Malawi0826_log.ldf',
     REPLACE, RECOVERY, STATS = 10;
GO
ALTER DATABASE [BrokerKnow_Malawi0826] SET COMPATIBILITY_LEVEL = 160;
GO
USE [BrokerKnow_Malawi0826];
SELECT 'Malawi0826' AS info,
    (SELECT COUNT(*) FROM dbo.Client)   AS Clients,
    (SELECT COUNT(*) FROM dbo.tbOrder)  AS Orders,
    (SELECT COUNT(*) FROM dbo.Contract) AS Contracts,
    (SELECT COUNT(*) FROM dbo.Payment)  AS Payments,
    (SELECT COUNT(*) FROM dbo.Security) AS Securities,
    (SELECT MAX(OrderDate) FROM dbo.tbOrder) AS NewestOrder;
GO
