/* Restore the 29 Jun Malawi dump as a SEPARATE database BrokerKnow_Malawi0629.
   Live BrokerKnow_Malawi0615 (:5260 prod) is NOT touched — kept as rollback. */
SET NOCOUNT ON;

IF DB_ID('BrokerKnow_Malawi0629') IS NOT NULL
BEGIN
    ALTER DATABASE [BrokerKnow_Malawi0629] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
    DROP DATABASE [BrokerKnow_Malawi0629];
END
GO

RESTORE DATABASE [BrokerKnow_Malawi0629]
FROM DISK = N'/tmp/malawi0629/Malawi290626.bak'
WITH MOVE N'BrokerKnow_dat' TO N'/var/opt/mssql/data/BrokerKnow_Malawi0629.mdf',
     MOVE N'BrokerKnow_log' TO N'/var/opt/mssql/data/BrokerKnow_Malawi0629_log.ldf',
     REPLACE, RECOVERY, STATS = 10;
GO

ALTER DATABASE [BrokerKnow_Malawi0629] SET COMPATIBILITY_LEVEL = 160;
GO

USE [BrokerKnow_Malawi0629];
SELECT 'restored' AS info,
    (SELECT COUNT(*) FROM dbo.Client)   AS Clients,
    (SELECT COUNT(*) FROM dbo.tbOrder)  AS Orders,
    (SELECT COUNT(*) FROM dbo.Contract) AS Contracts,
    (SELECT COUNT(*) FROM dbo.Payment)  AS Payments,
    (SELECT COUNT(*) FROM dbo.Security) AS Securities,
    (SELECT MAX(OrderDate) FROM dbo.tbOrder) AS NewestOrder;
GO
