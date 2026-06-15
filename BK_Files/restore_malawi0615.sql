/* Restore the 15 Jun morning Malawi dump as a SEPARATE database.
   Live BrokerKnow is NOT touched. Disposable; drop when done. */
SET NOCOUNT ON;

IF DB_ID('BrokerKnow_Malawi0615') IS NOT NULL
BEGIN
    ALTER DATABASE [BrokerKnow_Malawi0615] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
    DROP DATABASE [BrokerKnow_Malawi0615];
END
GO

RESTORE DATABASE [BrokerKnow_Malawi0615]
FROM DISK = N'/tmp/malawi0615/Malawi150626_Morning.bak'
WITH MOVE N'BrokerKnow_dat' TO N'/var/opt/mssql/data/BrokerKnow_Malawi0615.mdf',
     MOVE N'BrokerKnow_log' TO N'/var/opt/mssql/data/BrokerKnow_Malawi0615_log.ldf',
     REPLACE, RECOVERY, STATS = 10;
GO

ALTER DATABASE [BrokerKnow_Malawi0615] SET COMPATIBILITY_LEVEL = 160;
GO

USE [BrokerKnow_Malawi0615];
SELECT 'restored' AS info,
    (SELECT COUNT(*) FROM dbo.Client)   AS Clients,
    (SELECT COUNT(*) FROM dbo.tbOrder)  AS Orders,
    (SELECT COUNT(*) FROM dbo.Contract) AS Contracts,
    (SELECT COUNT(*) FROM dbo.Payment)  AS Payments,
    (SELECT COUNT(*) FROM dbo.Security) AS Securities,
    (SELECT MAX(OrderDate) FROM dbo.tbOrder) AS NewestOrder;
GO
