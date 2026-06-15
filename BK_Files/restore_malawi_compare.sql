/* restore_malawi_compare.sql — restore the Malawi dump as a SEPARATE database
   for comparison. The live BrokerKnow DB (powers cedarcapital/cedarclient
   logins) is NEVER referenced here, so current logins are unaffected.

   Source: Malawi (SQL 2000 / compat 80), logical files BrokerKnow_dat/_log.
   Target: BrokerKnow_Malawi0612, files under /var/opt/mssql/data/. */
SET NOCOUNT ON;

IF DB_ID('BrokerKnow_Malawi0612') IS NOT NULL
BEGIN
    ALTER DATABASE BrokerKnow_Malawi0612 SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
    DROP DATABASE BrokerKnow_Malawi0612;
END
GO

RESTORE DATABASE BrokerKnow_Malawi0612
FROM DISK = N'/var/opt/mssql/restore/Malawi120626_02.bak'
WITH
    MOVE N'BrokerKnow_dat' TO N'/var/opt/mssql/data/BrokerKnow_Malawi0612.mdf',
    MOVE N'BrokerKnow_log' TO N'/var/opt/mssql/data/BrokerKnow_Malawi0612_log.ldf',
    RECOVERY, REPLACE, STATS = 10;
GO

-- Bring the old compat level forward so it's queryable on SQL 2022.
ALTER DATABASE BrokerKnow_Malawi0612 SET COMPATIBILITY_LEVEL = 160;
GO

PRINT '===== restored OK — quick row counts (compare vs live BrokerKnow) =====';
USE BrokerKnow_Malawi0612;
SELECT 'Malawi0612' AS db,
       (SELECT COUNT(*) FROM dbo.Client      WHERE ISNULL(Deleted,0)=0) AS clients_live,
       (SELECT COUNT(*) FROM dbo.tbOrder)                                 AS orders_total,
       (SELECT COUNT(*) FROM dbo.Contract)                                AS contracts_total,
       (SELECT COUNT(*) FROM dbo.Payment)                                 AS payments_total,
       (SELECT COUNT(*) FROM dbo.Lot)                                     AS lots_total;
GO
