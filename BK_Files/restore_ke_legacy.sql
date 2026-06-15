-- Restore the Kenya legacy backup as a SEPARATE database (non-destructive).
-- Does NOT touch BrokerKnow_KE_Demo (the clean optimised :5264 DB).
-- NOTE: KE logical file names differ from Malawi/Rwanda:
--   data = BrokerknowAA_Data, log = BrokerknowAA_Log  (NOT BrokerKnow_dat/_log).
IF DB_ID('BrokerKnow_KE_Legacy') IS NOT NULL
BEGIN
    ALTER DATABASE [BrokerKnow_KE_Legacy] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
    DROP DATABASE [BrokerKnow_KE_Legacy];
END
GO
RESTORE DATABASE [BrokerKnow_KE_Legacy]
FROM DISK = N'/tmp/ke/Bk_KE_311211.bak'
WITH
    MOVE 'BrokerknowAA_Data' TO '/var/opt/mssql/data/BrokerKnow_KE_Legacy.mdf',
    MOVE 'BrokerknowAA_Log'  TO '/var/opt/mssql/data/BrokerKnow_KE_Legacy_log.ldf',
    REPLACE, RECOVERY, STATS = 10;
GO
ALTER DATABASE [BrokerKnow_KE_Legacy] SET COMPATIBILITY_LEVEL = 160;
GO
PRINT '=== Restored BrokerKnow_KE_Legacy (compat 160) ===';
GO
