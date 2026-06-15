-- Restore the Rwanda legacy backup as a SEPARATE database (non-destructive).
-- Does NOT touch BrokerKnow_RW_Demo (the current rwanda :5262 DB) — that stays
-- as the instant rollback. Mirrors the Malawi restore pattern.
-- Logical files in the dump are BrokerKnow_dat / BrokerKnow_log (same as Malawi).
IF DB_ID('BrokerKnow_RW_0612') IS NOT NULL
BEGIN
    ALTER DATABASE [BrokerKnow_RW_0612] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
    DROP DATABASE [BrokerKnow_RW_0612];
END
GO
RESTORE DATABASE [BrokerKnow_RW_0612]
FROM DISK = N'/tmp/rw0612/Rwanda120626.bak'
WITH
    MOVE 'BrokerKnow_dat' TO '/var/opt/mssql/data/BrokerKnow_RW_0612.mdf',
    MOVE 'BrokerKnow_log' TO '/var/opt/mssql/data/BrokerKnow_RW_0612_log.ldf',
    REPLACE, RECOVERY, STATS = 10;
GO
ALTER DATABASE [BrokerKnow_RW_0612] SET COMPATIBILITY_LEVEL = 160;
GO
PRINT '=== Restored BrokerKnow_RW_0612 (compat 160) ===';
GO
