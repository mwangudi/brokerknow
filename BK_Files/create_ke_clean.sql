-- Create the Kenya CLEAN target database (hardened schema target for migration).
-- The baseline script (schema_baseline_BrokerKnow.sql) is then applied with -d
-- against this DB; it creates the dbo tables + app schema + FKs + indexes.
IF DB_ID('BrokerKnow_KE_Clean') IS NOT NULL
BEGIN
    ALTER DATABASE [BrokerKnow_KE_Clean] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
    DROP DATABASE [BrokerKnow_KE_Clean];
END
GO
CREATE DATABASE [BrokerKnow_KE_Clean];
GO
ALTER DATABASE [BrokerKnow_KE_Clean] SET COMPATIBILITY_LEVEL = 160;
GO
PRINT '=== Created empty BrokerKnow_KE_Clean (compat 160) ===';
GO
