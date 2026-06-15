-- Verify the safety backup is restorable, then drop old BrokerKnow.
RESTORE VERIFYONLY FROM DISK = N'/var/opt/mssql/backups/BrokerKnow_precutover_20260615.bak';
GO
ALTER DATABASE [BrokerKnow] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
DROP DATABASE [BrokerKnow];
GO
PRINT 'Dropped BrokerKnow (safety .bak retained at /var/opt/mssql/backups/).';
GO
