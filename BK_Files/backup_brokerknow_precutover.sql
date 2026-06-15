-- Safety backup of the old (pre-cutover) BrokerKnow before dropping it.
-- This is the Malawi prod rollback; user confirmed fresh Malawi0615 data is OK.
-- Keeps a single .bak file as reversible insurance for a cooling-off period.
BACKUP DATABASE [BrokerKnow]
TO DISK = N'/var/opt/mssql/backups/BrokerKnow_precutover_20260615.bak'
WITH COPY_ONLY, COMPRESSION, INIT, STATS = 25,
     NAME = N'BrokerKnow pre-Malawi0615-cutover safety backup';
GO
