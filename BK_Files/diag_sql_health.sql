SET NOCOUNT ON;
SELECT name AS setting, CAST(value_in_use AS bigint) AS value_in_use
FROM sys.configurations
WHERE name IN ('max server memory (MB)', 'min server memory (MB)');
-- Recovery model per DB (FULL is required for the 15-min log backups to work)
SELECT name AS db, recovery_model_desc
FROM sys.databases
WHERE name IN ('BrokerKnow','BrokerKnow_Test','BrokerKnow_Clean','master');
GO
