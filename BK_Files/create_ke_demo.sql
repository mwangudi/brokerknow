/* create_ke_demo.sql — build the Axis Kenya demo DB as a SCHEMA-ONLY clone of
   BrokerKnow_Clean (full hardened schema: 119 tables, 22 FKs, indexes, 459
   programmables incl app.* views; ZERO data). Then make it writable for seeding.
   Idempotent: drops any prior BrokerKnow_KE_Demo first. */
IF DB_ID('BrokerKnow_KE_Demo') IS NOT NULL
BEGIN
    ALTER DATABASE BrokerKnow_KE_Demo SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
    DROP DATABASE BrokerKnow_KE_Demo;
END
GO
DBCC CLONEDATABASE (BrokerKnow_Clean, BrokerKnow_KE_Demo) WITH NO_STATISTICS, NO_QUERYSTORE;
GO
-- a freshly cloned DB is READ_ONLY; make it writable for seeding
ALTER DATABASE BrokerKnow_KE_Demo SET READ_WRITE WITH ROLLBACK IMMEDIATE;
ALTER DATABASE BrokerKnow_KE_Demo SET MULTI_USER;
GO
SELECT 'tables' AS k, COUNT(*) AS v FROM BrokerKnow_KE_Demo.sys.tables WHERE SCHEMA_NAME(schema_id) = 'dbo'
UNION ALL SELECT 'fks',   COUNT(*) FROM BrokerKnow_KE_Demo.sys.foreign_keys
UNION ALL SELECT 'views', COUNT(*) FROM BrokerKnow_KE_Demo.sys.views;
GO
