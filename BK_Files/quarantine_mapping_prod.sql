/* Precise mapping: which Tier-1 table is referenced by which object, so we
   quarantine ONLY tables that nothing references. READ-ONLY. */
SET NOCOUNT ON;
USE BrokerKnow;
GO
DECLARE @tier1 TABLE (name sysname PRIMARY KEY);
INSERT INTO @tier1 (name) VALUES
 ('Menus--'),('MenuGroups--'),('MenuGroupsBKP'),
 ('LevyContract2009Feb9'),('LevyContract20081126'),('LevyContract_20081120a'),
 ('LevyContract_20081120'),('OrdDetail2009Feb5'),('temp_MarchContractSchedule'),
 ('Users20081124'),('datastream_Market'),('excep_SummaryHoldings'),
 ('clientBalancesTemp'),('TempInstitutionMapping'),('temp01'),('PrimaryIssues');

PRINT '===== referenced Tier-1 tables -> referencing objects =====';
SELECT d.referenced_entity_name AS tier1_table,
       OBJECT_NAME(d.referencing_id) AS referenced_by,
       o.type_desc
FROM sys.sql_expression_dependencies d
JOIN sys.objects o ON o.object_id = d.referencing_id
WHERE d.referenced_entity_name IN (SELECT name FROM @tier1)
ORDER BY d.referenced_entity_name, referenced_by;

PRINT '===== Tier-1 tables with ZERO references (safe to quarantine now) =====';
SELECT t.name
FROM @tier1 t
WHERE NOT EXISTS (
    SELECT 1 FROM sys.sql_expression_dependencies d
    WHERE d.referenced_entity_name = t.name)
ORDER BY t.name;
GO
