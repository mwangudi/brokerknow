/* =====================================================================
   Tier-1 junk QUARANTINE — dependency pre-check. READ-ONLY.
   Target: BrokerKnow_Test. Confirms the unambiguous backup/temp tables have
   NO inbound FKs and are referenced by NO view/proc/function before we move
   them to a [trash] schema. Nothing is moved here.
   ===================================================================== */
SET NOCOUNT ON;
USE BrokerKnow_Test;
GO

DECLARE @tier1 TABLE (name sysname PRIMARY KEY);
INSERT INTO @tier1 (name) VALUES
 ('Menus--'),('MenuGroups--'),('MenuGroupsBKP'),
 ('LevyContract2009Feb9'),('LevyContract20081126'),('LevyContract_20081120a'),
 ('LevyContract_20081120'),('OrdDetail2009Feb5'),('temp_MarchContractSchedule'),
 ('Users20081124'),('datastream_Market'),('excep_SummaryHoldings'),
 ('clientBalancesTemp'),('TempInstitutionMapping'),('temp01'),('PrimaryIssues');

PRINT '===== Tier-1 tables present + row counts =====';
SELECT t.name, p.rows AS row_count, CAST(t.create_date AS date) AS created
FROM sys.tables t
JOIN sys.partitions p ON p.object_id=t.object_id AND p.index_id IN (0,1)
WHERE t.name IN (SELECT name FROM @tier1)
ORDER BY t.name;

PRINT '===== INBOUND FKs referencing any Tier-1 table (expect NONE) =====';
SELECT OBJECT_NAME(fk.parent_object_id) AS child_table, fk.name AS fk,
       OBJECT_NAME(fk.referenced_object_id) AS tier1_parent
FROM sys.foreign_keys fk
WHERE OBJECT_NAME(fk.referenced_object_id) IN (SELECT name FROM @tier1);

PRINT '===== OUTBOUND FKs from Tier-1 tables (informational) =====';
SELECT OBJECT_NAME(fk.parent_object_id) AS tier1_child, fk.name AS fk,
       OBJECT_NAME(fk.referenced_object_id) AS parent
FROM sys.foreign_keys fk
WHERE OBJECT_NAME(fk.parent_object_id) IN (SELECT name FROM @tier1);

PRINT '===== views / procs / functions referencing any Tier-1 table (expect NONE) =====';
SELECT DISTINCT OBJECT_NAME(d.referencing_id) AS referencing_object, o.type_desc
FROM sys.sql_expression_dependencies d
JOIN sys.objects o ON o.object_id = d.referencing_id
WHERE d.referenced_entity_name IN (SELECT name FROM @tier1);
GO
