/* migrate_schema_diff.sql — run against BrokerKnow_Clean.
   Decides whether the programmable build failures are a CLEAN-SCHEMA GAP
   (a table/column present in prod dbo but missing from clean = real bug)
   or pre-existing PROD ROT (views referencing long-dropped objects/columns). */
SET NOCOUNT ON;
USE BrokerKnow_Clean;
GO
PRINT '===== prod dbo TABLES missing from clean (expect NONE; quarantined live in [trash]) =====';
SELECT t.name AS prod_table
FROM BrokerKnow.sys.tables t
WHERE SCHEMA_NAME(t.schema_id) = 'dbo'
  AND NOT EXISTS (SELECT 1 FROM sys.tables ct WHERE ct.name = t.name AND SCHEMA_NAME(ct.schema_id) = 'dbo')
ORDER BY t.name;

PRINT '===== prod dbo COLUMNS missing from clean, on tables that DO exist in clean (expect NONE) =====';
SELECT t.name AS tbl, c.name AS col, ty.name AS typ
FROM BrokerKnow.sys.columns c
JOIN BrokerKnow.sys.tables t ON t.object_id = c.object_id
JOIN BrokerKnow.sys.types ty ON ty.user_type_id = c.user_type_id
WHERE SCHEMA_NAME(t.schema_id) = 'dbo'
  AND EXISTS (SELECT 1 FROM sys.tables ct WHERE ct.name = t.name AND SCHEMA_NAME(ct.schema_id) = 'dbo')
  AND NOT EXISTS (
        SELECT 1 FROM sys.columns cc
        JOIN sys.tables ct ON ct.object_id = cc.object_id
        WHERE ct.name = t.name AND SCHEMA_NAME(ct.schema_id) = 'dbo' AND cc.name = c.name)
ORDER BY t.name, c.name;

PRINT '===== roots of the failure cascades: do these exist in PROD at all? =====';
SELECT o.name, o.type_desc, SCHEMA_NAME(o.schema_id) AS sch
FROM BrokerKnow.sys.objects o
WHERE o.name LIKE '%ClientTransaction%'
   OR o.name IN ('CPayment','Order','ClientStatement')
ORDER BY o.name;

PRINT '===== is ClientStatement itself broken in PROD? (referenced entities that fail to resolve) =====';
BEGIN TRY
    SELECT referenced_entity_name AS refd, referenced_minor_name AS refd_col, is_caller_dependent
    FROM BrokerKnow.sys.dm_sql_referenced_entities('dbo.ClientStatement','OBJECT')
    WHERE referenced_id IS NULL OR referenced_minor_id = 0;
END TRY
BEGIN CATCH
    PRINT 'dm_sql_referenced_entities failed for dbo.ClientStatement: ' + ERROR_MESSAGE();
END CATCH
GO
