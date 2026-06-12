/* =====================================================================
   SCHEMA BASELINE GENERATOR — READ-ONLY. Target: BrokerKnow_Test.
   Emits ordered DDL for the CLEAN schema (dbo only, [trash] EXCLUDED) as the
   canonical target for the phased migration:
     1. schemas (app, etc.; not dbo/sys/trash)
     2. CREATE TABLE (columns, types w/ precision, IDENTITY, NULL/NOT NULL, DEFAULT)
     3. PRIMARY KEY constraints
     4. FOREIGN KEY constraints
     5. non-PK indexes
     6. views / procedures / functions (exact source via OBJECT_DEFINITION)
   Generates TEXT only — changes nothing. Capture with sqlcmd -y 0 -o file.
   ===================================================================== */
SET QUOTED_IDENTIFIER ON;   -- required for FOR XML PATH below
SET ANSI_NULLS ON;
SET NOCOUNT ON;
USE BrokerKnow_Test;
GO

IF OBJECT_ID('tempdb..#ddl') IS NOT NULL DROP TABLE #ddl;
CREATE TABLE #ddl (seq int IDENTITY(1,1), line nvarchar(max));

INSERT INTO #ddl(line) VALUES
 ('/* ============================================================'),
 ('   BrokerKnow CLEAN SCHEMA BASELINE (dbo; [trash] excluded)'),
 ('   Generated read-only from BrokerKnow_Test catalog.'),
 ('   Canonical target for the phased data migration.'),
 ('   ============================================================ */'),
 ('SET ANSI_NULLS ON;'),('SET QUOTED_IDENTIFIER ON;'),('GO');

/* ---- 1. schemas (exclude built-ins, dbo, trash) -------------------- */
INSERT INTO #ddl(line)
SELECT 'IF SCHEMA_ID(''' + s.name + ''') IS NULL EXEC(''CREATE SCHEMA ' + QUOTENAME(s.name) + ''');'
FROM sys.schemas s
WHERE s.name NOT IN ('dbo','trash','guest','sys','INFORMATION_SCHEMA')
  AND s.schema_id < 16384  -- exclude fixed db roles
  AND EXISTS (SELECT 1 FROM sys.objects o WHERE o.schema_id = s.schema_id);
INSERT INTO #ddl(line) VALUES ('GO');

/* ---- 2. CREATE TABLE for every dbo table --------------------------- */
INSERT INTO #ddl(line)
SELECT
  'CREATE TABLE [dbo].' + QUOTENAME(t.name) + ' (' + CHAR(13)+CHAR(10) +
  STUFF((
    SELECT ',' + CHAR(13)+CHAR(10) + '    ' + QUOTENAME(c.name) + ' ' +
      CASE
        WHEN ty.name IN ('decimal','numeric') THEN ty.name + '(' + CAST(c.precision AS varchar(5)) + ',' + CAST(c.scale AS varchar(5)) + ')'
        WHEN ty.name IN ('varchar','char','varbinary','binary') THEN ty.name + '(' + CASE WHEN c.max_length = -1 THEN 'max' ELSE CAST(c.max_length AS varchar(10)) END + ')'
        WHEN ty.name IN ('nvarchar','nchar') THEN ty.name + '(' + CASE WHEN c.max_length = -1 THEN 'max' ELSE CAST(c.max_length/2 AS varchar(10)) END + ')'
        WHEN ty.name IN ('datetime2','time','datetimeoffset') THEN ty.name + '(' + CAST(c.scale AS varchar(5)) + ')'
        ELSE ty.name
      END +
      CASE WHEN ic.is_identity = 1 THEN ' IDENTITY(' + CAST(CONVERT(bigint, ic.seed_value) AS varchar(20)) + ',' + CAST(CONVERT(bigint, ic.increment_value) AS varchar(20)) + ')' ELSE '' END +
      CASE WHEN c.is_nullable = 0 THEN ' NOT NULL' ELSE ' NULL' END +
      CASE WHEN dc.definition IS NOT NULL THEN ' CONSTRAINT ' + QUOTENAME(dc.name) + ' DEFAULT ' + dc.definition ELSE '' END
    FROM sys.columns c
    JOIN sys.types ty ON ty.user_type_id = c.user_type_id
    LEFT JOIN sys.identity_columns ic ON ic.object_id = c.object_id AND ic.column_id = c.column_id
    LEFT JOIN sys.default_constraints dc ON dc.parent_object_id = c.object_id AND dc.parent_column_id = c.column_id
    WHERE c.object_id = t.object_id
    ORDER BY c.column_id
    FOR XML PATH(''), TYPE).value('.','nvarchar(max)'), 1, 3, '')
  + CHAR(13)+CHAR(10) + ');' + CHAR(13)+CHAR(10) + 'GO'
FROM sys.tables t
WHERE t.schema_id = SCHEMA_ID('dbo')
ORDER BY t.name;

/* ---- 3. PRIMARY KEYS ---------------------------------------------- */
INSERT INTO #ddl(line) VALUES ('-- ===== PRIMARY KEYS =====');
INSERT INTO #ddl(line)
SELECT 'ALTER TABLE [dbo].' + QUOTENAME(t.name) + ' ADD CONSTRAINT ' + QUOTENAME(kc.name)
  + ' PRIMARY KEY ' + i.type_desc COLLATE DATABASE_DEFAULT + ' ('
  + STUFF((SELECT ', ' + QUOTENAME(c.name)
           FROM sys.index_columns ix JOIN sys.columns c ON c.object_id=ix.object_id AND c.column_id=ix.column_id
           WHERE ix.object_id=i.object_id AND ix.index_id=i.index_id ORDER BY ix.key_ordinal
           FOR XML PATH(''), TYPE).value('.','nvarchar(max)'),1,2,'') + ');' + CHAR(13)+CHAR(10) + 'GO'
FROM sys.key_constraints kc
JOIN sys.tables t ON t.object_id=kc.parent_object_id
JOIN sys.indexes i ON i.object_id=kc.parent_object_id AND i.index_id=kc.unique_index_id
WHERE kc.type='PK' AND t.schema_id=SCHEMA_ID('dbo')
ORDER BY t.name;

/* ---- 4. FOREIGN KEYS ---------------------------------------------- */
INSERT INTO #ddl(line) VALUES ('-- ===== FOREIGN KEYS =====');
INSERT INTO #ddl(line)
SELECT 'ALTER TABLE [dbo].' + QUOTENAME(OBJECT_NAME(fk.parent_object_id)) + ' ADD CONSTRAINT ' + QUOTENAME(fk.name)
  + ' FOREIGN KEY (' + QUOTENAME(pc.name) + ') REFERENCES [dbo].' + QUOTENAME(OBJECT_NAME(fk.referenced_object_id))
  + ' (' + QUOTENAME(rc.name) + ');' + CHAR(13)+CHAR(10) + 'GO'
FROM sys.foreign_keys fk
JOIN sys.foreign_key_columns fkc ON fkc.constraint_object_id=fk.object_id
JOIN sys.columns pc ON pc.object_id=fkc.parent_object_id AND pc.column_id=fkc.parent_column_id
JOIN sys.columns rc ON rc.object_id=fkc.referenced_object_id AND rc.column_id=fkc.referenced_column_id
WHERE OBJECT_SCHEMA_NAME(fk.parent_object_id)='dbo'
ORDER BY OBJECT_NAME(fk.parent_object_id), fk.name;

/* ---- 5. non-PK indexes -------------------------------------------- */
INSERT INTO #ddl(line) VALUES ('-- ===== INDEXES (non-PK) =====');
INSERT INTO #ddl(line)
SELECT 'CREATE ' + CASE WHEN i.is_unique=1 THEN 'UNIQUE ' ELSE '' END + i.type_desc COLLATE DATABASE_DEFAULT
  + ' INDEX ' + QUOTENAME(i.name) + ' ON [dbo].' + QUOTENAME(t.name) + ' ('
  + STUFF((SELECT ', ' + QUOTENAME(c.name) + CASE WHEN ix.is_descending_key=1 THEN ' DESC' ELSE '' END
           FROM sys.index_columns ix JOIN sys.columns c ON c.object_id=ix.object_id AND c.column_id=ix.column_id
           WHERE ix.object_id=i.object_id AND ix.index_id=i.index_id AND ix.is_included_column=0 ORDER BY ix.key_ordinal
           FOR XML PATH(''), TYPE).value('.','nvarchar(max)'),1,2,'') + ')'
  + ISNULL(' INCLUDE (' + STUFF((SELECT ', ' + QUOTENAME(c.name)
           FROM sys.index_columns ix JOIN sys.columns c ON c.object_id=ix.object_id AND c.column_id=ix.column_id
           WHERE ix.object_id=i.object_id AND ix.index_id=i.index_id AND ix.is_included_column=1 ORDER BY ix.index_column_id
           FOR XML PATH(''), TYPE).value('.','nvarchar(max)'),1,2,'') + ')','') + ';' + CHAR(13)+CHAR(10) + 'GO'
FROM sys.indexes i
JOIN sys.tables t ON t.object_id=i.object_id
WHERE t.schema_id=SCHEMA_ID('dbo') AND i.is_primary_key=0 AND i.type IN (1,2) AND i.is_unique_constraint=0
  AND i.name IS NOT NULL
ORDER BY t.name, i.name;

/* ---- 6. views / procedures / functions (exact source) ------------- */
INSERT INTO #ddl(line) VALUES ('-- ===== VIEWS / PROCEDURES / FUNCTIONS =====');
INSERT INTO #ddl(line)
SELECT OBJECT_DEFINITION(o.object_id) + CHAR(13)+CHAR(10) + 'GO'
FROM sys.objects o
WHERE o.type IN ('V','P','FN','IF','TF')
  AND OBJECT_SCHEMA_NAME(o.object_id) IN ('dbo','app')
  AND OBJECT_DEFINITION(o.object_id) IS NOT NULL
ORDER BY CASE o.type WHEN 'V' THEN 1 WHEN 'FN' THEN 2 WHEN 'IF' THEN 2 WHEN 'TF' THEN 2 ELSE 3 END, o.name;

/* ---- emit ---------------------------------------------------------- */
SELECT line FROM #ddl ORDER BY seq;
DROP TABLE #ddl;
GO
