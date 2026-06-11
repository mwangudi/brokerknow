/* =====================================================================
   BrokerKnow schema diagnostic — READ-ONLY.
   Target: BrokerKnow_Test (test DB on :5261). NEVER run on prod first.
   Safe: only SELECTs, catalog/DMV reads, and #temp tables. No writes to
   any user table, no DDL on user objects. Reversible (drops its #temps).

   Drives the in-place v2 modernization plan:
     1. heaviest tables           -> where perf/size matters
     2. missing-index DMV         -> indexes SQL Server itself wants
     3. _DPA_ cols w/o index      -> deterministic FK-join index candidates
     4. tables w/o primary key    -> integrity smell
     5. existing FK inventory     -> what's already enforced
     6. orphan / integrity audit  -> rows whose FK has no parent (core path)
   ===================================================================== */
SET NOCOUNT ON;
USE BrokerKnow_Test;   -- prod would be: USE BrokerKnow;
GO

/* ---- 0. context ---------------------------------------------------- */
SELECT 'CONTEXT' AS section,
       DB_NAME() AS database_name,
       (SELECT COUNT(*) FROM sys.tables) AS user_tables,
       (SELECT COUNT(*) FROM sys.foreign_keys) AS foreign_keys,
       CAST(SUM(a.total_pages) * 8.0 / 1024 AS decimal(10,1)) AS data_mb
FROM sys.allocation_units a;
GO

/* ---- 1. heaviest tables (top 25 by size) --------------------------- */
SELECT TOP 25
       'HEAVY_TABLES' AS section,
       t.name AS table_name,
       p.rows AS row_count,
       CAST(SUM(a.total_pages) * 8.0 / 1024 AS decimal(10,1)) AS total_mb
FROM sys.tables t
JOIN sys.indexes i      ON i.object_id = t.object_id
JOIN sys.partitions p   ON p.object_id = i.object_id AND p.index_id = i.index_id
JOIN sys.allocation_units a ON a.container_id = p.partition_id
WHERE i.index_id <= 1
GROUP BY t.name, p.rows
ORDER BY total_mb DESC;
GO

/* ---- 2. missing-index suggestions (DMV; since last restart) -------
   NOTE: populated by real query activity. On a quiet test DB this may be
   sparse — absence here is NOT proof the indexes aren't needed; section 3
   is the deterministic complement. */
SELECT TOP 25
       'MISSING_INDEX' AS section,
       OBJECT_NAME(mid.object_id) AS table_name,
       CAST(migs.avg_user_impact AS decimal(5,1)) AS avg_pct_improvement,
       migs.user_seeks + migs.user_scans AS uses,
       mid.equality_columns,
       mid.inequality_columns,
       mid.included_columns
FROM sys.dm_db_missing_index_details mid
JOIN sys.dm_db_missing_index_groups mig
     ON mig.index_handle = mid.index_handle
JOIN sys.dm_db_missing_index_group_stats migs
     ON migs.group_handle = mig.index_group_handle
WHERE mid.database_id = DB_ID()
ORDER BY migs.avg_user_impact * (migs.user_seeks + migs.user_scans) DESC;
GO

/* ---- 3. FK-like (_DPA_) columns that are NOT the leading key of any
          index -> prime non-clustered index candidates for joins ----- */
SELECT 'UNINDEXED_FK_COL' AS section,
       t.name AS table_name,
       c.name AS column_name
FROM sys.columns c
JOIN sys.tables t ON t.object_id = c.object_id
WHERE c.name LIKE '%[_]DPA[_]%'
  AND NOT EXISTS (
        SELECT 1 FROM sys.index_columns ic
        WHERE ic.object_id = c.object_id
          AND ic.column_id = c.column_id
          AND ic.key_ordinal = 1)
  AND t.name IN ('Client','tbOrder','OrdDetail','Lot','Contract','LevyContract',
                 'Payment','Journal','JournalEntry','Security','Agent','Broker',
                 'Holdings','LevySecurity','Voucher')
ORDER BY t.name, c.name;
GO

/* ---- 3b. total count of unindexed _DPA_ columns across ALL tables -- */
SELECT 'UNINDEXED_FK_TOTAL' AS section, COUNT(*) AS unindexed_dpa_columns
FROM sys.columns c
JOIN sys.tables t ON t.object_id = c.object_id
WHERE c.name LIKE '%[_]DPA[_]%'
  AND NOT EXISTS (
        SELECT 1 FROM sys.index_columns ic
        WHERE ic.object_id = c.object_id
          AND ic.column_id = c.column_id
          AND ic.key_ordinal = 1);
GO

/* ---- 4. tables with NO primary key (integrity smell) --------------- */
SELECT 'NO_PRIMARY_KEY' AS section, t.name AS table_name, p.rows AS row_count
FROM sys.tables t
JOIN sys.partitions p ON p.object_id = t.object_id AND p.index_id IN (0,1)
WHERE NOT EXISTS (
        SELECT 1 FROM sys.indexes i
        WHERE i.object_id = t.object_id AND i.is_primary_key = 1)
ORDER BY p.rows DESC;
GO

/* ---- 5. existing FK inventory -------------------------------------- */
SELECT 'EXISTING_FK' AS section,
       OBJECT_NAME(fk.parent_object_id)     AS child_table,
       OBJECT_NAME(fk.referenced_object_id) AS parent_table,
       fk.name AS fk_name,
       fk.is_disabled,
       fk.is_not_trusted
FROM sys.foreign_keys fk
ORDER BY child_table;
GO

/* =====================================================================
   6. ORPHAN / INTEGRITY AUDIT (core money path)
   Catalog-driven: a child column "X_DPA_" is matched to the parent table
   whose PRIMARY-KEY leading column is also "X_DPA_" (the legacy naming
   convention). We then count child rows whose non-NULL FK value has NO
   matching parent row. Joining on the PK alone means a soft-deleted
   parent still satisfies the join, so only TRUE orphans are flagged.
   Entity_DPA_ is excluded (polymorphic — resolved via EntityType_DPA_).
   ===================================================================== */
IF OBJECT_ID('tempdb..#pk')      IS NOT NULL DROP TABLE #pk;
IF OBJECT_ID('tempdb..#pairs')   IS NOT NULL DROP TABLE #pairs;
IF OBJECT_ID('tempdb..#orphans') IS NOT NULL DROP TABLE #orphans;

-- leading PK column per table
SELECT t.object_id, t.name AS table_name, c.name AS pk_col
INTO #pk
FROM sys.tables t
JOIN sys.indexes i        ON i.object_id = t.object_id AND i.is_primary_key = 1
JOIN sys.index_columns ic ON ic.object_id = i.object_id AND ic.index_id = i.index_id AND ic.key_ordinal = 1
JOIN sys.columns c        ON c.object_id = t.object_id AND c.column_id = ic.column_id;

-- child(col) -> parent(pk) candidate relationships on the core path
SELECT ct.name AS child_table, cc.name AS child_col,
       pk.table_name AS parent_table, pk.pk_col AS parent_pk
INTO #pairs
FROM sys.columns cc
JOIN sys.tables ct ON ct.object_id = cc.object_id
JOIN #pk pk        ON pk.pk_col = cc.name AND pk.object_id <> ct.object_id
WHERE cc.name LIKE '%[_]DPA[_]%'
  AND cc.name <> 'Entity_DPA_'                 -- polymorphic, skip
  AND ct.name IN ('Client','tbOrder','OrdDetail','Lot','Contract','LevyContract',
                  'Payment','Journal','JournalEntry','Security','Agent','Broker',
                  'Holdings','LevySecurity','Voucher');

CREATE TABLE #orphans (
    child_table sysname, child_column sysname,
    parent_table sysname, parent_pk sysname,
    nonnull_fk_rows int, orphan_rows int);

DECLARE @ct sysname, @cc sysname, @pt sysname, @pp sysname, @sql nvarchar(max);
DECLARE c CURSOR LOCAL FAST_FORWARD FOR SELECT * FROM #pairs;
OPEN c;
FETCH NEXT FROM c INTO @ct, @cc, @pt, @pp;
WHILE @@FETCH_STATUS = 0
BEGIN
    SET @sql = N'
      INSERT INTO #orphans
      SELECT ' + QUOTENAME(@ct,'''') + N',' + QUOTENAME(@cc,'''') + N','
               + QUOTENAME(@pt,'''') + N',' + QUOTENAME(@pp,'''') + N',
             COUNT(*),
             SUM(CASE WHEN p.' + QUOTENAME(@pp) + N' IS NULL THEN 1 ELSE 0 END)
      FROM ' + QUOTENAME(@ct) + N' ch
      LEFT JOIN ' + QUOTENAME(@pt) + N' p
             ON p.' + QUOTENAME(@pp) + N' = ch.' + QUOTENAME(@cc) + N'
      WHERE ch.' + QUOTENAME(@cc) + N' IS NOT NULL;';
    BEGIN TRY
        EXEC sp_executesql @sql;
    END TRY
    BEGIN CATCH
        INSERT INTO #orphans VALUES (@ct, @cc, @pt, @pp, -1, -1); -- -1 = check errored
    END CATCH
    FETCH NEXT FROM c INTO @ct, @cc, @pt, @pp;
END
CLOSE c; DEALLOCATE c;

SELECT 'ORPHAN_AUDIT' AS section, child_table, child_column,
       parent_table, parent_pk, nonnull_fk_rows, orphan_rows
FROM #orphans
ORDER BY orphan_rows DESC, child_table, child_column;

DROP TABLE #pk; DROP TABLE #pairs; DROP TABLE #orphans;
GO
