/* =====================================================================
   Phase-3 FK candidate analysis — READ-ONLY. Target: BrokerKnow_Test.
   For each child->parent relationship we want to enforce, report:
     - child column nullability
     - non-null row count
     - zero_rows         : rows where the FK value = 0 (legacy "none" sentinel)
     - parent_has_zero   : does the parent PK actually have a 0 row?
     - orphans_excl_zero : non-null, non-zero values with NO parent match
     - orphans_incl_zero : non-null values (incl 0) with NO parent match
   SAFE TO ENFORCE when orphans_incl_zero = 0 (every value, including any
   sentinel, resolves to a real parent). A column with 0-sentinels but no
   parent 0-row shows orphans_incl_zero>0 -> SKIP or handle the sentinel.
   No writes; only SELECTs into a #temp it drops at the end.
   ===================================================================== */
SET NOCOUNT ON;
USE BrokerKnow_Test;
GO

IF OBJECT_ID('tempdb..#cand') IS NOT NULL DROP TABLE #cand;
CREATE TABLE #cand (child sysname, col sysname, parent sysname, pk sysname);
INSERT INTO #cand VALUES
 ('Client','Branch_DPA_','Branch','Branch_DPA_'),
 ('Client','Class_DPA_','Class','Class_DPA_'),
 ('Client','Commission_DPA_','Commission','Commission_DPA_'),
 ('Client','Residency_DPA_','Residency','Residency_DPA_'),
 ('Client','Gender_DPA_','Gender','Gender_DPA_'),
 ('Client','Agent_DPA_','Agent','Agent_DPA_'),
 ('Client','EntityType_DPA_','EntityType','EntityType_DPA_'),
 ('tbOrder','Client_DPA_','Client','Client_DPA_'),
 ('tbOrder','Branch_DPA_','Branch','Branch_DPA_'),
 ('tbOrder','OrderType_DPA_','OrderType','OrderType_DPA_'),
 ('tbOrder','OrderSecType_DPA_','OrderSecType','OrderSecType_DPA_'),
 ('tbOrder','OrderHoldType_DPA_','OrderHoldType','OrderHoldType_DPA_'),
 ('tbOrder','Agent_DPA_','Agent','Agent_DPA_'),
 ('OrdDetail','Order_DPA_','tbOrder','Order_DPA_'),
 ('OrdDetail','Security_DPA_','Security','Security_DPA_'),
 ('Lot','OrdDetail_DPA_','OrdDetail','OrdDetail_DPA_'),
 ('Lot','Contract_DPA_','Contract','Contract_DPA_'),
 ('Lot','Broker_DPA_','Broker','Broker_DPA_'),
 ('LevyContract','Contract_DPA_','Contract','Contract_DPA_'),
 ('JournalEntry','Journal_DPA_','Journal','Journal_DPA_'),
 ('Payment','PayType_DPA_','PayType','PayType_DPA_'),
 ('Payment','EntityType_DPA_','EntityType','EntityType_DPA_'),
 ('Contract','Status_DPA_','Status','Status_DPA_');

CREATE TABLE #res (
  child sysname, col sysname, parent sysname,
  child_nullable bit, nonnull_rows int, zero_rows int,
  parent_has_zero bit, orphans_excl_zero int, orphans_incl_zero int);

DECLARE @ch sysname, @co sysname, @pt sysname, @pk sysname, @sql nvarchar(max);
DECLARE c CURSOR LOCAL FAST_FORWARD FOR SELECT child,col,parent,pk FROM #cand;
OPEN c; FETCH NEXT FROM c INTO @ch,@co,@pt,@pk;
WHILE @@FETCH_STATUS = 0
BEGIN
  SET @sql = N'
    INSERT INTO #res
    SELECT ' + QUOTENAME(@ch,'''') + N',' + QUOTENAME(@co,'''') + N',' + QUOTENAME(@pt,'''') + N',
      (SELECT c.is_nullable FROM sys.columns c WHERE c.object_id=OBJECT_ID(''dbo.'+@ch+N''') AND c.name='''+@co+N'''),
      (SELECT COUNT(*) FROM dbo.' + QUOTENAME(@ch) + N' WHERE ' + QUOTENAME(@co) + N' IS NOT NULL),
      (SELECT COUNT(*) FROM dbo.' + QUOTENAME(@ch) + N' WHERE ' + QUOTENAME(@co) + N' = 0),
      (CASE WHEN EXISTS (SELECT 1 FROM dbo.' + QUOTENAME(@pt) + N' WHERE ' + QUOTENAME(@pk) + N' = 0) THEN 1 ELSE 0 END),
      (SELECT COUNT(*) FROM dbo.' + QUOTENAME(@ch) + N' ch WHERE ch.' + QUOTENAME(@co) + N' IS NOT NULL AND ch.' + QUOTENAME(@co) + N' <> 0
          AND NOT EXISTS (SELECT 1 FROM dbo.' + QUOTENAME(@pt) + N' p WHERE p.' + QUOTENAME(@pk) + N' = ch.' + QUOTENAME(@co) + N')),
      (SELECT COUNT(*) FROM dbo.' + QUOTENAME(@ch) + N' ch WHERE ch.' + QUOTENAME(@co) + N' IS NOT NULL
          AND NOT EXISTS (SELECT 1 FROM dbo.' + QUOTENAME(@pt) + N' p WHERE p.' + QUOTENAME(@pk) + N' = ch.' + QUOTENAME(@co) + N'));';
  BEGIN TRY EXEC sp_executesql @sql; END TRY
  BEGIN CATCH INSERT INTO #res VALUES (@ch,@co,@pt,NULL,-1,-1,-1,-1,-1); END CATCH
  FETCH NEXT FROM c INTO @ch,@co,@pt,@pk;
END
CLOSE c; DEALLOCATE c;

SELECT child, col, parent, child_nullable AS nz, nonnull_rows, zero_rows,
       parent_has_zero AS pz, orphans_excl_zero AS orph_xz, orphans_incl_zero AS orph_iz,
       CASE WHEN orphans_incl_zero = 0 THEN 'SAFE'
            WHEN orphans_excl_zero = 0 AND zero_rows > 0 THEN 'SENTINEL-ONLY'
            ELSE 'HAS-ORPHANS' END AS verdict
FROM #res
ORDER BY verdict, child, col;

DROP TABLE #cand; DROP TABLE #res;
GO
