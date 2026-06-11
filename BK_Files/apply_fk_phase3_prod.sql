/* =====================================================================
   Phase-3 FOREIGN KEYS ? the 20 proven-safe relationships (0 orphans).
   Target: BrokerKnow_Test (:5261). TEST ONLY ? NOT prod.

   Design (legacy-write-safe):
     * NO CASCADE ? ON DELETE/UPDATE NO ACTION (default). An FK must never
       trigger cascading deletes the legacy app or our soft-delete model
       don't expect; it only blocks a hard-delete that WOULD orphan children.
     * WITH CHECK ? validates existing rows (all 0 orphans, so passes) and
       marks the FK "trusted" so the query optimiser can use it.
   Idempotent: skips an FK that already exists. Reversible:
     ALTER TABLE dbo.<child> DROP CONSTRAINT <fk_name>;
   EXCLUDED (reported separately): tbOrder.Agent_DPA_ (0-sentinel),
   LevyContract.Contract_DPA_ (5 orphans), JournalEntry.Journal_DPA_ (1).
   ===================================================================== */
SET NOCOUNT ON;
USE BrokerKnow;
GO

IF OBJECT_ID('tempdb..#fk') IS NOT NULL DROP TABLE #fk;
CREATE TABLE #fk (name sysname, child sysname, col sysname, parent sysname, pk sysname);
INSERT INTO #fk VALUES
 ('FK_Client_Branch','Client','Branch_DPA_','Branch','Branch_DPA_'),
 ('FK_Client_Class','Client','Class_DPA_','Class','Class_DPA_'),
 ('FK_Client_Commission','Client','Commission_DPA_','Commission','Commission_DPA_'),
 ('FK_Client_Residency','Client','Residency_DPA_','Residency','Residency_DPA_'),
 ('FK_Client_Gender','Client','Gender_DPA_','Gender','Gender_DPA_'),
 ('FK_Client_Agent','Client','Agent_DPA_','Agent','Agent_DPA_'),
 ('FK_Client_EntityType','Client','EntityType_DPA_','EntityType','EntityType_DPA_'),
 ('FK_tbOrder_Client','tbOrder','Client_DPA_','Client','Client_DPA_'),
 ('FK_tbOrder_Branch','tbOrder','Branch_DPA_','Branch','Branch_DPA_'),
 ('FK_tbOrder_OrderType','tbOrder','OrderType_DPA_','OrderType','OrderType_DPA_'),
 ('FK_tbOrder_OrderSecType','tbOrder','OrderSecType_DPA_','OrderSecType','OrderSecType_DPA_'),
 ('FK_tbOrder_OrderHoldType','tbOrder','OrderHoldType_DPA_','OrderHoldType','OrderHoldType_DPA_'),
 ('FK_OrdDetail_tbOrder','OrdDetail','Order_DPA_','tbOrder','Order_DPA_'),
 ('FK_OrdDetail_Security','OrdDetail','Security_DPA_','Security','Security_DPA_'),
 ('FK_Lot_OrdDetail','Lot','OrdDetail_DPA_','OrdDetail','OrdDetail_DPA_'),
 ('FK_Lot_Contract','Lot','Contract_DPA_','Contract','Contract_DPA_'),
 ('FK_Lot_Broker','Lot','Broker_DPA_','Broker','Broker_DPA_'),
 ('FK_Payment_EntityType','Payment','EntityType_DPA_','EntityType','EntityType_DPA_'),
 ('FK_Payment_PayType','Payment','PayType_DPA_','PayType','PayType_DPA_'),
 ('FK_Contract_Status','Contract','Status_DPA_','Status','Status_DPA_');

DECLARE @name sysname,@ch sysname,@co sysname,@pt sysname,@pk sysname,@sql nvarchar(max);
DECLARE c CURSOR LOCAL FAST_FORWARD FOR SELECT name,child,col,parent,pk FROM #fk;
OPEN c; FETCH NEXT FROM c INTO @name,@ch,@co,@pt,@pk;
WHILE @@FETCH_STATUS=0
BEGIN
  IF EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name=@name)
    PRINT 'exists  ' + @name;
  ELSE
  BEGIN
    SET @sql = N'ALTER TABLE dbo.' + QUOTENAME(@ch) + N' WITH CHECK ADD CONSTRAINT ' + QUOTENAME(@name)
      + N' FOREIGN KEY (' + QUOTENAME(@co) + N') REFERENCES dbo.' + QUOTENAME(@pt) + N'(' + QUOTENAME(@pk) + N');';
    BEGIN TRY EXEC sp_executesql @sql; PRINT 'created ' + @name; END TRY
    BEGIN CATCH PRINT 'FAILED  ' + @name + '  -> ' + ERROR_MESSAGE(); END CATCH
  END
  FETCH NEXT FROM c INTO @name,@ch,@co,@pt,@pk;
END
CLOSE c; DEALLOCATE c;
DROP TABLE #fk;
GO

PRINT '===== FK inventory now (excluding the 2 pre-existing app FKs) =====';
SELECT OBJECT_NAME(fk.parent_object_id) AS child, fk.name,
       OBJECT_NAME(fk.referenced_object_id) AS parent,
       fk.is_not_trusted, fk.delete_referential_action_desc AS on_delete
FROM sys.foreign_keys fk
WHERE fk.name LIKE 'FK[_]%'
ORDER BY child, fk.name;
SELECT COUNT(*) AS total_fks_in_db FROM sys.foreign_keys;
GO
