/* Cedar (Malawi) levy/account name corrections per Cedar Feedback (2026-07-29):
   BSE -> MSE (Malawi Stock Exchange) and CGT -> WHT (Withholding Tax).
   Display-only: the levy engine keys off SystemMaintained codes, not names, so
   this never affects contract/levy computation. IDEMPOTENT.
   RE-APPLY after each Malawi import (names revert with the fresh desktop dump). */
SET NOCOUNT ON; SET XACT_ABORT ON;
BEGIN TRAN;

-- Levy-account names shown in the Accounts Statement / Chart of Accounts (dbo.Entity, EntityType 6 = Levy).
UPDATE dbo.Entity SET EntityName = 'MSE' WHERE EntityType_DPA_ = 6 AND Entity_DPA_ = 1 AND EntityName = 'BSE';
UPDATE dbo.Entity SET EntityName = 'WHT' WHERE EntityType_DPA_ = 6 AND Entity_DPA_ = 7 AND EntityName = 'CGT';

-- Levy catalogue (Administration -> Levies Setup). SystemMaintained 101 = WHT/CGT slot.
UPDATE dbo.Levy
   SET LevyDescription = 'Withholding Tax', LevyShortName = 'WHT'
 WHERE SystemMaintained = 101 AND (LevyShortName = 'CGT' OR LevyDescription LIKE 'CGT%');

COMMIT;

SELECT 'Entity' AS tbl, CAST(EntityType_DPA_ AS varchar) + '/' + CAST(Entity_DPA_ AS varchar) AS id, EntityName AS name
FROM dbo.Entity WHERE EntityType_DPA_ = 6
UNION ALL
SELECT 'Levy', CAST(Levy_DPA_ AS varchar), LevyShortName + ' - ' + LevyDescription
FROM dbo.Levy WHERE SystemMaintained IN (25, 99, 100, 101)
ORDER BY tbl, id;
