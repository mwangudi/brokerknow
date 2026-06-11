/* PROD: for each duplicated Payment_DPA_, how many LIVE (Deleted=0/null) vs
   soft-deleted rows. Tells us if the test cleanup logic (delete only the
   soft-deleted dupes, keep the single live one) applies cleanly. READ-ONLY. */
SET NOCOUNT ON;
USE BrokerKnow;
GO
;WITH d AS (
    SELECT Payment_DPA_ FROM dbo.Payment
    WHERE Payment_DPA_ IS NOT NULL GROUP BY Payment_DPA_ HAVING COUNT(*)>1)
SELECT p.Payment_DPA_,
       COUNT(*) AS total_rows,
       SUM(CASE WHEN p.Deleted = 1 THEN 1 ELSE 0 END) AS soft_deleted,
       SUM(CASE WHEN p.Deleted = 0 OR p.Deleted IS NULL THEN 1 ELSE 0 END) AS live_rows
FROM dbo.Payment p JOIN d ON d.Payment_DPA_ = p.Payment_DPA_
GROUP BY p.Payment_DPA_
ORDER BY p.Payment_DPA_;

PRINT '===== overall: would the "delete soft-deleted dupes" plan leave exactly 1 row per key? =====';
;WITH d AS (
    SELECT Payment_DPA_ FROM dbo.Payment
    WHERE Payment_DPA_ IS NOT NULL GROUP BY Payment_DPA_ HAVING COUNT(*)>1),
g AS (
    SELECT p.Payment_DPA_,
           SUM(CASE WHEN p.Deleted = 0 OR p.Deleted IS NULL THEN 1 ELSE 0 END) AS live_rows
    FROM dbo.Payment p JOIN d ON d.Payment_DPA_ = p.Payment_DPA_
    GROUP BY p.Payment_DPA_)
SELECT
  SUM(CASE WHEN live_rows = 1 THEN 1 ELSE 0 END) AS keys_with_exactly_1_live,
  SUM(CASE WHEN live_rows = 0 THEN 1 ELSE 0 END) AS keys_with_0_live_NEEDS_REVIEW,
  SUM(CASE WHEN live_rows > 1 THEN 1 ELSE 0 END) AS keys_with_multiple_live_NEEDS_REVIEW
FROM g;
GO
