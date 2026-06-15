SET NOCOUNT ON;
USE BrokerKnow_KE_Clean;

PRINT '=== FULL KE Levy table (Active flag is the key) ===';
SELECT Levy_DPA_ AS Id, SystemMaintained AS SM, LevyDescription, LevyShortName,
       LevyType AS Typ, LevyAmount AS Rate, LevyActive AS Active,
       LevyAppSecurity AS Equity, LevyAppBond AS Bond, LevyBlock AS Blk, Vatable
FROM dbo.Levy
ORDER BY LevyDescription, LevyAmount DESC;

PRINT '=== Which levies FIRE on a new EQUITY trade? (engine InsertUserDefinedLevies path) ===';
-- LevyActive=1 AND SystemMaintained NOT IN reserved(11,12,25,99,100,101) AND LevyAppSecurity=1
SELECT Levy_DPA_ AS Id, SystemMaintained AS SM, LevyDescription, LevyAmount AS Rate,
       (SELECT COUNT(*) FROM dbo.LevySecurity s WHERE s.Levy_DPA_ = l.Levy_DPA_) AS scoped_secs
FROM dbo.Levy l
WHERE l.LevyActive = 1
  AND l.SystemMaintained IS NOT NULL
  AND l.SystemMaintained NOT IN (11,12,25,99,100,101)
  AND l.LevyAppSecurity = 1
ORDER BY LevyDescription, LevyAmount DESC;

PRINT '=== DOUBLE-CHARGE DETECTOR: same description, >1 ACTIVE equity levy ===';
SELECT LevyDescription, COUNT(*) AS active_equity_rows,
       STRING_AGG(CONCAT('SM', SystemMaintained, '@', LevyAmount), ' | ') AS rows_firing
FROM dbo.Levy
WHERE LevyActive = 1
  AND SystemMaintained NOT IN (11,12,25,99,100,101)
  AND LevyAppSecurity = 1
GROUP BY LevyDescription
HAVING COUNT(*) > 1;

PRINT '=== Which levies FIRE on a new BOND trade? ===';
SELECT Levy_DPA_ AS Id, SystemMaintained AS SM, LevyDescription, LevyAmount AS Rate
FROM dbo.Levy l
WHERE l.LevyActive = 1
  AND l.SystemMaintained IS NOT NULL
  AND l.SystemMaintained NOT IN (11,12,25,99,100,101)
  AND l.LevyAppBond = 1
ORDER BY LevyDescription, LevyAmount DESC;

PRINT '=== What did HISTORICAL KE contracts actually charge? (LevyName on LevyContract) ===';
SELECT TOP 25 LevyName, COUNT(*) AS times_used,
       CAST(AVG(LevyRate) AS decimal(10,5)) AS avg_rate
FROM dbo.LevyContract
WHERE Deleted = 0 OR Deleted IS NULL
GROUP BY LevyName
ORDER BY COUNT(*) DESC;
