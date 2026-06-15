SET NOCOUNT ON;
USE BrokerKnow_KE_Clean;

PRINT '=== Levy rows the engine relies on (SystemMaintained codes 11,12,25,99,100,101) ===';
SELECT Levy_DPA_, SystemMaintained, LevyDescription, LevyShortName, LevyType,
       LevyAmount, LevyAppSecurity, LevyAppBond, Vatable
FROM dbo.Levy
WHERE SystemMaintained IN (11,12,25,99,100,101)
ORDER BY SystemMaintained;

PRINT '=== ALL distinct SystemMaintained codes present in KE Levy data ===';
SELECT SystemMaintained, COUNT(*) AS n, MIN(LevyDescription) AS sample_desc
FROM dbo.Levy
GROUP BY SystemMaintained
ORDER BY SystemMaintained;

PRINT '=== Total Levy + LevySecurity + Commission rows migrated ===';
SELECT
  (SELECT COUNT(*) FROM dbo.Levy)        AS Levies,
  (SELECT COUNT(*) FROM dbo.LevySecurity) AS LevySecurity,
  (SELECT COUNT(*) FROM dbo.Commission)   AS Commissions;

PRINT '=== Sample of KE Commission structures (3-tier rates) ===';
SELECT TOP 5 Commission_DPA_, CommissionDescription, CommissionRate, CommissionMedian, CommissionUpper
FROM dbo.Commission ORDER BY Commission_DPA_;

PRINT '=== KE Branches (settlement/office context) ===';
SELECT Branch_DPA_, BranchDescription, DefaultSelection FROM dbo.Branch ORDER BY Branch_DPA_;
