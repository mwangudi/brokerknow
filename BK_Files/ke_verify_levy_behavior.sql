SET NOCOUNT ON;
USE BrokerKnow_KE_Clean;

PRINT '=== Do the engine reserved codes (11,12,25,99,100,101) exist in KE Levy? (0 = absent) ===';
SELECT SystemMaintained, COUNT(*) AS present
FROM dbo.Levy WHERE SystemMaintained IN (11,12,25,99,100,101)
GROUP BY SystemMaintained;

PRINT '=== KE levies: active flag + apply scope (drives the user-defined-levy path) ===';
SELECT SystemMaintained AS SM, LevyDescription, LevyShortName, LevyType,
       LevyAmount, LevyBlock, LevyAppSecurity, LevyAppBond, Vatable
FROM dbo.Levy ORDER BY SystemMaintained;

PRINT '=== Historical LevyContract rows migrated? (so existing KE contract notes render) ===';
SELECT COUNT(*) AS levycontract_rows,
       COUNT(DISTINCT Contract_DPA_) AS distinct_contracts
FROM dbo.LevyContract;

PRINT '=== Distinct levy descriptions actually used on historical KE contracts ===';
SELECT TOP 20 lc.LevyDescription, COUNT(*) AS n
FROM dbo.LevyContract lc
GROUP BY lc.LevyDescription
ORDER BY COUNT(*) DESC;
