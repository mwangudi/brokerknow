/* READ-ONLY: key uniqueness for the no-PK real tables, to see which can take
   a safe additive PK (Phase-2 style). */
SET NOCOUNT ON;
USE BrokerKnow_Test;
GO
SELECT 'BankAcc' AS tbl, 'BankAcc_DPA_' AS keycol,
  (SELECT COUNT(*) FROM dbo.BankAcc) AS rows_total,
  (SELECT COUNT(*) FROM dbo.BankAcc WHERE BankAcc_DPA_ IS NULL) AS null_keys,
  (SELECT COUNT(*) FROM (SELECT BankAcc_DPA_ FROM dbo.BankAcc WHERE BankAcc_DPA_ IS NOT NULL GROUP BY BankAcc_DPA_ HAVING COUNT(*)>1) d) AS dup_keys;
SELECT 'ClientBalances' AS tbl, 'client_DPA_' AS keycol,
  (SELECT COUNT(*) FROM dbo.ClientBalances) AS rows_total,
  (SELECT COUNT(*) FROM dbo.ClientBalances WHERE client_DPA_ IS NULL) AS null_keys,
  (SELECT COUNT(*) FROM (SELECT client_DPA_ FROM dbo.ClientBalances WHERE client_DPA_ IS NOT NULL GROUP BY client_DPA_ HAVING COUNT(*)>1) d) AS dup_keys;
GO
