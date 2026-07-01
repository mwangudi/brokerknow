/* Correct legacy data-entry swaps in dbo.BankAcc where the ACCOUNT NAME field
   holds a pure account NUMBER and the ACCOUNT NUMBER field holds the holder's
   NAME. Snapshots the affected rows first, swaps inside a guarded transaction.
   Only touches unambiguous swaps: name is all digits AND number contains letters. */
SET NOCOUNT ON;
SET XACT_ABORT ON;
SET QUOTED_IDENTIFIER ON;
SET ANSI_NULLS ON;

IF OBJECT_ID('dbo.BankAcc_swapfix_backup_20260701') IS NULL
    SELECT BankAcc_DPA_, BankAccName, BankAccNumber
    INTO dbo.BankAcc_swapfix_backup_20260701
    FROM dbo.BankAcc
    WHERE BankAccName NOT LIKE '%[^0-9]%' AND LTRIM(RTRIM(BankAccName)) <> ''
      AND BankAccNumber LIKE '%[A-Za-z]%';

DECLARE @n int = (SELECT COUNT(*) FROM dbo.BankAcc_swapfix_backup_20260701);

BEGIN TRAN;
UPDATE b
    SET BankAccName   = b.BankAccNumber,   -- RHS uses pre-update values (T-SQL swap)
        BankAccNumber = b.BankAccName
FROM dbo.BankAcc b
WHERE b.BankAccName NOT LIKE '%[^0-9]%' AND LTRIM(RTRIM(b.BankAccName)) <> ''
  AND b.BankAccNumber LIKE '%[A-Za-z]%';
DECLARE @updated int = @@ROWCOUNT;

PRINT CONCAT('Expected: ', @n, '  Updated: ', @updated);
IF @updated = @n AND @n > 0
BEGIN
    COMMIT;
    PRINT 'COMMITTED. Backup: dbo.BankAcc_swapfix_backup_20260701';
END
ELSE
BEGIN
    ROLLBACK;
    PRINT 'ROLLED BACK (count mismatch).';
END

-- Show the flagged client after the fix.
SELECT b.BankAcc_DPA_, c.ClientName, b.BankAccName, b.BankAccNumber
FROM dbo.BankAcc b LEFT JOIN dbo.Client c ON c.Client_DPA_ = b.Client_DPA_
WHERE b.Client_DPA_ = 5359;
