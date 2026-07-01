/* Clear the app's CDS/price UPLOAD STAGING buffers on axis_db_prod so the client
   can share fresh files. Snapshots each to a dated *_bak_20260701 table first
   (instant restore). Does NOT touch committed Contracts/Lots/Payments/prices.
   FK-safe order: PriceImportRows before PriceImportBatches. Transactional. */
SET NOCOUNT ON;
SET XACT_ABORT ON;

DECLARE @h0 int = (SELECT COUNT(*) FROM dbo.CdsImportedHoldings);
DECLARE @t0 int = (SELECT COUNT(*) FROM dbo.CdsImportedTrades);
DECLARE @pr0 int = (SELECT COUNT(*) FROM dbo.PriceImportRows);
DECLARE @pb0 int = (SELECT COUNT(*) FROM dbo.PriceImportBatches);

BEGIN TRAN;

IF OBJECT_ID('dbo.CdsImportedHoldings_bak_20260701') IS NULL
    SELECT * INTO dbo.CdsImportedHoldings_bak_20260701 FROM dbo.CdsImportedHoldings;
IF OBJECT_ID('dbo.CdsImportedTrades_bak_20260701') IS NULL
    SELECT * INTO dbo.CdsImportedTrades_bak_20260701 FROM dbo.CdsImportedTrades;
IF OBJECT_ID('dbo.PriceImportRows_bak_20260701') IS NULL
    SELECT * INTO dbo.PriceImportRows_bak_20260701 FROM dbo.PriceImportRows;
IF OBJECT_ID('dbo.PriceImportBatches_bak_20260701') IS NULL
    SELECT * INTO dbo.PriceImportBatches_bak_20260701 FROM dbo.PriceImportBatches;

DELETE FROM dbo.CdsImportedHoldings;
DELETE FROM dbo.CdsImportedTrades;
DELETE FROM dbo.PriceImportRows;      -- child first (FK -> PriceImportBatches)
DELETE FROM dbo.PriceImportBatches;

DECLARE @h1 int = (SELECT COUNT(*) FROM dbo.CdsImportedHoldings);
DECLARE @t1 int = (SELECT COUNT(*) FROM dbo.CdsImportedTrades);
DECLARE @pr1 int = (SELECT COUNT(*) FROM dbo.PriceImportRows);
DECLARE @pb1 int = (SELECT COUNT(*) FROM dbo.PriceImportBatches);

PRINT CONCAT('Holdings  : ', @h0, ' -> ', @h1);
PRINT CONCAT('Trades    : ', @t0, ' -> ', @t1);
PRINT CONCAT('PriceRows : ', @pr0, ' -> ', @pr1);
PRINT CONCAT('PriceBatch: ', @pb0, ' -> ', @pb1);

IF (@h1=0 AND @t1=0 AND @pr1=0 AND @pb1=0)
BEGIN COMMIT; PRINT 'COMMITTED: upload staging cleared (backups kept as *_bak_20260701).'; END
ELSE BEGIN ROLLBACK; PRINT 'ROLLED BACK: staging not fully cleared.'; END
