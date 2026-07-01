SET NOCOUNT ON;
-- Money control totals: fresh source (Malawi0701) vs migrated clean schema.
-- Raw table sums (all rows) => must match to the cent.
SELECT item,
       CONVERT(varchar(40), src, 1)   AS source_Malawi0701,
       CONVERT(varchar(40), clean, 1) AS clean_schema,
       CASE WHEN src = clean THEN 'MATCH' ELSE 'DIFF ***' END AS status
FROM (
    SELECT 'Payment.PaymentAmount' AS item,
           (SELECT SUM(PaymentAmount) FROM BrokerKnow_Malawi0701.dbo.Payment) AS src,
           (SELECT SUM(PaymentAmount) FROM BrokerKnow_Clean.dbo.Payment)      AS clean
    UNION ALL SELECT 'ClientBalances.CurrentBal',
           (SELECT SUM(CurrentBal) FROM BrokerKnow_Malawi0701.dbo.ClientBalances),
           (SELECT SUM(CurrentBal) FROM BrokerKnow_Clean.dbo.ClientBalances)
    UNION ALL SELECT 'Lot gross (qty*price)',
           (SELECT SUM(CAST(LotQty AS decimal(38,4)) * CAST(LotPrice AS decimal(38,6))) FROM BrokerKnow_Malawi0701.dbo.Lot),
           (SELECT SUM(CAST(LotQty AS decimal(38,4)) * CAST(LotPrice AS decimal(38,6))) FROM BrokerKnow_Clean.dbo.Lot)
    UNION ALL SELECT 'LevyContract.LevyAmount',
           (SELECT SUM(LevyAmount) FROM BrokerKnow_Malawi0701.dbo.LevyContract),
           (SELECT SUM(LevyAmount) FROM BrokerKnow_Clean.dbo.LevyContract)
) x;
