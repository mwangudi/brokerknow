CREATE VIEW dbo.ClientStatement
AS
SELECT     TOP 100 PERCENT COUNT(*) AS ClientTransaction_DPA_, a.Client_DPA_, a.TransDate, a.Ref, a.Particulars, a.Debit, a.Credit, 
                      CASE a.IsOpeningBalance WHEN 1 THEN CASE WHEN a.Balance >= 0 THEN CONVERT(NVARCHAR(400), a.Balance) 
                      + ' Cr' ELSE CONVERT(NVARCHAR(400), a.Balance) + ' Dr' END ELSE CASE WHEN SUM(b.Balance) >= 0 THEN CONVERT(NVARCHAR(400), 
                      SUM(b.Balance)) + ' Cr' ELSE CONVERT(NVARCHAR(400), ABS(SUM(b.Balance))) + ' Dr' END END AS Balance, a.IsOpeningBalance
FROM         (SELECT     TOP 100 PERCENT *
                       FROM          dbo.ClientTransactionList) a CROSS JOIN
                          (SELECT     TOP 100 PERCENT *
                            FROM          dbo.ClientTransactionList) b
WHERE     a.TransDate >= b.TransDate AND a.Client_DPA_ = b.Client_DPA_
GROUP BY a.Client_DPA_, a.TransDate, a.Ref, a.Particulars, a.Debit, a.Credit, a.Balance, a.IsOpeningBalance
ORDER BY a.Client_DPA_, a.IsOpeningBalance DESC, a.ClientTransaction_DPA_, a.Particulars, a.Ref, a.Debit, a.Credit, a.Balance


(1 rows affected)
