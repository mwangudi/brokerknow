SET NOCOUNT ON;
-- Compare the freshly-restored Malawi dump vs the LIVE prod BrokerKnow.
-- Read-only. Answers "are the orders/contracts/clients the same?".
SELECT m.metric,
       m.malawi_dump            AS [Malawi120626_02],
       l.live_brokerknow        AS [live BrokerKnow],
       (m.malawi_dump - l.live_brokerknow) AS [dump - live]
FROM (
    SELECT 'Clients (live)'  AS metric, 1 AS k, (SELECT COUNT(*) FROM BrokerKnow_Malawi0612.dbo.Client   WHERE ISNULL(Deleted,0)=0) AS malawi_dump
    UNION ALL SELECT 'Orders (tbOrder)',    2, (SELECT COUNT(*) FROM BrokerKnow_Malawi0612.dbo.tbOrder)
    UNION ALL SELECT 'Order details',       3, (SELECT COUNT(*) FROM BrokerKnow_Malawi0612.dbo.OrdDetail)
    UNION ALL SELECT 'Contracts',           4, (SELECT COUNT(*) FROM BrokerKnow_Malawi0612.dbo.Contract)
    UNION ALL SELECT 'Lots',                5, (SELECT COUNT(*) FROM BrokerKnow_Malawi0612.dbo.Lot)
    UNION ALL SELECT 'Payments',            6, (SELECT COUNT(*) FROM BrokerKnow_Malawi0612.dbo.Payment)
    UNION ALL SELECT 'Securities',          7, (SELECT COUNT(*) FROM BrokerKnow_Malawi0612.dbo.Security)
) m
JOIN (
    SELECT 1 AS k, (SELECT COUNT(*) FROM BrokerKnow.dbo.Client   WHERE ISNULL(Deleted,0)=0) AS live_brokerknow
    UNION ALL SELECT 2, (SELECT COUNT(*) FROM BrokerKnow.dbo.tbOrder)
    UNION ALL SELECT 3, (SELECT COUNT(*) FROM BrokerKnow.dbo.OrdDetail)
    UNION ALL SELECT 4, (SELECT COUNT(*) FROM BrokerKnow.dbo.Contract)
    UNION ALL SELECT 5, (SELECT COUNT(*) FROM BrokerKnow.dbo.Lot)
    UNION ALL SELECT 6, (SELECT COUNT(*) FROM BrokerKnow.dbo.Payment)
    UNION ALL SELECT 7, (SELECT COUNT(*) FROM BrokerKnow.dbo.Security)
) l ON l.k = m.k
ORDER BY m.k;
GO
-- Money control totals (do the financials match?)
SELECT 'Payment sum'  AS metric,
       (SELECT CAST(SUM(PaymentAmount) AS decimal(20,2)) FROM BrokerKnow_Malawi0612.dbo.Payment) AS malawi_dump,
       (SELECT CAST(SUM(PaymentAmount) AS decimal(20,2)) FROM BrokerKnow.dbo.Payment)            AS live_brokerknow
UNION ALL
SELECT 'Newest order date',
       (SELECT CAST(MAX(OrderDate) AS varchar) FROM BrokerKnow_Malawi0612.dbo.tbOrder) COLLATE DATABASE_DEFAULT,
       (SELECT CAST(MAX(OrderDate) AS varchar) FROM BrokerKnow.dbo.tbOrder) COLLATE DATABASE_DEFAULT;
GO
