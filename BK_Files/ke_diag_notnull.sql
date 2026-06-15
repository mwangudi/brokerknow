SET NOCOUNT ON;
USE BrokerKnow_KE_Clean;
-- Show the 5 columns that blocked the load (type + nullability) so we can relax them.
SELECT t.name AS tbl, c.name AS col, ty.name AS type,
       c.max_length, c.precision, c.scale, c.is_nullable
FROM sys.columns c
JOIN sys.tables t  ON t.object_id = c.object_id
JOIN sys.types ty  ON ty.user_type_id = c.user_type_id
WHERE (t.name='Account'         AND c.name='ReconStartDate')
   OR (t.name='Client'          AND c.name='CreditLimit')
   OR (t.name='PaymentRequests' AND c.name='FirstApproval')
   OR (t.name='Share'           AND c.name='SharePDate')
   OR (t.name='WebtbOrder'      AND c.name='Action')
ORDER BY t.name, c.name;

-- How many NULLs does the KE legacy source actually have in each?
SELECT 'Client.CreditLimit'  AS src_col, COUNT(*) AS null_rows FROM BrokerKnow_KE_Legacy.dbo.Client          WHERE CreditLimit   IS NULL
UNION ALL SELECT 'Account.ReconStartDate',          COUNT(*) FROM BrokerKnow_KE_Legacy.dbo.Account          WHERE ReconStartDate IS NULL
UNION ALL SELECT 'PaymentRequests.FirstApproval',   COUNT(*) FROM BrokerKnow_KE_Legacy.dbo.PaymentRequests   WHERE FirstApproval  IS NULL
UNION ALL SELECT 'Share.SharePDate',                COUNT(*) FROM BrokerKnow_KE_Legacy.dbo.Share            WHERE SharePDate     IS NULL
UNION ALL SELECT 'WebtbOrder.Action',               COUNT(*) FROM BrokerKnow_KE_Legacy.dbo.WebtbOrder       WHERE [Action]       IS NULL;
