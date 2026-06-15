SET NOCOUNT ON;
-- Does the raw Malawi dump have the tables/objects the NEW app needs?
-- If these are missing, the app's logins + screens break against it.
SELECT 'PortalUsers'      AS object_needed, CASE WHEN OBJECT_ID('BrokerKnow_Malawi0612.dbo.PortalUsers') IS NULL THEN 'MISSING' ELSE 'present' END AS malawi,
                                            CASE WHEN OBJECT_ID('BrokerKnow.dbo.PortalUsers') IS NULL THEN 'MISSING' ELSE 'present' END AS live
UNION ALL SELECT 'PortalRefreshTokens', CASE WHEN OBJECT_ID('BrokerKnow_Malawi0612.dbo.PortalRefreshTokens') IS NULL THEN 'MISSING' ELSE 'present' END, CASE WHEN OBJECT_ID('BrokerKnow.dbo.PortalRefreshTokens') IS NULL THEN 'MISSING' ELSE 'present' END
UNION ALL SELECT 'UserPageAccess',      CASE WHEN OBJECT_ID('BrokerKnow_Malawi0612.dbo.UserPageAccess') IS NULL THEN 'MISSING' ELSE 'present' END, CASE WHEN OBJECT_ID('BrokerKnow.dbo.UserPageAccess') IS NULL THEN 'MISSING' ELSE 'present' END
UNION ALL SELECT 'OrderHoldOptions',    CASE WHEN OBJECT_ID('BrokerKnow_Malawi0612.dbo.OrderHoldOptions') IS NULL THEN 'MISSING' ELSE 'present' END, CASE WHEN OBJECT_ID('BrokerKnow.dbo.OrderHoldOptions') IS NULL THEN 'MISSING' ELSE 'present' END
UNION ALL SELECT 'app.Clients (view)',  CASE WHEN OBJECT_ID('BrokerKnow_Malawi0612.app.Clients') IS NULL THEN 'MISSING' ELSE 'present' END, CASE WHEN OBJECT_ID('BrokerKnow.app.Clients') IS NULL THEN 'MISSING' ELSE 'present' END
UNION ALL SELECT 'PriceImportBatches',  CASE WHEN OBJECT_ID('BrokerKnow_Malawi0612.dbo.PriceImportBatches') IS NULL THEN 'MISSING' ELSE 'present' END, CASE WHEN OBJECT_ID('BrokerKnow.dbo.PriceImportBatches') IS NULL THEN 'MISSING' ELSE 'present' END;
GO
-- How many portal logins exist in each (0 in Malawi = nobody can sign in)
SELECT 'PortalUsers count' AS metric,
       CASE WHEN OBJECT_ID('BrokerKnow_Malawi0612.dbo.PortalUsers') IS NULL THEN -1
            ELSE (SELECT COUNT(*) FROM BrokerKnow_Malawi0612.dbo.PortalUsers) END AS malawi_dump,
       (SELECT COUNT(*) FROM BrokerKnow.dbo.PortalUsers) AS live;
GO
