/* rebrand_ke_emails.sql — rename the live Green Margin Capital demo logins
   from the old axis-kenya placeholder domain to greenmargin.demo.
   Idempotent: only updates rows that still have the old emails. */
SET QUOTED_IDENTIFIER ON;
SET ANSI_NULLS ON;
SET NOCOUNT ON;
USE BrokerKnow_KE_Demo;
GO
UPDATE dbo.PortalUsers SET Email = 'demo@greenmargin.demo'
 WHERE Email = 'demo@axis-kenya.demo';
UPDATE dbo.PortalUsers SET Email = 'admin@greenmargin.demo'
 WHERE Email = 'admin@axis-kenya.demo';
SELECT Email, FirstName, LastName, Role FROM dbo.PortalUsers ORDER BY Role;
GO
