SET NOCOUNT ON;
-- Does KE_Demo's synthetic demo client (ClientDpa=1001) collide with a REAL KE client?
SELECT Client_DPA_, ClientName FROM BrokerKnow_KE_Clean.dbo.Client WHERE Client_DPA_ = 1001;
-- KE_Demo logins to graft.
SELECT Id, Email, Role, Status, ClientDpa FROM BrokerKnow_KE_Demo.dbo.PortalUsers ORDER BY Id;
