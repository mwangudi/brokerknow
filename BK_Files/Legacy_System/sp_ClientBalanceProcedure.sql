



CREATE PROCEDURE ClientBalanceProcedure @clientDPA int

AS

SET NOCOUNT ON

--UPDATE ClientBalances SET CurrentBal=CurrentBalances.CurrentBal from ClientBalances inner join CurrentBalances on ClientBalances.Client_DPA_= CurrentBalances.Client_DPA_ Where CurrentBalances.Client_DPA_=@clientDPA

delete from ClientBalances where Client_DPA_=@clientDPA

Insert Into ClientBalances(Client_DPA_,CurrentBal) 

SELECT  ClientsStatement.Client_DPA_,   

SUM(ISNULL(ClientsStatement.Credit-ClientsStatement.Debit, 0)) + dbo.Client.ClientOpeningBal AS CurrentBal 

FROM (   

Select * from ClientStatement WHERE Client_DPA_ = @clientDPA

) ClientsStatement

INNER JOIN

                      dbo.Client ON ClientsStatement.Client_DPA_ = dbo.Client.Client_DPA_

WHERE     (dbo.Client.Deleted = 0)  AND dbo.Client.Client_DPA_= @clientDPA

GROUP BY ClientsStatement.Client_DPA_, dbo.Client.ClientOpeningBal





SET QUOTED_IDENTIFIER ON
