

CREATE PROCEDURE ClientTotalProcedure @clientDPA int

AS

--Update  ClientTotal SET Total = ClientTotals.Total from ClientTotal inner join ClientTotals on ClientTotal.client_DPA_=ClientTotals.Client_DPA_ where ClientTotals.Client_DPA_=@clientDPA



Delete From ClientTotal where Client_DPA_=@clientDPA

Insert Into ClientTotal(Client_DPA_,Total)



Select Client_DPA_,Sum(Total) as Total

from

(

SELECT     Client_DPA_, SUM(ISNULL(BalanceQty * OrdDetailPrice, 0)) AS Total

FROM         (SELECT     ordtbl.BalanceQty, CASE (dbo.OrdDetail.Best) WHEN 1 THEN dbo.datastream_SecurityPriceList.Price * 1.020825 ELSE CONVERT(float, 

                      dbo.OrdDetail.OrdDetailPrice) END AS OrdDetailPrice, dbo.tbOrder.Order_DPA_, dbo.tbOrder.Client_DPA_, dbo.tbOrder.OrderCanceled

FROM         dbo.OrdDetail INNER JOIN

                      dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN

                          (SELECT     CASE ISNULL(OrdDetailContractedQtyList.ContractQty, 1) 

                                                   WHEN 1 THEN dbo.OrdDetail.OrdDetailQty ELSE dbo.OrdDetail.OrdDetailQty - OrdDetailContractedQtyList.ContractQty END AS BalanceQty,

                                                    dbo.OrdDetail.OrdDetail_DPA_

                            FROM          dbo.OrdDetail LEFT OUTER JOIN

                                                   (SELECT     dbo.OrdDetail.OrdDetail_DPA_, SUM(dbo.Lot.LotQty) AS ContractQty

FROM         dbo.OrdDetail INNER JOIN

                      dbo.Lot ON dbo.OrdDetail.OrdDetail_DPA_ = dbo.Lot.OrdDetail_DPA_ 

inner join tbOrder on OrdDetail.Order_DPA_=tbOrder.Order_DPA_

WHERE     (dbo.Lot.Deleted = 0) AND (dbo.OrdDetail.Deleted = 0) and  tbOrder.Client_DPA_=@clientDPA

GROUP BY dbo.OrdDetail.OrdDetail_DPA_) OrdDetailContractedQtyList ON dbo.OrdDetail.OrdDetail_DPA_ = OrdDetailContractedQtyList.OrdDetail_DPA_

                            WHERE      (dbo.OrdDetail.Deleted = 0)) ordtbl ON ordtbl.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN

                      dbo.datastream_SecurityPriceList ON dbo.OrdDetail.Security_DPA_ = dbo.datastream_SecurityPriceList.Security_DPA_

WHERE     (dbo.OrdDetail.Deleted = 0) AND (dbo.tbOrder.Deleted = 0) AND (dbo.tbOrder.OrderType_DPA_ = 1) AND (ordtbl.BalanceQty > 0) AND 

                   (dbo.tbOrder.OrderCanceled = 0)  AND dbo.tbOrder.Client_DPA_=@clientDPA) a

GROUP BY Client_DPA_

union all

SELECT     Client_DPA_, SUM(PaymentAmount) AS Clienttotal

FROM         PaymentRequests

WHERE     (Processed_DPA_ IS NULL) AND (Deleted = 0) and Client_DPA_=@clientDPA

GROUP BY Client_DPA_) a 

group by Client_DPA_



