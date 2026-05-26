<!--#include file="../libroutines.asp"-->
<%
	const UDLName = "KBroker"
	const DataSource = "ApproveRequest"
	const DataEntity = "PaymentRequest"
	const DataEntityPlural = "PaymentRequests"
	const ActionFolder = "Operations"
	
	
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
   Dim ID
   Dim rsEdit
	
	action = ucase(Request.Form("delAction"))
	ID = Request("ID")
	
	Amount = Request("Amount")
	ClientID = Request("ClientDPA")
	PayDate = Request("PayDate")
	
	
	'Response.Write ClientID	
	'Response.End
	
		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for editing"
                		
                </script>
                <% response.end
        End If
        
        

	select case action 
		case "EXECUTE"
			Dim release
			dim ReleaseDate  
			
		
			Set conn = GetActiveConnection("KBroker")
	        
			'save data
			Dim userID
			Dim manualReleaseDate
			
			UserId=Session("UserID")
			
			manualReleaseDate = "GETDATE()"
			
			conn.execute ("Exec ClientBalanceProcedure " & ClientID)
			
			sqlStr = "SELECT     Client.Client_DPA_, Client.ClientCDSNo, Client.ClientName, tbOrder.Order_DPA_,lot.Contract_DPA_, Security.SecurityCode, ClientBalances.CurrentBal,  " & _
				"                       Lot.ContractNumber, Lot.LotGrossAmount - a.Levy AS NetAmount,  " & _
				"                       CASE WHEN ClientBalances.CurrentBal < CASE WHEN tbOrder.PayOption = 3 THEN CASE when tbOrder.PartialAmount >Lot.LotGrossAmount - a.Levy then Lot.LotGrossAmount - a.Levy else tbOrder.PartialAmount end ELSE Lot.LotGrossAmount - a.Levy END THEN " & _
				"                        ClientBalances.CurrentBal ELSE CASE WHEN tbOrder.PayOption = 3 THEN CASE when tbOrder.PartialAmount >Lot.LotGrossAmount - a.Levy then Lot.LotGrossAmount - a.Levy else tbOrder.PartialAmount end ELSE Lot.LotGrossAmount - a.Levy END END AS Amount, " & _
				"                        Contract.ContractSettlementDate,b.Balance " & _
				" FROM         OrdDetail INNER JOIN " & _
				"                       tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN " & _
				"                       Lot ON OrdDetail.OrdDetail_DPA_ = Lot.OrdDetail_DPA_ INNER JOIN " & _
				"                       Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ INNER JOIN " & _
				"                       ClientBalances ON Client.Client_DPA_ = ClientBalances.client_DPA_ INNER JOIN " & _
				"                       Security ON OrdDetail.Security_DPA_ = Security.Security_DPA_ LEFT OUTER JOIN " & _
				"                       PaymentRequests ON Lot.Contract_DPA_ = PaymentRequests.Contract_DPA_ INNER JOIN " & _
				"                           (SELECT     SUM(LevyAmount) AS Levy, Contract_DPA_ " & _
				"                             FROM          Levycontract " & _
				"                             WHERE      deleted = 0 " & _
				"                             GROUP BY Contract_DPA_) a ON Lot.Contract_DPA_ = a.Contract_DPA_ INNER JOIN " & _
				"                       Contract ON Lot.Contract_DPA_ = Contract.Contract_DPA_ INNER JOIN " & _
				"                           (SELECT     CASE WHEN tbOrder.PayOption = 3 THEN tbOrder.PartialAmount ELSE SUM(lot.LotGrossAmount) - SUM(d .LevyAmount)  " & _
				"                                                    END - SUM(isnull(PaymentRequests.PaymentAmount, 0)) AS Balance, tbOrder.Order_DPA_ " & _
				"                             FROM          OrdDetail INNER JOIN " & _
				"                                                    Lot ON OrdDetail.OrdDetail_DPA_ = Lot.OrdDetail_DPA_ INNER JOIN " & _
				"                                                    tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN " & _
				"                                                        (SELECT     SUM(LevyAmount) AS LevyAmount, contract_DPA_ " & _
				"                                                          FROM          LevyContract " & _
				"                                                          WHERE      deleted = 0 " & _
				"                                                          GROUP BY Contract_DPA_) d ON Lot.Contract_DPA_ = d .Contract_DPA_ LEFT OUTER JOIN " & _
				"                                                    PaymentRequests ON Lot.Contract_DPA_ = PaymentRequests.Contract_DPA_ " & _
				"                             WHERE      (tbOrder.Deleted = 0) AND (OrdDetail.Deleted = 0) AND (Lot.Deleted = 0) " & _
				"                             GROUP BY tbOrder.PartialAmount, tbOrder.Order_DPA_, tbOrder.PayOption) b ON tbOrder.Order_DPA_ = b.Order_DPA_" & _
				" WHERE     (OrdDetail.Deleted = 0) AND (Lot.Deleted = 0) AND (tbOrder.Deleted = 0) AND (Client.Deleted = 0) AND (tbOrder.OrderType_DPA_ = 2) AND  " & _
				"                       (tbOrder.PayOption <> 1) AND (PaymentRequests.Request_DPA_ IS NULL) and Lot.Contract_DPA_=" & ID
		
			
			'Response.Write sqlStr
			'Response.End			
			
			Set Rs = conn.Execute(sqlStr)
			
			if Not (Rs.eof or Rs.Bof) then
			
			narrative = "Outstanding Requests"
			
			if(CCur(Rs("Amount"))<CCur(Rs("Balance"))) then
			PaymentAmount = Rs("Amount")
			else
			PaymentAmount = Rs("Balance")
			end if
			
			sqlStr = "INSERT INTO [PaymentRequests] (Request_DPA_,Client_DPA_, " & _
					"PaymentAmount,RequestNarrative,RequestPayDate,ModifiedBy,CreatedBy,Contract_DPA_) " & _
					" SELECT " & " " & "iif(isnull(max(Request_DPA_)),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'PaymentRequests'),max(Request_DPA_) + 1)" & " " & " as Request_DPA_" & _									
					"," & "" & Rs("Client_DPA_") & "" & " as Client_DPA_" & _
					"," & " " & PaymentAmount & " " & " as PaymentAmount" & _						                                    
					"," & "'" & narrative & "'" & " as RequestNarrative" & _				
					"," & "'" & Rs("ContractSettlementDate") & "'" & " as RequestPayDate" & _
					"," & " " & UserId & " " & " as ModifiedBy" & _	
					"," & " " & UserId & " " & " as CreatedBy" & _
					"," & " " & Rs("Contract_DPA_") & " " & " as Contract_DPA_" & _						                                    																							                                    																								                                    																							                                    																		
					" FROM [PaymentRequests]"
		
							
		'response.write sqlStr
		
		'response.end
			conn.BeginTrans
					conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
					conn.execute ("Exec ClientTotalProcedure " & Rs("Client_DPA_"))							
					conn.execute ("Exec ClientBalanceProcedure " & Rs("Client_DPA_"))
			conn.CommitTrans
			conn.Close
			Set conn = Nothing
		end if
		
			response.redirect "SaleRequests.asp"	
    end select
    	
%>

