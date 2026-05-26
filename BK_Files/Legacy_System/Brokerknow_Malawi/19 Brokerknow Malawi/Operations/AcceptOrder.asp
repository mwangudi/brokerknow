<!--#include file="../libroutines.asp"-->
<%
	const UDLName = "KBroker"
	const DataSource = "AcceptOrder"
	const DataEntity = "Order"
	const DataEntityPlural = "Orders"
	const ActionFolder = "Operations"
	
	
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
   Dim ID
   Dim rsEdit
	
	action = ucase(Request.Form("delAction"))
	ID = Request("ID")

		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for editing"
                		
                </script>
                <% response.end
        End If
        
        

	select case action 
		case "EXECUTE"
			Dim accept
			dim ReleaseDate
			Dim guid
			Dim guidStr
	        
			accept = Request.Form("Accept")

		
			Set conn = GetActiveConnection("KBroker")
	        
			'save data
			sqlStr = "SELECT * FROM WebtbOrder WHERE (Order_DPA_ = " & ID & ")"
			Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
			If (rs.EOF Or rs.BOF) Then%>
			 			<script language = 'vbscript'>
			 					ShowMessage "The specified order cannot be found."
				         		window.history.back			
			 			</script>
			 			<% response.end
			End If
			
			set guid = server.createobject("NDUtils.CGUID")
			 guidStr = guid.GenerateGUID
				 
			 sqlStr = "INSERT INTO [tbOrder] (OrderDate,OrderHold,OrderRef,Order_DPA_,Order_EIT_,Branch_DPA_,OrderSecType_DPA_,Client_DPA_,OrderType_DPA_,OrderAutoReleaseDate,OrderHoldType_DPA_,OrderCompounded) SELECT " & _
					 "#" & rs.fields("OrderDate") & "#" & " as OrderDate" & _
					 "," & " " & 1 & " " & " as OrderHold" & _
			         "," & "'" & rs.fields("OrderRef") & "'" & " as OrderRef" & _ 
			         "," & " " & "iif(isnull(max([Order_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'tbOrder'),max([Order_DPA_]) + 1)" & " " & " as Order_DPA_" & _
			         "," & "'" & guidStr & "'" & " as Order_EIT_" & _ 
			         "," & " " & Session("Branch_DPA_") & " " & " as Branch_DPA_" & _
			         "," & " " & rs.fields("OrderSecType_DPA_") & " " & " as OrderSecType_DPA_" & _
			         "," & " " & rs.fields("Client_DPA_") & " " & " as Client_DPA_" & _
			         "," & " " & rs.fields("OrderType_DPA_") & " " & " as OrderType_DPA_" & _
			         "," & " NULL " & " as OrderAutoReleaseDate" & _
			         "," & " " & 2 & " " & " as OrderHoldType_DPA_" & _
			         "," & " " & abs(cint(rs.fields("OrderCompounded"))) & " " & " as OrderCompounded" & _
			         " FROM [tbOrder]"
			 Set conn = GetActiveConnection("KBroker")
			 conn.BeginTrans
					sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
					conn.Execute sqlStr
				     
					'obtain header key value
					sqlStr = "SELECT [tbOrder.Order_DPA_] FROM [tbOrder] WHERE [tbOrder.Order_EIT_] = " & "'" & guidStr & "'"
				     
				     Dim orderRS
					Set orderRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
					If (orderRS.EOF Or orderRS.BOF) Then%>
			         			<script language = 'vbscript'>
			         					ShowMessage "A serious error has been encountered while saving the data. Try saving again"
				         					
			         			</script>
			         			<% response.end
					End If
				     
					'save detail data
					sqlStr = "INSERT INTO [OrdDetail] (OrdDetailCertNo,OrdDetailPrice,OrdDetailQty,OrdDetailValidity" & _
							",OrdDetail_DPA_,Order_DPA_,Security_DPA_) SELECT " & _
							"''" & " as OrdDetailCertNo" & _
							"," & "'" & rs.fields("OrdDetailPrice") & "'" & " as OrdDetailPrice" & _
							"," & " " & CDbl(rs.fields("OrdDetailQty")) & " " & " as OrdDetailQty" & _
							"," & "#" & rs.fields("OrdDetailValidity") & "#" & " as OrdDetailValidity" & _
							"," & " " & "iif(isnull(max([OrdDetail_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'OrdDetail'),max([OrdDetail_DPA_]) + 1)" & " " & " as OrdDetail_DPA_" & _
							"," & " " & orderRS.Fields("Order_DPA_") & " " & " as Order_DPA_" & _
							"," & " " & rs.fields("Security_DPA_") & " " & " as Security_DPA_" & _
							" FROM [OrdDetail]"
				     
					sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
					conn.Execute = sqlStr
					
					sqlStr = "UPDATE WebtbOrder SET Accepted = 1 WHERE(Order_DPA_ = " & RS.Fields("Order_DPA_") & ")"
					conn.Execute = sqlStr
			conn.CommitTrans
			conn.Close
			Set conn = Nothing

			response.redirect "OnlineOrderList.asp"	
    end select
    	
%>

