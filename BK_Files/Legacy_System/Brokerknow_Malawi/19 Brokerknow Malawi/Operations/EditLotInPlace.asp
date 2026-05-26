<!--#include file="../libroutines.asp"-->

<%
	const UDLName = "KBroker"
	const DataSource = "EditLot"
	const DataEntity = "Lot"
	const DataEntityPlural = "Lots"
	const ActionFolder = "Operations"
	
	const LinkedIndependent = 1
   const LinkedDependent = 2
	
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim guidStr 
	Dim guid
	Dim ID
	Dim ItemID
	Dim rsEdit
	Dim IDHolder
	Dim IDArray
	
	IDHolder = Request("ID")
	If (Trim(IDHolder) = "") or (IDHolder = "0") Then%>
			<script language = 'vbscript'>
                	ShowMessage "Please select an Order item for Lot allocation"
                	window.self.close
			</script>
			<%response.end
	End If
	IDArray = split(IDHolder,"<->")
	ID = IDArray(lbound(IDArray))
	ItemID = IDArray(ubound(IDArray))
	
   	Dim slip
	Dim broker
	Dim tDate
	Dim qty
	Dim price
	Dim orderType
	dim orderSecType
	Dim orderIsSaleType
	Dim securityID
	Dim commission
	
		
	broker = Request.Form("cboBrokerInPlace")
	tDate = Trim(Request.Form("TDate"))
	tDate = tDate & " " & Time	
	slip = Request.Form("SlipNo")
	qty = Request.Form("Quantity")
	price = Request.Form("Price")
	orderType = Request.Form("OrderType")
	orderSecType = Request.Form("OrderSecType")
	orderIsSaleType = cbool(Request.Form("OrderIsSaleType"))
	securityID = Request.Form("SecurityID")
	commission = Request.Form("Commission")
	'validate Broker
	If Trim(Broker) = "" Then%>
			<script language = 'vbscript'>
				ShowMessage "Please specify the Broker"
				window.history.back();
			</script>
			<% response.end
	End If
	
	'validate Slip
	If Trim(Slip) = "" Then%>
			<script language = 'vbscript'>
				ShowMessage "Please specify the Slip No."
				window.history.back();
			</script>
			<% response.end
	End If
	'ensure Slip is numeric
	If (Not IsNumeric(Slip)) Then%>
		<script language = 'vbscript'>
			ShowMessage "Slip No. must be numeric"
			window.history.back();
		</script>
		<% response.end
	End If
	
	'validate Estimated Price
	If Trim(Price) = "" Then%>
		<script language = 'vbscript'>
				ShowMessage "Please specify the Price "
				window.history.back();
		</script>
		<% response.end
	End If
	'validate Estimated Quantity
	If Trim(qty)= "" Then%>
		<script language = 'vbscript'>
				ShowMessage "Please specify the Quantity "
				window.history.back();
		</script>
		<% response.end
	End If
'ensure Order Detail Estimated Price is numeric
	If (Price <> "") And (Not IsNumeric(Price)) Then%>
		<script language = 'vbscript'>
			ShowMessage "Order Detail Estimated Price must be numeric"
			window.history.back();
		</script>
		<% response.end
	End If
	'ensure Order Detail Estimated Quantity is numeric
	If (Qty <> "") And (Not IsNumeric(Qty)) Then%>
		<script language = 'vbscript'>
			ShowMessage "Order Detail Estimated Quantity must be numeric"
			window.history.back();
		</script>
		<% response.end
	End If
	
	Set conn = GetActiveConnection("KBroker")    
	'calculate amounts
	Dim grossAmount  'this is the amount before application of Levies
	Dim levyRS
	Dim blocks
	
	'grossAmount = price * qty     

	'calculate amounts						
	if orderSecType = "Fixed" then ' "F" is FIXED security
		grossAmount = (price * qty) / 100
	else
		grossAmount = price * qty
	end if		
			  
		
	conn.BeginTrans
			if itemID = "0" then
					'save contract
					set guid = server.createobject("NDUtils.CGUID")
					guidStr = guid.GenerateGUID
						
					sqlStr = "INSERT INTO [Contract] (Contract_DPA_, Contract_EIT_, Status_DPA_) " & _
							"SELECT " & " " & "iif(isnull(max([Contract_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Contract'),max([Contract_DPA_]) + 1)" & " " & " as Contract_DPA_" & _
							"," & "'" & guidStr & "'" & " as Contract_EIT_" & _
							"," & " " & 1 & " " & " as Status_DPA_" & _
							" FROM [Contract]"
					
					conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
					
					'obtain contract number
					sqlStr = "SELECT Contract_DPA_ FROM [Contract] WHERE [Contract.Contract_EIT_] = " & "'" & guidStr & "'"
					
					Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
					If (rs.EOF Or rs.BOF) Then%>
				         		<script language = 'vbscript'>
				         				ShowMessage "A serious error has been encountered while saving the data. Try saving again"
							         	
				         		</script>
				         		<% response.end
					End If
					
					'save lot
					sqlStr = "INSERT INTO [Lot] (Lot_DPA_,Contract_DPA_,OrdDetail_DPA_,LotPrice" & _
							",LotQty,LotSlipNo,LotTDate,Broker_DPA_,ContractNumber,LotGrossAmount)" & _
							" SELECT " & " " & "iif(isnull(max([Lot_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Lot'),max([Lot_DPA_]) + 1)" & " " & " as Lot_DPA_" & _
							"," & " " & rs.fields("Contract_DPA_") & " " & " as Contract_DPA_" & _
							"," & " " & ID & " " & " as OrdDetail_DPA_" & _
							"," & " " & price & " " & " as LotPrice" & _
							"," & " " & qty & " " & " as LotQty" & _
							"," & " " & slip & " " & " as LotSlipNo" & _
							"," & "#" & FormatDate(tDate) & "#" & " as LotTDate" & _
							"," & " " & broker & " " & " as Broker_DPA_" & _
							"," & "'" & left(orderType,1) & rs.fields("Contract_DPA_") & "'" & " as ContractNumber" & _								
							"," & " " & grossAmount & " " & " as LotGrossAmount" & _
							" FROM [Lot]"
					
					conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
				
					'apply relevant levies
					
					Dim cond
					if orderSecType = "Fixed" then
						cond = "WHERE LevyAppBond = 1"
					else
						cond = "WHERE LevyAppSecurity = 1"
					end if
					
					sqlStr = "SELECT * FROM [LevyList] " & cond
					Set levyRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
					
					Dim levyAmount
					If Not (levyRS.EOF Or levyRS.BOF) Then
							levyRS.MoveFirst
							Do Until levyRS.EOF
									if levyRS.Fields("LevyType") = "P" Then
											levyAmount = CCur((levyRS.Fields("LevyAmount")/100.00) * grossAmount)
											LevyRatePercentage = levyRS.Fields("LevyAmount") & "%"
									else
											blocks = Int(grossAmount / levyRS.Fields("LevyBlock"))
											if (grossAmount mod levyRS.Fields("LevyBlock")) <> 0 then
													blocks = blocks + 1
											end if
											levyAmount = CCur(blocks * levyRS.Fields("LevyAmount"))
											LevyRatePercentage = levyRS.Fields("LevyAmount") & " for every " & levyRS.Fields("LevyBlock")
									end if
									sqlStr = "INSERT INTO [LevyContract] (LevyContract_DPA_,Contract_DPA_,LevyName,LevyAmount,LevyRate,LevyBlock,SystemMaintained,LevyShortName,LevyRatePercentage) SELECT " & " " & "iif(isnull(max([LevyContract_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'LevyContract'),max([LevyContract_DPA_]) + 1)" & " " & " as LevyContract_DPA_" & _
											"       ," & " " & rs.Fields("Contract_DPA_") & " " & " as Contract_DPA_" & _
											"       ," & "'" & levyRS.Fields("LevyDescription") & "'" & " as LevyName" & _
											"       ," & " " & RoundPoint05(levyAmount) & " " & " as LevyAmount" & _
											"       ," & " " & levyRS.Fields("LevyAmount") & " " & " as LevyRate" & _
											"       ," & " " & levyRS.Fields("LevyBlock") & " " & " as LevyBlock" & _
											"       ," & " " & levyRS.Fields("SystemMaintained") & " " & " as SystemMaintained" & _
											"       ," & " '" & levyRS.Fields("LevyShortName") & "' " & " as LevyShortName" & _												
											"       ," & " '" & LevyRatePercentage  & "' " & " as LevyRatePercentage" & _												
											"        FROM [LevyContract]"
									conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
									levyRS.MoveNext
							Loop
					end if
					'consider the transfer fee
					if Not(orderIsSaleType) then
						Dim transRS
						sqlStr = "SELECT * FROM [SecTransFeeListLatest] WHERE Security_DPA_ = " & securityID
						Set transRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
						sqlStr = "INSERT INTO [LevyContract] (LevyContract_DPA_,Contract_DPA_,LevyName,LevyAmount,LevyRate,LevyBlock) SELECT " & " " & "iif(isnull(max([LevyContract_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'LevyContract'),max([LevyContract_DPA_]) + 1)" & " " & " as LevyContract_DPA_" & _
								"       ," & " " & rs.Fields("Contract_DPA_") & " " & " as Contract_DPA_" & _
								"       ," & "'Transfer Fee'" & " as LevyName" & _
								"       ," & " " & transRS.Fields("Fee") & " " & " as LevyAmount" & _
								"       ," & " " & transRS.Fields("Fee") & " " & " as LevyRate" & _
								"       ," & " " & 0 & " " & " as LevyBlock" & _
								"        FROM [LevyContract]"
						conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
					end if
					'Apply broker commission
					levyAmount = CCur((commission/100.00) * grossAmount)
					sqlStr = "INSERT INTO [LevyContract] (LevyContract_DPA_,Contract_DPA_,LevyName,LevyAmount,LevyRate,LevyBlock,LevyShortName,LevyRatePercentage) SELECT " & " " & "iif(isnull(max([LevyContract_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'LevyContract'),max([LevyContract_DPA_]) + 1)" & " " & " as LevyContract_DPA_" & _
							"       ," & " " & rs.Fields("Contract_DPA_") & " " & " as Contract_DPA_" & _
							"       ," & "'Broker Commission'" & " as LevyName" & _
							"       ," & " " & RoundPoint05(levyAmount) & " " & " as LevyAmount" & _
							"       ," & " " & commission & " " & " as LevyRate" & _
							"       ," & " " & 0 & " " & " as LevyBlock" & _
							"       ," & " 'Commission' " & " as LevyShortName" & _
							"       ," & " '" & commission & "%" & "' " & " as LevyRatePercentage" & _
							"        FROM [LevyContract]"
					conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
			else
					'edit detail data
					sqlStr = "UPDATE Lot SET LotPrice = " & price & "," & _
							" LotQty = " & qty & ", LotSlipNo = " & slip & "," & _
							" LotTDate = #" & FormatDate(tDate) & "#," & _
							" Broker_DPA_ = " & broker & _
							", LotGrossAmount = " & grossAmount & _
							" WHERE Lot_DPA_=" & itemID
					conn.Execute SQLServerFormat(HandleQuote(sqlStr))
					
					'retrieve contract number
					sqlStr = "SELECT Contract_DPA_ FROM LotList WHERE LotList.Lot_DPA_=" & itemID
					Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
					If (rs.EOF Or rs.BOF) Then%>
				         		<script language = 'vbscript'>
				         				ShowMessage "A serious error has been encountered while saving the data. Try saving again"
							         	window.history.back();
				         		</script>
				         		<% response.end
					End If
					
					'update levies
					
					sqlStr = "SELECT * FROM LevyContractList WHERE Contract_DPA_= " & rs.fields("Contract_DPA_")
					Set levyRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
					If (levyRS.EOF Or levyRS.BOF) Then%>
				         		<script language = 'vbscript'>
				         				ShowMessage "The levies for the selected contract are missing. This is a serious database corruption."
							         	window.history.back();
				         		</script>
				         		<% response.end
					End If
					do until levyRS.eof
							if levyRS.Fields("LevyBlock")<>0 then
									'block type levy
									blocks = Int(grossAmount / levyRS.Fields("LevyBlock"))
									if (grossAmount mod levyRS.Fields("LevyBlock")) <> 0 then
											blocks = blocks + 1
									end if
									levyAmount = CCur(blocks * levyRS.Fields("LevyRate"))
							elseif levyRS.Fields("LevyName") = "Transfer Fee" then
									'transfer fee
									levyAmount = levyRS.Fields("LevyRate")
							else
									'broker commission or other percentage-based commission
									levyAmount = CCur((levyRS.Fields("LevyRate")/100.00) * grossAmount)
							end if
							
							sqlStr = "UPDATE LevyContract SET LevyAmount = " & RoundPoint05(levyAmount) & _
									" WHERE LevyContract_DPA_ = " & levyRS.Fields("LevyContract_DPA_")
							conn.Execute SQLServerFormat(HandleQuote(sqlStr))
							levyRS.movenext
					loop
			end if 
				
	conn.CommitTrans    
	conn.Close
	Set conn = Nothing
	response.redirect DataEntity & "List.asp"
	Response.End 		
   	
%>
