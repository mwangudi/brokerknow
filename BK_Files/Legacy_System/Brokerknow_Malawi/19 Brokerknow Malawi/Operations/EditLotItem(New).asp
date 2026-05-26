<!--#include file="../libroutines.asp"-->

<%
	const UDLName = "KBroker"
	const DataSource = "EditLot"
	const DataEntity = "Lot"
	const DataEntityPlural = "Lots"
	const ActionFolder = "Operations"
	
	const LinkedIndependent = 1
    const LinkedDependent = 2
	
	Dim action
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
	
	
	action = ucase(Request.Form("action"))
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
	
	ItemID = Request.Form("ItemID")
	if action <> "" then
			If Trim(ItemID) = "" Then%>
					<script language = 'vbscript'>
            				ShowMessage "No item specified"                				
					</script>
					<%response.end
			End If
	end if
	select case action 
		case "EXECUTE_DETAIL"
   			Dim slip
			Dim broker
			Dim tDate
			Dim qty
			Dim price
			Dim orderType
			Dim orderIsSaleType
			Dim securityID
			Dim commission	
			Dim orderSecType
			Dim agentCommission
			Dim staffCommission
	        
	        if itemID = "-1" then
					broker = Request.Form("cboBroker")
					tDate = Trim(Request.Form("txtTDate"))
					tDate = tDate & " " & Time
					slip = Request.Form("txtSlip")
					qty = Request.Form("txtQty")
					price = Request.Form("txtPrice")					
			else
					broker = Request.Form("cboBrokerInPlace")
					tDate = Trim(Request.Form("Date"))
					tDate = tDate & " " & Time
					slip = Request.Form("Slip")
					qty = Request.Form("Quantity")
					price = Request.Form("Price")
			end if
			
			orderType = Request.Form("txtOrderType")
			orderIsSaleType = cbool(Request.Form("txtOrderIsSaleType"))
			securityID = Request.Form("txtSecurityID")
			commission = Request.Form("txtCommission")
			orderSecType = Request.Form("txtInstrument")
			agentCommission = Request.Form("txtAgentCommission")
			staffCommission = Request.Form("txtStaffCommission")
			
			'validate Broker
			If Trim(Broker) = "" Then%>
				    <script language = 'vbscript'>
				        ShowMessage "Please specify the Broker"
				        
				    </script>
				    <% response.end
			End If
			
			'validate Slip
			If Trim(Slip) = "" Then%>
				    <script language = 'vbscript'>
				        ShowMessage "Please specify the Slip No."
				        
				    </script>
				    <% response.end
			End If
			'ensure Slip is numeric
			If (Not IsNumeric(Slip)) Then%>
				<script language = 'vbscript'>
					ShowMessage "Slip No. must be numeric"
					
				</script>
				<% response.end
			End If
			
			'validate Estimated Price
			If Trim(Price) = "" Then%>
				<script language = 'vbscript'>
						ShowMessage "Please specify the Price "
					    
				</script>
				<% response.end
			End If
			'validate Estimated Quantity
			If Trim(qty)= "" Then%>
				<script language = 'vbscript'>
						ShowMessage "Please specify the Quantity "
					    
				</script>
				<% response.end
			End If
		'ensure Order Detail Estimated Price is numeric
			If (Price <> "") And (Not IsNumeric(Price)) Then%>
				<script language = 'vbscript'>
					ShowMessage "Price must be a number"
					
				</script>
				<% response.end
			End If
			'ensure Order Detail Estimated Quantity is numeric
			If (Qty <> "") And (Not IsNumeric(Qty)) Then%>
				<script language = 'vbscript'>
					ShowMessage "Quantity must be a number"
					
				</script>
				<% response.end
			End If
			
			'ensure date is valid format
			If Not IsDate(tDate) Then%>
				<script language = 'vbscript'>
					ShowMessage "Date must be a valid date"					
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
					if itemID = "-1" then
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
							
							sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))

							conn.Execute sqlStr
						
							'apply relevant levies
							
							Dim cond
							if orderSecType = "Fixed" then ' "F" is FIXED security
								cond = "WHERE LevyAppBond = 1 AND LevyActive = 1"
							else
								cond = "WHERE (LevyAppSecurity = 1) AND (LevyActive = 1) AND (" & _
									" Levy_DPA_ IN (SELECT Levy_DPA_ FROM LevySecurity " & _
									" WHERE Security_DPA_ = " & securityID & "))"
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
											sqlStr = "INSERT INTO [LevyContract] (LevyContract_DPA_,Contract_DPA_,LevyName,LevyAmount,LevyRate,LevyBlock, SystemMaintained,LevyShortName,LevyRatePercentage) SELECT " & " " & "iif(isnull(max([LevyContract_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'LevyContract'),max([LevyContract_DPA_]) + 1)" & " " & " as LevyContract_DPA_" & _
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
								sqlStr = "INSERT INTO [LevyContract] (LevyContract_DPA_,Contract_DPA_,LevyName,LevyAmount,LevyRate,LevyBlock,LevyShortName,SystemMaintained) SELECT " & " " & "iif(isnull(max([LevyContract_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'LevyContract'),max([LevyContract_DPA_]) + 1)" & " " & " as LevyContract_DPA_" & _
										"       ," & " " & rs.Fields("Contract_DPA_") & " " & " as Contract_DPA_" & _
										"       ," & "'Transfer Fee'" & " as LevyName" & _
										"       ," & " " & transRS.Fields("Fee") & " " & " as LevyAmount" & _
										"       ," & " " & transRS.Fields("Fee") & " " & " as LevyRate" & _
										"       ," & " " & 0 & " " & " as LevyBlock" & _
										"       ," & " 'Transfer' " & " as LevyShortName" & _
										"       ,13 " & " as SystemMaintained" & _	
										"        FROM [LevyContract]"
								conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
							end if
							'Apply broker commission
							levyAmount = CCur((commission/100.00) * grossAmount)
							sqlStr = "INSERT INTO [LevyContract] (LevyContract_DPA_,Contract_DPA_,LevyName,LevyAmount,LevyRate,LevyBlock,LevyShortName,LevyRatePercentage,SystemMaintained) SELECT " & " " & "iif(isnull(max([LevyContract_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'LevyContract'),max([LevyContract_DPA_]) + 1)" & " " & " as LevyContract_DPA_" & _
									"       ," & " " & rs.Fields("Contract_DPA_") & " " & " as Contract_DPA_" & _
									"       ," & "'Broker Commission'" & " as LevyName" & _
									"       ," & " " & RoundPoint05(levyAmount) & " " & " as LevyAmount" & _
									"       ," & " " & commission & " " & " as LevyRate" & _
									"       ," & " " & 0 & " " & " as LevyBlock" & _
									"       ," & " 'Commission' " & " as LevyShortName" & _
									"       ," & " '" & commission & "%" & "' " & " as LevyRatePercentage" & _	
									"       ,11 " & " as SystemMaintained" & _	
									"        FROM [LevyContract]"
							conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
							
							'Apply agent commission
							Dim tmpLevy 'broker commission
						
							tmpLevy = CCur((commission/100.00) * grossAmount)
							levyAmount = CCur((agentCommission/100.00) * tmpLevy)
							sqlStr = "INSERT INTO [LevyContract] (LevyContract_DPA_,Contract_DPA_,LevyName,LevyAmount,LevyRate,LevyBlock,LevyShortName,LevyRatePercentage,SystemMaintained) SELECT " & " " & "iif(isnull(max([LevyContract_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'LevyContract'),max([LevyContract_DPA_]) + 1)" & " " & " as LevyContract_DPA_" & _
									"       ," & " " & rs.Fields("Contract_DPA_") & " " & " as Contract_DPA_" & _
									"       ," & "'Agent Commission'" & " as LevyName" & _
									"       ," & " " & RoundPoint05(levyAmount) & " " & " as LevyAmount" & _
									"       ," & " " & agentCommission * (commission/100.00) & " " & " as LevyRate" & _
									"       ," & " " & 0 & " " & " as LevyBlock" & _
									"       ," & " 'Agent' " & " as LevyShortName" & _
									"       ," & " '" & agentCommission & "%" & "' " & " as LevyRatePercentage" & _
									"       ,12 " & " as SystemMaintained" & _	
									"        FROM [LevyContract]"
							conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
							
							'Apply account manager commission
							tmpLevy = CCur((commission/100.00) * grossAmount)
							levyAmount = CCur((staffCommission/100.00) * tmpLevy)
							sqlStr = "INSERT INTO [LevyContract] (LevyContract_DPA_,Contract_DPA_,LevyName,LevyAmount,LevyRate,LevyBlock,LevyShortName,LevyRatePercentage,SystemMaintained) SELECT " & " " & "iif(isnull(max([LevyContract_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'LevyContract'),max([LevyContract_DPA_]) + 1)" & " " & " as LevyContract_DPA_" & _
									"       ," & " " & rs.Fields("Contract_DPA_") & " " & " as Contract_DPA_" & _
									"       ," & "'Staff Commission'" & " as LevyName" & _
									"       ," & " " & RoundPoint05(levyAmount) & " " & " as LevyAmount" & _
									"       ," & " " & staffCommission * (commission/100.00) & " " & " as LevyRate" & _
									"       ," & " " & 0 & " " & " as LevyBlock" & _
									"       ," & " 'Staff' " & " as LevyShortName" & _
									"       ," & " '" & staffCommission & "%" & "' " & " as LevyRatePercentage" & _
									"       ,8 " & " as SystemMaintained" & _	
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
							         			
				         				</script>
				         				<% response.end
							End If
							
							'update levies
							
							sqlStr = "SELECT * FROM LevyContractList WHERE Contract_DPA_= " & rs.fields("Contract_DPA_")
							Set levyRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
							If (levyRS.EOF Or levyRS.BOF) Then%>
				         				<script language = 'vbscript'>
				         						ShowMessage "The levies for the selected contract are missing. This is a serious database corruption."
							         			
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
			WriteDialogRelocateScript "EditLotItem.asp?ID=" & IDHolder
			Response.End
			
		case "EXECUTE_DELETE"
			Set conn = GetActiveConnection("KBroker")
			
			'obtain levies, contract and lot to be deleted
			sqlStr = "SELECT * FROM LevyContractList WHERE Lot_DPA_=" & ItemID
			Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
			If (rs.EOF Or rs.BOF) Then%>
				        <script language = 'vbscript'>
				         		ShowMessage "No Levies were found for this Lot. This is a serious database corruption."
							    
				        </script>
				        <% response.end
			End If
			Dim contractID
			Dim lotID
			
			contractID = rs.fields("Contract_DPA_")
			lotID = rs.fields("Lot_DPA_")
			conn.BeginTrans
				'delete levies
					do until rs.eof
							DeleteItem "LevyContract","LevyContract_DPA_",rs.fields("LevyContract_DPA_")
							rs.movenext
					loop
				'delete lot
					DeleteItem "Lot","Lot_DPA_",lotID
				'delete contract
					DeleteItem "Contract","Contract_DPA_",contractID
			conn.CommitTrans
			conn.Close
			Set conn = Nothing
			WriteDialogRelocateScript "EditLotItem.asp?ID=" & IDHolder
			Response.End
   	
   	case else
   			sqlStr = "SELECT * FROM OrdDetailList WHERE OrdDetailList.OrdDetail_DPA_= " & ID
		   	
   			Set conn = GetActiveConnection("KBroker")
   			set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
   			If rs.EOF Or rs.BOF Then%>
					<script language = 'vbscript'>
                			window.self.ShowMessage "The selected Order item cannot be retrieved for lot allocation"
		                	
					</script>
					<% response.end
			End If
   	
   	end select
   		
   	
%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit <%=DataEntity%> Item</title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
 <script language='vbscript'>
			function ItemSelected(itemID)
					
 					frm<%=DataSource%>Item.elements("ItemID").value = itemID
			end function
			
			
			function SaveInPlaceEdit()
				    Dim myOwnerFrame				
					'UpdateID
					Set window.parent.dialogArguments.opener.parent.frames("footer").editDocOpener = window.self
					frm<%=DataSource%>Item.target = "deleteFrame" 					
					frm<%=DataSource%>Item.submit
			end function
		</script>
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
		
		<!-- ActiveUI stylesheet and scripts -->
		<link href="../runtime/classic/activeui.css" rel="stylesheet" type="text/css">
		<script src="../runtime/activeui.js"></script>
		<!-- Include patches here -->
		<script src="../runtime/paging1.js"></script>
		<!-- grid format -->
		<style> 
			.active-controls-grid {height: 100%; font: menu;}
			.active-row-highlight .active-row-cell {background-color: skyblue}
		    
		    
		     	
			.active-column-0 {width: 50px;}
			.active-column-1 {width: 70px;}
			.active-column-2 {width: 120px;}
			.active-column-3 {width: 80px;}
			.active-column-4 {width: 80px;}
			.active-column-5 {width: 200px;}
			.active-column-6 {width: 100px;}
			
			
			
			.active-grid-row,
			.active-grid-row.active-list-item,
			.active-scroll-left .active-list-item {height: 22px;}
			
			
			.active-selection-true, .active-selection-true .active-row-cell {
				color: blue!important;
				background-color: bisque!important;
				}
		</style>

</head>

<body Class="Dialog" SCROLL="NO">	
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<form name = 'frm<%=DataSource%>Item' id='frmMain' method = 'post' action = '<%=DataSource%>Item.asp' OnSubmit="UpdateID();">

<SCRIPT language="JavaScript">
	var calTDate;
	function changeDateInterface(selCol){
		try{
			calTDate = new ctlSpiffyCalendarBox('calTDate', 'frm<%=DataSource%>Item', 'txtTDate', 'cmdTDate','<%= FormatDate(Date) %>', 1); 
			calTDate.readonly = false;
			calTDate.returnOutStringOnWrite(); 
			//var parentDiv = document.all.item("txtTDate").parentNode;
			//parentDiv.innerHTML = calTDate.writeControl();
			//if (selCol==null || selCol == "undefined"){
				document.all.item("txtTDate").outerHTML = calTDate.writeControl();
		//	}
		//	else{
		//		document.all.item(selCol).outerHTML = calTDate.writeControl();
				
		//	}	
			//parentDiv.style.zIndex = 10;
			//parentDiv.childNodes(1).style.zIndex = 10	;
		}
		
		catch(e){}	
		
		document.all.item('txtSlip').focus();
	}
	
	document.body.onload = changeDateInterface;
	
	//function ShowGrid()
	//{
	///	ShowMessage(document.all.item("GridCell").innerText);
	//}
</SCRIPT>
 
  
 <%
 Set conn = GetActiveConnection("KBroker")
 
	Dim rowCount
	Dim brokerList
	Dim quote 
	
	quote =  chr(34)
	brokerList = GetBrokerList("cboBroker")
	       
    sqlStr = "SELECT * FROM LotList WHERE LotList.OrdDetail_DPA_= " & ID
       
    Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
    If Not(rs.EOF Or rs.BOF) Then
		'store the Order Type%>
        <input type = 'hidden' name ='txtOrderType' id = 'txtOrderType' size='9' value = '<%=rs.Fields("OrdDetailType")%>'>
        <input type = 'hidden' name ='txtInstrument' id = "txtInstrument" size='9' value = '<%=rs.Fields("OrdDetailSecType")%>'>
        <input type = 'hidden' name ='txtOrderIsSaleType' id = "txtOrderIsSaleType" size="20" value = '<%=rs.Fields("OrderTypeSale")%>'>
        <input type = 'hidden' name ='txtSecurityID' id = "txtSecurityID" size="20" value = '<%=rs.Fields("Security_DPA_")%>'>
        <input type = 'hidden' name ='txtCommission' id = "txtCommission" size="20" value = '<%=rs.Fields("CommissionRate")%>'>
        <input type = 'hidden' name ='txtAgentCommission' id = "txtAgentCommission" size="20" value = '<%=rs.Fields("AgentCommission")%>'>
        <input type = 'hidden' name ='txtStaffCommission' id = "txtStaffCommission" size="20" value = '<%=rs.Fields("StaffCommission")%>'>
		<!-- grid data -->
		<% 'row data
		
		rowCount = 0 
		rs.MoveFirst
		Do Until rs.EOF 

'======================= Begin_Alter_Across_Entities =================================
			'row ID 
			
			rowData = rowData & quote & rs.Fields("Lot_DPA_") & quote & " : " 
			
			'row data 
			rowData = rowData & "[" 
			rowData = rowData & quote & rs.Fields("Lot_DPA_") & quote & "," 
			rowData = rowData & quote & rs.Fields("LotSlipNo") & quote & ","
			rowData = rowData & quote & FormatDate(rs.Fields("LotTDate")) & quote & ","
			rowData = rowData & quote & rs.Fields("LotQty") & quote & ","
			rowData = rowData & quote & rs.Fields("LotPrice") & quote & ","
			rowData = rowData & quote & rs.Fields("BrokerCode") & " : " & rs.Fields("BrokerName") & quote & ","
			rowData = rowData & quote & " " & quote  
			rowData = rowData & "]" 
			
			rowIDs = rowIDs & quote & rs.Fields("Lot_DPA_") & quote 
			rowCount = rowCount + 1
		
		
			rs.MoveNext 
			
			
				'build the row IDs array
				rowIDs = rowIDs & "," 
				rowData = rowData & ","	
			
'======================= End_Alter_Across_Entities =================================

			
		Loop

		
		rs.MoveFirst
		%><script language="javascript">
			//update the quantity balance
			try{
				window.parent.frames("header").document.frmMain.elements("txtBalance").value = '<%= FormatNum(rs.Fields("BalanceQty")) %>';
			}
			catch(e){}	
		 </script><%
	End if
	
		'row ID 	
		rowData = rowData & quote & -1 & quote & " : " 
				
		'row data 
		rowData = rowData & "[" 
		rowData = rowData & quote & "New Line" & quote & "," 
		'rowData = rowData & quote & "<input type = 'text' name ='txtSlip' id = 'txtSlip' size='9' onChange = 'AddRowInProgress();event.cancelBubble=true;' OnClick='event.cancelBubble=true;'>" & quote & ","
		'rowData = rowData & quote & "<input type='text' name='txtTDate' size=40 value='" & FormatDate(Date) & "' onChange = 'AddRowInProgress();event.cancelBubble=true;' OnClick='event.cancelBubble=true;'>" & quote & ","
		'rowData = rowData & quote & "<input type = 'text' name ='txtQty' id = 'txtQty' size='9' onChange = 'AddRowInProgress();event.cancelBubble=true;' OnClick='event.cancelBubble=true;'>" & quote & ","
		'rowData = rowData & quote & "<input type = 'text' name ='txtPrice' id = 'txtPrice' size='9' onChange = 'AddRowInProgress();event.cancelBubble=true;' OnClick='event.cancelBubble=true;'>" & quote & ","
		rowData = rowData & quote & "<input type = 'text' name ='txtSlip' id = 'txtSlip' size='7'  OnClick='event.cancelBubble=true;'>" & quote & ","
		rowData = rowData & quote & "<input type='text' name='txtTDate' size=40 value='" & FormatDate(Date) & "' OnClick='event.cancelBubble=true;'>" & quote & ","
		rowData = rowData & quote & "<input type = 'text' name ='txtQty' id = 'txtQty' size='9' OnClick='event.cancelBubble=true;'>" & quote & ","
		rowData = rowData & quote & "<input type = 'text' name ='txtPrice' id = 'txtPrice' size='9' OnClick='event.cancelBubble=true;'>" & quote & ","
		rowData = rowData & quote & brokerList & quote  & ","
		rowData = rowData & quote & "<input type=button value='Add' Class=Buttons OnClick='JavaScript: AddRowInProgress();'>&nbsp;&nbsp;<input type='reset' value='Cancel' Class=Buttons>" & quote 
		rowData = rowData & "]" 
	

		'build the row IDs array 
		rowIDs = rowIDs & quote & -1 & quote 
		rowCount = rowCount + 1
		
'======================= Begin_Alter_Across_Entities =================================%> 
		<script language="javascript">
			//column titles 
			var colCount = 7;
			var colNames = ["", "Slip",  
					 "Date", "Quantity","Price","Broker", ""];
			
			var myColumns = ["Lot No", "Slip",  
					 "Date", "Quantity","Price","Broker", ""];
		</script>
<%'======================= End_Alter_Across_Entities =================================%>			
		<script language="javascript">
		
			
			//data
			var myData = {<%=rowData%>}; 
			var myRowIDs = [<%=rowIDs%>]; 
			
			
			//editing
			var inPlaceEdit = false;
			var addInProgress = false;
			var clickedRowID = -1; 
			var dataChanged = false;
			var prevRow = -1;//the row currently under in-place edit
			
			function EditInPlaceDataChanged()
			{
				dataChanged = true;
			}
			
			function AddRowInProgress()
			{
				addInProgress = true;
			}
			
			var currentBrokerName = "";
			var RowEditFn = function(src)
			{
				var rowIndex = src.getProperty("row/index");
				var i;
				if (rowIndex==0) return;
				for(i = 0; i < colCount; i++)
				{
					if(colNames[i] != "")
					{
						if(prevRow >= 0)
						{
							if(colNames[i]=="Broker")
							{
								myData[prevRow][i] = currentBrokerName;
							}
							else
							{
								myData[prevRow][i] = document.frmMain.elements(colNames[i]).value;
							}
						}
						if(colNames[i]=="Broker")
						{	
							currentBrokerName =  myData[rowIndex][i];
							myData[rowIndex][i] = inPlaceList;
							
						}
						else
						{
							if(colNames[i]=="Date")
							{
								currentDate =  myData[rowIndex][i];
								myData[rowIndex][i] = "<INPUT TYPE='text' NAME='" + colNames[i] + "' ID='" + colNames[i] + "' VALUE='<%=FormatDate(Date)%>' onChange = 'EditInPlaceDataChanged();event.cancelBubble=true;' OnClick='event.cancelBubble=true;'>";
								//show the calendar
								changeDateInterface(colNames[i]);
				
								
							}
							else {
								myData[rowIndex][i] = "<INPUT TYPE='text' NAME='" + colNames[i] + "' ID='" + colNames[i] + "' VALUE='" + myData[rowIndex][i] + "' onChange = 'EditInPlaceDataChanged();event.cancelBubble=true;' OnClick='event.cancelBubble=true;'>";
							}	
						}
					}
				}
				
				myData[rowIndex][colCount - 1] = "<INPUT TYPE='button' class='Buttons' VALUE='Save' onClick = 'SaveInPlaceEdit();event.cancelBubble=true;'>&nbsp;<INPUT TYPE='button' class='Buttons' VALUE='Cancel' onClick = 'cancelEditRow();event.cancelBubble=true;'>";
				inPlaceEdit = true;
				prevRow = rowIndex;
				grid.refresh();
				//select the appropriate item
				var secList = document.frmMain.elements("cboBrokerInPlace");
				for (i=0; i < secList.options.length; i++) {
					if(secList.options(i).text == currentBrokerName)
					{
							secList.options(i).selected = true;
					}
				}
				
				
			}
			
			function cancelEditRow(){
					var i;
							for(i = 0; i < colCount; i++)
							{
								if(colNames[i] != "")
								{
									if(colNames[i]=="Broker")
									{
										myData[prevRow][i] = currentBrokerName;
									}
									else
									{
										myData[prevRow][i] = document.frmMain.elements(colNames[i]).value;
									}
								}
							}
							myData[prevRow][colCount - 1] = "";						
							inPlaceEdit = false;
							prevRow = -1;
							grid.refresh();
			}
			
			var RowChangeFn = function(src)
			{
				if(inPlaceEdit || addInProgress)
				{
					if(dataChanged || addInProgress)
					{
						ItemSelected(prevRow);
						SaveInPlaceEdit();
					}
					else
					{
						if(prevRow != clickedRowID)
						{
							cancelEditRow();
						}
					}
				}
			}
			
			var HandleClick = function(src)
			{
				clickedRowID = src.getProperty("row/index");
				ItemSelected(clickedRowID);
			}
			
			var headerID;
			
			function UpdateID(){
				headerID = window.parent.frames["header"].document.all.item("ID").value;
				document.all.item("ID").value = headerID;
			}
			
			function restoreID(){
				document.all.item("ID").value = headerID;
			}
			
			function HandleDeleteAction()
			{
					document.frmMain.elements("action").value = "Execute_Delete"
					SaveInPlaceEdit();
					document.frmMain.elements("action").value = "Execute_Detail"
			}
			
			//get ready for in-place edit
			var inPlaceList = "<%=GetBrokerList("cboBrokerInPlace")%>"
		</script> 
		
		<script language="javascript"> 

			// create ActiveUI Grid javascript object 
			var grid = new Active.Controls.Grid; 
			
			
			// set rows ids 
			grid.setRowValues(myRowIDs); 
			

			// set number of columns 
			grid.setColumnCount(colCount); 

			// provide cells and headers text 
			grid.setDataText(function(i, j){return myData[i][j]}); 
			grid.setColumnText(function(i){return myColumns[i]}); 

			// set click action handler 
			grid.setAction("click", HandleClick); 
			grid.setAction("dblclick", RowEditFn); 
			grid.setAction("selectionChanged", RowChangeFn);
			
			//stripes 
			var alternate = function(){ return this.getProperty("row/order") % 2 ? "gainsboro" : "white";} 
			var row = new Active.Templates.Row; row.setStyle("background", alternate); 
			row.setEvent("onmouseover", "mouseover(this, 'active-row-highlight')"); 
			row.setEvent("onmouseout", "mouseout(this, 'active-row-highlight')"); 
			grid.setTemplate("row", row); 
			var column = new Active.Templates.Text; 
			column.setStyle("border-right", "1px solid white");  
			grid.setTemplate("column", column);  grid.setRowHeaderWidth("0px"); 
			
			//disable sort
			grid.getTemplate("top/item").setEvent("onmousedown", null);
			
			// write grid html to the page 
			document.write(grid); 
			
			//let grid be aware of composite layout
			grid.getLayoutTemplate().action("adjustSize");
		</script> 
        <%
       
  function GetBrokerList(listName)
		Dim secList
		
		secList = "<select name = '" & listName & "' id = '" & listName & "' size='1' "
		secList = secList & "OnClick='event.cancelBubble=true;' "  
		secList = secList & "onChange='event.cancelBubble=true;' " 
		secList = secList & "onKeypress='return (dodefaultaction()==\""\""); ' "  
		secList = secList & "onKeydown='return (dodefaultaction()==\""\"");event.cancelBubble=true;' "  
		secList = secList & "onKeyup='return (change(" & listName & "));' "  
		secList = secList & "onfocus='txtval = \""\"";inputIsItemCode = 1;' "  
		secList = secList & "onblur='txtval = \""\"";inputIsItemCode = 1;'>"
		
		secList = secList & "<option selected SearchCode = '0' SearchText = ''  value = ''></option>"
		
        sqlStr = "SELECT * FROM [BrokerList] Order By BrokerCode"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF
                        secList = secList & "<option SearchCode = '" & rs.Fields("BrokerCode") & "' SearchText = '" & rs.Fields("BrokerName") & "'  value = '" & rs.Fields("Broker_DPA_") & "'>" & rs.Fields("BrokerNameEx") & "</option>"
                        rs.MoveNext
                Loop
        End If
	    secList = secList & "</select>"
	    GetBrokerList = secList
  end function
  
  function DeleteItem(EntityName,KeyField,DelItemID)
		dim delRS
		'find out whether any child records exist
		sqlStr = "SELECT Child,DeletionMessage,ParentKey FROM [_Parent_Child_Links_] WHERE (Parent = '" & EntityName & "') AND (ChildType = " & LinkedIndependent & ")"
		Set delRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		If Not (delRS.BOF Or delRS.EOF) Then
				Dim childRS
				Dim tableName
				
				delRS.MoveFirst
				Do Until rs.EOF
                		tableName = delRS.Fields("Child")
						sqlStr = "SELECT TOP 1 * FROM [" & tableName & "] WHERE " & KeyField & " = " & DelItemID
						Set childRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
						If Not (childRS.BOF Or childRS.EOF) Then%>
                				<script language = 'vbscript'>
                					ShowMessage "<%=delRS.Fields("DeletionMessage")%>"
                					
                				</script>
                				<%response.end
						End If
						delRS.MoveNext
				Loop
		End If
		
		'delete from database
		sqlStr = "DELETE FROM [" & EntityName & "] WHERE " & KeyField & " = " & DelItemID
		conn.Execute SQLServerFormat(HandleQuote(sqlStr))
  end function
 %>
 <table border="0" width="100%" ID="Table1">
<tr><td>
<input type = 'hidden' name ='ItemID' id = 'ItemID'>
<input type = 'hidden' name ='ID' id = 'ID' value="<%= IDHolder %>">
<input type = 'hidden' name ='action' id = 'action' value="Execute_Detail">
</td>
</tr>
</table>
</form>
</body>
</html>
