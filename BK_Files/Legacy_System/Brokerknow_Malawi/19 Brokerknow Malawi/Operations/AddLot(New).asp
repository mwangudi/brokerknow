<!--#include file="../libroutines.asp"-->
<%
	
	'======================= Begin_Alter_Across_Entities =================================
		
		const UDLName = "KBroker"
		const DataSource = "AddLot"
		const DataEntity = "Lot"
		const DataEntityPlural = "Lots"
		const ActionFolder = "Operations"
'======================= End_Alter_Across_Entities =================================
	
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
   Dim guidStr 
   Dim guid 
	
	action = ucase(Request.Form("action"))
	
	if action = "EXECUTE" then
		Dim buttonAction
		Dim reloadRequired
		
		reloadRequired = false
		buttonAction = Trim(Trim(Ucase(Request.Form("cmdAdd"))))
		if buttonAction = "SAVE" then
				ID = Request("ID")

				If Trim(ID) = "" Then%>
						<script language = 'vbscript'>
                				ShowMessage "No Order item was specified for Lot allocation"
                				window.self.close
						</script>
						<%response.end
				End If
				
				Dim slip
				Dim broker
				Dim tDate
				Dim qty
				Dim price
				Dim orderType
				Dim orderIsSaleType
				Dim securityID
				Dim commission
				Dim agentCommission
				Dim staffCommission
			 
				broker = Request.Form("cboBroker")
				tDate = Trim(Request.Form("txtTDate"))
				tDate = tDate & " " & Time				
				slip = Request.Form("txtSlip")
				qty = Request.Form("txtQty")
				price = Request.Form("txtPrice")
				orderType = Request.Form("txtOrderType")
				orderSecType = Request.Form("txtInstrument")
				orderIsSaleType = cbool(Request.Form("txtOrderIsSaleType"))
				securityID = Request.Form("txtSecurityID")
				commission = Request.Form("txtCommission")
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
						ShowMessage "Order Detail Estimated Price must be numeric"
						
				    </script>
				    <% response.end
				End If
				'ensure Order Detail Estimated Quantity is numeric
				If (Qty <> "") And (Not IsNumeric(Qty)) Then%>
				    <script language = 'vbscript'>
						ShowMessage "Order Detail Estimated Quantity must be numeric"
						
				    </script>
				    <% response.end
				End If
				         
				 'save contract
				 set guid = server.createobject("NDUtils.CGUID")
				 guidStr = guid.GenerateGUID
				 
				 sqlStr = "INSERT INTO [Contract] (Contract_DPA_, Contract_EIT_, Status_DPA_) " & _
				         "SELECT " & " " & "iif(isnull(max([Contract_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Contract'),max([Contract_DPA_]) + 1)" & " " & " as Contract_DPA_" & _
				         "," & "'" & guidStr & "'" & " as Contract_EIT_" & _
				         "," & " " & 1 & " " & " as Status_DPA_" & _
				         " FROM [Contract]"
				 Set conn = GetActiveConnection("KBroker")
				 conn.BeginTrans
						conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
				     
						'obtain contract number
						sqlStr = "SELECT [Contract.Contract_DPA_] FROM [Contract] WHERE [Contract.Contract_EIT_] = " & "'" & guidStr & "'"
				     
						Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
						If (rs.EOF Or rs.BOF) Then%>
				         			<script language = 'vbscript'>
				         					ShowMessage "A serious error has been encountered while saving the data. Try saving again"
				         					
				         			</script>
				         			<% response.end
						End If
				     
				     
						'calculate amounts
						Dim grossAmount  'this is the amount before application of Levies
					    
						
						if orderSecType = "Fixed" then ' "F" is FIXED security
							grossAmount = (price * qty) / 100
						else
							grossAmount = price * qty
						end if
						
						
						'save lot
						sqlStr = "INSERT INTO [Lot] (Lot_DPA_,Contract_DPA_,OrdDetail_DPA_,LotPrice" & _
								",LotQty,LotSlipNo,LotTDate,Broker_DPA_,ContractNumber, LotGrossAmount)" & _
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
						Dim levyRS
						Dim cond
						if orderSecType = "Fixed" then 
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
										if levyRS.Fields("LevyType") = "P" Then ' "P" is for PERCENTAGE ie Levies that are a percentage of Gross
												levyAmount = CCur((levyRS.Fields("LevyAmount")/100.00) * grossAmount)
												LevyRatePercentage = levyRS.Fields("LevyAmount") & "%"
										else
												Dim blocks
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
				conn.CommitTrans
				
				WriteDialogRelocateScript "EditLot.asp?ID=" & ID
				Response.end
		end if

	end if
	
		
	Dim IDHolder
	Dim IDArray
	Dim ItemID
	
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
	
	if (ItemID <> 0) then
			Response.redirect "EditLot.asp?ID=" & ID
			Response.end
	end if
   	sqlStr = "SELECT * FROM OrdDetailList WHERE OrdDetailList.OrdDetail_DPA_= " & ID
   	
   	Set conn = GetActiveConnection("KBroker")
   	set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
   	If rs.EOF Or rs.BOF Then%>
            <script language = 'vbscript'>
                	window.self.ShowMessage "The selected Order item cannot be retrieved for lot allocation"
                	
            </script>
            <% response.end
    End If
    
%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add <%=DataEntity%></title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
<SCRIPT language="JavaScript">
	var calTDate=new ctlSpiffyCalendarBox("calTDate", "frm<%=DataSource%>", "txtTDate","cmdTDate","<%=FormatDate(Date)%>",1);
</SCRIPT>
<!--END CALENDAR -->

<script >
		var validNavigate = false;
		function ReleaseRecord()
		{
			if(!validNavigate)
			{
 				event.returnValue = "Please use the cancel button to close the dialog"
 			}
		}
		
		function AllowedNavigation()
		{
			validNavigate = true;
		}
</script>
</head>

<body Class="Dialog">
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>

<form name = 'frm<%=DataSource%>' method = 'post' action = '<%=DataSource%>.asp' target = 'deleteFrame' id='frmMain' OnSubmit="JavaScript: UpdateDialogHandle();">


<table border="0" width="100%" height="265">
  <tr>
    <td width="17%" height="25">Order No</td>
    <td width="83%" height="25"><input readonly = 'true' class=readonly   STYLE="width: 100px; text-align: right" type = 'text' name ='txtOrderNo' id = 'txtOrderNo' size="20" value = '<%=rs.Fields("Order_DPA_")%>'></td>
  </tr>
  <tr>
    <td width="17%" height="25">Item No</td>
    <td width="83%" height="25"><input readonly = 'true' class=readonly   STYLE="width: 100px; text-align: right" type = 'text' name ='txtItemNo' id = 'txtItemNo' size="20" value = '<%=rs.Fields("OrdDetail_DPA_")%>'></td>
  </tr>
  <tr>
    <td width="17%" height="25">Order Type</td>
    <td width="83%" height="25"><input readonly = 'true' class=readonly  STYLE="width: 150px"   type = 'text' name ='txtOrderType' id = 'txtOrderType' size="20" value = '<%=rs.Fields("OrdDetailType")%>'>
    <input readonly = 'true' class=readonly  type = 'hidden' name ='txtOrderIsSaleType' id = "txtOrderIsSaleType" size="20" value = '<%=rs.Fields("OrderTypeSale")%>'></td>
  </tr>
  <tr>
    <td width="17%" height="25">Instrument</td>
    <td width="83%" height="25"><input readonly = 'true' class=readonly STYLE="width: 150px"  type = 'text' name ='txtInstrument' id = 'txtInstrument' size="20" value = '<%=rs.Fields("OrdDetailSecType")%>'></td>
  </tr>
   <tr>
    <td width="17%" height="25">Client</td>
    <td width="83%" height="25">
<input readonly = 'true' class=readonly  type = 'text' name ='txtClient'  STYLE="width: 300px" id = 'txtClient' size="20" value = '<%=rs.Fields("OrdDetailClient")%>'>
<input type = 'hidden' name ='txtCommission' id = "txtCommission" size="20" value = '<%=rs.Fields("CommissionRate")%>'></td>
<input type = 'hidden' name ='txtAgentCommission' id = "txtAgentCommission" size="20" value = '<%=rs.Fields("AgentCommission")%>'></td>
<input type = 'hidden' name ='txtStaffCommission' id = "txtStaffCommission" size="20" value = '<%=rs.Fields("StaffCommission")%>'></td>
  </tr>

  <tr>
    <td width="17%" height="25">Security</td>
    <td width="83%" height="25"><input readonly = 'true'  STYLE="width: 300px"   class=readonly  type = 'text' name ='txtSecurity' id = 'txtSecurity' size="20" value = '<%=rs.Fields("OrdDetailSecurity")%>'>
    <input type = 'hidden' name ='txtSecurityID' id = "txtSecurityID" size="20" value = '<%=rs.Fields("Security_DPA_")%>'></td>
  </tr>
  <tr>
    <td width="17%" height="25">Balance</td>
    <td width="83%" height="25"><input readonly = 'true' STYLE="width: 150px; text-align: right"  class=readonly  type = 'text' name ='txtBalance' id = 'txtBalance' size="20" value = '<%= FormatNum(rs.Fields("BalanceQty")) %>'></td>
  </tr>
  <tr>
  <td colspan = '2' height="58">
  
  <table border="0" width="100%">
    <tr>
      <td width="23%"><b><font color="#000080">CDS Ref</font></b></td>
      <td width="13%"><b><font color="#000080">Date&nbsp;</font></b></td>
      <td width="17%"><b><font color="#000080">Qantity</font></b></td>
      <td width="14%"><b><font color="#000080">Price</font></b></td>
      <td width="33%"><b><font color="#000080">Broker</font></b></td>
    </tr>
    <tr>
      <td width="23%" valign="top"><input type = 'text' name ='txtSlip' id = 'txtSlip' size="15">&nbsp;</td>
      <td width="13%" valign="top"><SCRIPT language="JavaScript">calTDate.writeControl();</SCRIPT></td>
      <td width="17%" valign="top"><input type = 'text' name ='txtQty' id = 'txtQty' size="11"></td>
      <td width="14%" valign="top"><input type = 'text' name ='txtPrice' id = 'txtPrice' size="10"></td>
      <td width="33%" valign="top"><select name="cboBroker" id="cboBroker" size="1" 
			onKeypress="return (dodefaultaction()==''); " 
			onKeydown="return (dodefaultaction()==''); " 
			onKeyup="return (change(cboBroker));" 
			onfocus="txtval = '';inputIsItemCode = 1;" 
			onblur="txtval = '';inputIsItemCode = 1;">
          <option selected SearchCode = "0" SearchText = ""  value=""></option>
          <%
		Set conn = GetActiveConnection("KBroker")
        sqlStr = "SELECT * FROM [BrokerList] Order By BrokerCode"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF%>
					<option SearchCode = "<%=rs.Fields("BrokerCode")%>" SearchText = "<%=rs.Fields("BrokerName")%>" value = '<%=rs.Fields("Broker_DPA_")%>'><%=rs.Fields("BrokerNameEx")%></option>
					<%rs.MoveNext
                Loop
        End If
          %>
        </select>
      </td>
    </tr>
  </table>
  
  </td>
  </tr>
  <tr>
    <td width="100%" colspan=2 align=right>
		<input type = 'submit' class=buttons name ='cmdAdd' id = 'cmdAdd' value=" Save " onclick = "AllowedNavigation()">
    	&nbsp; <input type = 'button' class=buttons name ='cmdClose' id = "cmdClose" value=" Close " onclick = "window.self.close();">
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
	</td>
  </tr>
</table>

</form></body>

</html>
