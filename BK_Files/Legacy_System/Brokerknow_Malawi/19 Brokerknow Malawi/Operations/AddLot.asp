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
	
	UserId=Session("UserID")
	
	action = ucase(Request.Form("action"))
	
	if action = "EXECUTE" then


		Dim buttonAction
		Dim reloadRequired
		
		reloadRequired = false
		buttonAction = Trim(Trim(Ucase(Request.Form("buttonAction"))))
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
				Dim varBalanceQty
				Dim price
				Dim orderType
				Dim orderIsSaleType
				Dim securityID
				Dim commission
				Dim agentCommission
				Dim staffCommission
				Dim volComm
				Dim volBound
				Dim minComm
				Dim cma
				Dim imobRate
				Dim secImob
				Dim regularComm
			 	Dim interbank
				Dim ContractDPA
				Dim custodian
				Dim client
				Dim cliententity
				Dim ClientClass
				Dim ContractSettlementDate 

				volComm = Request.Form("txtVolumeCommission")
				volBound = ccur(Request.Form("txtVolumeBoundary"))
				minComm = ccur(Request.Form("txtMinimumCommission"))
				cma = Request.Form("txtCMA")
				imobRate = Request.Form("txtPostImmobilisedRate")
				secImob = Request.Form("txtSecurityImmobilised")
				broker = Request.Form("cboBroker")
				tDate = Trim(Request.Form("txtTDate"))
				tDate = tDate & " " & Time				
				slip = Request.Form("txtSlip")
				qty = Request.Form("txtQty")
				varBalanceQty = Request.Form("txtBalance")
				price = Request.Form("txtPrice")
				orderType = Request.Form("txtOrderType")
				orderSecType = Request.Form("txtInstrument")
				orderIsSaleType = cbool(Request.Form("txtOrderIsSaleType"))
				securityID = Request.Form("txtSecurityID")
				agentCommission = Request.Form("txtAgentCommission")
				staffCommission = Request.Form("txtStaffCommission")
				regularComm = Request.Form("txtCommission")
				interbank=Cint(Request.Form("txtinterbank"))       
       			custodian=Cint(Request.Form("txtcustodian"))       
       			client =Request.Form("txtClientDPA")       
       			cliententity =Cint(Request.Form("txtEntityDPA"))       
       			clientClass =Cint(Request.Form("txtClass"))
				ContractSettlementDate =trim(Request.Form("txtSettleDate"))

       			Set TimeLimitRs = CreateObject("ADODB.Recordset")   						        
   				TimeLimitRs.CursorLocation = adUseClient
   				
   				'Set CustodianRS	= CreateObject("ADODB.Recordset")   						        
   				'CustodianRS.CursorLocation = adUseClient
   				   				
				Set conn = GetActiveConnection("KBroker")
			     
				'validate Settlement Date
				 If Trim(ContractSettlementDate) = "" Then%>
				         <script language = 'vbscript'>
				         		ShowMessage "Please specify the Settlement Date."
				         		
				         </script>
				         <% response.endReloadPage(ID)
				 End If

				 'validate Broker
				 If Trim(Broker) = "" Then%>
				         <script language = 'vbscript'>
				         		ShowMessage "Please specify the Broker"				         		
				         </script>					
				         <% ReloadPage(ID) 
						 response.end 
						 
				 End If
				 
				 'validate Slip
				 If Trim(Slip) = "" Then%>
				         <script language = 'vbscript'>
				         		ShowMessage "Please specify the Ref No."
				         		
				         </script>
				         <%ReloadPage(ID) 
						 response.end 
				 End If
				 'ensure Slip is numeric
				'If (Not IsNumeric(Slip)) Then%>
				    <script language = 'vbscript'>
						'ShowMessage "Slip No. must be numeric"
						
				    </script>
				    <% 'response.end
				'End If
				
				'validate Estimated Price
				If Trim(Price) = "" Then%>
				    <script language = 'vbscript'>
				         	ShowMessage "Please specify the Price "
				         	
				    </script>
				    <% ReloadPage(ID) 
						 response.end 
				End If
				'validate Estimated Quantity
				If Trim(qty)= "" Then%>
				    <script language = 'vbscript'>
				         	ShowMessage "Please specify the Quantity "
				         	
				    </script>
				    <% ReloadPage(ID) 
						 response.end 
				End If
			'ensure Order Detail Estimated Price is numeric
				If (Price <> "") And (Not IsNumeric(Price)) Then%>
				    <script language = 'vbscript'>
						ShowMessage "Order Detail Estimated Price must be numeric"
						
				    </script>
				    <% ReloadPage(ID) 
						 response.end 
				End If
				'ensure Order Detail Estimated Quantity is numeric
				If (Qty <> "") And (Not IsNumeric(Qty)) Then%>
				    <script language = 'vbscript'>
						ShowMessage "Order Detail Estimated Quantity must be numeric"
						
				    </script>
				    <% ReloadPage(ID) 
						 response.end 
				End If

				
			'ensure balance qty on order is not negative
			If cdbl(varBalanceQty) <= 0 Then%>
				<SCRIPT LANGUAGE="JAVASCRIPT">					
					alert ("Order already filled. Please place a new order.");
				</SCRIPT>
				<% 
				ReloadPage(ID) 
						 response.end 
			End If

			'ensure balance qty is not exceeded
			If cdbl(varBalanceQty) < cdbl(qty) Then%>
				<SCRIPT LANGUAGE="JAVASCRIPT">					
					alert ("Specified quantity exceeds balance quantity on order.");
				</SCRIPT>
				<% 
				ReloadPage(ID) 
						 response.end 
			End If

				sqlStr = "execute cont_CreateContract  " & ID & ", " & broker & ", " & price & ", " & qty & ", '" & slip & "', '" & FormatDate(tDate) & "', '" & ContractSettlementDate & "', " & UserId & ""

				conn.BeginTrans
					conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
				conn.CommitTrans

				%>
						<SCRIPT LANGUAGE="JAVASCRIPT">
							//alert('hapa');
							//window.parent.parent.frames['maininfo'].location.reload();
							//window.opener.location= window.opener.location;
							//alert('hapa2');
							this.location='EditLot.asp?ID=<%=ID%>';
							//window.parent.location.href=EditLot.asp?ID=<%=ID%>
						</SCRIPT>
						<%
				
				'WriteDialogRelocateScript2 "EditLot.asp?ID=" & ID
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
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add <%=DataEntity%></title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
<%
'set default settlement date
'Calculate Holidays

 'DefaultDate =conn.execute("SELECT     CONVERT(varchar(20), CASE WHEN DATEPART(dw, GETDATE() - 0) = 2 THEN getdate() - 0 + 4 WHEN 'DATEPART(dw, GETDATE() - 0) = 1 THEN getdate() - 0 + 7 - 3 WHEN DATEPART(dw, GETDATE() - 0) = 7 THEN getdate() - 0 + 7 - 2 ELSE 'getdate() - 0 + 6 END, 106) AS SettlementDate")(0)' FormatDate(dateadd("d",5,Now()))

 DefaultDate =formatdate(conn.execute("select   getdate() - 0 + 7 ")(0))
%>
<SCRIPT language="JavaScript">
	var calTDate = new ctlSpiffyCalendarBox("calTDate", "frm<%=DataSource%>", "txtTDate","cmdTDate","<%=FormatDate(Date)%>",1);
	var calDate = new ctlSpiffyCalendarBox("calDate", "frm<%=DataSource%>", "txtSettleDate","cmdSettleDate","<%=DefaultDate%>",1);
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
		function hideButton()
		{
		 document.getElementById('cmdAdd').style.display='none';
		}

		function forceSubmit()
	{
		setOpener();
		//var targetPage = window.dialogArguments.opener.document.all.item("frmMain").elements("EditPage").value
				
		document.frm<%=DataSource%>.method='post';
		document.frm<%=DataSource%>.target='_self';
		document.frm<%=DataSource%>.submit();	
		
	}
	
	function setOpener()
	{

		window.self.opener = window.dialogArguments.opener;
				//alert(window.dialogArguments.opener.location);
	}

</script>
</head>

<body Class="Dialog" onLoad="javascript: setOpener()">
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
		<input type = 'hidden' name ='txtAgentCommission' id = "txtAgentCommission" size="20" value = '<%=rs.Fields("AgentCommission")%>'></td>
		<input type = 'hidden' name ='txtStaffCommission' id = "txtStaffCommission" size="20" value = '<%=rs.Fields("StaffCommission")%>'></td>
		<input type = 'hidden' name ='txtCommission' id = "txtCommission" size="20" value = '<%=rs.Fields("CommissionRate")%>'></td>
		<input type = 'hidden' name ='txtVolumeCommission' id = "txtVolumeCommission" size="20" value = '<%=rs.Fields("VolumeRate")%>'></td>
		<input type = 'hidden' name ='txtVolumeBoundary' id = "txtVolumeBoundary" size="20" value = '<%=rs.Fields("VolumeBoundary")%>'></td>
		<input type = 'hidden' name ='txtMinimumCommission' id = "txtMinimumCommission" size="20" value = '<%=rs.Fields("MinimumCommission")%>'></td>
		<input type = 'hidden' name ='txtCMA' id = "txtCMA" size="20" value = '<%=rs.Fields("CMARegulated")%>'></td>
		<input type = 'hidden' name ='txtPostImmobilisedRate' id = "txtPostImmobilisedRate" size="20" value = '<%=rs.Fields("PostImmobilisedRate")%>'></td>
		<input type = 'hidden' name ='txtSecurityImmobilised' id = "txtSecurityImmobilised" size="20" value = '<%=rs.Fields("SecurityImmobilised")%>'></td>
		<input type = 'hidden' name ='txtClientDPA' id = "txtClientDPA" size="20" value = '<%=rs.Fields("Client_DPA_")%>'></td>
		<input type = 'hidden' name ='txtEntityDPA' id = "txtEntityDPA" size="20" value = '<%=rs.Fields("EntityType_DPA_")%>'></td>
		<input type = 'hidden' name ='txtClass' id = "txtClass" size="20" value = '<%=rs.Fields("Class")%>'></td>

  <% 
	if rs.Fields("InterBank")=true then
	%>
	<input type = 'hidden' name ='txtinterbank' id = "txtinterbank" size="20" value = '1'></td>
	<%
	else
	%>
	<input type = 'hidden' name ='txtinterbank' id = "txtinterbank" size="20" value = '0'></td>
	<%
	end if
	%>
	 
	<% 
	
	if rs.Fields("IsCustodian")=true then
	%>
	<input type = 'hidden' name ='txtcustodian' id = "txtcustodian" size="20" value = '1'></td>
	<%
	else
	%>
	<input type = 'hidden' name ='txtcustodian' id = "txtcustodian" size="20" value = '0'></td>
	<%
	end if
	%>

 
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
    <td width="17%" height="25">Settlement Date</td>
    <td width="83%" height="25"><SCRIPT language="JavaScript">calDate.writeControl();</SCRIPT></td>
  </tr>

  <tr>
  <td colspan = '2' height="58">
  
  <table border="0" width="100%">
    <tr>
      <td width="23%"><b><font color="#000080">Ref</font></b></td>
      <td width="13%"><b><font color="#000080">Date&nbsp;</font></b></td>
      <td width="17%"><b><font color="#000080">Quantity</font></b></td>
      <td width="14%"><b><font color="#000080">Price</font></b></td>
      <td width="33%"><b><font color="#000080">Broker</font></b></td>
    </tr>
    
    <tr>
      <td width="23%" valign="top"><input type = 'text' name ='txtSlip' id = 'txtSlip' size="15">&nbsp;</td>
      <SCRIPT language="javascript">
      document.all.item('txtSlip').focus();
    </script>
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
		<input type = 'button' class=buttons name ='cmdAdd' id = 'cmdAdd' value=" Save " onclick = "forceSubmit();AllowedNavigation();hideButton();">
    	&nbsp; <input type = 'button' class=buttons name ='cmdClose' id = "cmdClose" value=" Close " onclick = "window.self.close();">
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
		<input type = 'hidden' name ='buttonAction' id = 'action' value="Save">
	</td>
  </tr>
</table>

</form>
</body>

</html>