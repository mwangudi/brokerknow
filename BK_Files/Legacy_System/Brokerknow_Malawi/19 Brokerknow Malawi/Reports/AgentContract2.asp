<!--#include file="../libroutinesTEST.asp"-->
<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Sale Contract</title>  
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
<SCRIPT language=VBScript src="../scripts/reports.vbs"></SCRIPT>
<SCRIPT language=Javascript src="../scripts/reports.js"></SCRIPT>

	<style media="print">
		@page {
			@top{font-family: Helvetica, Arial, sans-serif;
				font-size: 150%;
				font-weight: bolder;
				text-align: left;
				content: "<%= FormatDate(Date) %>";			
			}
			
			margin-left: 2cm;
			margin-right: 2cm;
			margin-top: 1cm;    
			margin-bottom: 1cm;
			size: portrait;
			
			br.newpage{
				page-break-before:always;
			}
			
			tr.pageNumbering{
				display:none;
			}
		}

	</style>
</head>

<body Class="Reports">



<%

genReport = Request.Form("genReport")
selectedContractDate = Request.Form("txtDate")

If genReport <> "1" Or selectedContractDate = "" Then%>
	<Script Language="JavaScript">
		report_SetBodyClass();
		function validateForm(frm){			
			if (frm.txtDate.value==''){
				alert("Select a date");
				frm.txtDate.focus();
				return;
			}
			
			frm.target = '_self';			
			frm.submit();
		}
		
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "txtDate","cmdDate","<%= FormatDate(Date) %>",1);
	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="AgentContract2.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">	
		<table>
			<tr>
				<td>Select date:</td>
				<td>
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>
			</tr>
			
			<tr>
				<td colspan=2><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... ">&nbsp;&nbsp; </td>
			</tr>
		</table>
		
	</form>
	
	<%
	Response.End
End If

%>

<% DrawPageFunctions True, True, True, True %>

<p id="toPDFOrient" name="toPDFOrient" value="P" style="display:none;">P
<p id="toPDF" name="toPDF">

<%
	Set conn = GetActiveConnection("KBroker")
	Set Rs = CreateObject("ADODB.Recordset")						        
	sqlStr = "SELECT * FROM AgentContracts WHERE LotTDate = '" & FormatDate(selectedContractDate) & "' ORDER BY ContractNumber"		
	
	
	Rs.CursorLocation = adUseClient	
	Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
	
	If rs.EOF Or rs.BOF Then%>
		<Script Language="JavaScript">
			alert("There are no contracts based on the specified criterion.")
			window.parent.history.go(-1);			
		</Script>
		<%Set Rs = Nothing
		Set Conn = Nothing
		Response.End
	End If
	
	If adkljdakfj then
	ContractDPA = rs("Contract_DPA_")
	
	''RECALCULATE THE COMMISSIONS
	'Client Contracts
	sqlStr = "UPDATE LevyContract" & _
		" SET LevyAmount = ProperCommission3" & _
		" FROM (SELECT ShowProperCommissionForCompoundedContractsForEachLevyEntry.LotTDate, " & _
		" ShowProperCommissionForCompoundedContractsForEachLevyEntry.ProperCommissionRate, " & _
		" ShowProperCommissionForCompoundedContractsForEachLevyEntry.ProperCommission3, LevyContract.LevyContract_DPA_" & _
		" FROM ShowProperCommissionForCompoundedContractsForEachLevyEntry INNER JOIN" & _
		" (SELECT cast('"& FormatDate(Rs("LotTDate")) &"' AS datetime) AS LastDate) LastDateTransactions ON " & _
		" ShowProperCommissionForCompoundedContractsForEachLevyEntry.LotTDate = LastDateTransactions.LastDate INNER JOIN" & _
		" LevyContract ON " & _
		" ShowProperCommissionForCompoundedContractsForEachLevyEntry.LevyContract_DPA_ = LevyContract.LevyContract_DPA_" & _
		" WHERE (LevyContract.Contract_DPA_ = "& ContractDPA &")) A" & _
		" WHERE LevyContract.LevyContract_DPA_ = A.LevyContract_DPA_"
					
	'conn.Execute SQLServerFormat(HandleQuote(sqlStr))
		
	Response.Write sqlstr
	Response.End 
						
	'Agent Contracts
	sqlStr = "UPDATE    LevyContract" & _
		" SET LevyAmount = ProperAgentAmount" & _
		" FROM (SELECT LevyContract.Contract_DPA_, LevyContract_1.LevyContract_DPA_ AS AgentLevyContract_DPA_, " & _
		" ShowProperCommissionForCompoundedContractsForEachLevyEntry.ProperCommission3 AS BrokerAmount, " & _
		" LevyContract_1.LevyAmount AS AgentAmount, Commission.CommissionRate, " & _
		" ROUND(Commission.CommissionRate / 100 * ShowProperCommissionForCompoundedContractsForEachLevyEntry.ProperCommission3, 2) " & _
		" AS ProperAgentAmount, OrdDetail.OrdDetail_DPA_" & _
		" FROM (SELECT cast('"& FormatDate(Rs("LotTDate")) &"' AS datetime) AS LastDate) LastDateTransactions INNER JOIN" & _
		" Lot INNER JOIN" & _
		" OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN" & _
		" ShowProperCommissionForCompoundedContractsForEachLevyEntry INNER JOIN" & _
		" LevyContract ON " & _
		" ShowProperCommissionForCompoundedContractsForEachLevyEntry.LevyContract_DPA_ = LevyContract.LevyContract_DPA_ INNER JOIN" & _
		" LevyContract LevyContract_1 ON LevyContract.Contract_DPA_ = LevyContract_1.Contract_DPA_ ON " & _
		" Lot.Contract_DPA_ = LevyContract.Contract_DPA_ ON CAST(FLOOR(CAST(LastDateTransactions.LastDate AS float)) AS datetime)" & _
		" = CAST(FLOOR(CAST(Lot.LotTDate AS float)) AS datetime) CROSS JOIN" & _
		" Client INNER JOIN" & _
		" tbOrder ON Client.Client_DPA_ = tbOrder.Client_DPA_ INNER JOIN" & _
		" Agent ON Client.Agent_DPA_ = Agent.Agent_DPA_ INNER JOIN" & _
		" Commission ON Agent.Commission_DPA_ = Commission.Commission_DPA_" & _
		" WHERE (LevyContract_1.SystemMaintained = 12) AND (LevyContract.Contract_DPA_ = "& ContractDPA &")) A" & _
		" WHERE LevyContract.LevyContract_DPA_ = A.AgentLevyContract_DPA_"
							
	'conn.Execute SQLServerFormat(HandleQuote(sqlStr))
	
	Response.Write sqlstr
	Response.End 
	end if			
							
	Set groupRs = Conn.Execute("SELECT ContractNumber FROM AgentContracts WHERE LotTDate = '" & FormatDate(selectedContractDate) & "' GROUP BY ContractNumber ")
	
	Set levyOrderRs = Conn.Execute("SELECT * FROM LevyOrderList ORDER BY LevyOrder")
	Dim pageNumber
	
	pageNumber = 0
	Do Until groupRs.EOF
	
		Rs.Filter = "ContractNumber = '" & groupRs.Fields("ContractNumber").Value & "'"
	
	totalGross = Rs.Fields("LotGrossAmount").Value
	
	sqlStr = "SELECT     tbOrder.*, Client.ClientAddr FROM Lot INNER JOIN " & _
			" OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN " & _
			" tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN " & _
			" Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ " & _
			" WHERE Lot.Lot_DPA_ = " & Rs.Fields("Lot_DPA_").Value
		
	
	Set temprs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
	
	If Not (temprs.EOF Or temprs.BOF) Then
		orderRef = temprs.Fields("orderRef").Value
		if isnull(orderRef) or trim(orderRef) = "" then
				orderRef = temprs.Fields("Order_DPA_").Value
		else
				orderRef = temprs.Fields("Order_DPA_").Value & "/" & orderRef
		end if
		clientAddress = Replace(temprs.Fields("ClientAddr").Value, vbCrLf, ",")		
	End If
	
	Set temprs = Nothing
	
	If Trim(UCase(Rs.Fields("OrdDetailType").Value)) = "PURCHASE" Then
		IsPurchase = True
	Else
		IsPurchase = False
	End If
	
	pageNumber = pageNumber + 1
%>
<table border="0" cellspacing=0 cellpadding=0 class="ReportsTable" align ="center" width="90%"> 
	<THEAD>
	<tr class="pageNumbering">
		<td align="left" colspan=2>
			<FONT FACE=ARIAL SIZE=2><B>Page <%=pageNumber%>	</B></FONT>	
		</td>		
	</tr>
	<tr>
		<td align="center" colspan=2 height="160px">
			<Img Src="../data/photos/aaprintlogo.jpg">			
		</td>		
	</tr>
	<THEAD>  
	<tr>
		<td align="center" colspan="2">
			<FONT FACE=ARIAL SIZE=3><B>
				<%= UCase(Rs.Fields("OrdDetailType").Value) %> CONTRACT <%= Rs.Fields("ContractNumber").Value %>
			</B></FONT>
		</td>
		
	</tr>  
	<!--<tr>
		<td colspan=2>&nbsp;</td>
	</tr>-->
	<tr>
		<td nowrap width="80%" valign="top">			
			<table cellspacing=0 cellpadding=0 border=0  align ="center" width="100%">
				<tr>
					<td nowrap><%= Rs.Fields("AgentName").Value %> </td>					
				</tr>
				<tr>					
					<td nowrap><%= Replace(Rs.Fields("AgentAddr").Value, vbCrLf, ",") %></td>
				</tr>
				<tr>					
					<td nowrap><b>Account:</b> <%= Rs.Fields("OrdDetailClient").Value %></td>
				</tr>
			</table>
		</td>
		<td valign="top" width="20%" align="right">
			<table cellspacing=0 cellpadding=0 border=0  align ="center" width="90%">
				<tr>
					<td nowrap>Date:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%= FormatDate(Rs.Fields("LotTDate").Value) %></td>
				</tr>
				<tr>
					<td nowrap>Order Ref:&nbsp;&nbsp;&nbsp;<%= orderRef %></td>					
				</tr>
				<tr>
					<td nowrap>CDS Ref:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%= Rs.Fields("LotSlipNo").Value %></td>					
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td colspan=2>Dear Sir/Madam</td>
	</tr>
	<tr>
		<td colspan=2>&nbsp;</td>
	</tr>
	<tr>
		<td colspan=2><U><b>RE: <%= UCase(Rs.Fields("OrdDetailType").Value) %> OF SECURITIES<b></U></td>
	</tr>
	<tr>
		<td colspan=2>&nbsp;</td>
	</tr>
	<tr>
		<td colspan=2><PRE><font face=Arial size=2>We wish to advise that we have <% If IsPurchase Then Response.Write "bought" Else Response.Write "sold"  %> the following securities as per your instruction. Kindly arrange
to complete the transaction as per the details given below: </font></PRE></td>
	</tr>
	
	<tr>
		<td colspan=2>&nbsp;</td>
	</tr>
</table>	

<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" align ="center" width="90%">	
	
	<tr>
		<td>Security</td>
		<td align=right><%= Rs.Fields("OrdDetailSecurity").Value %></td>
	</tr>
	<tr>
		<td><%if lcase(rs.Fields("OrdDetailSecType")) = "fixed" Then%>
					Face Value
			<%else%>
					Quantity
			<%end if%></td>
		<td align=right> <%= FormatNum(Rs.Fields("LotQty").Value) %> </td>
	</tr>
	<tr>
		<td>Price</td>
		<td align=right> 
		<%if lcase(rs.Fields("OrdDetailSecType")) = "fixed" Then%>
				<%= FormatNumEx(Rs.Fields("LotPrice").Value,4) %>
		<%else%>
				<%= FormatNum(Rs.Fields("LotPrice").Value) %>
		<%end if%></td>
	</tr>
    <tr>
		<td>Gross</td>
		<td align=right><b> <%= FormatNum(totalGross) %> </b> </td>
	</tr>
	<tr>
		<TD COLSPAN=2 VALIGN="TOP">
		<TABLE align ="center" width="100%" CELLSPACING="0" CELLPADDING="4">
		<%
		dim brokerCommission
		totalLevies = 0
		totalContractStamps = 0
		levyArray = SortLevies(rs, levyOrderRs)
		transferFeeVal = 0
		contractStampsVal = 0
		
		
		For i = 1 To UBound(levyArray)			
					
			thisLevyName = Trim(levyArray(i, 0))
			
			
			If levyArray(i, 2) = 10 Then
				totalContractStamps = levyArray(i, 1)
				totalContractStampsName = levyArray(i, 0)
			End If			
			
			
			'ignore the below line as the stored levy percentage
			'is not very reliable....
			'levyPerc = levyArray(i, 3)
			
			'recalculate percentage (really getting tired with this...)
			levyPerc = (levyArray(i, 1) / totalGross) * 100			
			'test = levyArray(i, 1)
			'response.write "test =" & test
			
			If (levyArray(i, 2) = 11)  Then
				brokerCommission = levyArray(i, 1)
			End If

			If (levyArray(i, 2) = 11) and (levyArray(i, 1) <= 100) Then
				levyPerc = "Minimum"
				brokerRate = formatnum(0)
			Else
				brokerRate = levyPerc
			End If
						
			'grab agent commission here: to be used laters..
			If levyArray(i, 2) = 12 Then
				AgentCommission = levyArray(i, 1)
				AgentRate = levyPerc 'levyArray(i, 3)
				AgentRateDescription = levyArray(i, 0)
				'empty the levy name variable so as not to display
				thisLevyName = ""
			End If
			
			'do not display staff commission  HARD CODING!!!!!!
			If (levyArray(i, 2) = 8) Then
				thisLevyName = ""
			End If
			
			'grab the broker rate here
			'If levyArray(i, 2) = 11 Then
			'	brokerRate = levyPerc
			'End If
			
			'transfer fee
			If levyArray(i, 2) = 13  Then				
				transferFeeDesc = thisLevyName
				transferFeeVal = levyArray(i, 1)
				thisLevyName = ""
			End If
			
			If levyArray(i, 2) = 10 Then
				contractStampsDesc = thisLevyName
				contractStampsVal = levyArray(i, 1)
				thisLevyName = ""
			End If
			
			
			If thisLevyName <> "" Then
				'total up levies only when displayed
				totalLevies = totalLevies + levyArray(i, 1) 
				if levyPerc <> "Minimum" then levyPerc = FormatNumber(levyPerc,3) & "%"
		%>
				<TR>		
					<td nowrap width="10%" STYLE="PADDING-LEFT: 0PX"><%= thisLevyName %></td>
					<td ALIGN="LEFT">&nbsp;&nbsp;<%= levyPerc %></td>
					<td align=right><%= FormatNum(levyArray(i, 1)) %></td>		
				</TR>
		<%	End If
		Next%>
		</TABLE>	
		</TD>
	</tr>
	
	<tr>
		<td>&nbsp;</td>
		<td align=right>
			<hr width=70px height=1px>
			<%= Replace(Space(Len(FormatNum(totalLevies)) * 4), Space(1), "&nbsp;") %>
		</td>
	</tr>
	<tr>
		<td>Total Commission and Levies</td>
		<td align=right>	
			<%If IsPurchase Then%>
				<b><%= FormatNum(totalLevies) %></b>
			<%Else%>	
				<b>(<%= FormatNum(totalLevies) %>)</b>
			<%End If%>
		</td>
	</tr>
	<%if trim(transferFeeDesc) <> "" then%>
			<tr>
				<td><%= transferFeeDesc %></td>
				<td align=right>	
					 <%= FormatNum(transferFeeVal) %>		
				</td>
			</tr>
	<%end if%>
	
	<%if trim(contractStampsDesc) <> "" then%>
			<tr>
				<td><%= contractStampsDesc %></td>
				<td align=right>	
					 <%= FormatNum(contractStampsVal) %>		
				</td>
			</tr>
	<%end if%>
	<tr>
		<td>&nbsp;</td>
		<td align=right> 
			<%= Replace(Space(Len(FormatNum(totalLevies)) * 4), Space(1), "&nbsp;") %>
			<hr width=70px height=1px>
		</td>
	</tr>
	<tr>
		<td>GROSS AMOUNT</td>
		<td align=right><b>
			<%
			totalLevies = totalLevies + transferFeeVal + contractStampsVal
			If IsPurchase Then
				grossAmount = totalGross + totalLevies				
			Else
				grossAmount = totalGross - totalLevies				
			End If
			
			Response.Write FormatNum(grossAmount) 
			%>
			
			</b></td>
	</tr>
	<tr>	
		<td COLSPAN=2 VALIGN="TOP">
			<TABLE width="100%" CELLSPACING="0" CELLPADDING="4">
				<TR>		
					<td nowrap width="10%" STYLE="PADDING-LEFT: 0PX">Returnable Commission</td>
					<td ALIGN="LEFT">&nbsp;&nbsp;<%
					'Trim to nearest whole number. Change By Muchiri
					
					if brokerRate= 0 then 
						Response.Write "0 %"
					else
						'Response.write FormatNum(FormatNumEx(AgentRate * 100/brokerRate,0)) & "%"
						Response.write FormatNum(AgentCommission/brokerCommission*100) & "%"
					End if	
					 %>
					
					
					</td>
					<td align=right><%= FormatNum(AgentCommission) %></td>		
				</TR>		
		</table>
		</td>	
	</tr>
	<tr>
		<td>NET AMOUNT</td>
		<td align=right><b>
			<% Dim netAmount
				
				If IsPurchase Then
					netAmount = grossAmount - AgentCommission				
				Else
					netAmount = grossAmount + AgentCommission
				End If
			%>
			
			 <%=FormatNum(netAmount) %>
			
			</b></td>
	</tr>
	<!--<tr>
		<td colspan=2>&nbsp;</td>
	</tr>
	<tr>
		<td colspan=2>&nbsp;</td>
	</tr>-->
	
	<tr>
		<td>
		<PRE><font face=Arial size=2>
       <i>For and on behalf of</i>
African Alliance Malawi Securities

Sign..........................................
		</FONT></PRE>			
		</td>
		<td align="left" valign="top">			
			<img src="../images/stamp.gif" border="0" style="position: absolute;z-index: 2">
			<BR>
			&nbsp;<SPAN style="position: absolute;z-Index: 10"><b>KShs&nbsp;&nbsp;<%= FormatNum(totalContractStamps) %></b></SPAN>
			<br><br>
			Revenue Stamps prepaid
			
		</td>
	</tr>
</table>
	
<%
		Rs.Cancel
	groupRs.MoveNext
	'important!
		If Not groupRs.EOF Then %>
			<br class="newpage">
	<%	End If
		
	Loop
	
	Set groupRs = Nothing
	Set rs = Nothing
	Set Conn = Nothing
	
	Function SortLevies(srcRs, orderRs)
		Dim returnArray()
		
		ReDim returnArray(srcRs.RecordCount, 3)
		
		'careful, recordcount is 1- based
		arrCounter = 0 
		If Not (orderRs.EOF Or orderRs.BOF) Then
			Do Until orderRs.EOF
					isExistingLevy = False
					Do Until srcRs.EOF
							If StrComp(orderRs.Fields("LevyName").Value, srcRs.Fields("LevyShortName").Value, vbTextCompare) = 0 Then
								arrCounter = arrCounter + 1
								returnArray(arrCounter, 0) = srcRs.Fields("LevyName").Value
								returnArray(arrCounter, 1) = srcRs.Fields("LevyAmount").Value
								returnArray(arrCounter, 2) = srcRs.Fields("SystemMaintained").Value
								returnArray(arrCounter, 3) = srcRs.Fields("LevyRatePercentage").Value
								Exit Do	
							End If
						
						srcRs.MoveNext
					 Loop	
					
					srcRs.MoveFirst					
					orderRs.MoveNext
			  Loop	
			
			orderRs.MoveFirst
			
			If arrCounter <> UBound(returnArray) Then
				'take care of the rest unordered items
				Do Until srcRs.EOF
					thisName = srcRs.Fields("LevyName").Value
					
					'check if it's in array first
					existsInArray = False
					For k = 1 To arrCounter
						If StrComp(thisName, returnArray(k, 0), vbTextCompare) = 0 Then
							existsInArray = True
							Exit For
						End If
					Next
					
					If Not existsInArray Then
						arrCounter = arrCounter + 1
						returnArray(arrCounter, 0) = srcRs.Fields("LevyName").Value
						returnArray(arrCounter, 1) = srcRs.Fields("LevyAmount").Value
						returnArray(arrCounter, 2) = srcRs.Fields("SystemMaintained").Value	
						returnArray(arrCounter, 3) = srcRs.Fields("LevyRatePercentage").Value	
					End If
					
					srcRs.MoveNext
				Loop	
				srcRs.MoveFirst
			End If
			
			
		Else
			Do Until srcRs.EOF
				arrCounter = arrCounter + 1
				returnArray(arrCounter, 0) = srcRs.Fields("LevyName").Value
				returnArray(arrCounter, 1) = srcRs.Fields("LevyAmount").Value
				returnArray(arrCounter, 2) = srcRs.Fields("SystemMaintained").Value		
				returnArray(arrCounter, 3) = srcRs.Fields("LevyRatePercentage").Value
				srcRs.MoveNext
			Loop	
			srcRs.MoveFirst
		End If
		
		SortLevies = returnArray
		
	End Function
%>

</body>

</html>
