<!--#include file="../libroutines.asp"-->
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
			@top{font-family: Times New Roman, Times New Roman, Times New Roman;
				font-size: 140%;
				font-weight: bolder;
				text-align: left;
				content: "<%= FormatDate(Date) %>";			
			}
			
			margin-left: 2cm;
			margin-right: 2cm;
			margin-top: 1cm;    
			margin-bottom: 2cm;
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
	<form method="POST" action="AgentContract.asp" Name="frmMain" id="frmMain">
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

<% DrawPageFunctions True, True, True %>


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
	
	Set groupRs = Conn.Execute("SELECT ContractNumber FROM AgentContracts WHERE LotTDate = '" & FormatDate(selectedContractDate) & "' GROUP BY ContractNumber")
	
	Set levyOrderRs = Conn.Execute("SELECT * FROM LevyOrderList ORDER BY LevyOrder")
	Dim pageNumber
	
	pageNumber = 0
	Do Until groupRs.EOF
	
		Rs.Filter = "ContractNumber = '" & groupRs.Fields("ContractNumber").Value & "'"
	
	'totalGross = Rs.Fields("LotQty").Value * Rs.Fields("LotPrice").Value
	totalGross = Rs.Fields("LotGrossAmount").Value
	
	sqlStr = "SELECT tbOrder.*, Client.ClientAddr FROM Lot INNER JOIN " & _
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
		clientAddress = Replace(temprs.Fields("ClientAddr").Value, Chr(13), ",")		
	End If
	
	Set temprs = Nothing
	
	If Trim(UCase(Rs.Fields("OrdDetailType").Value)) = "PURCHASE" Then
		IsPurchase = True
	Else
		IsPurchase = False
	End If
	
	pageNumber = pageNumber + 1
%>

<table border="0" cellspacing=0 cellpadding=0 class="ReportsTable" width="90%" align="center" valign="top"> 
<tr><td>
<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" width="100%" align="center"> 
	<THEAD>
	<tr class="pageNumbering">
		<td align="left" colspan=2>
			<FONT FACE=Times New Roman SIZE=2><B>Page <%=pageNumber%>	</B></FONT>	
		</td>		
	</tr>
	<tr>
		<td align="right" colspan=2 height="160px">
			<Img Src="../data/photos/aaprintlogo.jpg">
		</td>		
	</tr>
	<THEAD>  
	<tr>
		<td align="center" colspan="2">
			<FONT FACE=Times New Roman SIZE=3><B>
				<%= UCase(Rs.Fields("OrdDetailType").Value) %> CONTRACT <%= Rs.Fields("ContractNumber").Value %>
			</B></FONT>
		</td>
		
	</tr>  
	 	
	<tr><td>
	<tr><td colspan="2"><table border=1 width="100%" cellspacing="0" cellpadding="0" class="ReportsTable">
	<tr><td>
	<table border=0 width="100%">
	<tr>
		<td nowrap width="65%" valign="top">			
			<table cellspacing=0 cellpadding=0 border=0  width="100%">
				<tr>
					<td nowrap><%= Rs.Fields("AgentName").Value %> </td>					
				</tr>
				<tr>					
					<td nowrap><%= Replace(Rs.Fields("AgentAddr").Value, Chr(13), ",") %></td>
				</tr>
				<tr>					
					<td nowrap><br><b>Account:&nbsp;&nbsp;</b><%= Rs.Fields("OrdDetailClient").Value %>&nbsp;&nbsp;[<%= Rs.Fields("Client_DPA_").Value %>]</td>
				</tr>
			</table>
		</td>
		<td valign="top">
			<table cellspacing=0 cellpadding=2 border=0  width="100%">
				<tr>
					<td nowrap align="right">Date:&nbsp;</td>
					<td align="left"><%= FormatDate(Rs.Fields("LotTDate").Value) %></td>
				</tr>
				<tr>
					<td nowrap align="right">Order Ref:&nbsp;</td>
					<td align="left"><%= orderRef %></td>					
				</tr>
				<tr>
					<td nowrap align="right">Slip Nos:&nbsp;</td>
					<td align="left"> <%= Rs.Fields("LotSlipNo").Value %></td>					
				</tr>
				<tr><td colspan="3">&nbsp;</tr></td>
			</table>
		</td>
	</tr>
	</table></td></tr>
	</td></tr>
	</table></td></tr>
	<tr>
		<td colspan=2><p>&nbsp;&nbsp;&nbsp;Dear Sir/Madam,</p></td>
	</tr>
	
	<tr>
		<td colspan=2 align="center"><b><u>RE: <%= UCase(Rs.Fields("OrdDetailType").Value) %> OF SECURITIES<b></U></td>
	</tr>	
</table>	
</td></tr>
<tr><td>
<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" width="100%">	
	<tr>
		<td align=right colspan=3>KShs&nbsp;</td>
	</tr>
	<tr>
		<td align="left" colspan="3"><b><%= Rs.Fields("OrdDetailSecurity").Value %></b></td>
	</tr>
	
		<TD COLSPAN=3 VALIGN="TOP">
		<TABLE WIDTH="100%" CELLSPACING="0" CELLPADDING="3" border=0>
		<tr>
		<td width="8%">&nbsp;</td>
		<td STYLE="PADDING-LEFT: 0PX"><%if lcase(rs.Fields("OrdDetailSecType")) = "fixed" Then%>Face Value
			<%else%>Quantity
			<%end if%></td>
		<td align=right colspan="2"> <%= FormatNum(Rs.Fields("LotQty").Value) %></td>
	</tr>
	<tr>
		<td width="8%">&nbsp;</td>
		<td STYLE="PADDING-LEFT: 0PX">Price</td>
		<td align=right colspan="2"> 
		<%if lcase(rs.Fields("OrdDetailSecType")) = "fixed" Then%>
				<%= FormatNumEx(Rs.Fields("LotPrice").Value,4) %>
		<%else%>
				<%= FormatNum(Rs.Fields("LotPrice").Value) %>
		<%end if%></td>
	</tr>
    <tr>
		<td width="8%">&nbsp;</td>
		<td STYLE="PADDING-LEFT: 0PX">Gross</td>
		<td align=right colspan="2"><b> <%= FormatNum(totalGross) %></b> </td>
	</tr>
	<tr>
		<%
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
			If levyArray(i, 2) = 11 Then
				brokerRate = levyPerc
			End If
			
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
				levyPerc = FormatNum(levyPerc) & "%"
				
				if(Trim(thisLevyName)="Broker Commission") then
				LevyAmount=FormatNumEx(levyArray(i, 1),2)
				else
				LevyAmount=FormatNum(levyArray(i, 1))
				end if
				%>
				<TR>
					<td width="8%">&nbsp;</td>		
					<td nowrap width="10%" STYLE="PADDING-LEFT: 0PX"><%= thisLevyName %></td>
					<% if(Cdbl(LevyAmount)=Cdbl(rs("MinimumCommission")) and Trim(thisLevyName)="Broker Commission") then
					%>
					<td ALIGN="LEFT" STYLE="PADDING-LEFT: 0PX">&nbsp;&nbsp;Minimum</td>
					<%
					else
					%>
						<td ALIGN="LEFT" STYLE="PADDING-LEFT: 0PX">&nbsp;&nbsp;<%= levyPerc %></td>
					<% end if%>
					<td align=right><%= LevyAmount %></td>		
				</TR>
		<%	End If
		Next%>
		<tr>
		<td width="8%">&nbsp;</td>
		<td>&nbsp;</td>
		<td align=right colspan="2">
			<u>
			<%= Replace(Space(Len(FormatNum(totalLevies)) * 4), Space(1), "&nbsp;") %></u>
		</td>
	</tr>
	<tr>
		<td width="8%">&nbsp;</td>
		<td STYLE="PADDING-LEFT: 0PX">Total Commission and Levies</td>
		<td align=right colspan="2">	
			<%If IsPurchase Then%>
				<b><%= FormatNum(totalLevies) %></b>
			<%Else%>	
				<b>(<%= FormatNum(totalLevies) %>)</b>
			<%End If%>
		</td>
	</tr>
	<%if trim(transferFeeDesc) <> "" then%>
			<tr>
				<td width="8%">&nbsp;</td>
				<td STYLE="PADDING-LEFT: 0PX"><%= transferFeeDesc %></td>
				<td align=right colspan="2">	
					 <%= FormatNum(transferFeeVal) %>
				</td>
			</tr>
	<%end if%>
	
	<%if trim(contractStampsDesc) <> "" then%>
			<tr>
				<td width="8%">&nbsp;</td>
				<td STYLE="PADDING-LEFT: 0PX"><%= contractStampsDesc %></td>
				<td align=right colspan="2">	
					 <%= FormatNum(contractStampsVal) %>
				</td>
			</tr>
	<%end if%>
	<tr>
		<td width="8%">&nbsp;</td>
		<td>&nbsp;</td>
		<td align=right colspan="2"><u> 
			<%= Replace(Space(Len(FormatNum(totalLevies)) * 4), Space(1), "&nbsp;") %></u>
		</td>
	</tr>
	<tr>
		<td width="8%">&nbsp;</td>
		<td STYLE="PADDING-LEFT: 0PX">GROSS AMOUNT</td>
		<td align=right colspan="2"><b>
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
		<TR>	
			<td width="8%">&nbsp;</td>	
			<td nowrap width="10%" STYLE="PADDING-LEFT: 0PX">Returnable Commission</td>
			<td ALIGN="LEFT">&nbsp;&nbsp;<%= RoundPoint05(AgentRate * 100/brokerRate) & "%" %></td>
			<td align=right><%= FormatNum(AgentCommission) %></td>		
		</TR>
		</TABLE>	
		</TD>
	</tr>
	
	
	
	<tr>
		<td widht="50%"><b>NET AMOUNT</b></td>
		<td align=right nowrap><b>
			<% Dim netAmount
				
				If IsPurchase Then
					netAmount = grossAmount - AgentCommission				
				Else
					netAmount = grossAmount + AgentCommission
				End If
			%>
			
			 <%=FormatNum(netAmount) %>
			</b>&nbsp;</td>
	</tr>
	<tr>
		<td colspan=2>&nbsp;</td>
	</tr>
		
	<tr>
	<table width="100%" border=0 cellpadding=0 cellspacing=0>
	<td width="6%" STYLE="PADDING-LEFT: 0PX">&nbsp;</td>
		<td align="left" valign="top">			
			<img src="../images/stamp.gif" border="0" style="position: absolute;z-index: 2">
			<BR>
			&nbsp;<SPAN style="position: absolute;z-Index: 10"><b>KShs&nbsp;&nbsp;<%= FormatNum(totalContractStamps) %></b></SPAN>
			<br><br><br>
			Revenue Stamps prepaid
			
		</td>
	</tr>
	
	<tr>
	<tr><td colspan="2">&nbsp;<br>&nbsp
	</td><tr>
	<td width="6%" STYLE="PADDING-LEFT: 0PX">&nbsp;</td>	
		<td><i>	<font face=Times New Roman size=2>For and on behalf of</i><br>
<span lang="en-us">Dyer and Blair Investment Bank Ltd</span><br>
Sign..........................................
		</FONT></PRE>			
		</td>
		
	</tr>
	<tr>
		<td colspan=3><!--#Include file="DirectorFooter.asp"--></td>
	</tr>
	</table>
	</td></tr>
</table>
</td></tr>
</table>	
<%
		Rs.Cancel
	groupRs.MoveNext
	'important!
		If Not groupRs.EOF Then %>
			<BR class="newpage">
	<%	End If
		
	Loop
	
	Set groupRs = Nothing
	Set rs = Nothing
	Set Conn = Nothing
	
	Function SortLevies(srcRs, orderRs)
		Dim returnArray()
		
		ReDim returnArray(srcRs.RecordCount, 3)
		
		'careful, recordcount is 1-based
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