<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Balance Sheet</title>

	<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
	<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
	<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
	<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
	<SCRIPT language=VBScript src="../scripts/reports.vbs"></SCRIPT>
	<SCRIPT language=Javascript src="../scripts/reports.js"></SCRIPT>
	<SCRIPT language=Javascript src="../scripts/accountList.js"></SCRIPT>



	<style media="print">
		@page {
			@top{font-family: Helvetica, Arial, sans-serif;
				font-size: 150%;
				font-weight: bolder;
				text-align: left;
				content: "<%= FormatDate(Date) %>";			
			}
			
			margin-left: 0cm;
			margin-right: 0cm;
			margin-top: 1cm;    
			margin-bottom: 2cm;
			size: portrait;
			
			br.newpage{
				page-break-before:always;
			}
			
			
		}

	</style>

</head>

<body Class="Reports">
<!--#include file="../libroutines.asp"-->

<%

genReport = Request.Form("genReport")
'selectedBank = Request.Form("cboAccount")
selectedFromDate = Request.Form("transFromDate")
selectedToDate = Request.Form("txtToDate")
SelectedType=Request.Form("cboEntity")
FirstDay=DateSerial(Year(Date), Month(Date)-1 + iOffset, 1)
thistype=Request.Form("Selectedtype")

If genReport <> "1" Or Not IsDate(selectedFromDate) Then%>
	<Script Language="JavaScript">
		report_SetBodyClass();
		
		function validateForm(frm)
		{						
			frm.target = '_self';			
			frm.submit();
		}
		
		
		function evaluateEntity(Val, Entity)
		{      	
	  	FetchAccounts1(Entity)
		}
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "transFromDate","cmdDate","<%= FormatDate(FirstDay) %>",1);
		var cal1=new ctlSpiffyCalendarBox("cal1", "frmMain", "txtToDate","cmdDate","<%= FormatDate(Date) %>",1);

	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="BalanceSheet.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">
		<% currentEntityType=5 %>
		<table>
			<tr>
				<td colspan="2">Select date from:</td>
				<td>
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>				
			</tr>
			<tr>
				<td colspan="2">To date:</td>
				<td>
					<SCRIPT language="JavaScript">cal1.writeControl();</SCRIPT>	
				</td>
				
			</tr>			

			<tr>
				<td colspan="3"><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... ">&nbsp;&nbsp;</td>
			</tr>
		</table>
		
	</form>
	
	<%
	'Set rs = Nothing
	'Set Conn = Nothing
	Response.End
End If

%>
<% DrawPageFunctions True, True, True %>
<%
	Dim i
	Dim Conn
	Dim Rs

	Set conn = GetActiveConnection("KBroker")
	Set Rs = CreateObject("ADODB.Recordset")	
	Rs.CursorLocation = adUseClient	
	%>
	

	<table border="0" cellspacing="2" cellpadding="2" style="font-family: Arial Narrow" width="100%">
		<tr>
		<td width="10%" nowrap><font face="Impact" size="4">BALANCE SHEET</font></td>
		<td width="60%" nowrap align="right"><font face="Impact" size="3"><%= Session("CompanyName") %>
		    </font></td>
		</tr>
	</table>
	<br>
	<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
		<tr>
		<td width="1%"><b>Date: </b></td>
		<td width="48%"><%= FormatDate(Date) %>
		</td>
		</tr>
	</table>
	<br>
	<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow; LEFT-MARGIN:100PX" width="100%">
		<tr>
		<td style="border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" colspan="2"><b><font face="Arial Narrow" size="3">Fixed Assets</font></b></td>
		<td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" align="right"><b><font face="Arial Narrow" size="3">&nbsp;</font></b></td>
		<td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" align="right"><b><font face="Arial Narrow" size="3">&nbsp;</font></b></td>
		</tr>
		<%
	sqlStr = "FixedAssetsProc '" & FormatDate(selectedFromDate) & "', '" & FormatDate(selectedToDate) & "' "
	'Response.write(sqlStr)
	'Response.end

	'Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
	'Rs.Filter = "Client_DPA_ LIKE '" & selectedClient & "' AND TransDate >= '" & FormatDate(selectedFromDate) & "'"
	
	Rs.CursorLocation = adUseClient	
	Rs.Open SQLServerFormat((sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic

	'for i = 1 to 9
	'	Set Rs =  Rs.NextRecordset 
	'next
	
	Dim fixedAssets
    Do Until Rs.EOF 
		if(Rs.Fields("Balance").Value <> 0) then
		%>
		<tr>
		  <td colspan="2"><%= Rs.Fields("AccountName").Value %>
		  </td>
		  <td align="right">&nbsp;
		  </td>
		  <td align="right"><%=FormatNum(Rs.Fields("Balance").Value)%>
		  </td>
		</tr>
		<%
		end if
				If trim(Rs.Fields("SubTotal").Value) <> "" Then%>
					<tr>
						<td colspan="2">&nbsp;
						</td>
						<td align="right">&nbsp;
						</td>
						<td align="right"><%=FormatNum(Rs.Fields("SubTotal").Value)%>
						</td>
					</tr>				
				<%
				fixedAssets = Rs.Fields("SubTotal").Value
				exit do
				End If   	
		Rs.MoveNext
		%>

		
		<%
	Loop%>

	<tr>
		<td style="border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" colspan="2"><b><font face="Arial Narrow" size="3">Current Assets</font></b></td>
		<td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" align="right"><b><font face="Arial Narrow" size="3">&nbsp;</font></b></td>
		<td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" align="right"><b><font face="Arial Narrow" size="3">&nbsp;</font></b></td>
		</tr>
		<%
	sqlStr = "CurrentAssetsProc '" & FormatDate(selectedFromDate) & "','" & FormatDate(selectedToDate) & "' "
	Rs.Close
	Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
	'Rs.Filter = "Client_DPA_ LIKE '" & selectedClient & "' AND TransDate >= '" & FormatDate(selectedFromDate) & "'"
	
	'for i = 1 to 9
	'	Set Rs =  Rs.NextRecordset 
	'next
	
	Dim currentAssets
    Do Until Rs.EOF 
		if(Rs.Fields("Balance").Value<>0) then
		%>
		<tr>
		  <td colspan="2"><%= Rs.Fields("AccountName").Value %>
		  </td>
		  <td align="right"><%=FormatNum(Rs.Fields("Balance").Value)%>
		  </td>
		  <td align="right">&nbsp;
		  </td>
		</tr>
		<%	
		end if
				If trim(Rs.Fields("SubTotal").Value) <> "" Then%>
					<tr>
						<td colspan="2">&nbsp;
						</td>
						<td align="right"><%=FormatNum(Rs.Fields("SubTotal").Value)%>
						</td>
						<td align="right">&nbsp;
						</td>
					</tr>				
				<%
				currentAssets = Rs.Fields("SubTotal").Value
				exit do
				End If   	
		Rs.MoveNext

	Loop
	%>
	
	<tr>
		<td style="border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" colspan="2"><b><font face="Arial Narrow" size="3">Current Liabilities</font></b></td>
		<td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" align="right"><b><font face="Arial Narrow" size="3">&nbsp;</font></b></td>
		<td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" align="right"><b><font face="Arial Narrow" size="3">&nbsp;</font></b></td>
		</tr>
		<%
	sqlStr = "CurrentLiabilitiesProc '" & FormatDate(selectedFromDate) & "','" & FormatDate(selectedToDate) & "' "
	Rs.Close
	Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
	'Rs.Filter = "Client_DPA_ LIKE '" & selectedClient & "' AND TransDate >= '" & FormatDate(selectedFromDate) & "'"
	
	'for i = 1 to 9
	'	Set Rs =  Rs.NextRecordset 
	'next
	
	Dim currentLiabilities
    Do Until Rs.EOF
		if(Rs.Fields("Balance").Value<>0) then
		%>
		<tr>
		  <td colspan="2"><%= Rs.Fields("AccountName").Value %>
		  </td>
		  <td align="right"><%=FormatNum(Rs.Fields("Balance").Value)%>
		  </td>
		  <td align="right">&nbsp;
		  </td>
		</tr>
		<%
		end if
				If trim(Rs.Fields("SubTotal").Value) <> "" Then%>
					<tr>
						<td colspan="2">&nbsp;
						</td>
						<td align="right"><%=FormatNum(Rs.Fields("SubTotal").Value)%>
						</td>
						<td align="right">&nbsp;
						</td>
					</tr>				
				<%
				currentLiabilities = Rs.Fields("SubTotal").Value
				exit do
				End If   	
		Rs.MoveNext
		
	Loop
	
	Dim netCurrentAssets
	Dim totalAssets
	netCurrentAssets = CCur(currentAssets) + CCur(currentLiabilities)
	totalAssets = CCur(fixedAssets) + CCur(netCurrentAssets)
	%>
	
	<tr>
		<td colspan="2"><b><font face="Arial Narrow" size="3">Net Current Assets</font></b>
		</td>
		<td align="right">&nbsp;
		</td>
		<td align="right"><%=FormatNum(netCurrentAssets)%>
		</td>
	</tr>
	
	<tr>
		<td colspan="2"><b><font face="Arial Narrow" size="3">Total Assets</font></b>
		</td>
		<td align="right">&nbsp;
		</td>
		<td align="right"><%=FormatNum(totalAssets)%>
		</td>
	</tr>
	
	<tr>
		<td colspan="2"><b><font face="Arial Narrow" size="3">&nbsp;</font></b>
		</td>
		<td align="right">&nbsp;
		</td>
		<td align="right">&nbsp;
		</td>
	</tr>
	
	<tr>
		<td style="border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" colspan="2"><b><font face="Arial Narrow" size="3">Long Term Liabilities</font></b></td>
		<td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" align="right"><b><font face="Arial Narrow" size="3">&nbsp;</font></b></td>
		<td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" align="right"><b><font face="Arial Narrow" size="3">&nbsp;</font></b></td>
		</tr>
		<%
	sqlStr = "LongTermLiabilitiesProc '" & FormatDate(selectedFromDate) & "','" & FormatDate(selectedToDate) & "' "
	Rs.Close
	Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
	'Rs.Filter = "Client_DPA_ LIKE '" & selectedClient & "' AND TransDate >= '" & FormatDate(selectedFromDate) & "'"
	
	'for i = 1 to 9
	'	Set Rs =  Rs.NextRecordset 
	'next
	
	Dim longTermLiabilities
	Dim pnlSummary
    Do Until Rs.EOF 
		if(Rs.Fields("Balance").Value<>0) then
		%>
		<tr>
		  <td colspan="2"><%= Rs.Fields("AccountName").Value %>
		  </td>
		  <td align="right"><%=FormatNum(Rs.Fields("Balance").Value)%>
		  </td>
		  <td align="right">&nbsp;
		  </td>
		</tr>
		<%	
		end if
				If trim(Rs.Fields("SubTotal").Value) <> "" Then
						longTermLiabilities = Rs.Fields("SubTotal").Value
						exit do
				End If   	
		Rs.MoveNext

	Loop
	
	sqlStr = "PnLProc '" & FormatDate(selectedFromDate) & "','" & FormatDate(selectedToDate) & "' "
	Rs.Close
	Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
	'Rs.Filter = "Client_DPA_ LIKE '" & selectedClient & "' AND TransDate >= '" & FormatDate(selectedFromDate) & "'"
	
	'for i = 1 to 9
	'	Set Rs =  Rs.NextRecordset 
	'next
	
	RS.MoveLast
	pnlSummary = Rs.Fields("GrandTotal").Value%>
	<tr>
		<td colspan="2">Profit / Loss
		</td>
		<td align="right">&nbsp;
		</td>
		<td align="right"><%=FormatNum(pnlSummary)%>
		</td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;
		</td>
		<td align="right">&nbsp;
		</td>
		<td align="right"><%=FormatNum(pnlSummary + longTermLiabilities)%>
		</td>
	</tr>
	</table>
	<%
Set Rs = Nothing
Set Conn = Nothing%>

</body>

</html>
