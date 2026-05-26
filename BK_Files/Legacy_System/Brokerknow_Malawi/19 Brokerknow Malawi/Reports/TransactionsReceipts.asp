<html>
<head>
<title>Transactions Receipts</title>
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
	<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
	<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
	<!--CALENDAR -->
	<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
	<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
	<SCRIPT language=VBScript src="../scripts/reports.vbs"></SCRIPT>
	<SCRIPT language=Javascript src="../scripts/reports.js"></SCRIPT>
	
	<style media="print">
		@page {
			margin-left: 1cm;
			margin-right: 1cm;
			margin-top: 1cm;    
			margin-bottom: 1cm;
			writing-mode: tb-rl;
			height: 90%;
			margin: 10% 0%;			
			
			br.newpage{
				page-break-before:always;
			}
		}
	</style>
</head>

<body Class="Reports">

<Script Language="JavaScript">
	function HideRemindSelectLandscape(){
		try{			
			document.getElementById('landRem').style.display = 'none';
		}	
		catch(e){}	
	}
	function ShowRemindSelectLandscape(){
		try{			
			document.getElementById('landRem').style.display = '';
		}	
		catch(e){}	
	}
	window.onbeforeprint = HideRemindSelectLandscape;
	window.onafterprint = ShowRemindSelectLandscape;
</Script>

<!--#include file="../libroutinesTEST.asp"-->
<%

genReport = Request.Form("genReport")
selectedTradeDate = Request.Form("txtDate")

If genReport <> "1" Or Not IsDate(selectedTradeDate) Then
	%>
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
	<form method="POST" action="TransactionsReceipts.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">	
		<table>
			<tr>
				<td>Select Date</td>
				<td>
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>
			</tr>
			<tr>
				<td colspan=2><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... " id=Button1 name=Button1>&nbsp;&nbsp;</td>
			</tr>
		</table>
	</form>
	<%
	Set rs = Nothing
	Set Conn = Nothing
	
	Response.End
End If
%>

<%
DrawPageFunctions True, True, True, True
headerDescription = FormatDateFull(selectedTradeDate)
%>

<p id="toPDFOrient" name="toPDFOrient" value="L" style="display:none;">L
<p id="toPDF" name="toPDF">

<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
	<tr>
		<td nowrap><b><font face="Arial Narrow" size="4">Receipt Transactions</font></b></td>
		<td nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>
	</tr>	
	<tr>
		<td COLSPAN=2><font face="Arial" size="2">Date:  <%= headerDescription %></font></td>
	</tr>
	<tr>
		<td COLSPAN=2><font face="Arial" size="2">&nbsp;</font></td>
	</tr>
</table>				

<table border="0" width="100%" cellPadding="2" cellSpacing=0>
	<tr bgColor="#000000">
		<td bgColor="#000000" nowrap align="left"><b><font color="#FFFFFF">Date</font></b></td>
		<td bgColor="#000000" nowrap align="left"><b><font color="#FFFFFF">No</font></b></td>
		<td bgColor="#000000" nowrap align="left"><b><font color="#FFFFFF">Narrative</font></b></td>
		<td bgColor="#000000" nowrap align="right"><b><font color="#FFFFFF">Amount</font></b></td>
		<td bgColor="#000000" nowrap align="left"><b><font color="#FFFFFF">Mode</font></b></td>
		<td bgColor="#000000" nowrap align="left"><b><font color="#FFFFFF">Bank</font></b></td>
		<td bgColor="#000000" nowrap align="left"><b><font color="#FFFFFF">Del</font></b></td>
		<td bgColor="#000000" nowrap align="left"><b><font color="#FFFFFF">Changed By</font></b></td>
		<td bgColor="#000000" nowrap align="left"><b><font color="#FFFFFF">Modified</font></b></td>
	</tr>
	<%
	Set conn = GetActiveConnection("KBroker")
	sqlStr = "SELECT [User] AS ChangedBy, [Mod Date] AS ModDate, [Payment Date] AS PaymentDate, [Receipt No] AS ReceiptNo, Narrative, Amount, Mode, Bank, Del" & _
	" FROM TransactionsList_Receipt WHERE [Payment Date] = '" & FormatDate(selectedTradeDate) & "'"
	
	Set Rs = CreateObject("ADODB.Recordset")		
	'Rs.CursorLocation = adUseClient		
	'Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
	'Rs.Filter = "LotTDate = '" & FormatDate(selectedTradeDate) & "'"
	Set Rs = conn.execute(SqlStr)
		
	If Not(rs.EOF Or rs.BOF) Then
		Do Until rs.EOF
			%>
			<tr>
				<td nowrap align="left"><font size="1"><%=Day(rs.Fields("PaymentDate")) & " " & MonthName(Month(rs.Fields("PaymentDate")), True) & " " & Year(rs("PaymentDate"))%></font></td>
				<td nowrap align="left"><font size="1"><%=rs("ReceiptNo")%></font></td>
				<td nowrap align="left"><font size="1"><%=rs("Narrative")%></font></td>
				<td nowrap align="right"><font size="1"><%=FormatNumber(rs("Amount"),2)%></font></td>
				<td nowrap align="left"><font size="1"><%=rs("Mode")%></font></td>
				<td nowrap align="left"><font size="1"><%=rs("Bank")%></font></td>
				<td nowrap align="left"><font size="1"><%=rs("Del")%></font></td>
				<td nowrap align="left"><font size="1"><%=rs("ChangedBy")%></font></td>
				<td nowrap align="left"><font size="1"><%=Day(rs.Fields("ModDate")) & " " & MonthName(Month(rs.Fields("ModDate")), True) & " " & Year(rs("ModDate"))%></font></td>
			</tr>      
			<%
			
			TotalAmount = TotalAmount + rs("Amount")
			
			rs.MoveNext
		Loop
		%>
		<tr>
			<td nowrap colspan=3 align="left"><font size="1">&nbsp;</font></td>
			<td nowrap align="right" style="font-weight:bold;border-top:1 solid black;border-bottom:1 solid black;"><font size="1"><b><%=FormatNumber(TotalAmount,2)%></b></font></td>
			<td nowrap colspan=5 align="left"><font size="1">&nbsp;</font></td>
		</tr>
		<%
	else
		%>
		<script language = 'javascript'>
			alert ("No transactions found using the specified criteria");
			window.parent.history.go(-1);          		
		</script>
		<%
		Set Rs = Nothing
		Set Conn = Nothing
		Response.end
	End if
	%>
</table>
 
</body>

</html>