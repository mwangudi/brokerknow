<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Profit and Loss Account</title>

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
	<form method="POST" action="ProfitAndLoss1.asp" Name="frmMain" id="frmMain">
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
	
	<%Set rs = Nothing
	Set Conn = Nothing
	Response.End
End If

%>

<% DrawPageFunctions True, True, True %>
<%
	Set conn = GetActiveConnection("KBroker")
	Set Rs = CreateObject("ADODB.Recordset")						        	
	Rs.CursorLocation = adUseClient	
	
	Dim Subtotal
	Dim total
	
	Subtotal=0
	total=0
	sqlStr="SELECT SUM(ISNULL(Credit-Debit, 0)) AS CurrentBal, Entity_DPA_ as AccountCode,'Broker Commission' as AccountName FROM  dbo.BrokerCommissionStatement" & _
	       " WHERE (TransDate between '" & FormatDate(selectedFromDate) & "' and '" & FormatDate(CDate(selectedToDate)) & "')  GROUP BY Entity_DPA_"
	
	sqlStr=sqlStr & " Union All " & _
			"SELECT SUM(DB_BankAccountStatement.Credit - DB_BankAccountStatement.Debit) AS CurrentBal, Account.AccountCode as AccountCode,Account.AccountName" & _
		 	" FROM  dbo.DB_BankAccountStatement INNER JOIN dbo.Account ON dbo.DB_BankAccountStatement.BankAccount_DPA_ = dbo.Account.Account_DPA_" & _
		 	" WHERE  (Account.AccountTypeLevel1 = 1) and (dbo.DB_BankAccountStatement.TransDate between '" & FormatDate(selectedFromDate) & "' and '" & FormatDate(CDate(selectedToDate)) & "')" & _			
		 	" GROUP BY Account.AccountCode,Account.AccountName"  
	
	'Response.write(sqlStr)
	'Response.end

	Set Rs = Conn.Execute(sqlStr) %>	
	<table border="0" cellspacing="2" cellpadding="2" style="font-family: Arial Narrow" width="100%">
		<tr>
		<td width="10%" nowrap><font face="Impact" size="3">PROFIT and LOSS
		    STATEMENT</font></td>
		<td width="60%" nowrap align="right"><font face="Impact" size="3"><%= Session("CompanyName") %>
		    </font></td>
		</tr>
	</table>
	<br>
	<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">		
		<tr>
		<td width="1%"><b>Between: </b></td>
		<td width="48%"><b><%=FormatDate(selectedFromDate)%> &nbsp; And &nbsp;<%=FormatDate(selectedToDate)%></b>
		</td>
		</tr>
	</table>
	<br>
	<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow; LEFT-MARGIN:100PX" width="100%">		
		<tr>
		<td style="border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"><b><font face="Arial Narrow" size="3">Income</font></b></td>
		<td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" align="right"><b><font face="Arial Narrow" size="3">&nbsp;</font></b></td>
		<td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" align="right"><b><font face="Arial Narrow" size="3">&nbsp;</font></b></td>
		</tr>
		<tr>
		  <td><b>Account&nbsp;Code</b></td>
		  <td><b>Account&nbsp;Name</b></td>
		  <td align="right"><b>Amount</b></td>
		</tr>
		
	<%
    Do Until Rs.EOF 
		%>
		<tr>
		  <td><%= Rs.Fields("AccountCode").Value %>
		  </td>
		  <td><%= Rs.Fields("AccountName").Value %>
		  </td>
		  <td align="right"><%=FormatNum(Abs(Rs.Fields("CurrentBal").Value))%>
		  </td>
		</tr>
		<%
		Subtotal=Subtotal+Abs(Rs.Fields("CurrentBal").Value)
		Rs.MoveNext		
	Loop
	%>
	<tr>
	    <td>&nbsp;</td>
	    <td align="Right"><b>Sub Total</b></td>
	    <td align="right"><%=FormatNum(Subtotal)%>
	    </td>
	</tr>
	<%
	Total=Total+Subtotal
	Subtotal=0
	
	sqlStr = "SELECT SUM(DB_BankAccountStatement.Credit - DB_BankAccountStatement.Debit) AS CurrentBal, Account.AccountCode as AccountCode,Account.AccountName" & _
		 " FROM  dbo.DB_BankAccountStatement INNER JOIN dbo.Account ON dbo.DB_BankAccountStatement.BankAccount_DPA_ = dbo.Account.Account_DPA_" & _
		 " WHERE  (Account.AccountTypeLevel1 = 2) and (dbo.DB_BankAccountStatement.TransDate between '" & FormatDate(selectedFromDate) & "' and '" & FormatDate(CDate(selectedToDate)) & "')" & _			
		 " GROUP BY Account.AccountCode,Account.AccountName,DB_BankAccountStatement.IsOpeningBalance, dbo.Account.AccountOpeningBal"  

	Set Rs = Conn.Execute(sqlStr)		
	%>
	<tr><td colspan="3">&nbsp;</td></tr>	
	<tr>
		<td style="border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" colspan="2"><b><font face="Arial Narrow" size="3">Expenditure</font></b></td>
		<td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" align="right"><b><font face="Arial Narrow" size="3">&nbsp;</font></b></td>
		<td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" align="right"><b><font face="Arial Narrow" size="3">&nbsp;</font></b></td>
		</tr>
		<tr>
		  <td><b>Account&nbsp;Code</b></td>
		  <td><b>Account&nbsp;Name</b></td>
		  <td align="right"><b>Amount</b></td>
		</tr>
		
			<%
    Do Until Rs.EOF 
		%>
		<tr>
		  <td><%= Rs.Fields("AccountCode").Value %>
		  </td>
		  <td><%= Rs.Fields("AccountName").Value %>
		  </td>
		  <td align="right"><%=FormatNum(Rs.Fields("CurrentBal").Value)%>
		  </td>
		</tr>
		
		<%
		Subtotal=Subtotal+Rs.Fields("CurrentBal").Value
		Rs.MoveNext
		
	Loop
	%>
	<tr>
	    <td>&nbsp;</td>
	    <td align="Right"><b>Sub Total</b></td>
	    <td align="right"><%=FormatNum(Subtotal)%>
	    </td>
	</tr>

	<%
	Total=Total+Subtotal	
	%>       
	<tr>
	    <td colspan="3">&nbsp;</td>	    
	</tr>
	
	<tr>
	    <td>&nbsp;</td>
	    <td align="Right"><b>Profit/(Loss)</b></td>
	    <td style="border: 1px solid #000000" valign="top" align="right"><%=FormatNum(Total)%>
	    </td>
	</tr>
	<tr>
	    <td colspan="3">&nbsp;</td>	    
	</tr>
	<%
	
	Subtotal=0
	
	sqlStr="SELECT SUM(ISNULL(Balance, 0)) AS CurrentBal FROM  dbo.BrokerCommissionStatement" & _
	       " WHERE (TransDate between '" & FormatDate(selectedFromDate) & "' and '" & FormatDate(CDate(selectedToDate)) & "') and IsOpeningBalance=1"
	
	sqlStr= sqlStr & " Union All	SELECT SUM(DB_BankAccountStatement.CreditBal - DB_BankAccountStatement.Debit) AS CurrentBal" & _
		   " FROM  dbo.DB_BankAccountStatement INNER JOIN dbo.Account ON dbo.DB_BankAccountStatement.BankAccount_DPA_ = dbo.Account.Account_DPA_" & _
		   " WHERE  (Account.AccountTypeLevel1 = 2 or Account.AccountTypeLevel1 = 1) and (dbo.DB_BankAccountStatement.TransDate between '" & FormatDate(selectedFromDate) & "' and '" & FormatDate(CDate(selectedToDate)) & "')" & _			
		   " and IsOpeningBalance=1"  
	
	'Response.Write(sqlStr)
	'Response.End
	
	Set Rs = Conn.Execute(sqlStr)		
	%>
	<tr><td colspan="3">&nbsp;</td></tr>	
	<tr>
		<td style="border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" colspan="2"><b><font face="Arial Narrow" size="3">Balances</font></b></td>
		<td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" align="right"><b><font face="Arial Narrow" size="3">&nbsp;</font></b></td>
		<td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" align="right"><b><font face="Arial Narrow" size="3">&nbsp;</font></b></td>
		</tr>
		<tr>
		  <td><b>&nbsp;</b></td>
		  <td><b>Particulars</b></td>
		  <td align="right"><b>Amount</b></td>
		</tr>
		
			<%
    Do Until Rs.EOF 		
		Subtotal=Subtotal+Rs.Fields("CurrentBal").Value		
		Rs.MoveNext
		
	Loop
	%>
		<tr>
		  <td>&nbsp;</td>
		  <td>Opening&nbsp;Balance</td>
		  <td align="right"><%=FormatNum(Subtotal)%>
		  </td>
		</tr>		
		
		<tr>
		  <td>&nbsp;</td>
		  <td>Profit/(Loss)</td>
		  <td align="right"><%=FormatNum(Total)%>
		  </td>
		</tr>
	<tr>
	    <td>&nbsp;</td>
	    <td align="Right"><b>Closing&nbsp;Balance</b></td>
	    <td style="border: 1px solid #000000" valign="top" align="right"><%=FormatNum(Subtotal+Total)%>
	    </td>
	</tr>
	</table>
	
	<%
Set Rs = Nothing
Set Conn = Nothing%>

</body>

</html>
