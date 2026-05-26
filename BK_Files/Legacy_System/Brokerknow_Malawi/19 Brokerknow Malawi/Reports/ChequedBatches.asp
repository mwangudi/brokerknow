 <html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Paid Up Batches</title>
	<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
	<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
	<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
	<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
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
			margin-right: 5cm;
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
<% DrawPageFunctions True, True, True %>

<%
	Set conn = GetActiveConnection("KBroker")
	Set Rs = CreateObject("ADODB.Recordset")						        
	
	sqlStr = "SELECT * FROM ForwardsList where Batch_No is null"
	
	sqlStr = "SELECT     Payment.PaymentAmount as Amount, Account.AccountName, Offerings.Batch_No, Payment.PaymentPDate, Payment.TimeChanged,  " & _
				"                       Users.Surname + ' ' + Users.OtherNames AS ChangedBy,Security.SecurityName,Payment.PaymentReference " & _
				" FROM         Payment INNER JOIN " & _
				"                       Offerings ON Payment.Batch_No = Offerings.Batch_No INNER JOIN " & _
				"                       Account ON Payment.BankAccount_DPA_ = Account.Account_DPA_ INNER JOIN " & _
				"                       Users ON Payment.ChangedBy = Users.UserID " & _
				" INNER JOIN Security ON Offerings.Offering = Security.Security_DPA_ where payment.Deleted=0"
	
	'Response.write(sqlStr)
	'Response.end

	Rs.CursorLocation = adUseClient	
	Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
	
	If rs.EOF Or rs.BOF Then%>
		<Script Language="JavaScript">
			alert("The specified Batch No does not have any transaction using the specified date criterion")
			window.history.go(-1);
		</Script>
		<%Set Rs = Nothing
		Set Conn = Nothing
		Response.End
	End If
	
%>	

<table border="0" cellspacing="2" cellpadding="2" style="font-family: Arial Narrow" width="100%">
    <tr>
		<td width="10%" nowrap><font face="Impact" size="4">Paid Up Batches</font></td>
      <td width="60%" nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>
      
    </tr>

  </table>
<br>
<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
    <tr>
      <td width="1%"><b>Date:</b></td>
      <td width="48%"><%= FormatDate(Date) %></td>
    </tr>
	<tr>
      <td width="1%"><b>&nbsp;</b></td>
      <td width="48%"><font face="Arial Narrow" size="3"><b><%= Rs("SecurityName") %></b></font></td>
    </tr>
</table>
<BR>


  <table border="0" cellspacing="3" cellpadding="0" style="font-family: Arial Narrow; LEFT-MARGIN:100PX"  width="100%">
    <tr>      
      <td style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"  nowrap><b><font face="Arial Narrow" size="2">Batch No</font></b></td>
      <td align="right" style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"><b><font face="Arial Narrow" size="2">Amount</font></b></td>
      <td style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"  nowrap><b><font face="Arial Narrow" size="2">Date:</font></b></td>      
	  <td style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" nowrap><b><font face="Arial Narrow" size="2">Bank:</font></b></td>
      <td style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" nowrap><b><font face="Arial Narrow" size="2">Cheque No:</font></b></td>   
      <td style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" nowrap><b><font face="Arial Narrow" size="2">Paid By</font></b></td>	  
    </tr>

   <%
   totalquantity=0
   totalpayable=0
   do until Rs.eof
	%>
	<tr>		
		<td><%=Rs("Batch_No")%></td>
		<td align="right"><%=Rs("Amount")%></td>
		<td nowrap><%=FormatDate(RS("PaymentPDate"))%></td>		
		<td nowrap><%=Mid(RS("AccountName"),1,30)%></td>
		<td nowrap><%=RS("PaymentReference")%></td>		
		<td><%=Rs("ChangedBy")%></td>
	</tr>
	<%	
	totalpayable = totalpayable + Rs("Amount")
	
	Rs.MoveNext
   loop

   %>
   <tr>
		<td><b>Total</b></td>
		<td align="right"><%=FormatNum(totalpayable)%></td>		
		<td colspan="4">&nbsp;</td>		
   </tr>
  </table>
   
<%Set Rs = Nothing
Set Conn = Nothing%>   
</body>

</html>