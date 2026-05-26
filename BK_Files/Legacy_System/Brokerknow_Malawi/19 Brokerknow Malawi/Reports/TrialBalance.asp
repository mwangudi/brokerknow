<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Trial Balance</title>

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
	<form method="POST" action="TrialBalance.asp" Name="frmMain" id="frmMain">
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

<% DrawPageFunctions True, True, True

'sqlStr = "TrialBalanceProc"

Set Conn = GetActiveConnection("KBroker")
Set Rs = CreateObject("ADODB.Recordset")		
  
sqlStr = "Exec TrialBalanceProc '" & FormatDate(selectedFromDate) & "', ' " & FormatDate(selectedToDate) & "' "

Rs.CursorLocation = adUseClient	
Rs.Open SQLServerFormat((sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic

 If Rs.EOF Or Rs.BOF Then%>
		<Script Language="JavaScript">	
			ShowMessage('No data was found using the criteria entered.');
			window.parent.history.go(-1);
		</Script>
		<%Set Conn = Nothing
		Set Rs = Nothing
		Response.End
  End If


 %>

<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
     <tr>
		<td nowrap><b><font face="Arial Narrow" size="4"><%= report_description %></font></b></td>
		<td nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>
	</tr>	
       <tr>
		  <td COLSPAN=2><font face="Arial" size="2">&nbsp;</font></td>
	</tr>
</table>			
<br>
	<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
		<tr>
		<td width="1%"><b>Date Between: </b></td>
		<td width="48%"><%= FormatDate(selectedFromDate) %>&nbsp;&nbsp;<%= FormatDate(selectedToDate) %>
		</td>
		</tr>
	</table>
	<br>

    <table border="0" width="100%" cellPadding="2" cellSpacing=0>
    <tr bgColor="#000000">
		<td><b><font color="#FFFFFF">Account Code</font></b></td>
		<td><b><font color="#FFFFFF">Account Name</font></b></td>
		<td align="right"><b><font color="#FFFFFF">Debit</font></b></td>
		<td align="right"><b><font color="#FFFFFF">Credit</font></b></td>
	</tr>
	
	<%
	DebitTotal=0
	CreditTotal=0
	Do Until rs.EOF
	DebitTotal=DebitTotal+Rs.Fields("Debit").Value
	CreditTotal=CreditTotal+Rs.Fields("Credit").Value
	%>
        	<tr>
        		<td ><%=Rs.Fields("Account Code").Value%></td>
				<td ><%=Rs.Fields("AccountName").Value%></td>
        		<td align="right"><%=FormatNum(Rs.Fields("Debit").Value)%></td>
        		<td align="right"><%=FormatNum(Rs.Fields("Credit").Value)%></td>      				 
	        </tr>
	 <%  
	Rs.MoveNext
	 Loop

	 If trim(DebitTotal) <> "" Then%>
			<tr>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
			</tr>
			<tr>
				<td><b>Total</b></td>
				<td>&nbsp;</td>
				<td align="right"><%= FormatNum(DebitTotal) %></td>
				<td align="right"><%= FormatNum(CreditTotal) %></td>
			</tr>
	 <%end if
	 %>	  	
  </table>
  
	 <%	 
	 Set Rs = Nothing
	 Set Conn = Nothing
     %>	
</body>

</html>