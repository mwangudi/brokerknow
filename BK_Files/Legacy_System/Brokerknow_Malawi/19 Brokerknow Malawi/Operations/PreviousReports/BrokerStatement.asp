<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Client Statement</title>
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

<%

genReport = Request.Form("genReport")
selectedBroker = Request.Form("cboBroker")
selectedFromDate = Request.Form("transFromDate")

If genReport <> "1" Or selectedBroker = "" Or Not IsDate(selectedFromDate) Then%>
	<Script Language="JavaScript">
		report_SetBodyClass();
		function validateForm(frm){			
			if (frm.cboBroker.selectedIndex < 0){
				alert("Select a broker");
				frm.cboBroker.focus();
				return;
			}
			
			frm.target = '_self';			
			frm.submit();
		}
		
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "transFromDate","cmdDate","<%= FormatDate(DateAdd("d", -90, Date)) %>",1);
	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="BrokerStatement.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">
		<table>
			<tr>
				<td>Broker: </td>
				<td><select name = 'cboBroker' id = 'cboBroker' size="1">
					<option selected value = ''></option>
					<%
					        Set conn = GetActiveConnection("KBroker")
					        
					        sqlStr = "SELECT * FROM Broker ORDER BY BrokerName"
					        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
					        If Not (rs.EOF Or rs.BOF) Then
					                rs.MoveFirst
					                Do Until rs.EOF%>
					                        <option value = '<%=rs.Fields("Broker_DPA_")%>'><%=rs.Fields("BrokerName")  %></option>
					                        <%rs.MoveNext
					                Loop
					        End If
					%>

					    </select>
				</td>
			</tr>
			<tr>
				<td>Select date from:</td>
				<td>
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>
			</tr>
			<tr>
				<td colspan=2><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... ">&nbsp;&nbsp;</td>
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
	sqlStr = "SELECT * FROM BrokerStatement"
	'sqlStr = "ClientStatement"
	Rs.CursorLocation = adUseClient	
	Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
	Rs.Filter = "Broker_DPA_ LIKE '" & selectedBroker & "' AND TransDate >= '" & FormatDate(selectedFromDate) & "'"
	
	
	If rs.EOF Or rs.BOF Then%>
		<Script Language="JavaScript">
			alert("The specified agent does not have any transaction using the specified date criterion")
			window.history.go(-1);
		</Script>
		<%Set Rs = Nothing
		Set Conn = Nothing
		Response.End
	End If
	
	Set rsBroker = Conn.Execute ("SELECT * FROM Broker WHERE Broker_DPA_ = " & selectedBroker)
	
	If Not (rsBroker.EOF Or rsBroker.BOF) Then
		accountDesc = rsBroker.Fields("BrokerName").Value
		accountAddress = rsBroker.Fields("BrokerAddr").Value
	End If
	
	Set rsBroker = Nothing
	
	isOpeningBalance = CBool(Rs.Fields("IsOpeningBalance").Value)
	
	If Not IsOpeningBalance Then
		'get latest prev balance
		
		Set cloneRs = Rs.Clone(adLockOptimistic)
		cloneRs.Filter = ""
		cloneRs.Filter = "Broker_DPA_ LIKE '" & selectedBroker & "' AND TransDate < '" & FormatDate(selectedFromDate) & "'"
		cloneRs.Sort = "TransDate DESC"
		If cloneRs.EOF Or cloneRs.BOF Then
			cloneRs.Filter = ""
			cloneRs.Filter = "Broker_DPA_ LIKE '" & selectedBroker & "' AND TransDate >= '" & FormatDate(selectedFromDate) & "'"
			cloneRs.Sort = "TransDate DESC"
			If cloneRs.EOF Or cloneRs.BOF Then%>
				<Script Language="JavaScript">
					alert("A problem occured when calculating the opening balance");
					window.history.go(-1);
				</Script>
			<%	Set cloneRs = Nothing
				Set Rs = Nothing
				Set Conn = Nothing
				Response.End
			End If
		End If
		
		OpeningBalance = cloneRs.Fields("Balance").Value
		
	
	Else
		OpeningBalance =  Rs.Fields("Balance").Value		
		
	End If
	
	
%>	

<table border="0" cellspacing="2" cellpadding="2" style="font-family: Arial Narrow" width="100%">
    <tr>
		<td width="10%" nowrap><font face="Impact" size="4">BROKERS STATEMENT</font></td>
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
      <td width="1%"><b>Account:</b></td>
      <td width="48%"><%= accountDesc %></td>
    </tr>

    <tr>
      <td width="1%"><b><font size="2" face="Arial">&nbsp;</font></b></td>
      <td width="48%"><%= accountAddress %></td>
    </tr>

</table>
<BR>


  <table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow; LEFT-MARGIN:100PX"  width="100%">
    <tr>
      <td style="border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"><b><font face="Arial Narrow" size="3">Date</font></b></td>
      <td style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"><b><font face="Arial Narrow" size="3">Ref:</font></b></td>
      <td style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"><b><font face="Arial Narrow" size="3">Particulars:</font></b></td>
      <td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"  align=right><b><font face="Arial Narrow" size="3">Debit:</font></b></td>
      <td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" align=right><b><font face="Arial Narrow" size="3">Credit:</font></b></td>
      <td align="right" style="border-right-style: solid; border-right-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"  align=right><b><font face="Arial Narrow" size="3">Balance</font></b></td>
    </tr>

    <%If Not IsOpeningBalance Then%>
		<tr>	
		  <td><%= FormatDate(selectedFromDate) %></td>
		  <td>&nbsp;</td>
		  <td>Opening Balance</td>
		  <td align="right">&nbsp;</td>
		  <td align="right">&nbsp;</td>
		  <td align="right"><%= FormatNum(cloneRs.Fields("Balance").Value) %></td>
		</tr>    
    <%	runningBal = CreditDebitValueRev(cloneRs.Fields("Balance").Value)
		OpeningBalance = cloneRs.Fields("Balance").Value
		Set cloneRs = Nothing
    End If
    
    totalDebits = 0
    totalCredits = 0
    
    Do Until Rs.EOF
		totalDebits = totalDebits + Rs.Fields("Debit").Value 
		totalCredits = totalCredits + Rs.Fields("Credit").Value%>
		<tr>	
		  <td><%= FormatDate(Rs.Fields("TransDate").Value) %></td>
		  <td><%= Rs.Fields("Ref").Value %></td>
		  <td><%= Rs.Fields("Particulars").Value %></td>
		  <td align="right"><% If Rs.Fields("Debit").Value <> "0" Then 
									Response.Write FormatNum(Rs.Fields("Debit").Value) 
							   End If			%>
		  </td>
		  <td align="right"><% If Rs.Fields("Credit").Value <> "0" Then
									Response.Write FormatNum(Rs.Fields("Credit").Value) 
							   End If			
							%>
		  </td>
		  <td align="right">
			<% If Not IsOpeningBalance Then
					runningBal = runningBal + (Rs.Fields("Credit").Value - Rs.Fields("Debit").Value)					
					Response.Write  FormatNum(CreditDebitValue(runningBal)) 
				Else	
					runningBal = Rs.Fields("Balance").Value					
					Response.Write FormatNum(Rs.Fields("Balance").Value)
				End If 
				lastDate = Rs.Fields("TransDate").Value%>
		  </td>
		</tr>
	
	<%	Rs.MoveNext
	Loop
	%>
	
		<tr>	
		  <td><%= FormatDate(lastDate) %></td>
		  <td>&nbsp;</td>
		  <td>Closing Balance</td>
		  <td align="right">&nbsp;</td>
		  <td align="right">&nbsp;</td>
		  <td align="right"><%= CreditDebitValue(FormatNum(runningBal)) %></td>
		</tr>  
	
    <tr>
      <td colspan="6" align="right" style="border-bottom-style: solid; border-bottom-width: 1">
        &nbsp;&nbsp;&nbsp; </td>

    </tr>

    <tr>
      <td colspan="5" align="right">
        Opening Balance:</td>

      <td align="right"><%= FormatNum(CreditDebitValueRev(OpeningBalance)) %></td>
    </tr>

    <tr>
      <td colspan="5" align="right">less Total Debits:</td>

      <td align="right"><%= FormatNum(0 - totalDebits) %></td>
    </tr>

    <tr>
      <td colspan="5" align="right">add Total Credits:</td>

      <td align="right"> <%= FormatNum(totalCredits) %></td>
    </tr>

    <tr>
      <td colspan="6" align="right" style="border-bottom-style: solid; border-bottom-width: 1">
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>

    </tr>

    <tr>
      <td colspan="2">Current: <%= CreditDebitValue(FormatNum(runningBal)) %></td>
      <td>30-60 Days: 
		
      </td>
      <td>Over 60 Days: </td>
      <td align="right"><b>Total Balance:</b></td>
      <td align="right"><b><%= CreditDebitValue(FormatNum(runningBal)) %></b></td>
    </tr>

  </table>
   
<%Set Rs = Nothing
Set Conn = Nothing%>   
</body>

</html>
