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
selectedClient = Request.Form("cboClient")
selectedFromDate = Request.Form("transFromDate")

If genReport <> "1" Or selectedClient = "" Or Not IsDate(selectedFromDate) Then%>
	<Script Language="JavaScript">
		report_SetBodyClass();
		function validateForm(frm){			
			if (frm.cboClient.selectedIndex < 0){
				alert("Select a client");
				frm.cboClient.focus();
				return;
			}
			
			frm.target = '_self';			
			frm.submit();
		}
		
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "transFromDate","cmdDate","<%= FormatDate(DateAdd("d", -90, Date)) %>",1);
	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="ClientStatement.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">
		<table>
			<tr>
				<td>Client: </td>
				<td><select name = 'cboClient' id = 'cboClient' size="1">
					<option selected value = ''></option>
					<%
					        Set conn = GetActiveConnection("KBroker")
					        
					        sqlStr = "SELECT * FROM Client ORDER BY ClientName"
					        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
					        If Not (rs.EOF Or rs.BOF) Then
					                rs.MoveFirst
					                Do Until rs.EOF%>
					                        <option value = '<%=rs.Fields("Client_DPA_")%>'><%=rs.Fields("ClientName")  %></option>
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
	sqlStr = "SELECT * FROM ClientStatement WHERE Client_DPA_ = '" & selectedClient & "' AND TransDate >= '" & FormatDate(selectedFromDate) & "'"
	'Response.Write sqlstr
	'Response.End
	'sqlStr = "ClientStatement"
	Rs.CursorLocation = adUseClient	
	Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
	'Rs.Filter = "Client_DPA_ = '" & selectedClient & "' AND TransDate >= '" & FormatDate(selectedFromDate) & "'"
	
	
	If rs.EOF Or rs.BOF Then%>
		<Script Language="JavaScript">
			alert("The specified client does not have any transaction using the specified date criterion")
			window.history.go(-1);
		</Script>
		<%Set Rs = Nothing
		Set Conn = Nothing
		Response.End
	End If
	
	Set rsClient = Conn.Execute ("SELECT * FROM Client WHERE Client_DPA_ = " & selectedClient)
	
	If Not (rsClient.EOF Or rsClient.BOF) Then
		accountDesc = rsClient.Fields("ClientName").Value
		accountAddress = rsClient.Fields("ClientAddr").Value
	End If
	
	Set rsClient = Nothing
	
	
	isOpeningBalance = CBool(Rs.Fields("IsOpeningBalance").Value)
	
	

	If Not IsOpeningBalance Then
		'get latest prev balance
		
		'Set cloneRs = Rs.Clone(adLockOptimistic)
		'cloneRs.Filter = ""
		'cloneRs.Filter = "Client_DPA_ LIKE '" & selectedClient & "' AND TransDate < '" & FormatDate(selectedFromDate) & "'"
		'cloneRs.Sort = "TransDate DESC"
		'If cloneRs.EOF Or cloneRs.BOF Then
		'	cloneRs.Filter = ""
		'	cloneRs.Filter = "Client_DPA_ LIKE '" & selectedClient & "' AND TransDate >= '" & FormatDate(selectedFromDate) & "'"
		'	cloneRs.Sort = "TransDate DESC"
		'	If cloneRs.EOF Or cloneRs.BOF Then%>
				<Script Language="JavaScript">
		//			alert("A problem occured when calculating the opening balance");
		//			window.history.go(-1);
				</Script>
			<%'	Set cloneRs = Nothing
		'		Set Rs = Nothing
		'		Set Conn = Nothing
		'		Response.End
		'	End If
		'End If

		sqlStr = "SELECT Client_DPA_, '" & FormatDate(selectedFromDate) & "' AS TransDate, '' AS REF, 'Opening Balance' AS Particulars, " & _
				" 0 AS Debit, 0 AS Credit, 1 AS IsOpeningBalance, SUM(Credit - Debit) " & _
				" AS Balance FROM  ClientStatement " & _
				" WHERE     (Client_DPA_ = " & selectedClient & ") AND (TransDate < '" & FormatDate(selectedFromDate) & "')" & _
				" GROUP BY Client_DPA_"
		
		Set cloneRs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
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
		OpeningBalance = cloneRs.Fields("Balance").Value
		
	
	Else
		OpeningBalance =  Rs.Fields("Balance").Value		
		
	End If
	
	
%>	

<table border="0" cellspacing="2" cellpadding="2" style="font-family: Arial Narrow" width="100%">
    <tr>
		<td width="10%" nowrap><font face="Impact" size="4">CLIENTS STATEMENT</font></td>
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
		  <td><font size="1"><%= Day(selectedFromDate) & " " & MonthName(Month(selectedFromDate), True) & " " & Right(Year(selectedFromDate),2) %></font></td>
		  <td><font size="1">&nbsp;</font></td>
		  <td><font size="1">Opening Balance</font></td>
		  <td align="right"><font size="1">&nbsp;</font></td>
		  <td align="right"><font size="1">&nbsp;</font></td>
		  <td align="right"><font size="1"><%= CreditDebitValue(FormatNum(cloneRs.Fields("Balance").Value)) %></font></td>
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
		  <td><font size="1"><%= Day(rs.Fields("TransDate")) & " " & MonthName(Month(rs.Fields("TransDate")), True) & " " & Right(Year(rs.Fields("TransDate")),2) %></font></td>
		  <td><font size="1"><%= Rs.Fields("Ref").Value %></font></td>
		  <td><font size="1"><%= Rs.Fields("Particulars").Value %></font></td>
		  <td align="right"><font size="1"><% If Rs.Fields("Debit").Value <> "0" Then 
									Response.Write FormatNum(Rs.Fields("Debit").Value) 
							   End If			%>
            </font>
		  </td>
		  <td align="right"><font size="1"><% If Rs.Fields("Credit").Value <> "0" Then
									Response.Write FormatNum(Rs.Fields("Credit").Value) 
							   End If			
							%>
            </font>
		  </td>
		  <td align="right">
            <font size="1">
			<% If Not IsOpeningBalance Then
					runningBal = runningBal + (Rs.Fields("Credit").Value - Rs.Fields("Debit").Value)					
					Response.Write  FormatNum(CreditDebitValue(runningBal)) 
				Else	
					runningBal = Rs.Fields("Balance").Value					
					Response.Write FormatNum(Rs.Fields("Balance").Value)
				End If 
				lastDate = Rs.Fields("TransDate").Value%>
            </font>
		  </td>
		</tr>
	
	<%	Rs.MoveNext
	Loop
	%>
	
		<tr>	
		  <td><font size="1"><%= Day(lastDate) & " " & MonthName(Month(lastDate), True) & " " & Right(Year(lastDate),2) %></font></td>
		  <td><font size="1">&nbsp;</font></td>
		  <td><font size="1">Closing Balance</font></td>
		  <td align="right"><font size="1">&nbsp;</font></td>
		  <td align="right"><font size="1">&nbsp;</font></td>
		  <td align="right"><font size="1"><%= CreditDebitValue(FormatNum(runningBal)) %></font></td>
		</tr>  
	
    <tr>
      <td colspan="6" align="right" style="border-bottom-style: solid; border-bottom-width: 1">
        &nbsp;&nbsp;&nbsp; </td>

    </tr>

    <tr>
      <td colspan="5" align="right">
        <font size="1">
        Opening Balance:</font></td>

      <td align="right"><font size="1"><%= FormatNum(CreditDebitValueRev(OpeningBalance)) %></font></td>
    </tr>

    <tr>
      <td colspan="5" align="right"><font size="1">less Total Debits:</font></td>

      <td align="right"><font size="1"><%= FormatNum(0 - totalDebits) %></font></td>
    </tr>

    <tr>
      <td colspan="5" align="right"><font size="1">add Total Credits:</font></td>

      <td align="right"> <font size="1"> <%= FormatNum(totalCredits) %></font></td>
    </tr>

    <tr>
      <td colspan="6" align="right" style="border-bottom-style: solid; border-bottom-width: 1">
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>

    </tr>

    <tr>
      <td colspan="2"><font size="1">Current: <%= CreditDebitValue(FormatNum(runningBal)) %></font></td>
      <td><font size="1">30-60 Days:</font> 
		
      </td>
      <td><font size="1">Over 60 Days:</font> </td>
      <td align="right"><b><font size="1">Total Balance:</font></b></td>
      <td align="right"><b><font size="1"><%= CreditDebitValue(FormatNum(runningBal)) %></font></b></td>
    </tr>

  </table>
   
<%Set Rs = Nothing
Set Conn = Nothing%>   
</body>

</html>
