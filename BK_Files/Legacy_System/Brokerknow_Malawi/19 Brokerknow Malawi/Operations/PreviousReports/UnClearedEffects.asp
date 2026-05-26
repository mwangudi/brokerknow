<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Uncleared Effects</title>

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
selectedBank = Request.Form("cboBank")
selectedFromDate = Request.Form("transFromDate")
selectedToDate = Request.Form("txtToDate")

FirstDay=DateSerial(Year(Date), Month(Date)-1 + iOffset, 1)

If genReport <> "1" Or selectedBank = "" Or Not IsDate(selectedFromDate) Then%>
	<Script Language="JavaScript">
		report_SetBodyClass();
		
		function validateForm(frm){			
			if (frm.cboBank.selectedIndex < 0){
				alert("Select a bank");
				frm.cboBank.focus();
				return;
			}
			
			frm.target = '_self';			
			frm.submit();
		}
		
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "transFromDate","cmdDate","<%= FormatDate(FirstDay) %>",1);
		var cal1=new ctlSpiffyCalendarBox("cal1", "frmMain", "txtToDate","cmdDate","<%= FormatDate(Date) %>",1);

	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="UnClearedEffects.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">
		<table>
			<tr>
				<td>Account: </td>
				<td><input type = 'text' name ='txtClientCode' id = 'txtClientCode' size="10" onBlur="txtval = this.value; selectItem(cboBank);"></td>
				<td><select name = 'cboBank' id = "cboBank" size="1" 
    					onchange='UpdateCode(true,cboBank,txtClientCode)'
					onKeypress="return (dodefaultaction()==''); " 
					onKeydown="return (dodefaultaction()==''); " 
					onKeyup="return (FilterData(this,1,UpdateCode(change(cboBank,0),cboBank,txtClientCode)));" 
					onfocus="txtval = '';inputIsItemCode = 1;" 
					onblur="txtval = '';inputIsItemCode = 1;">
					<option selected SearchCode = "" SearchText = ""  value = ''></option>
					<%

					        Set conn = GetActiveConnection("KBroker")
					        
					        sqlStr = "SELECT * FROM Account WHERE (AccountTypeLevel1 = 7) AND (NOT (AccountName LIKE N'%CASH%'))ORDER BY AccountName"
							'Response.write(sqlStr)
							'Response.end

					        Set rs = conn.Execute(sqlStr)
					        If Not (rs.EOF Or rs.BOF) Then
					                rs.MoveFirst
					                Do Until rs.EOF%>					                        
									<option SearchCode = "<%=rs.Fields("AccountCode")%>" SearchText = "<%=rs.Fields("AccountName")%>" value = '<%=rs.Fields("Account_DPA_")%>'><%=Mid(rs.Fields("AccountName"),1,30)%></option>
					                        <%rs.MoveNext
					                Loop
					        End If
					%>

					    </select>
				</td>
			</tr>

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
				<td colspan="3"><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... " id=Button1 name=Button1>&nbsp;&nbsp;</td>
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

'Response.write(selectedBank)
'Response.end

	Set conn = GetActiveConnection("KBroker")
	Set Rs = CreateObject("ADODB.Recordset")						        
	sqlStr = "SELECT * FROM DB_Reconciliation where BankAccount_DPA_ =" & selectedBank & " AND (ReconcileDate is null ) and TransDate between '" & FormatDate(CDate(selectedFromDate)) & "' and '" & FormatDate(CDate(selectedToDate)+1) & "'"		
	Rs.CursorLocation = adUseClient	
	Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
	'Rs.Filter = ""
	
	If rs.EOF Or rs.BOF Then%>
		<Script Language="JavaScript">
			alert("The specified bank does not have any transaction using the specified date criterion")
			window.history.go(-1);
		</Script>
		<%Set Rs = Nothing
		Set Conn = Nothing
		Response.End
	End If
	
	Set rsBank = Conn.Execute ("SELECT * FROM Account WHERE Account_DPA_ = " & selectedBank)
	If Not (rsBank.EOF Or rsBank.BOF) Then
		accountDesc = rsBank.Fields("AccountName").Value  & "&nbsp;[" & rsBank.Fields("AccountCode").value & "]"
		'accountAddress = rsBank.Fields("ClientAddr").Value
	End If
	
	Set rsBank = Nothing
	
	isOpeningBalance = CBool(Rs.Fields("IsOpeningBalance").Value)
	
	If Not IsOpeningBalance Then
		'get latest prev balance
		
		sqlStr = "SELECT CASE (DB_Reconciliation.IsOpeningBalance) WHEN 0 THEN SUM(ISNULL(dbo.DB_Reconciliation.Credit - dbo.DB_Reconciliation.Debit, 0))" & _ 
                 " ELSE 0 END + dbo.Account.AccountOpeningBal AS CurrentBal, dbo.DB_Reconciliation.BankAccount_DPA_" & _
				 " FROM  dbo.DB_Reconciliation INNER JOIN dbo.Account ON dbo.DB_Reconciliation.BankAccount_DPA_ = dbo.Account.Account_DPA_" & _
				 " WHERE  not (dbo.DB_Reconciliation.ReconcileDate is null) and TransDate between '" & FormatDate(CDate(selectedFromDate)) & "' and '" & FormatDate(CDate(selectedToDate)+1) & "' GROUP BY dbo.DB_Reconciliation.BankAccount_DPA_,DB_Reconciliation.IsOpeningBalance, dbo.Account.AccountOpeningBal having dbo.DB_Reconciliation.BankAccount_DPA_=" & selectedBank  		
		
		Set cloneRs = Conn.Execute(sqlStr)
		
		if not(cloneRs.eof and cloneRs.bof)	then
		OpeningBalance =  cloneRs.Fields("CurrentBal").Value		
		else
		OpeningBalance=0
		end if
	Else
		OpeningBalance =  Rs.Fields("Balance").Value		
		
	End If
	
	
%>	


<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="600">
    <tr>
      <td width="10%" nowrap><font face="Impact" size="4">UNCLEARED EFFECTS</font></td>
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


</table>
<BR>



  <table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow; LEFT-MARGIN:100PX"  width="800">      
      <tr>
      <td style="border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" width="100"><b><font face="Arial Narrow" size="3">Date</font></b></td>
      <td style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"><b><font face="Arial Narrow" size="3">Ref:</font></b></td>
      <td style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"><b><font face="Arial Narrow" size="3">Receipt No:</font></b></td>
      <td style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"><b><font face="Arial Narrow" size="3">Particulars:</font></b></td>
      <td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"  align=right><b><font face="Arial Narrow" size="3">Debit:</font></b></td>
      <td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" align=right><b><font face="Arial Narrow" size="3">Credit:</font></b></td>
      <td align="right" style="border-right-style: solid; border-right-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"  align=right><b><font face="Arial Narrow" size="3">Balance</font></b></td>
    </tr>
    
    <%If Not IsOpeningBalance Then%>
		<tr>	
		  <td width="100"><%= FormatDate(selectedFromDate) %></td>
		  <td>&nbsp;</td>
		  <td>&nbsp;</td>
		  <td>Opening total of all cleared effects</td>
		  <td align="right">&nbsp;</td>
		  <td align="right">&nbsp;</td>
		  <td align="right"><%= FormatNum(CreditDebitValue(OpeningBalance)) %></td>
		</tr>    
    <%	runningBal = CreditDebitValueRev(OpeningBalance)
		OpeningBalance = OpeningBalance
		Set cloneRs = Nothing
    End If
    
    totalDebits = 0
    totalCredits = 0
    
    Do Until Rs.EOF
		totalDebits = totalDebits + Rs.Fields("Debit").Value 
		totalCredits = totalCredits + Rs.Fields("Credit").Value%>
		<tr>	
		  <td width="100"><%= FormatDate(Rs.Fields("TransDate").Value) %></td>
		  <td><%= Rs.Fields("Ref").Value %></td>
		  <% if(Rs.Fields("ReceiptNo").Value =0 ) then %>
		  <td>&nbsp;</td>
		  <% else %>
		  <td><%= Rs.Fields("ReceiptNo").Value %></td>
		  <% end if %>
		  <td><%= Rs.Fields("Particulars").Value %></td>
		  <td align="right"><% If Rs.Fields("Debit").Value <> "0" Then
									Response.Write FormatNum(Rs.Fields("Debit").Value)
							   End If %>
		  </td>
		  <td align="right"><%If Rs.Fields("Credit").Value Then
									Response.Write FormatNum(Rs.Fields("Credit").Value) 
							  End If%>
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
	
	<%	
		Rs.MoveNext
	Loop
	
	%>
	
		<tr>	
		  <td><%= FormatDate(lastDate) %></td>
		  <td>&nbsp;</td>
		  <td>&nbsp;</td>
		  <td>Closing Balance</td>
		  <td align="right">&nbsp;</td>
		  <td align="right">&nbsp;</td>
		  <td align="right"><%= CreditDebitValue(FormatNum(runningBal)) %></td>
		</tr>  
	
    <tr>
      <td colspan="7" align="right" style="border-bottom-style: solid; border-bottom-width: 1">
        &nbsp;&nbsp;&nbsp; </td>

    </tr>

    <tr>
      <td colspan="6" align="right">
        Opening total of all cleared effects:</td>

      <td align="right"><%= FormatNum(CreditDebitValueRev(OpeningBalance)) %></td>
    </tr>

    <tr>
      <td colspan="6" align="right">less Total Debits:</td>

      <td align="right"><%= FormatNum(0 - totalDebits) %></td>
    </tr>

    <tr>
      <td colspan="6" align="right">add Total Credits:</td>

      <td align="right"> <%= FormatNum(totalCredits) %></td>
    </tr>

    <tr>
      <td colspan="7" align="right" style="border-bottom-style: solid; border-bottom-width: 1">
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>

    </tr>

    <tr>
      <td colspan="2">Current: <%= CreditDebitValue(FormatNum(runningBal)) %></td>
      <td>30-60 Days: 
		
      </td>
      <td>Over 60 Days: </td>
      <td colspan="2">
        <p align="right"><b>Total Balance:</b></td>
      <td align="right"><b><%= CreditDebitValue(FormatNum(runningBal)) %></b></td>
    </tr>


  </table>
   
</body>

</html>