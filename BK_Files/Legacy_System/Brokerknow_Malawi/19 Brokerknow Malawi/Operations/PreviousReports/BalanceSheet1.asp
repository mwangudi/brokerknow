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
	<form method="POST" action="BalanceSheet1.asp" Name="frmMain" id="frmMain">
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
	
	Debit=0
	Credit=0

	Subtotal=0
	total=0
	
	NetTotal=0

	sqlStr=	"SELECT CASE (DB_BankAccountStatement.IsOpeningBalance) WHEN 0 THEN SUM(ISNULL(dbo.DB_BankAccountStatement.Credit - dbo.DB_BankAccountStatement.Debit, 0))" & _ 
                 " ELSE 0 END + dbo.Account.AccountOpeningBal AS CurrentBal, Account.AccountCode as AccountCode,Account.AccountName" & _
		" FROM  dbo.DB_BankAccountStatement INNER JOIN dbo.Account ON dbo.DB_BankAccountStatement.BankAccount_DPA_ = dbo.Account.Account_DPA_" & _
		" WHERE  (Account.AccountTypeLevel1 = 7) and (dbo.DB_BankAccountStatement.TransDate between '" & FormatDate(selectedFromDate) & "' and '" & FormatDate(CDate(selectedToDate)) & "')" & _			
		" GROUP BY Account.AccountCode,Account.AccountName, DB_BankAccountStatement.IsOpeningBalance, dbo.Account.AccountOpeningBal  having (CASE (DB_BankAccountStatement.IsOpeningBalance) WHEN 0 THEN SUM(ISNULL(dbo.DB_BankAccountStatement.Credit - dbo.DB_BankAccountStatement.Debit, 0))" & _ 
                 " ELSE 0 END + dbo.Account.AccountOpeningBal) <>0 order by CurrentBal Asc"  

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
	<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow; LEFT-MARGIN:100PX" width="450">
		<tr>
		  <td><b>Account&nbsp;Code</b></td>
		  <td><b>Account&nbsp;Name</b></td>
		  <td align="right"><b>Amount</b></td>
		</tr>
		
		<tr>
		<td style="border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"><b><font face="Arial Narrow" size="3">Bank&nbsp;or&nbsp;Cash</font></b></td>
		<td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" align="right"><b><font face="Arial Narrow" size="3">&nbsp;</font></b></td>
		<td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" align="right"><b><font face="Arial Narrow" size="3">&nbsp;</font></b></td>
		</tr>
		
	<%
	Debit=0
	Credit=0
    Do Until Rs.EOF
	
		if(Rs("CurrentBal")<0) then
		Debit=1
		Credit=0
		else
		Credit=1
		Debit=0
		end if	

		if(Debit<>Debit1) then
			if(Debit=1) then
			%>
		<tr>
	    	<td><b>DEBITS</b></td>
	    	<td align="Right">&nbsp;</td>
	    	<td align="right">&nbsp;
	    	</td>
		</tr>

		<%		
			else
			%>
		<tr>
	    	<td><b>CREDITS</b></td>
	    	<td align="Right">&nbsp;</td>
	    	<td align="right">&nbsp;
	    	</td>
		</tr>

		<%	
			end if
		end if

		%>
		
		<tr>
		  <td><%= Rs.Fields("AccountCode").Value %>
		  </td>
		  <td><%= Mid(Rs.Fields("AccountName").Value,1,30) %>
		  </td>
		  <td align="right"><%=FormatNum(Abs(Rs.Fields("CurrentBal").Value))%>
		  </td>
		</tr>
		<%
		Subtotal=Subtotal+Rs.Fields("CurrentBal").Value
		Debit1=Debit
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
	AccountType1=""
	AccountType2=""

	Total=Total+Subtotal
	Subtotal=0
	
	sqlStr = "SELECT Top 100 Percent SUM(ISNULL(dbo.StatementList1.Credit - dbo.StatementList1.Debit, 0)) + dbo.Client.ClientOpeningBal AS CurrentBal, dbo.StatementList1.Client_DPA_ as AccountCode,Client.ClientName as AccountName,'Clients' as AccountType" & _
		 " FROM  dbo.StatementList1 INNER JOIN dbo.Client ON dbo.StatementList1.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN                      EntityType ON Client.EntityType_DPA_ = EntityType.EntityType_DPA_" & _
                 " WHERE (dbo.Client.Deleted = 0) AND (dbo.StatementList1.TransDate between '" & FormatDate(selectedFromDate) & "' and '" & FormatDate(CDate(selectedToDate)) & "') and EntityType.AccountType_DPA_=4" & _
		 "  GROUP BY dbo.StatementList1.Client_DPA_,Client.ClientName, dbo.Client.ClientOpeningBal HAVING      (SUM(ISNULL(StatementList1.Credit - StatementList1.Debit, 0)) + Client.ClientOpeningBal <> 0) "

	sqlStr =sqlStr & " Union all SELECT Top 100 Percent SUM(ISNULL(dbo.AgentStatement1.Credit - dbo.AgentStatement1.Debit, 0)) + dbo.Agent.AgentOpeningBal AS CurrentBal, dbo.AgentStatement1.Agent_DPA_ as AccountCode,Agent.AgentName as AccountName,'Agents' as AccountType" & _
			 " FROM  dbo.AgentStatement1 INNER JOIN dbo.Agent ON dbo.AgentStatement1.Agent_DPA_ = dbo.Agent.Agent_DPA_ inner join EntityType on Agent.EntityType_DPA_=EntityType.EntityType_DPA_" & _
			 " WHERE (dbo.Agent.Deleted = 0) AND (dbo.AgentStatement1.TransDate between '" & FormatDate(selectedFromDate) & "' and '" & FormatDate(CDate(selectedToDate)) & "') and Entitytype.AccountType_DPA_=4" & _
			 "  GROUP BY dbo.AgentStatement1.Agent_DPA_,Agent.AgentName, dbo.Agent.AgentOpeningBal having (SUM(ISNULL(dbo.AgentStatement1.Credit - dbo.AgentStatement1.Debit, 0)) + dbo.Agent.AgentOpeningBal)<>0"
		
	sqlStr =sqlStr & " Union all SELECT Top 100 Percent SUM(ISNULL(dbo.BrokerStatement1.Credit - dbo.BrokerStatement1.Debit, 0)) + dbo.Broker.BrokerOpeningBal AS CurrentBal, dbo.BrokerStatement1.Broker_DPA_ as AccountCode,Broker.BrokerName as AccountName,'Brokers' as AccountType" & _
			 " FROM  dbo.BrokerStatement1 INNER JOIN dbo.Broker ON dbo.BrokerStatement1.Broker_DPA_ = dbo.Broker.Broker_DPA_ inner join EntityType on Broker.Entitytype_DPA_=Entitytype.Entitytype_DPA_" & _
			 " WHERE (dbo.BrokerStatement1.TransDate between '" & FormatDate(selectedFromDate) & "' and '" & FormatDate(CDate(selectedToDate)) & "') and Entitytype.Accounttype_DPA_=4" & _
			 "  GROUP BY dbo.BrokerStatement1.Broker_DPA_,Broker.BrokerName, dbo.Broker.BrokerOpeningBal having (SUM(ISNULL(dbo.BrokerStatement1.Credit - dbo.BrokerStatement1.Debit, 0)) + dbo.Broker.BrokerOpeningBal)<>0"
	
	sqlStr =sqlStr & " Union all SELECT Top 100 Percent SUM(ISNULL(Balance, 0)) AS CurrentBal, Entity_DPA_ as AccountCode,EntityName as AccountName,'Levies' as AccountType FROM  dbo.CommissionStatement" & _
			 " WHERE (TransDate between '" & FormatDate(selectedFromDate) & "' and '" & FormatDate(CDate(selectedToDate)) & "')  GROUP BY Entity_DPA_,EntityName having (SUM(ISNULL(Credit - Debit, 0)))<0"		
	
	sqlStr =sqlStr & " Union all SELECT Top 100 Percent SUM(ISNULL(dbo.OwnerStatement.Credit - dbo.OwnerStatement.Debit, 0)) + dbo.Owner.OwnerOpeningBal AS CurrentBal, dbo.OwnerStatement.Owner_DPA_ as AccountCode,Owner.OwnerFName as AccountName,'Account Managers' as AccountType" & _
			 " FROM  dbo.OwnerStatement INNER JOIN dbo.Owner ON dbo.OwnerStatement.Owner_DPA_ = dbo.Owner.Owner_DPA_ inner join EntityType on Owner.EntityType_DPA_=EntityType.Entitytype_DPA_" & _
			 " WHERE (dbo.OwnerStatement.TransDate between '" & FormatDate(selectedFromDate) & "' and '" & FormatDate(CDate(selectedToDate)) & "') and Entitytype.AccountType_DPA_=4" & _
			 " GROUP BY dbo.OwnerStatement.Owner_DPA_,owner.OwnerFName, dbo.Owner.OwnerOpeningBal having (SUM(ISNULL(dbo.OwnerStatement.Credit - dbo.OwnerStatement.Debit, 0)) + dbo.Owner.OwnerOpeningBal)<>0"
	
	sqlStr =sqlStr & " Union all SELECT Top 100 Percent  SUM(ISNULL(case(IsOpeningBalance) when 1 then Balance else Credit - Debit end , 0)) AS CurrentBal, Client_DPA_ as AccountCode,Entity.EntityName as AccountName,Entity.EntityName as AccountType FROM  dbo.CDSControlStatement inner join Entity on CDSControlStatement.Client_DPA_=Entity.Entity_DPA_  inner join Entitytype on Entity.EntityType_DPA_=Entitytype.Entitytype_DPA_" & _
			 " WHERE (TransDate between '" & FormatDate(selectedFromDate) & "' and '" & FormatDate(CDate(selectedToDate)) & "') and EntityType.AccountType_DPA_=4 GROUP BY Client_DPA_,Entity.EntityName having (SUM(ISNULL(Credit - Debit, 0)))<>0"				
			
	sqlStr =sqlStr & " Union all SELECT Top 100 Percent CASE (DB_BankAccountStatement.IsOpeningBalance) WHEN 0 THEN SUM(ISNULL(dbo.DB_BankAccountStatement.Credit - dbo.DB_BankAccountStatement.Debit, 0))" & _ 
                 " ELSE 0 END + dbo.Account.AccountOpeningBal AS CurrentBal, Account.AccountCode as AccountCode,Account.AccountName,'Nominal' as AccountType" & _
			 " FROM  dbo.DB_BankAccountStatement INNER JOIN dbo.Account ON dbo.DB_BankAccountStatement.BankAccount_DPA_ = dbo.Account.Account_DPA_" & _
			 " WHERE  (Account.AccountTypeLevel1 = 4) and (dbo.DB_BankAccountStatement.TransDate between '" & FormatDate(selectedFromDate) & "' and '" & FormatDate(CDate(selectedToDate)) & "')" & _			
			 " GROUP BY Account.AccountCode,Account.AccountName,DB_BankAccountStatement.IsOpeningBalance, dbo.Account.AccountOpeningBal having CASE (DB_BankAccountStatement.IsOpeningBalance) WHEN 0 THEN SUM(ISNULL(dbo.DB_BankAccountStatement.Credit - dbo.DB_BankAccountStatement.Debit, 0))" & _ 
                 " ELSE 0 END + dbo.Account.AccountOpeningBal <>0 Order by AccountType,CurrentBal Asc"  
		
	'Response.write(sqlStr)
	'Response.end

	Set Rs = Conn.Execute(sqlStr)		
	%>	
	<tr>
		<td style="border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" colspan="2"><b><font face="Arial Narrow" size="3">Current Assets</font></b></td>
		<td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" align="right"><b><font face="Arial Narrow" size="3">&nbsp;</font></b></td>
		<td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" align="right"><b><font face="Arial Narrow" size="3">&nbsp;</font></b></td>
		</tr>
			<%
		Debit=0
		Debit=0
		AccountTypeSubTotal=0
		First=0

    	Do Until Rs.EOF
		AccountType1=Rs("AccountType")
		
		if(Rs("CurrentBal")<0) then
		Debit=1
		Credit=0
		else
		Credit=1
		Debit=0
		end if		
			
		if(AccountType1<>AccountType2)  then
		if(First<>0) then
				%>
				<tr>
					<td>&nbsp;</td>
					<td align="Right"><b><%=Ucase(AccountType2)%></b>&nbsp;<b>Sub Total</b></td>
					<td align="right"><%=FormatNum(AccountTypeSubTotal)%>
					</td>
				</tr>

				<%
				AccountTypeSubTotal=0
				First=0
			end if
		%>
		<tr>
	    	<td><b><%=Ucase(AccountType1)%></b></td>
	    	<td align="Right">&nbsp;</td>
	    	<td align="right">&nbsp;
	    	</td>
		</tr>

		<%
		First=1
		end if

		if(Debit<>Debit1) then
			if(Debit=1) then
			%>
		<tr>
	    	<td><b>Debits</b></td>
	    	<td align="Right">&nbsp;</td>
	    	<td align="right">&nbsp;
	    	</td>
		</tr>

		<%		
			else
			%>
		<tr>
	    	<td><b>Credits</b></td>
	    	<td align="Right">&nbsp;</td>
	    	<td align="right">&nbsp;
	    	</td>
		</tr>

		<%	
			end if
		end if

		%>
		<tr>
		  <td><%= Rs.Fields("AccountCode").Value %>
		  </td>
		  <td><%= mid(Rs.Fields("AccountName").Value,1,30) %>
		  </td>
		  <td align="right"><%=FormatNum(Abs(Rs.Fields("CurrentBal").Value))%>
		  </td>
		</tr>
		
		<%
		Subtotal=Subtotal+Rs.Fields("CurrentBal").Value
		AccountTypeSubTotal=AccountTypeSubTotal + Rs.Fields("CurrentBal").Value
		AccountType2=AccountType1
		Debit1=Debit
		Rs.MoveNext
		
	Loop
	if(First<>0) then
				%>
				<tr>
					<td>&nbsp;</td>
					<td align="Right"><b><%=Ucase(AccountType2)%></b>&nbsp;<b>Sub Total</b></td>
					<td align="right"><%=FormatNum(AccountTypeSubTotal)%>
					</td>
				</tr>

				<%
				AccountTypeSubTotal=0
				First=0
			end if
		%>
	<tr><td>&nbsp;</td></tr>
	<tr>
	    <td>&nbsp;</td>
	    <td align="Right"><b>Sub Total</b></td>
	    <td align="right"><%=FormatNum(Subtotal)%>
	    </td>
	</tr>

	<%
	Total=Total+Subtotal	
	Subtotal=0
	
	
	sqlStr=" SELECT SUM(ISNULL(Credit-Debit, 0)) AS CurrentBal, Entity_DPA_ as AccountCode,'Computers' as AccountName FROM  dbo.ComputerStatement" & _
	       " WHERE (TransDate between '" & FormatDate(selectedFromDate) & "' and '" & FormatDate(CDate(selectedToDate)) & "') GROUP BY Entity_DPA_"
				
			
	Set Rs = Conn.Execute(sqlStr)		
	%>	
	<tr>
		<td style="border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" colspan="2"><b><font face="Arial Narrow" size="3">Fixed Assets</font></b></td>
		<td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" align="right"><b><font face="Arial Narrow" size="3">&nbsp;</font></b></td>
		<td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" align="right"><b><font face="Arial Narrow" size="3">&nbsp;</font></b></td>
		</tr>
			<%
    Do Until Rs.EOF 
		%>
		<tr>
		  <td><%= Rs.Fields("AccountCode").Value %>
		  </td>
		  <td><%= Mid(Rs.Fields("AccountName").Value,1,30) %>
		  </td>
		  <td align="right"><%=FormatNum(Abs(Rs.Fields("CurrentBal").Value))%>
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
	AccountType1=""
	AccountType2=""

	Total=Total+Subtotal	
	Subtotal=0
	
	sqlStr = "SELECT Top 100 Percent SUM(ISNULL(dbo.StatementList1.Credit - dbo.StatementList1.Debit, 0)) + dbo.Client.ClientOpeningBal AS CurrentBal, dbo.StatementList1.Client_DPA_ as AccountCode,Client.ClientName as AccountName,'Clients' as AccountType" & _
		 " FROM  dbo.StatementList1 INNER JOIN dbo.Client ON dbo.StatementList1.Client_DPA_ = dbo.Client.Client_DPA_ inner join EntityType on Client.Entitytype_DPA_=EntityType.EntityType_DPA_" & _
                 " WHERE (dbo.Client.Deleted = 0) AND (dbo.StatementList1.TransDate between '" & FormatDate(selectedFromDate) & "' and '" & FormatDate(CDate(selectedToDate)) & "') and EntityType.AccountType_DPA_=5" & _
		 "  GROUP BY dbo.StatementList1.Client_DPA_,Client.ClientName, dbo.Client.ClientOpeningBal having (SUM(ISNULL(dbo.StatementList1.Credit - dbo.StatementList1.Debit, 0)) + dbo.Client.ClientOpeningBal)<>0"

	sqlStr =sqlStr & " Union all SELECT  Top 100 Percent SUM(ISNULL(dbo.AgentStatement1.Credit - dbo.AgentStatement1.Debit, 0)) + dbo.Agent.AgentOpeningBal AS CurrentBal, dbo.AgentStatement1.Agent_DPA_ as AccountCode,Agent.AgentName as AccountName,'Agents' as AccountType" & _
			 " FROM  dbo.AgentStatement1 INNER JOIN dbo.Agent ON dbo.AgentStatement1.Agent_DPA_ = dbo.Agent.Agent_DPA_ inner join Entitytype on Agent.Entitytype_DPA_=Entitytype.Entitytype_DPA_ " & _
			 " WHERE (dbo.Agent.Deleted = 0) AND (dbo.AgentStatement1.TransDate between '" & FormatDate(selectedFromDate) & "' and '" & FormatDate(CDate(selectedToDate)) & "') and Entitytype.AccountType_DPA_=5" & _
			 "  GROUP BY dbo.AgentStatement1.Agent_DPA_,Agent.AgentName, dbo.Agent.AgentOpeningBal having (SUM(ISNULL(dbo.AgentStatement1.Credit - dbo.AgentStatement1.Debit, 0)) + dbo.Agent.AgentOpeningBal)<>0"
		
	sqlStr =sqlStr & " Union all SELECT Top 100 Percent SUM(ISNULL(dbo.BrokerStatement1.Credit - dbo.BrokerStatement1.Debit, 0)) + dbo.Broker.BrokerOpeningBal AS CurrentBal, dbo.BrokerStatement1.Broker_DPA_ as AccountCode,Broker.BrokerName as AccountName,'Brokers' as AccountType" & _
			 " FROM  dbo.BrokerStatement1 INNER JOIN dbo.Broker ON dbo.BrokerStatement1.Broker_DPA_ = dbo.Broker.Broker_DPA_ inner join EntityType on Broker.EntityType_DPA_=EntityType.EntityType_DPA_" & _
			 " WHERE (dbo.BrokerStatement1.TransDate between '" & FormatDate(selectedFromDate) & "' and '" & FormatDate(CDate(selectedToDate)) & "') and EntityType.AccountType_DPA_=5" & _
			 "  GROUP BY dbo.BrokerStatement1.Broker_DPA_,Broker.BrokerName, dbo.Broker.BrokerOpeningBal having (SUM(ISNULL(dbo.BrokerStatement1.Credit - dbo.BrokerStatement1.Debit, 0)) + dbo.Broker.BrokerOpeningBal)<>0"
	
	sqlStr =sqlStr & " Union all SELECT Top 100 Percent SUM(ISNULL(Balance, 0)) AS CurrentBal, Entity_DPA_ as AccountCode,EntityName as AccountName,'Levies' as AccountType FROM  dbo.CommissionStatement" & _
			 " WHERE (TransDate between '" & FormatDate(selectedFromDate) & "' and '" & FormatDate(CDate(selectedToDate)) & "')  GROUP BY Entity_DPA_,EntityName having (SUM(ISNULL(Credit - Debit, 0)))<>0"		
	
	sqlStr =sqlStr & " Union all SELECT Top 100 Percent SUM(ISNULL(dbo.OwnerStatement.Credit - dbo.OwnerStatement.Debit, 0)) + dbo.Owner.OwnerOpeningBal AS CurrentBal, dbo.OwnerStatement.Owner_DPA_ as AccountCode,Owner.OwnerFName as AccountName,'Account Managers' as AccountType" & _
			 " FROM  dbo.OwnerStatement INNER JOIN dbo.Owner ON dbo.OwnerStatement.Owner_DPA_ = dbo.Owner.Owner_DPA_ inner join Entitytype on Owner.EntityType_DPA_=EntityType.EntityType_DPA_" & _
			 " WHERE (dbo.OwnerStatement.TransDate between '" & FormatDate(selectedFromDate) & "' and '" & FormatDate(CDate(selectedToDate)) & "') and Entitytype.AccountType_DPA_=5" & _
			 " GROUP BY dbo.OwnerStatement.Owner_DPA_,owner.OwnerFName, dbo.Owner.OwnerOpeningBal having (SUM(ISNULL(dbo.OwnerStatement.Credit - dbo.OwnerStatement.Debit, 0)) + dbo.Owner.OwnerOpeningBal)<>0"
	
	sqlStr =sqlStr & " Union all SELECT Top 100 Percent SUM(ISNULL(case(IsOpeningBalance) when 1 then Balance else Credit - Debit end , 0)) AS CurrentBal, Client_DPA_ as AccountCode,Entity.EntityName as AccountName,Entity.EntityName as AccountType FROM  dbo.CDSControlStatement inner join Entity on CDSControlStatement.Client_DPA_=Entity.Entity_DPA_ inner join EntityType on Entity.Entitytype_DPA_=Entitytype.Entitytype_DPA_" & _
			 " WHERE (TransDate between '" & FormatDate(selectedFromDate) & "' and '" & FormatDate(CDate(selectedToDate)) & "') and Entitytype.AccountTYpe_DPA_=5 GROUP BY Client_DPA_,Entity.EntityName having (SUM(ISNULL(Credit - Debit, 0)))<>0"				
			
	sqlStr =sqlStr & " Union all SELECT Top 100 Percent CASE (DB_BankAccountStatement.IsOpeningBalance) WHEN 0 THEN SUM(ISNULL(dbo.DB_BankAccountStatement.Credit - dbo.DB_BankAccountStatement.Debit, 0))" & _ 
                 " ELSE 0 END + dbo.Account.AccountOpeningBal AS CurrentBal, Account.AccountCode as AccountCode,Account.AccountName,'Nominal' as AccountType" & _
			 " FROM  dbo.DB_BankAccountStatement INNER JOIN dbo.Account ON dbo.DB_BankAccountStatement.BankAccount_DPA_ = dbo.Account.Account_DPA_" & _
			 " WHERE  (Account.AccountTypeLevel1 = 5) and (dbo.DB_BankAccountStatement.TransDate between '" & FormatDate(selectedFromDate) & "' and '" & FormatDate(CDate(selectedToDate)) & "')" & _			
			 " GROUP BY Account.AccountCode,Account.AccountName,DB_BankAccountStatement.IsOpeningBalance, dbo.Account.AccountOpeningBal having (CASE (DB_BankAccountStatement.IsOpeningBalance) WHEN 0 THEN SUM(ISNULL(dbo.DB_BankAccountStatement.Credit - dbo.DB_BankAccountStatement.Debit, 0))" & _ 
                 " ELSE 0 END + dbo.Account.AccountOpeningBal) <>0 Order by AccountType,CurrentBal Asc"  
			 
	Set Rs = Conn.Execute(sqlStr)		
	%>	
	<tr>
		<td style="border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" colspan="2"><b><font face="Arial Narrow" size="3">Current Liabilities</font></b></td>
		<td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" align="right"><b><font face="Arial Narrow" size="3">&nbsp;</font></b></td>
		<td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" align="right"><b><font face="Arial Narrow" size="3">&nbsp;</font></b></td>
		</tr>
			<%
			Debit=0
			Credit=0
			Debit1=0
			Debit2=0
			PreviousDebit=0
			PreviousCredit=0
			AccountTypeSubTotal=0
    Do Until Rs.EOF
		AccountType1=Rs("AccountType")
		
		if(AccountType1<>AccountType2)  then
			if(First<>0) then
				%>
				<tr>
					<td>&nbsp;</td>
					<td align="Right"><b><%=Ucase(AccountType2)%></b>&nbsp;<b>Sub Total</b></td>
					<td align="right"><%=FormatNum(AccountTypeSubTotal)%>
					</td>
				</tr>

				<%
				AccountTypeSubTotal=0
				First=0
			end if		
		
		Debit=0
		Credit=0
		PreviousDebit=0
		PreviousCredit=0
		%>
		<tr>
	    	<td><b><%=Ucase(AccountType1)%></b></td>
	    	<td align="Right">&nbsp;</td>
	    	<td align="right">&nbsp;
	    	</td>
		</tr>

		<%
		First=1
		end if
		
		if(Cdbl(Rs("CurrentBal"))<0) then
		Debit=1
		Credit=0
		else
		Credit=1
		Debit=0
		end if		
		
		'if(Rs.Fields("AccountCode").Value=100630) then
		'Response.write(Debit2)
		'Response.write(Debit1)
		'Response.end
		'end if

		if((Cint(Debit)<>Cint(PreviousDebit)) or (Cint(Credit)<>Cint(PreviousCredit))) then		
			if(Cint(Debit)=1) then
			%>
		<tr>
	    	<td><b>Debits</b></td>
	    	<td align="Right">&nbsp;</td>
	    	<td align="right">&nbsp;
	    	</td>
		</tr>

		<%		
			else
			%>
		<tr>
	    	<td><b>Credits</b></td>
	    	<td align="Right">&nbsp;</td>
	    	<td align="right">&nbsp;
	    	</td>
		</tr>

		<%	
			end if
		end if

		%>
		
		<tr>
		  <td><%= Rs.Fields("AccountCode").Value %>
		  </td>
		  <td><%= Mid(Rs.Fields("AccountName").Value,1,30) %>
		  </td>
		  <td align="right"><%=FormatNum((Abs(Rs.Fields("CurrentBal").Value)))%>
		  </td>
		</tr>
		
		<%
		Subtotal=Subtotal+(Rs.Fields("CurrentBal").Value)
		AccountTypeSubTotal=AccountTypeSubTotal + Rs.Fields("CurrentBal").Value
		AccountType2=AccountType1		

		PreviousDebit=Debit
		PreviousCredit=Credit
		Rs.MoveNext
		
	Loop
	if(First<>0) then
				%>
				<tr>
					<td>&nbsp;</td>
					<td align="Right"><b><%=Ucase(AccountType2)%></b>&nbsp;<b>Sub Total</b></td>
					<td align="right"><%=FormatNum(AccountTypeSubTotal)%>
					</td>
				</tr>

				<%
				AccountTypeSubTotal=0
				First=0
			end if
		%>
	<tr><td>&nbsp;</td></tr>
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
		<td style="border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" colspan="2"><b><font face="Arial Narrow" size="3">Long Term Liabilities</font></b></td>
		<td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" align="right"><b><font face="Arial Narrow" size="3">&nbsp;</font></b></td>
		<td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" align="right"><b><font face="Arial Narrow" size="3">&nbsp;</font></b></td>
		</tr>
		<tr>
		  <td>&nbsp;</td>
		  <td>&nbsp;</td>
		  <td>&nbsp;</td>
		</tr>
		
			<%

	Profit=0
	Gross=0

	sqlStr="SELECT SUM(ISNULL(Credit-Debit, 0)) AS CurrentBal, Entity_DPA_ as AccountCode,'Broker Commission' as AccountName FROM  dbo.BrokerCommissionStatement" & _
	       " WHERE (TransDate between '" & FormatDate(selectedFromDate) & "' and '" & FormatDate(CDate(selectedToDate)) & "')  GROUP BY Entity_DPA_"
	
	sqlStr=sqlStr & " Union All " & _
			"SELECT SUM(DB_BankAccountStatement.Credit - DB_BankAccountStatement.Debit) AS CurrentBal, Account.AccountCode as AccountCode,Account.AccountName" & _
		 	" FROM  dbo.DB_BankAccountStatement INNER JOIN dbo.Account ON dbo.DB_BankAccountStatement.BankAccount_DPA_ = dbo.Account.Account_DPA_" & _
		 	" WHERE  (Account.AccountTypeLevel1 = 1) and (dbo.DB_BankAccountStatement.TransDate between '" & FormatDate(selectedFromDate) & "' and '" & FormatDate(CDate(selectedToDate)) & "')" & _			
		 	" GROUP BY Account.AccountCode,Account.AccountName,DB_BankAccountStatement.IsOpeningBalance, dbo.Account.AccountOpeningBal"  

	Set Rs = Conn.Execute(sqlStr) 
	%>
		<tr>
		<td style="border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"><b><font face="Arial Narrow" size="3">Profit&nbsp;&&nbsp;Loss</font></b></td>
		<td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" align="right"><b><font face="Arial Narrow" size="3">&nbsp;</font></b></td>
		<td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" align="right"><b><font face="Arial Narrow" size="3">&nbsp;</font></b></td>
		</tr>
		
	<%
    Do Until Rs.EOF 
		Gross=Gross + Abs(Rs.Fields("CurrentBal").Value)
		Rs.MoveNext		
	Loop	
	
	sqlStr = "SELECT SUM(DB_BankAccountStatement.Credit - DB_BankAccountStatement.Debit) AS CurrentBal, Account.AccountCode as AccountCode,Account.AccountName" & _
		 " FROM  dbo.DB_BankAccountStatement INNER JOIN dbo.Account ON dbo.DB_BankAccountStatement.BankAccount_DPA_ = dbo.Account.Account_DPA_" & _
		 " WHERE  (Account.AccountTypeLevel1 = 2) and (dbo.DB_BankAccountStatement.TransDate between '" & FormatDate(selectedFromDate) & "' and '" & FormatDate(CDate(selectedToDate)) & "')" & _			
		 " GROUP BY Account.AccountCode,Account.AccountName,DB_BankAccountStatement.IsOpeningBalance, dbo.Account.AccountOpeningBal"  

	Set Rs = Conn.Execute(sqlStr)		
	
    	Do Until Rs.EOF 
		
	Gross=Gross+Rs.Fields("CurrentBal").Value
		Rs.MoveNext
		
	Loop
	
	sqlStr="SELECT SUM(ISNULL(Balance, 0)) AS CurrentBal FROM  dbo.BrokerCommissionStatement" & _
	       " WHERE (TransDate between '" & FormatDate(selectedFromDate) & "' and '" & FormatDate(CDate(selectedToDate)) & "') and IsOpeningBalance=1"
	
	sqlStr= sqlStr & " Union All	SELECT SUM(DB_BankAccountStatement.CreditBal - DB_BankAccountStatement.Debit) AS CurrentBal" & _
		   " FROM  dbo.DB_BankAccountStatement INNER JOIN dbo.Account ON dbo.DB_BankAccountStatement.BankAccount_DPA_ = dbo.Account.Account_DPA_" & _
		   " WHERE  (Account.AccountTypeLevel1 = 2 or Account.AccountTypeLevel1 = 1) and (dbo.DB_BankAccountStatement.TransDate between '" & FormatDate(selectedFromDate) & "' and '" & FormatDate(CDate(selectedToDate)) & "')" & _			
		   " and IsOpeningBalance=1"  
	
	Set Rs = Conn.Execute(sqlStr)		
	
    	Do Until Rs.EOF 
		
		Gross=Gross+Rs.Fields("CurrentBal").Value
		Rs.MoveNext	
		Loop
	 %>		
		
		<tr>
		  <td>&nbsp;</td>
		  <% if(Gross<0) then %>
		  <td><b>LOSS</b></td>
		  <% else %>
		  <td><b>PROFIT</b></td>
		  <% end if%>
		  <td align="right"><%=FormatNum(Gross)%>
		  </td>
		</tr>	
	
	<% Total=Total+Gross %>	
		<tr>
		  <td>&nbsp;</td>
		  <td>&nbsp;</td>
		  <td>&nbsp;</td>
		</tr>
		
		<tr>
		  <td>&nbsp;</td>
	    	  <td align="Right"><b>Net Total</b></td>	    
		  <td align="right"><%=FormatNum(Total)%></td>
		</tr>
		
	</table>
	
	<%
Set Rs = Nothing
Set Conn = Nothing%>

</body>

</html>
