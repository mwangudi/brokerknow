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

'Set Rs = CreateObject("ADODB.Recordset")	

sqlStr="SELECT  TBal.* INTO #tableA FROM ( " & _
 " SELECT      " & _
 " 	'******' + CONVERT(VARCHAR(500),dbo.FullEntityTypeList.EntityType_DPA_) AS [Account Code],  	 " & _
 " 	dbo.FullEntityTypeList.EntityTypeName AS AccountName,  " & _
 " 	CASE  " & _
 " 		WHEN SUM(dbo.EntityTransactionList.Balance) < 0 THEN 0 - SUM(dbo.EntityTransactionList.Balance)  " & _
 " 		ELSE 0 END AS Debit, " & _
 " 	CASE  " & _
 " 		WHEN SUM(dbo.EntityTransactionList.Balance) >= 0 THEN SUM(dbo.EntityTransactionList.Balance)  " & _
 " 		ELSE 0 END AS Credit,Cast(Floor(Cast(Min(EntityTransactionList.TransDate) as Float)) as DateTime) as TransDate	              " & _
 " FROM         dbo.Entity INNER JOIN " & _
 "                       dbo.FullEntityTypeList ON dbo.Entity.EntityType_DPA_ = dbo.FullEntityTypeList.EntityType_DPA_ INNER JOIN " & _
 "                       dbo.EntityTransactionList ON dbo.Entity.Entity_DPA_ = dbo.EntityTransactionList.Entity_DPA_ " & _
 " GROUP BY dbo.FullEntityTypeList.EntityTypeAccountType, dbo.FullEntityTypeList.EntityTypeName, dbo.FullEntityTypeList.EntityType_DPA_ " & _
 " ) TBal " & _
 "  " & _
 "  " & _
 " INSERT INTO #tableA " & _
 " SELECT  " & _
 " 	'******2' AS [Account Code],  " & _
 " 	(SELECT EntityTypeName FROM dbo.FullEntityTypeList WHERE EntityType_DPA_ = 2) AS AccountName,Case When Sum(AgentStatement.Balance) <0 then 0-Sum(AgentStatement.Balance) else 0 end as Debit,Case When Sum(AgentStatement.Balance)>=0 then Sum(AgentStatement.Balance) else 0 end as Credit,TransDate from AgentStatement Group by TransDate" & _ 
 " INSERT INTO #tableA " & _
 " SELECT  " & _
 " 	'******3' AS [Account Code],  " & _
 " 	(SELECT EntityTypeName FROM dbo.FullEntityTypeList WHERE EntityType_DPA_ = 3) AS AccountName, 		 " & _
 " 	CASE  " & _
 " 		WHEN SUM(Credit-Debit) < 0 THEN 0 - SUM(Credit-Debit)  " & _
 " 		ELSE 0 END AS Debit, " & _
 " 	CASE  " & _
 " 		WHEN SUM(Credit-Debit) >= 0 THEN SUM(Credit-Debit)  " & _
 " 		ELSE 0 END AS Credit,Cast(Floor(Cast(TransDate as Float)) as Datetime) as TransDate	 " & _
 " 	 " & _
 " 	FROM BrokerStatement Group by Cast(Floor(Cast(TransDate as Float)) as Datetime) " & _ 
 " INSERT INTO #tableA " & _
 "  " & _
 " SELECT 	 " & _
 " 	'******1' AS [Account Code],  " & _
 " 	(SELECT EntityTypeName FROM dbo.FullEntityTypeList WHERE EntityType_DPA_ = 1) AS AccountName,CASE  " & _
 " 		WHEN SUM(StatementList.Balance) < 0 THEN 0 - SUM(StatementList.Balance)  " & _
 " 		ELSE 0 END AS Debit, " & _
 " 	CASE  " & _
 " 		WHEN SUM(StatementList.Balance) >= 0 THEN SUM(StatementList.Balance)  " & _
 " 		ELSE 0 END AS Credit,TransDate from StatementList Group by TransDate " & _
 " 	" & _ 
 " INSERT INTO #tableA " & _
 " SELECT  " & _
 " 	'******7' AS [Account Code],  " & _
 " 	(SELECT EntityTypeName FROM dbo.FullEntityTypeList WHERE EntityType_DPA_ = 7) AS AccountName,  " & _
 " 	CASE  " & _
 " 		WHEN SUM(Credit-Debit) < 0 THEN 0 - SUM(Credit-Debit)  " & _
 " 		ELSE 0 END AS Debit, " & _
 " 	CASE  " & _
 " 		WHEN SUM(Credit-Debit) >= 0 THEN SUM(Credit-Debit)  " & _
 " 		ELSE 0 END AS Credit,Cast(Floor(Cast(TransDate as Float)) as DateTime) as TransDate " & _
 " FROM  OwnerStatement Group by Cast(Floor(Cast(TransDate as Float)) as DateTime) " & _
 " INSERT INTO #tableA " & _
 " SELECT  Distinct " & _
 " 	'------' + CONVERT(VARCHAR(500),Entity.Entity_DPA_) AS [Account Code],  		 " & _
 " 	Entity.EntityName AS AccountName, " & _
 " 	CASE  " & _
 " 		WHEN SUM(Balance) < 0 THEN 0 - SUM(Balance)  " & _
 " 		ELSE 0 END AS Debit, " & _
 " 	CASE  " & _
 " 		WHEN SUM(Balance) >= 0 THEN SUM(Balance)  " & _
 " 		ELSE 0 END AS Credit,Cast(Floor(Cast(TransDate as Float)) as DateTime) as TransDate		 " & _
 " FROM         BrokerCommissionStatement inner join Entity on BrokerCommissionStatement.Entity_DPA_=Entity.Entity_DPA_ Group by Entity.Entity_DPA_,Entity.EntityName,Cast(Floor(Cast(TransDate as Float)) as DateTime) INSERT INTO #tableA SELECT Distinct '------' + CONVERT(VARCHAR(500),Entity.Entity_DPA_) AS [Account code],Entity.EntityName AS AccountName, CASE WHEN SUM(Balance) < 0 THEN 0- SUM(Balance) ELSE 0 END as Debit, CASE WHEN SUM(Balance) >= 0 THEN SUM(Balance) ELSE 0 END AS Credit,TransDate From CommissionStatement inner join Entity on CommissionStatement.Entity_DPA_ =Entity.Entity_DPA_ Group By Entity.Entity_DPA_,Entity.EntityName,TransDate INSERT INTO #tableA SELECT '------' + CONVERT(VARCHAR(500),Entity.Entity_DPA_) AS [Account Code],Entity.EntityName AS AccountName,CASE WHEN SUM(Balance) < 0 THEN 0 - SUM(Balance) else 0 END AS Debit, CASE WHEN SUM(Balance) >= 0 THEN SUM(Balance) ELSE 0 End AS Credit,TransDate From CDSControlStatement inner join Entity on CDSControlStatement.Client_DPA_=Entity.Entity_DPA_ Group by Entity_DPA_,Entity.EntityName,TransDate " & _  
 " INSERT INTO #tableA " & _
 " SELECT      " & _
 " 	CONVERT(VARCHAR(500),dbo.AccountList.Account_DPA_) AS [Account Code],   " & _
 " 	dbo.AccountList.AccountName,  " & _
 " 	CASE  " & _
 " 		WHEN SUM(dbo.NominalTransactionList.Balance) < 0 THEN 0 - SUM(dbo.NominalTransactionList.Balance)  " & _
 " 		ELSE 0 END AS Debit, " & _
 " 	CASE  " & _
 " 		WHEN SUM(dbo.NominalTransactionList.Balance) >= 0 THEN SUM(dbo.NominalTransactionList.Balance)  " & _
 " 		ELSE 0 END AS Credit,Cast(Floor(Cast(NominalTransactionList.TransDate as Float)) as Datetime) as TransDate		 " & _
 " FROM         dbo.AccountList INNER JOIN " & _
 "                       dbo.NominalTransactionList ON dbo.AccountList.Account_DPA_ = dbo.NominalTransactionList.Account_DPA_ " & _
 " GROUP BY dbo.AccountList.AccountTypeLevel1,dbo.AccountList.AccountName, dbo.AccountList.Account_DPA_,Cast(Floor(Cast(NominalTransactionList.TransDate as Float)) as Datetime) " & _
 "  " & _
 "  " & _
 " INSERT INTO #tableA " & _
 " SELECT     " & _
 " 	CONVERT(VARCHAR(500),dbo.AccountList.AccountCode) AS [Account Code],  	 " & _
 " 	dbo.AccountList.AccountName,  " & _
 " 	CASE  " & _
 " 		WHEN SUM(dbo.DB_BankAccountStatement.Balance) < 0 THEN 0 - SUM(DB_BankAccountStatement.Balance)  " & _
 " 		ELSE 0 END AS Debit, " & _
 " 	CASE  " & _
 " 		WHEN SUM(dbo.DB_BankAccountStatement.Balance) >= 0 THEN SUM(dbo.DB_BankAccountStatement.Balance)  " & _
 " 		ELSE 0 END AS Credit,Cast(Floor(Cast(Db_BankAccountStatement.TransDate as Float)) as DateTime) as TransDate 			 " & _
 " FROM         dbo.DB_BankAccountStatement INNER JOIN " & _
 "                       dbo.AccountList ON dbo.DB_BankAccountStatement.BankAccount_DPA_ = dbo.AccountList.Account_DPA_ " & _
 " GROUP BY dbo.AccountList.AccountTypeLevel1,dbo.AccountList.AccountName, dbo.AccountList.AccountCode,Cast(Floor(Cast(Db_BankAccountStatement.TransDate as Float)) as DateTime) " & _
 "  " & _
 " SELECT  IDENTITY(int, 1,1)  AS SequenceID, #tableA.* " & _
 " 	 INTO #tableB FROM #tableA ORDER BY [Account Code] " & _
 "  " & _ 
 "  " & _
 " select T.[Account Code], T.AccountName,CASE WHEN Sum(T.Credit-T.Debit) <0  then 0-Sum(T.Credit-T.Debit) else 0 end as Debit,CASE WHEN Sum(T.Credit-T.Debit)>=0 then Sum(T.Credit-T.Debit) else 0 end as Credit " & _
 " from #tableB T  Where T.TransDate " & _
 " <= '" & FormatDate(selectedFromDate) & "'" & _
 " Group by T.[Account Code], T.AccountName " & _
 "   order by T.[Account Code] " 
     
'Response.write(sqlStr)
'Response.end

'Set Rs = Conn.Execute(sqlStr)

sqlStr = "Exec TrialBalanceProc '" & FormatDate(selectedFromDate) & "',' " & FormatDate(selectedToDate) & "' "

'Response.write(sqlStr)
'Response.end

Rs.CursorLocation = adUseClient	
Rs.Open SQLServerFormat((sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic

'Dim i 
	
'for i = 1 to 9
'	Set Rs =  Rs.NextRecordset 
'next

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