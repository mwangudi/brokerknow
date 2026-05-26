<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Annual Budget</title>

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
			margin-left: 2cm;
			margin-right: 5cm;
			margin-top: 1cm;    
			margin-bottom: 2cm;
			writing-mode: tb-rl;
			height: 80%;
			margin: 10% 0%;						
			br.newpage{
				page-break-before:always;
			}		
		}		 
		
	</style>
<Script language='Javascript'>
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "transFromDate","cmdDate","<%= FormatDate(Date) %>",1);
		var cal1=new ctlSpiffyCalendarBox("cal1", "frmMain", "txtToDate","cmdDate","<%= FormatDate(Date) %>",1);

</Script>
</head>

<body Class="Reports">

<!--#include file="../libroutines.asp"-->

<%


genReport = Request.Form("genReport")
selectedFromDate = Request.Form("transFromDate")

If genReport = "" Then%>
	<Script Language="JavaScript">
		document.body.className = 'dialog';
		
		function validateForm(frm){						
			frm.target = '_self';			
			frm.submit();
		}	
	
	</Script>
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="AnnualBudget.asp" Name="frmMain" id="frmMain">	
			<table>			
			<tr>
				<td>Select date:</td>
				<td>
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>				
			</tr>
			<tr>			
				<td colspan=2>
				<input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.getElementById('frmMain'))" Value=" Generate... ">&nbsp;&nbsp; <input type="Button" class="Buttons" Value=" Close " OnClick="JavaScript: window.parent.self.close();"></td>
			</tr>			
			<input type="hidden" name="genReport" value="1">
		</table>
		
	</form>
	
	<%Set rs = Nothing
	Set Conn = Nothing
	Response.End
End If

%>


<% DrawPageFunctions True, True, True


Set Conn = GetActiveConnection("KBroker")

 %>

<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
     <tr>
		<td nowrap><b><font face="Arial Narrow" size="4">ANNUAL BUDGET</font></b></td>
		<td nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>
	</tr>	
       <tr>
		  <td COLSPAN=2><font face="Arial" size="2">&nbsp;</font></td>
	</tr>
	<tr>
		  <td COLSPAN=2><font face="Arial" size="2"><b>For the Year &nbsp;<%=ucase(Year((CDate(SelectedFromDate))))%></b></font></td>
	</tr>
</table>			



    <table border="0" width="400" cellPadding="2" cellSpacing=0>
    <tr bgColor="#000000">
		<td><b><font color="#FFFFFF">Code</font></b></td>
		<td><b><font color="#FFFFFF">Account</font></b></td>
		<td align="right"><b><font color="#FFFFFF">Actual</font></b></td>
		<td align="right"><b><font color="#FFFFFF">Budget</font></b></td>
		<td align="right"><b><font color="#FFFFFF">Variance</font></b></td>
	</tr>
	
	<%
        sqlStr = "SELECT * FROM [BudgetReport] ORDER BY AccountTYpeParent_DPA_,AccountType_DPA_"
	
	'Response.write(sqlStr)
	'Response.end        
	
	set cloneRs=Conn.Execute(sqlStr)
	
	AccountType1=""
	AccountType2=""

	Actual=0
	Budget=0
	Variance=0
	Quarter=0
	
	FirstDate=CDate("01/01/" & Year(CDate(SelectedFromDate)))
	SecondDate=CDate("31/12/" & Year(CDate(SelectedFromDate)))

	Do while cloneRs.eof=false
	AccountType1=CloneRs("AccountTYpeParent_DPA_")
	
		if(Month(CDate(SelectedFromDate))<=3) then
			Quarter=cloneRs("Quarter1")			
		end if
		if(Month(CDate(SelectedFromDate))>3 and Month(CDate(SelectedFromDate))<=6) then
			Quarter=cloneRs("Quarter2")

			FirstDate=CDate("01/04/" & Year(CDate(SelectedFromDate)))
			SecondDate=CDate("30/06/" & Year(CDate(SelectedFromDate)))
		end if
		if(Month(CDate(SelectedFromDate))>6 and Month(CDate(SelectedFromDate))<=9) then
			Quarter=cloneRs("Quarter3")
			FirstDate=CDate("01/07/" & Year(CDate(SelectedFromDate)))
			SecondDate=CDate("30/09/" & Year(CDate(SelectedFromDate)))
		end if
		if(Month(CDate(SelectedFromDate))>9 and Month(CDate(SelectedFromDate))<=12) then
			FirstDate=CDate("01/10/" & Year(CDate(SelectedFromDate)))
			SecondDate=CDate("31/12/" & Year(CDate(SelectedFromDate)))
		end if
	
	'if(AccountType1=1 or AccountType1=2) then
	'	if(Quarter=0) then
	'	Budget=0
	'	else
	'	Budget=Quarter/3
	'	end if
	'else
	Budget=Quarter
	'end if
	
	if(AccountType1=1 or AccountType1=2) then
	sqlStr="SELECT isnull(SUM(DB_BankAccountStatement.Credit - DB_BankAccountStatement.Debit),0) AS CurrentBal FROM         DB_BankAccountStatement INNER JOIN   Account ON DB_BankAccountStatement.BankAccount_DPA_ = Account.Account_DPA_ " & _
		 	
		 	" WHERE  (Account.AccountTypeLevel2 = " & cloneRs("AccountType_DPA_") & ") and (dbo.DB_BankAccountStatement.TransDate  between '" & CDate(FirstDate) & "' and '" & CDate(SecondDate) & "')" 			
	else
	FirstDate=CDate("01/01/" & Year(CDate(SelectedFromDate)))

	sqlStr="SELECT isnull(SUM(Balance),0) AS CurrentBal	 FROM         DB_BankAccountStatement INNER JOIN   Account ON DB_BankAccountStatement.BankAccount_DPA_ = Account.Account_DPA_ " & _
		 	
		 	" WHERE  (Account.AccountTypeLevel2 = " & cloneRs("AccountType_DPA_") & ") and (dbo.DB_BankAccountStatement.TransDate  between '" & CDate(FirstDate) & "' and '" & CDate(SecondDate) & "')" 	
	end if
	'Response.write(sqlStr)
	'Response.end
	
	Set Rs=Conn.Execute(sqlStr)

	if not (Rs.Eof and Rs.Bof) then
		Actual=Rs("CurrentBal")
		else
		Actual=0
	end if	
	
		Variance=Actual-Budget

	        if (AccountType2<>AccountType1) then
		 	%><tr><td colspan="5"><b><%=CloneRs("AccountTypeParentName")%></b></td></tr><%
			end if			
		%>
		<tr>
		<td><%=cloneRs("AccountType_DPA_")%></td>
		<td><%=cloneRs("AccountTypeName")%></td>
		<td align="right"><%=FormatNum(Actual)%></td>
		<td align="right"><%=FormatNum(Budget)%></td>
		<td align="right"><%=FormatNum(Variance)%></td>
		</tr>
	<%		
	Actuals=Actuals+Actual
	Budgets=Budgets+Budget
	Variances=Variances+Variance

	 AccountType2=AccountType1
	 cloneRs.MoveNext
	Loop	  
	%>
		<tr>
		<td colspan="2" align="right"><b>Totals</b></td>
		<td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b><%=FormatNum(Actuals)%></font></b></td>
		<td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b><%=FormatNum(Budgets)%></b></font></td>      
		<td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b><%=FormatNum(Variances)%></b></font></td>     		
		</tr>	  	
  </table>
  
	 <%	 
	 Set Rs = Nothing
	 Set Conn = Nothing
     %>	
</body>

</html>