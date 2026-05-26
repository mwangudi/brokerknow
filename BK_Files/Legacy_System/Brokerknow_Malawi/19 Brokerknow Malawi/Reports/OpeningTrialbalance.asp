<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Accounts Statement</title>

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
	<form method="POST" action="OpeningTrialBalance.asp" Name="frmMain" id="frmMain">	
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

<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="450">
     <tr>
		<td nowrap><b><font face="Arial Narrow" size="4">OPENING BALANCES</font></b></td>
		<td nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>
	</tr>	
       <tr>
		  <td COLSPAN=2><font face="Arial" size="2">&nbsp;</font></td>
	</tr>
	<tr>
		  <td COLSPAN=2><font face="Arial" size="2"><b>AS OF&nbsp;<%=FormatDate(SelectedFromDate)%></b></font></td>
	</tr>
</table>			



    <table border="0" width="450" cellPadding="2" cellSpacing=0>
    <tr bgColor="#000000">
		<td><b><font color="#FFFFFF">Account&nbsp;Code</font></b></td>
		<td width="300"><b><font color="#FFFFFF">Account&nbsp;Name</font></b></td>
		<td align="right"><b><font color="#FFFFFF">Debit</font></b></td>
		<td align="right"><b><font color="#FFFFFF">Credit</font></b></td>
	</tr>
	
	<%
        sqlStr = "SELECT * FROM [FullEntityTypeList] ORDER BY EntityType_DPA_"
	
	'Response.write(sqlStr)
	'Response.end
        
	Credit=0
	Debit=0

	set rs=Conn.Execute(sqlStr)
	
	Do while rs.eof=false		
		Select Case(rs("EntityType_DPA_"))
      		case 1						        	
		sqlStr = "SELECT distinct dbo.Client.ClientOpeningBal AS CurrentBal, dbo.StatementList1.Client_DPA_ as AccountCode,Client.ClientName as AccountName" & _
				 " FROM  dbo.StatementList1 INNER JOIN dbo.Client ON dbo.StatementList1.Client_DPA_ = dbo.Client.Client_DPA_" & _
                                 " WHERE (dbo.Client.Deleted = 0) AND (dbo.StatementList1.TransDate <= '" & FormatDate(selectedFromDate) & "')" 
				 
		case 2
		sqlStr = "SELECT distinct dbo.Agent.AgentOpeningBal AS CurrentBal, dbo.AgentStatement1.Agent_DPA_ as AccountCode,Agent.AgentName as AccountName" & _
				 " FROM  dbo.AgentStatement1 INNER JOIN dbo.Agent ON dbo.AgentStatement1.Agent_DPA_ = dbo.Agent.Agent_DPA_" & _
				 " WHERE (dbo.Agent.Deleted = 0) AND (dbo.AgentStatement1.TransDate <= '" & FormatDate(selectedFromDate) & "')" 				 
		case 3
		
		sqlStr = "SELECT distinct dbo.Broker.BrokerOpeningBal AS CurrentBal, dbo.BrokerStatement1.Broker_DPA_ as AccountCode,Broker.BrokerName as AccountName" & _
				 " FROM  dbo.BrokerStatement1 INNER JOIN dbo.Broker ON dbo.BrokerStatement1.Broker_DPA_ = dbo.Broker.Broker_DPA_" & _
				 " WHERE (dbo.BrokerStatement1.TransDate <= '" & FormatDate(selectedFromDate) & "')" 
				 
		case 4
		sqlStr="SELECT SUM(ISNULL(Balance, 0)) AS CurrentBal, Entity_DPA_ as AccountCode,'Broker Commission' as AccountName FROM  dbo.BrokerCommissionStatement" & _
			   " WHERE (TransDate <= '" & FormatDate(selectedFromDate) & "') and IsOpeningBalance=1 GROUP BY Entity_DPA_"
		case 5
		sqlStr = "SELECT SUM(DB_BankAccountStatement.CreditBal - DB_BankAccountStatement.Debit) AS CurrentBal, Account.AccountCode as AccountCode,Account.AccountName" & _
			 " FROM  dbo.DB_BankAccountStatement INNER JOIN dbo.Account ON dbo.DB_BankAccountStatement.BankAccount_DPA_ = dbo.Account.Account_DPA_" & _
			 " WHERE  (dbo.DB_BankAccountStatement.TransDate <= '" & FormatDate(selectedFromDate) & "') and IsOpeningBalance=1 Group by AccountCode,AccountName" 						 
				 
		case 6
		sqlStr="SELECT SUM(ISNULL(Balance, 0)) AS CurrentBal, Entity_DPA_ as AccountCode,EntityName as AccountName FROM  dbo.CommissionStatement" & _
			   " WHERE (TransDate <= '" & FormatDate(selectedFromDate) & "') and IsOpeningBalance=1  GROUP BY Entity_DPA_,EntityName"		
		case 7
		sqlStr = "SELECT dbo.Owner.OwnerOpeningBal AS CurrentBal, dbo.OwnerStatement.Owner_DPA_ as AccountCode,Owner.OwnerFName as AccountName" & _
				 " FROM  dbo.OwnerStatement INNER JOIN dbo.Owner ON dbo.OwnerStatement.Owner_DPA_ = dbo.Owner.Owner_DPA_" & _
				 " WHERE (dbo.OwnerStatement.TransDate <= '" & FormatDate(selectedFromDate) & "')" 
				 
		case 8
		sqlStr="SELECT Balance AS CurrentBal, Client_DPA_ as AccountCode,'CDS Clearing Account' as AccountName FROM  dbo.CDSControlStatement" & _
			   " WHERE (TransDate <= '" & FormatDate(selectedFromDate) & "') and IsOpeningBalance=1"				
		
		case 11
		sqlStr="SELECT SUM(ISNULL(Credit - Debit, 0)) AS CurrentBal, Entity_DPA_ as AccountCode,'Computers' as AccountName FROM  dbo.ComputerStatement" & _
			   " WHERE (TransDate <= '" & FormatDate(selectedFromDate) & "') GROUP BY Entity_DPA_"
		
		end select 	    
		
		'Response.write(sqlStr)
		'Response.end
		
                if(rs("EntityType_DPA_")=2) then
				'Response.write(sqlStr)
				'Response.end
				end if
 
		Set cloneRs = Conn.Execute(sqlStr)
	        	'if not(cloneRs.eof and cloneRs.bof) then
		 	%><tr><td colspan="4"><b><%=rs("EntityTypeName")%></b></td></tr><%
			'end if
		Do Until cloneRs.EOF 
			if(cloneRs.Fields("CurrentBal").Value<>0) then
			%>
        		<tr>
        			<td><%=cloneRs.Fields("AccountCode").Value%></td>
				<td width="300"><%=Mid(cloneRs.Fields("AccountName").Value,1,30)%></td>
				<% if(cloneRs.Fields("CurrentBal").Value<0) then 
				Debit=Debit+Abs(cloneRs.Fields("CurrentBal").Value)
				%>
        			<td align="right"><%=FormatNum(Abs(cloneRs.Fields("CurrentBal").Value))%></td>
				<td>&nbsp;</td>
			 	<% else 
				Credit=Credit+cloneRs.Fields("CurrentBal").Value
				%>
				<td>&nbsp;</td>
        			<td align="right"><%=FormatNum(cloneRs.Fields("CurrentBal").Value)%></td>      				 
				<% end if %>
	        	</tr>
	 		<%
			end if 
	 	
	 cloneRs.MoveNext
	 Loop
	 
	 Rs.MoveNext
	Loop	  
	%>
		<tr>
		<td colspan="2" align="right"><b>Totals</b></td>
		<td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b><%=FormatNum(Debit)%></font></b></td>
		<td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b><%=FormatNum(Credit)%></b></font></td>      
		</tr>	  	
  </table>
  
	 <%	 
	 Set Rs = Nothing
	 Set Conn = Nothing
     %>	
</body>

</html>