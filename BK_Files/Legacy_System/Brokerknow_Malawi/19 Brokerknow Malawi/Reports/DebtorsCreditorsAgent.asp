<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Debtors and Creditors</title>
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
<!--#include file="../libroutinesTEST.asp"-->

<%

genReport = Request.Form("genReport")

selectedFromDate = Request.Form("transFromDate")
clientType = Request.Form("clientType")
Agent_DPA_ = Request.Form("cboAgent")

If genReport <> "1" Or Not IsDate(selectedFromDate) Then%>
	<Script Language="JavaScript">
		report_SetBodyClass();
		
		function validateForm(frm){			
			frm.target = '_self';			
			frm.submit();
		}
		
		
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "transFromDate","cmdDate","<%= FormatDate(Date) %>",1);
	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="DebtorsCreditorsAgent.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">
		<table>
		
			<tr>
				<td>Start of:</td>
				<td>
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>
			</tr>
			
			<tr>
				<td>Select Client Type</td>
				<td>
					<select id="clientType" name="clientType">
					<option value=''>ALL</option>
					<option value='1'>Custodian</option>
					<option value='0'>Non - Custodian</option>
					
					</select>	
				</td>
			</tr>
			
			<tr>
				<td>Agent: </td>
				<td><select name = 'cboAgent' id = 'cboAgent' size="1">
					<option selected value = ''></option>
					<%
					        Set conn = GetActiveConnection("KBroker")
					        
					        sqlStr = "SELECT Agent_DPA_,AgentName FROM Agent ORDER BY AgentName"
					        Set rs = conn.Execute(sqlStr)
					        If Not (rs.EOF Or rs.BOF) Then
					                rs.MoveFirst
					                Do Until rs.EOF%>
					                        <option value = '<%=rs.Fields("Agent_DPA_")%>'><%=rs.Fields("AgentName")  %></option>
					                        <%rs.MoveNext
					                Loop
					        End If
					%>

					    </select>
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

<% DrawPageFunctions True, True, True, True %>

<%
	Dim conn 
	Dim sqlStr
	Dim rs
	
	Set Rs = CreateObject("ADODB.Recordset")
	Set RsDebtors = CreateObject("ADODB.Recordset")
								        
	Rs.CursorLocation = adUseClient	
        
	'Evaluate Client Type 
	if trim(clientType) = "" then
		clientTypeSQL = ""
	elseif trim(clientType) =  1 then
		clientTypeSQL = " AND dbo.client.Iscustodian = 1 "
	else
		clientTypeSQL = " AND dbo.client.Iscustodian <> 1 "
	end if

	If Len(Agent_DPA_) > 0 Then
		DebtorsSQL = " HAVING (b.CurrentBal < 0) AND (Client.Agent_DPA_ = "& Agent_DPA_ &") ORDER BY dbo.Client.ClientName "
	Else
		DebtorsSQL = " HAVING (b.CurrentBal < 0) ORDER BY dbo.Client.ClientName "
	End If
	
	If Len(Agent_DPA_) > 0 Then
		CreditorsSQL =  " HAVING (b.CurrentBal > 0) AND (Client.Agent_DPA_ = "& Agent_DPA_ &") ORDER BY dbo.Client.ClientName "
	Else
		CreditorsSQL =  " HAVING (b.CurrentBal > 0) ORDER BY dbo.Client.ClientName "
	End If
	
	SQL = " SELECT DISTINCT TOP 100 PERCENT a.Client_DPA_, MAX(a.TransDate) AS LastDate, b.CurrentBal AS Balance, dbo.Client.ClientName " & _ 
		" FROM (SELECT * FROM dbo.tblStatementList where Transdate <= '" & selectedFromDate & "' ) a  " & _ 
		" INNER JOIN dbo.Client ON a.Client_DPA_ = dbo.Client.Client_DPA_  " & _ 
		" INNER JOIN (SELECT     SUM(ISNULL(dbo.tblStatementList.Credit - dbo.tblStatementList.Debit, 0)) + dbo.Client.ClientOpeningBal AS CurrentBal, dbo.tblStatementList.Client_DPA_  " & _ 
		" FROM dbo.tblStatementList INNER JOIN  " & _ 
		" dbo.Client ON dbo.tblStatementList.Client_DPA_ = dbo.Client.Client_DPA_  " & _ 
		" WHERE (dbo.Client.Deleted = 0 AND cast(floor(cast(Transdate as float)) as datetime) <= '" & selectedFromDate & "')  " & _ 
		" GROUP BY dbo.tblStatementList.Client_DPA_, dbo.Client.ClientOpeningBal)b ON dbo.Client.Client_DPA_ = b.Client_DPA_   " & _ 
		" WHERE (dbo.Client.Deleted = 0)  " & clientTypeSQL
		
	If Len(Agent_DPA_) > 0 Then
		SQL = SQL & " GROUP BY a.Client_DPA_, b.CurrentBal, dbo.Client.ClientName, Client.Agent_DPA_ " 
	Else
		SQL = SQL & " GROUP BY a.Client_DPA_, b.CurrentBal, dbo.Client.ClientName  " 
	End If
					
	'sqlStr = "SELECT * FROM [Creditors] WHERE LastDate <= '" & selectedFromDate & "'" 
	'sqlStr2 = "SELECT * FROM [Debtors] WHERE LastDate <= '" & selectedFromDate & "'" '

	sqlstr = SQL & CreditorsSQL
	sqlstr2 = SQL & DebtorsSQL
		
''Agent_DPA_
'Response.Write sqlstr & "<br><br>"

'Response.Write sqlstr2 & "<br>"

'Response.End 
		
	Set conn = GetActiveConnection("KBroker")
	conn.CommandTimeout = 0
     'Conn.Execute("ClientBalancesDelete")
     'Conn.Execute("ClientBalancesProcedure")
    
    'Rs.Open sqlStr, conn.ConnectionString, adOpenKeyset, adLockOptimistic	
    Set rs = conn.execute(sqlStr)  
        
    'RsDebtors.Open sqlStr2, conn.ConnectionString, adOpenKeyset, adLockOptimistic	
    Set RsDebtors = conn.execute(sqlStr2)	
		
	If (rs.EOF Or rs.BOF) And (RsDebtors.EOF Or RsDebtors.BOF)  Then
            %>
            <Script Language="JavaScript">
				alert("The report did not find any values available");
				window.parent.history.go(-1);					
            </Script>
            <% 
            Set Rs = Nothing
            Set RsDebtors = Nothing
            Set Conn = Nothing
            Response.End
    End If
        
    rs.MoveFirst
    RsDebtors.MoveFirst    
  
  if not (RsDebtors.EOF Or RsDebtors.BOF) then
  'Show Debtors
%>
<p id="toPDFOrient" name="toPDFOrient" value="P" style="display:none;">P
<p id="toPDF" name="toPDF">
<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
	<%
	Set rsAgent = Conn.Execute ("SELECT * FROM Agent WHERE Agent_DPA_ = " & Agent_DPA_)
	
	If Not (rsAgent.EOF Or rsAgent.BOF) Then
		accountDesc = rsAgent.Fields("AgentName").Value
		accountAddress = rsAgent.Fields("AgentAddr").Value
	End If
	
	Set rsAgent = Nothing
	%>
	<tr>
      <td nowrap colspan=2><font size="2" face="Arial">Agent:&nbsp;<b><%= accountDesc %></b><br><%= accountAddress %></font></td>
    </tr>
    
    <tr>
      <td nowrap colspan=2><font size="2" face="Arial">&nbsp;</font></td>
    </tr>
    
    <tr>
      <td width="10%" nowrap><b><font face="Arial" size="4"><b>Debtors</b></font></b></td>
      <td nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>
      
    </tr>

    <tr>
      <td nowrap colspan=2><font size="2" face="Arial">as at: <%= FormatDate(selectedFromDate) %></font></td>
    </tr>
  </table>
<br>
 <table border="0" cellspacing="0" cellpadding="3" style="font-family: Arial Narrow; LEFT-MARGIN:100PX" width="700">
    <tr>
      <td bgcolor="#000000" width="50"><b><font color="#FFFFFF" face="Arial Narrow" size="3" >Code</font></b></td>
      <td bgcolor="#000000" width="300"><b><font color="#FFFFFF" face="Arial Narrow" size="3" >Client</font></b></td>
      <td bgcolor="#000000" width="100"><b><font color="#FFFFFF" face="Arial Narrow" size="3" >Telephone</font></b></td>
      <td align="right" bgcolor="#000000" width="100"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Balance</font></b></td>
      <td align="left" bgcolor="#000000" width="100"><b><font color="#FFFFFF" face="Arial Narrow" size="3" >Last Date</font></b></td>
    </tr>
    
    <%
    totalBal = 0
    Do Until RsDebtors.EOF
		totalBal = totalBal + RsDebtors.Fields("Balance").Value 
		thisClientID = RsDebtors.Fields("Client_DPA_").Value
		thisClientName = ""
		thisClientTel = ""
		Set getClientRsDebtors = Conn.Execute("SELECT * FROM Client WHERE Client_DPA_ = " & thisClientID) 
		If Not (getClientRsDebtors.EOF Or getClientRsDebtors.BOF) Then
			thisClientName = getClientRsDebtors.Fields("ClientName").Value
			thisClientTel = getClientRsDebtors.Fields("ClientOfficeTel").Value
		End If
		
		Set getClientRsDebtors = Nothing
		
		balance = FormatNum(Cdbl(trim(RsDebtors.Fields("Balance"))))
		
			If balance <> 0 Then%>

			<tr>
			  <td align="left" width="50"><%= thisClientID %></td>
			  <td align="left" width="300"><%= Mid(thisClientName,1,30) %></td>
			  <td width="100"><%= thisClientTel %></td>
			  <td align="right" width="100"><%= balance %></td>
			  <td align="left" width="100"><%= FormatDate(RsDebtors.Fields("LastDate").Value) %></td>
			</tr>
    
    <%		End If
		RsDebtors.MoveNext
    Loop%>

    <tr>
      <td colspan="5" width="575">
        &nbsp;</td>

    </tr>
    

    <tr>      
	  <td width="50">
      </td>	
      <td width="300"><b>Total Balance: </b></td>
      <td width="100">
      </td>
      <td style="border-style: solid; border-width: 1" align="right" width="100"><%= FormatNum(totalBal) %></td>
      <td width="100">
      </td>
    </tr>

  </table>
<%
 end if
 
         
 if not (rs.EOF Or rs.BOF) then
 'Show Creditors
%>
<br>
<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
    <tr>
      <td width="10%" nowrap><b><font face="Arial" size="4"><b>Creditors</b></font></b></td>
      <td nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>
      
    </tr>

    <tr>
      <td nowrap colspan=2><font size="2" face="Arial">as at: <%= FormatDate(selectedFromDate) %></font></td>
    </tr>

  </table>
<br>
 <table border="0" cellspacing="0" cellpadding="3" style="font-family: Arial Narrow; LEFT-MARGIN:100PX" width="700">
    <tr>
      <td bgcolor="#000000" width="50"><b><font color="#FFFFFF" face="Arial Narrow" size="3" >Code</font></b></td>
      <td bgcolor="#000000" width="300"><b><font color="#FFFFFF" face="Arial Narrow" size="3" >Client</font></b></td>
      <td bgcolor="#000000" width="100"><b><font color="#FFFFFF" face="Arial Narrow" size="3" >Telephone</font></b></td>
      <td align="right" bgcolor="#000000" width="100"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Balance</font></b></td>
      <td align="left" bgcolor="#000000" width="100"><b><font color="#FFFFFF" face="Arial Narrow" size="3" >Last Date</font></b></td>
    </tr>
    
    <%
    totalBal = 0
    Do Until Rs.EOF
		totalBal = totalBal + Rs.Fields("Balance").Value 
		thisClientID = Rs.Fields("Client_DPA_").Value
		thisClientName = ""
		thisClientTel = ""
		Set getClientRs = Conn.Execute("SELECT * FROM Client WHERE Client_DPA_ = " & thisClientID) 
		If Not (getClientRs.EOF Or getClientRs.BOF) Then
			thisClientName = getClientRs.Fields("ClientName").Value
			thisClientTel = getClientRs.Fields("ClientOfficeTel").Value
		End If
		
		Set getClientRs = Nothing
		balance = FormatNum(Cdbl(trim(Rs.Fields("Balance"))))
		
			If balance <> 0 Then%>

			<tr>
			  <td align="left" width="50"><%= thisClientID %></td>
			  <td align="left" width="300"><%= Mid(thisClientName,1,30) %></td>
			  <td width="100"><%= thisClientTel %></td>
			  <td align="right" width="100"><%= balance%></td>
			  <td align="left" width="100"><%= FormatDate(Rs.Fields("LastDate").Value) %></td>
			</tr>
    
    <%		End If
		Rs.MoveNext
    Loop%>

    <tr>
      <td colspan="5" width="575">
        &nbsp;</td>

    </tr>
    
    <tr>      
	  <td width="50">
      </td>	
      <td width="300"><b>Total Balance: </b></td>
      <td width="100">
      </td>
      <td style="border-style: solid; border-width: 1" align="right" width="100"><%= FormatNum(totalBal) %></td>
      <td width="100">
      </td>
    </tr>

  </table>
<%
 end if
%>
</body>

</html>