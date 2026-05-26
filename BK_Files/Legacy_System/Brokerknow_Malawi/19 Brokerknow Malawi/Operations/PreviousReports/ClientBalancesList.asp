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

Selectedtype=Request.Form("cboselect")
selectedAgent = Request.Form("cboagent")
selectedAccount = Request.Form("cboaccount")

selectedFromDate = Request.Form("transFromDate")

If genReport <> "1" Then%>
	<Script Language="JavaScript">
		report_SetBodyClass();
		function validateForm(frm){			
			//if (frm.cboagent.selectedIndex < 0)
			//{
			//	alert("Select agent");
			//	frm.cboagent.focus();
			//	return;
			//}
			
			frm.target = '_self';			
			frm.submit();
		}
		
		function ChangeClient(selecttype)
			{
			var thetype=selecttype.value;			
			if(thetype=='Account')
				{
				document.getElementById("cboaccount").style.display = ""
				document.getElementById("cboagent").style.display = "none"
				}
			else
				{
				document.getElementById("cboaccount").style.display = "none"
				document.getElementById("cboagent").style.display = ""
				}
			}
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "transFromDate","cmdDate","<%= FormatDate(DateAdd("d", -90, Date)) %>",1);
	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="ClientBalancesList.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">
		<table>
		<tr><td>Account/Agent:</td>
			<td>
				<select name="cboSelect" onchange='ChangeClient(this)'>			
					<option selected SearchCode = "0" SearchText = "Account" value = 'Account'>Account Manager</option>			
					<option value='Agent'>Agent</option>			
				</select>
    		</td>
					
		</tr>
		<tr>
			<td>Select: </td>
				<td><select name = 'cboaccount' id = 'cboaccount' size="1">
					<option selected value = ''></option>
					<%					
					dim ClientName
					dim NameClient
					        Set conn = GetActiveConnection("KBroker")
					        
					        sqlStr = "SELECT Distinct OwnerName,Owner_DPA_ FROM FullClientList where not(ownerName is null) order by Owner_DPA_"
					        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
					        If Not (rs.EOF Or rs.BOF) Then
					                rs.MoveFirst
					                Do Until rs.EOF
					                ClientName=rs.Fields("OwnerName")
					                NameClient=rs.Fields("Owner_DPA_") & " " & Mid(ClientName,1,20)
					                %>
					                        <option value = '<%=rs.Fields("Owner_DPA_")%>'><%=NameClient%></option>
					                        <%rs.MoveNext
					                Loop
					        End If
					%>

					    </select>
									
				<select name = 'cboagent' id = 'cboagent' size="1" style="display:none">
					<option selected value = ''></option>
					<%					
					'dim ClientName
					'dim NameClient
					        Set conn = GetActiveConnection("KBroker")
					        
					        sqlStr = "SELECT distinct AgentName,Agent_DPA_ FROM FullClientList where not(Agent_DPA_ is null) order by Agent_DPA_"
					        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
					        If Not (rs.EOF Or rs.BOF) Then
					                rs.MoveFirst
					                Do Until rs.EOF
					                ClientName=rs.Fields("agentName")
					                NameClient=rs.Fields("agent_DPA_") & " " & Mid(ClientName,1,20)
					                %>
					                        <option value = '<%=rs.Fields("agent_DPA_")%>'><%=NameClient%></option>
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

Dim BalanceTitle

	Set conn = GetActiveConnection("KBroker")
	Set Rs = CreateObject("ADODB.Recordset")
	
	if(Selectedtype="Account") then		
		BalanceTitle="Account Manager Balances"
		accountDesc="Account Manager"				        
		sqlStr = "SELECT * FROM BalanceAccountList WHERE Owner_DPA_ = " & selectedAccount & " AND (Day(TransDate) = Day(#" & CDate(FormatDate(selectedFromDate)) & "#))"
	else
		BalanceTitle="Agent Balances"
		accountDesc="Agent"				        
				        
		sqlStr = "SELECT * FROM BalanceAgentList WHERE agent_DPA_ = " & selectedAgent & " AND (Day(TransDate) = Day(#" & CDate(FormatDate(selectedFromDate)) & "#))"
	end if
	
	sqlstr = SQLServerFormat(HandleQuote(sqlStr))
	
	Rs.CursorLocation = adUseClient	
	Rs.Open sqlstr, conn.ConnectionString, adOpenKeyset, adLockOptimistic
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

%>
<table border="0" cellspacing="2" cellpadding="2" style="font-family: Arial Narrow" width="100%">
    <tr>
		<td width="10%" nowrap><font face="Impact" size="4"><%=BalanceTitle%></font></td>
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
      <td width="48%"><b><%= accountDesc %></b></td>
      <td width="1%">&nbsp;</td>
    </tr>

    <tr>
      <td width="1%"><b><font size="2" face="Arial">&nbsp;</font></b></td>
      <td width="48%"><%= accountAddress %></td>
    </tr>

</table>
<BR>


  <table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow; LEFT-MARGIN:100PX"  width="100%">
    <tr>
      <td style="border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"><b><font face="Arial Narrow" size="3">Client Code</font></b></td>
      <td style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"><b><font face="Arial Narrow" size="3">Client Name</font></b></td>
      <td style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"><b><font face="Arial Narrow" size="3">Balance</font></b></td>
      <td align="left" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"  align=right><b><font face="Arial Narrow" size="3">Contacts</font></b></td>
    </tr>
    <%
    do while rs.eof=false
    %>
    <tr>
      <td><font face="Arial Narrow" size="3"><%=rs("Client_DPA_")%></font></td>
      <td><font face="Arial Narrow" size="3"><%=rs("ClientName")%></font></td>
      <td><font face="Arial Narrow" size="3"><%=rs("Balance")%></font></td>
      <td align="left"><font face="Arial Narrow" size="3"><%=rs("Contacts")%></font></td>
    </tr>    
    <%
    rs.moveNext
    loop
    %>
</table>   
</body>

</html>