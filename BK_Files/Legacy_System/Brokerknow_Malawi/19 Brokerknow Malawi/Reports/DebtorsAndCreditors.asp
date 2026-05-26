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
	<form method="POST" action="DebtorsAndCreditors.asp" Name="frmMain" id="frmMain">
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

<p id="toPDFOrient" name="toPDFOrient" value="P" style="display:none;">P
<p id="toPDF" name="toPDF">

<%
   Dim conn 
   Dim sqlStr
   Dim rs
	Set Rs = CreateObject("ADODB.Recordset")						        
	Rs.CursorLocation = adUseClient	
	
	'Evaluate Client Type 

   if trim(clientType) = "" then
    clientTypeSQL = ""
   elseif trim(clientType) =  1 then
    clientTypeSQL = " AND dbo.client.Iscustodian = 1 "
   else
    clientTypeSQL = " AND dbo.client.Iscustodian <> 1 "
   end if

	    'sqlStr = "SELECT * FROM [Debtors] WHERE LastDate <= '" & selectedFromDate & "'" ' 
	    
        DebtorsSQL = " HAVING (b.CurrentBal < 0) ORDER BY dbo.Client.ClientName "

		SQL =   " SELECT DISTINCT TOP 100 PERCENT a.Client_DPA_, MAX(a.TransDate) AS LastDate, b.CurrentBal AS Balance, dbo.Client.ClientName,dbo.Client.CreditLimit " & _ 
					" FROM (SELECT * FROM dbo.ClientTransactionList where Transdate <= '" & selectedFromDate & "' ) a  " & _ 
					" INNER JOIN dbo.Client ON a.Client_DPA_ = dbo.Client.Client_DPA_  " & _ 
					" INNER JOIN (SELECT     SUM(ISNULL(dbo.StatementList.Credit - dbo.StatementList.Debit, 0)) + dbo.Client.ClientOpeningBal AS CurrentBal, dbo.StatementList.Client_DPA_  " & _ 
					" FROM dbo.StatementList INNER JOIN  " & _ 
					" dbo.Client ON dbo.StatementList.Client_DPA_ = dbo.Client.Client_DPA_  " & _ 
					" WHERE (dbo.Client.Deleted = 0 AND Transdate <= '" & selectedFromDate & "')  " & _ 
					" GROUP BY dbo.StatementList.Client_DPA_, dbo.Client.ClientOpeningBal)b ON dbo.Client.Client_DPA_ = b.Client_DPA_   " & _ 
					" WHERE (dbo.Client.Deleted = 0)  " & clientTypeSQL & _ 
					" GROUP BY a.Client_DPA_, b.CurrentBal, dbo.Client.ClientName,dbo.Client.CreditLimit  " 
        sqlStr = SQL & DebtorsSQL
         
         
         'Response.Write sqlStr 
         
		Set conn = GetActiveConnection("KBroker")
       
        Conn.Execute("ClientBalancesProcedure")

        Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic		
		
		If rs.EOF Or rs.BOF Then
                %>
                <Script Language="JavaScript">
					alert("The report did not find any values available");
					window.parent.history.go(-1);					
                </Script>
                <% Set Rs = Nothing
                Set Conn = Nothing
                Response.End
        End If
        
        rs.MoveFirst
        
%>

<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
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
      <td bgcolor="#000000" width="200"><b><font color="#FFFFFF" face="Arial Narrow" size="3" >Client</font></b></td>
      <td bgcolor="#000000" width="80"><b><font color="#FFFFFF" face="Arial Narrow" size="3" >Telephone</font></b></td>
      <td align="right" bgcolor="#000000" width="80"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Balance</font></b></td>
      <td align="right" bgcolor="#000000" width="80"><b><font color="#FFFFFF" face="Arial Narrow" size="3">CreditLimit</font></b></td>
      <td align="right" bgcolor="#000000" width="80"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Excess</font></b></td>
      <td align="left" bgcolor="#000000" width="80"><b><font color="#FFFFFF" face="Arial Narrow" size="3" >Last Date</font></b></td>
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
		
			If Rs.Fields("Balance").Value <> 0 Then%>
			<tr>
			  <td align="left" width="50" nowrap><%= thisClientID %></td>
			  <td align="left" width="200" nowrap><%= thisClientName %></td>
			  <td width="80" nowrap><%= thisClientTel %></td>
			  <td align="right" width="80" nowrap><%= FormatNum(Rs.Fields("Balance").Value) %></td>
			  <td align="right" width="80" nowrap><%= FormatNum(Rs.Fields("CreditLimit").Value) %></td>
			  <td align="right" width="80" nowrap><%= FormatNum((Rs.Fields("Balance").Value + Rs.Fields("CreditLimit").Value) ) %></td>
			  <td align="right" width="80" nowrap><%= FormatDate(Rs.Fields("LastDate").Value) %></td>
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

</body>

</html>