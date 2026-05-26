<!--#include file="../libroutines.asp"-->

<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Portfolio BS</title>  
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
<SCRIPT language=VBScript src="../scripts/reports.vbs"></SCRIPT>
<SCRIPT language=Javascript src="../scripts/reports.js"></SCRIPT>



<style media="print">
		@page {
				size: landscape;
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
</head>

<body Class="Reports">



<%
'FirstDay=DateSerial(Year(Date), Month(Date) + iOffset, 1)
FirstDay=DateSerial(Year(Date), Month(Date)-1 + iOffset, 1)

genReport = Request.Form("genReport")
selectedClient = Request.Form("cboClient")

selectedFromDate = Request.Form("txtFromDate")
selectedToDate = Request.Form("txtToDate")

If genReport <> "1" Then%>
		<Script Language="JavaScript">
		report_SetBodyClass();
		function validateForm(frm){			
			//if (frm.txtDate.value==''){
			//	alert("Select a date");
			//	frm.txtDate.focus();
			//	return;
			//}
			
			frm.target = '_self';			
			frm.submit();
		}
		
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "txtFromDate","cmdDate","<%= FormatDate(FirstDay) %>",1);
		var cal1=new ctlSpiffyCalendarBox("cal1", "frmMain", "txtToDate","cmdDate","<%= FormatDate(Date) %>",1);
	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="ClientCommissions.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">	
		<table width="80%">			
			<tr>
				<td>Client: </td>
				<td><input type = 'text' name ='txtClientCode' id = 'txtClientCode' size="10" onBlur="txtval = this.value; selectItem(cboClient);"></td>
				<td><select name = 'cboClient' id = "cboClient" size="1" 
    				onchange='UpdateCode(true,cboClient,txtClientCode)'
					onKeypress="return (dodefaultaction()==''); " 
					onKeydown="return (dodefaultaction()==''); " 
					onKeyup="return (FilterData(this,1,UpdateCode(change(cboClient,0),cboClient,txtClientCode)));" 
					onfocus="txtval = '';inputIsItemCode = 1;" 
					onblur="txtval = '';inputIsItemCode = 1;">
					<option selected SearchCode = "" SearchText = ""  value = ''></option>
					<%
					dim ClientName
					dim NameClient
					        Set conn = GetActiveConnection("KBroker")
					        
					        sqlStr = "SELECT * FROM FullClientList order by ClientName"
					        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
					        If Not (rs.EOF Or rs.BOF) Then
					                rs.MoveFirst
					                Do Until rs.EOF
					                ClientName=rs.Fields("ClientName")
					                NameClient=Mid(ClientName,1,30)
					                %>					                        
					                        <option SearchCode = "<%=rs.Fields("Client_DPA_")%>" SearchText = "<%=NameClient%>" value = '<%=rs.Fields("Client_DPA_")%>'><%=NameClient%></option>

					                        <%rs.MoveNext
					                Loop
					        End If
					%>

					    </select>
				</td>
			</tr>

			<tr>
				<td>From Date:</td>
				<td>
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>
			</tr>
			<tr>
				<td>To date:</td>
				<td>
					<SCRIPT language="JavaScript">cal1.writeControl();</SCRIPT>	
				</td>
			</tr>
			
			<tr>
				<td colspan=2 align="Center"><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... ">&nbsp;&nbsp;</td>
			</tr>
		</table>
		
	</form>
	
	<%
	Response.End
End If

%>


<% DrawPageFunctions True, True, True %>

<%

	If Trim(selectedClient)="" Then%>
		<Script Language="JavaScript">
			alert("Please select The Account Manager")
			window.history.go(-1);
		</Script>
		<%
		Response.End
	End If	
	
	Set conn = GetActiveConnection("KBroker")
	Set Rs = CreateObject("ADODB.Recordset")						        
	sqlStr = "SELECT * FROM ClientCommissions WHERE Client_DPA_ = " & selectedClient & " AND TransDate Between '" & FormatDate(CDate(selectedFromDate)) & "' and '" & FormatDate(CDate(selectedToDate)) & "'"	
	
	'Response.write(sqlStr)
	'Response.end
	
	Rs.CursorLocation = adUseClient	
	Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
	
	Set getClientRs = Conn.Execute("SELECT * FROM ClientList WHERE Client_DPA_ = " & selectedClient) 
		If Not (getClientRs.EOF Or getClientRs.BOF) Then
			thisOwnerName = getClientRs.Fields("ClientName").Value & "&nbsp;" & "[" & getClientRs.Fields("Client_DPA_").Value & "]"			
		End If
		
	
	If rs.EOF Or rs.BOF Then%>
		<Script Language="JavaScript">
			alert("The specified Client does not have any transaction using the specified date criterion")
			window.history.go(-1);
		</Script>
		<%Set Rs = Nothing
		Set Conn = Nothing
		Response.End
	End If	
	
	Dim Total
%>	

<table border="0" cellspacing="2" cellpadding="2" style="font-family: Arial Narrow" width="100%">
    <tr>
		<td width="10%" nowrap><font face="Impact" size="4">CLIENT COMMISSIONS</font></td>
      <td width="60%" nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>
      
    </tr>

  </table>
<br>
<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
    <tr>
      <td width="1%"><b>Date:</b></td>
      <td width="48%"><b>From:</b><%= FormatDate(selectedFromDate) %>&nbsp;<b>To:</b><%= FormatDate(selectedToDate) %></td>
    </tr>

    <tr>
      <td width="1%"><b>Account:</b></td>
      <td width="48%"><%= thisOwnerName %></td>
    </tr>
      
    <tr>
      <td width="1%"><b>Agent:</b></td>
      <td width="48%"><%= rs("AgentName") %>&nbsp;[<%=rs("Agent_DPA_")%>]</td>
    </tr>   	
</table>
<BR>


  <table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow; LEFT-MARGIN:100PX"  width="570">
    <tr>
      <td width="150"><b><font face="Arial Narrow" size="2">Traded</font></b></td>
	<td width="50"><b><font face="Arial Narrow" size="2">Type</font></b></td>
      <td width="70"><b><font face="Arial Narrow" size="2">Security</font></b></td>
      <td width="50"><b><font face="Arial Narrow" size="2">Broker</font></b></td>      
	<td width="100"><b><font face="Arial Narrow" size="2">Contract</font></b></td>      
	<td width="30"><b><font face="Arial Narrow" size="2">Slip</font></b></td>      
	<td align="right" width="30"><b><font face="Arial Narrow" size="2">Price</font></b></td>      
	<td align="right" width="50"><b><font face="Arial Narrow" size="2">Quantinty</font></b></td>
	<td align="right" width="50"><b><font face="Arial Narrow" size="2">Comm</font></b></td>      
	<td align="right" width="100"><b><font face="Arial Narrow" size="2">Gross</font></b></td>      
    </tr>
    <%
    Do while rs.eof=false
    Total=Total+Rs("LevyAmount")
    TotalGross=TotalGross+Rs("LotGrossAmount")
    TotalQty=TotalQty+Rs("LotQty")
    %>   
	<tr>	
	  <td width="150"><font size="1"><%=FormatDate(Rs("TransDate"))%></font></td>
	  <td width="50"><font size="1"><%=Rs("OrderTypeDescription")%></font></td>
	  <td width="70"><font size="1"><%=Rs("SecurityCode")%></font></td>
	  <td width="50"><font size="1"><%=Rs("BrokerName")%></font></td>
	  <td width="100"><font size="1"><%=Rs("ContractNumber")%></font></td>
	  <td width="30"><font size="1"><%=Rs("LotSlipNo")%></font></td>
	  <td align="right" width="30"><font size="1"><%=FormatNum(Rs("LotPrice"))%></font></td>
	  <td align="right" width="50"><font size="1"><%=FormatNum(Rs("LotQty"))%></font></td>
	  <td align="right" width="50"><font size="1"><%=FormatNum(Rs("LevyAmount"))%></font></td>
	  <td align="right" width="100"><font size="1"><%=FormatNum(Rs("LotGrossAmount"))%></font></td>		  
	</tr>    
    <%	Rs.MoveNext
	Loop
	%>
	<tr>	
	  <td colspan="10"><font size="1">&nbsp;</font></td>	  
	</tr> 
	<tr>	
	  <td width="68" colspan="3"><font size="1">&nbsp;</font></td>
	  <td width="40" colspan="3"><font size="1">&nbsp;</font></td>
	  <td width="150"><font size="2"><b>Total</b></font></td>
	  <td align=right style="border-style: solid; border-color: #000000; border-width: 1" height="30px"><%= FormatNum(TotalQty) %></td>
	  <td align=right style="border-style: solid; border-color: #000000; border-width: 1" height="30px"><%= FormatNum(Total) %></td>
	  <td align=right style="border-style: solid; border-color: #000000; border-width: 1" height="30px"><%= FormatNum(TotalGross) %></td>	  	  
	</tr>    
    
  </table>
   
<%Set Rs = Nothing
Set Conn = Nothing%>   
</body>

</html>