<html>

<head>
<title>Sent Batch Items</title>
  
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
			
		tr.pageNumbering{
			display:none;
		}
	}

</style>
</head>

<body Class="Reports">

<!--#include file="../libroutinesTEST.asp"-->
<%

genReport = Request.Form("genReport")
selectedFromDate = Request.Form("txtDate")
selectedToDate = Request.Form("txtDate2")
timeLimit = Request.Form("timeLimit")

If genReport <> "1" Or Not IsDate(selectedFromDate) Then%>
	<Script Language="JavaScript">
		report_SetBodyClass();
		
		function validateForm(frm){			
			if (frm.txtDate.value==''){
				alert("Select a date");
				frm.txtDate.focus();
				return;
			}
			
			frm.target = '_self';			
			frm.submit();
		}
		
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "txtDate","cmdDate","<%= FormatDate(Date) %>",1);
		var cal2=new ctlSpiffyCalendarBox("cal2", "frmMain", "txtDate2","cmdDate2","<%= FormatDate(Date) %>",1);
	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="SentBatchItems.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">	
		<table>
			<tr>
				<td>From Date:</td>
				<td>
				<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>
			</tr>
			<tr>
				<td>To Date:</td>
				<td>
				<SCRIPT language="JavaScript">cal2.writeControl();</SCRIPT>	
				</td>
			</tr>
			<tr>
				<td colspan=2><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... " id=Button1 name=Button1>&nbsp;&nbsp;</td>
			</tr>
		</table>
	</form>
	<%
	Set rs = Nothing
	Set Conn = Nothing
	Response.End
End If

%>

<% DrawPageFunctions True, True, True, True %>

<%
	Dim conn 
	Dim sqlStr
	Dim Rs
	   
	Set conn = GetActiveConnection("KBroker")

	sqlStr = "SELECT SendBatchItems.SendBatchItemID, Client.Client_DPA_, Client.ClientName, Client.ClientEmail, SendBatchItems.Frequency, SendBatchItems.Report, " & _
	" SendBatchItems.DateSent, SendBatchItems.TimeSent" & _
	" FROM SendBatchItems INNER JOIN" & _
	" Client ON SendBatchItems.Client_DPA_ = Client.Client_DPA_" & _
	" WHERE SendBatchItems.DateSent BETWEEN '"& selectedFromDate &"' AND '"& selectedToDate &"'"
	Set Rs = Conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        
    %>
	<p id="toPDFOrient" name="toPDFOrient" value="P" style="display:none;">P
	<p id="toPDF" name="toPDF">
		
	<table border="0" cellspacing=2 cellpadding=2 align="center">
	<tr>
		<td colspan=6><b><font face="Arial Narrow" size="4">Sent Batch Reports</font></b></td>
	</tr>		

	<tr bgcolor="#000000">
		<td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial" size="2">&nbsp;</font></b></td>
		<td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial" size="2">Code</font></b></td>
		<td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial" size="2">Name</font></b></td>	
		<td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial" size="2">Email</font></b></td>	
		<td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial" size="2">Frequency</font></b></td>
		<td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial" size="2">Report</font></b></td>
		<td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial" size="2">Sent</font></b></td>
	</tr>
	<%
	If Not (Rs.EOF Or Rs.BOF) Then
			num = 1
			Do Until rs.EOF
				Select Case rs.Fields("Report")
					Case "ClientStatement"
						theReport = "Client Statement"
						
					Case "ClientContract"
						theReport = "Client Contract"
						
					Case "ClientContractCompounded"
						theReport = "Client Contract Compounded"
						
					Case "HoldingsValuation"
						theReport = "Holdings Valuation"
						
					Case Else
						theReport = ""
				End Select
				%>
				<tr>
					<td><%=num%>.</td>
					<td><%=rs.Fields("Client_DPA_")%></td>
					<td><%=rs.Fields("ClientName")%></td>
					<td><%=rs.Fields("ClientEmail")%></td>
					<td><%=rs.Fields("Frequency")%></td>
					<td><%=theReport%></td>
					<td><%=FormatDate(rs.Fields("DateSent")) & " " & FormatDateTime(rs.Fields("TimeSent"), vbLongTime)%></td>                   
				</tr>
	            <%
	            num = num + 1
				Rs.MoveNext   
			Loop
	Else
		%>
		<tr>
			<td colspan="7" align="left"><font color="black" face="Arial" size="2">&nbsp;No Records.</font></td>
		</tr>
		<%			
	End if
        
	Set Rs = Nothing
	conn.Close
	Set conn = Nothing
    %>
</table>

</body>

</html>
