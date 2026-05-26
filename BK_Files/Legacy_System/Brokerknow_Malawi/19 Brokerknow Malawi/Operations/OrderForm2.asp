<!--#include file="../libroutines.asp"-->
<%
Const report_ViewName = "OrderForm"
Const reportPage = "OrderForm.asp"
Const headerColCount = 3
Const groupingHeaderCol = 0
Const reportTitle = "Order Form"

OrderID = Request.QueryString("order_id")

Set Conn = GetActiveConnection("KBroker")

set Rs = conn.execute("SELECT * FROM OrderForm WHERE (Order_DPA_ = " & OrderID & ")")

if (Rs.bof or Rs.eof) then
	%>
	<script language = 'vbscript'>
    		ShowMessage "Invalid data sent. Please try again"
	</script>
	<%
	Response.End
end if
%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Order Form</title>
<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<SCRIPT language=VBScript src="../scripts/reports.vbs"></SCRIPT>
 <SCRIPT language=Javascript src="../scripts/reports.js"></SCRIPT>
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">

<style media="print">
	@page 
		{
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

<body onLoad="javascript: printReportDoc()">
<%
'OrderID
DrawPageFunctions2 true, true, true, "EditOrder.asp?action=first&From=Report&ID=" & OrderID
%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="57">
	<tr>
		<td width="66%" colspan="2" height="19"><b><font face="Arial" size="6">&nbsp;</font></b></td>
		<td width="34%" rowspan="3" height="57">
		<p align="center"></td>
	</tr>
	<tr>
		<td width="66%" colspan="2" height="19"><b><font face="Arial" size="5"><%=Rs("OrderTypeDescription") & "&nbsp;" & Rs("Order_DPA_")%> </font></b></td>
	</tr>
	<tr>
		<td width="17%" height="19"><font face="Arial">&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; <b><%=FormatDate(Date)%></b></font></td>
		<td width="49%" height="19"><font face="Arial"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=Hour(Time()) & ":" & Minute(Time())%> </b>
		</font> </td>
	</tr>
</table>
<br>&nbsp;
<br>&nbsp;
<br>&nbsp;
<table border="00" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
	<tr>
		<td width="100%" colspan="2"><font face="Arial"><b>Client:</b></font></td>
	</tr>
	<tr>
		<td width="40%"><font face="Arial"><%=Rs("ClientName")%></font></td>
		<td width="60%"><font face="Arial">[<%=Rs("Client_DPA_")%>]</font></td>
	</tr>
	<tr>
		<td width="40%"><font face="Arial"><%=Rs("ClientAddr")%></font></td>
		<td width="60%">&nbsp;</td>
	</tr>
	<tr>
		<td width="40%"><font face="Arial">Tel&nbsp;:&nbsp;<%=Rs("ClientOfficeTel")%></font></td>
		<td width="60%">&nbsp;</td>
	</tr>
	<tr>
		<td width="100%" colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td width="100%" colspan="2"><font face="Arial"><b>Agent</b></font></td>
	</tr>
	<tr>
		<td width="40%"><font face="Arial"><%=Rs("AgentName")%></font></td>
		<td width="60%"><font face="Arial">[<%=Rs("Agent_DPA_")%>]</font></td>
	</tr>
</table>
<br>&nbsp;
<br>&nbsp;
<br>&nbsp;
<table border="0" cellpadding="0" cellspacing="3" style="border-collapse: collapse" bordercolor="#111111" width="100%">
	<tr>
		<td width="35%"><font face="Arial"><b>Security</b></font></td>
		<td align="left" width="12%"><font face="Arial"><b>Quantity</b></font></td>
		<td align="left" width="12%"><font face="Arial"><b>Price</b></font></td>
		<td align="left" width="12%"><font face="Arial"><b>Amount</b></font></td>
		<td width="12%"><font face="Arial"><b>Valid Until</b></font></td>
		<td width="17%"><font face="Arial"><b>Certificate</b></font></td>
	</tr>
	<tr>
		<td width="100%" colspan="5">&nbsp;</td>
	</tr>
	<%
	dim AmtDep
	
	AmtDep=0
	
	do until Rs.eof
		if isNull(Rs("OrdDetailValidity")) then
			if Len(Rs("OrdDetailValidity")) <> 0 then
				Validity = FormatDate(Rs("OrdDetailValidity"))
			else
				Validity = "Until Execution"
			end if
		else
			Validity = "Until Execution"
		end if
		%>
		<tr>
			<%
			if(Cint(Rs("OrderSecType_DPA_"))=1) then
			%>
			<td width="35%"><%=Rs("SecurityName")%>&nbsp;<%=Rs("BondDescription")%></td>
			<%
			else
			%>
			<td width="35%"><%=Rs("SecurityName")%></td>
			<%
			end if
			%>
			<%if(Rs("Amount")<>0) then
			AmtDep=1
			%>
			<td align="left" width="12%">AMT DEP*</td>
			<td align="left" width="12%">Best</td>			
			<td align="left" width="12%"><%=FormatNum(Rs("Amount"))%></td>
			<%
			else
			%>
			<td align="left" width="12%"><%=Rs("OrdDetailQty")%></td>
			<td align="left" width="12%"><%=Rs("OrdDetailPrice")%></td>
			<td align="left" width="12%">&nbsp;</td>
			<% 
			end if
			%>			
			<td width="12%"><%=Validity%></td>
			<td width="17%"><%=Rs("OrdDetailCertNo")%></td>
		</tr>
		<%
		Rs.movenext
	loop
	%>
	<tr>
		<td width="100%" colspan="5">&nbsp;</td>
	</tr>
	<%
	if(Cint(AmtDep)=1) then
	%>
	<tr>
		<td width="100%" colspan="5"></td>
	</tr>
	<tr>
		<td width="100%" colspan="5">* Amount Dependent: For best orders the quantity is dependent on the current market price.</td>
	</tr>
	<%
	end if
	%>
	<tr>
		<td width="100%" colspan="5">&nbsp;</td>
	</tr>
	<tr>
		<td width="100%" colspan="5"><font face="Arial">Please call us after 3 days to confirm the 
		status of your offer. We shall endeavour to execute your orders on a best 
		offer basis.</font></td>
	</tr>
	<tr>
		<td width="100%" colspan="5">&nbsp;</td>
	</tr>
</table>
<br>&nbsp;
<br>&nbsp;
<br>&nbsp;
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="101%">
	<tr>
		<td width="10%"><font face="Arial"><b>Signature</b></font></td>
		<td width="40%">__________________________________</td>
		<td width="10%"><font face="Arial"><b>&nbsp;&nbsp;&nbsp;&nbsp;Date</b></font></td>
		<td width="40%">__________________________________</td>
	</tr>
	<tr>
		<td width="101%" colspan="4">&nbsp;</td>
	</tr>
	<tr>
		<td width="10%"><font face="Arial"><b>Customer Service</b></font></td>
		<td width="40%">__________________________________</td>
		<td width="10%">&nbsp;</td>
		<td width="40%">&nbsp;</td>
	</tr>
	<tr>
		<td width="101%" colspan="4">&nbsp;</td>
	</tr>
	<tr>
		<td width="101%" colspan="4"><font face="Arial">IMPORTANT</font></td>
	</tr>
	<tr>
		<td width="101%" colspan="4">
		<ol>
			<li><font face="Arial">If payment is by cheque the order will be 
			executed after the cheque clears</font></li>
			<li><font face="Arial">We shall not be held responsible for any 
			dividends/bonus due to delay in the execution and delivery of documents 
			to our office in good time</font></li>
		</ol>
		</td>
	</tr>
	<tr>
		<td width="101%" colspan="4">&nbsp;</td>
	</tr>
</table>
</body>

</html>