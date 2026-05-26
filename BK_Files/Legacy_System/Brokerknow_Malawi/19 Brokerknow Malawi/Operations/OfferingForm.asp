<!--#include file="../libroutines.asp"-->
<%
Const report_ViewName = "OfferingForm"
Const reportPage = "OfferingForm.asp"
Const headerColCount = 3
Const groupingHeaderCol = 0
Const reportTitle = "IPO Forward Form"

OrderID = Request.QueryString("ID")
'returnPath = Request.QueryString("returnPath")
returnPath ="EditForward.asp"
OrdDetailID = Request.QueryString("ID")

'Always return ID as OrdDetail_DPA_ Otherwise Default to Order_DPA_
if trim(OrdDetailID) = "" then
 ID = OrderID
Else
 ID = OrdDetailID
end if

'Get the Redirect Page
if trim(returnPath) = "" then
 returnPath = "EditForward.asp"
end if

Set Conn = GetActiveConnection("KBroker")

set Rs = conn.execute("SELECT top 4 * FROM OfferingForm WHERE (Offering_DPA_ = " & ID & ")")

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
			
		margin-left: 1cm;
		margin-right: 2cm;
		margin-top: 0cm;    
		margin-bottom: 0cm;
		size: portrait;
			
		br.newpage{
			page-break-before:always;
		}
	}
</style>

</head>

<body>
<%
DrawPageFunctions2 true, true, true, returnPath & "?ID=" & ID

pageNumber =1
p=0
i = 0
page1=0
page2=0
pagenumbers=0
m=0
pagenumbers=int(rs.recordcount/5)

if(rs.recordcount mod 5)<> 0 then
pagenumbers=pagenumbers+1
end if

do until p=pagenumbers
page1=pageNumber
if(page1<>page2) then

%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
<tr>
<td colspan="2" align="center" valign="top">
			<!--#include file="../Reports/Header.asp"--></td>
</tr>
<td colspan="2" height="20">&nbsp;</td>
</tr>
<tr>
	<td colspan="2" height="10" align="center"><b><font face="Arial" size="5"><%=" FORWARD " & " " & rs("Offering_DPA_")%> </font></b></td>
</tr>
</tr>
<td colspan="2" >&nbsp;</td>
</tr>
</td>
</tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
	<td colspan="6">&nbsp;&nbsp;&nbsp;<font size="2"><b>CLIENT DETAILS</font></b></td>
</tr>
<tr>
	<td colspan="6" valign="top"><hr width="100%" color="black"></td>
</tr>
<tr>
	<td width="10%">&nbsp;&nbsp;&nbsp;<b>NAME:</b></td>
	<td COLSPAN="2"><%=ucase(Rs("ClientName"))%></td>
	<td><b>CLIENT CDS.:</b>&nbsp;&nbsp;&nbsp;<%=Rs("ClientCDSNo")%></td>
	<td></td>
</tr>
<tr><td colspan="6">&nbsp;</td></tr>
<tr><td colspan="6">
<TABLE cellspacing="0" width="100%">
<tr>
	<td width="120" style="border-right: #000000 1px inset; border-top: #000000 1px inset; border-bottom: #000000 1px inset" >&nbsp;&nbsp;CONTACT NUMBERS:</td>
	<td width="5" style="border-right: #000000 1px inset; border-top: #000000 1px inset; border-bottom: #000000 1px inset" >&nbsp;W&nbsp;</td>
	<td  width="200" style="border-right: #000000 1px inset; border-top: #000000 1px inset; border-bottom: #000000 1px inset">&nbsp;<%=Rs("ClientOfficeTel")%></td>
	<td width="5" style="border-right: #000000 1px inset; border-top: #000000 1px inset; border-bottom: #000000 1px inset" >&nbsp;F&nbsp;</td>
	<td style="border-right: #000000 1px inset; border-top: #000000 1px inset; border-bottom: #000000 1px inset" >&nbsp;<%=Rs("ClientFax")%></td>
	<td width="5"style="border-right: #000000 1px inset; border-top: #000000 1px inset; border-bottom: #000000 1px inset" >&nbsp;M&nbsp;</td>
	<td style="border-top: #000000 1px inset; border-bottom: #000000 1px inset" >&nbsp;<%=Rs("ClientCellTel")%></td>
</tr>
</table>
</td></tr>
<tr>
	<td>&nbsp;</td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
</tr>	
<tr>
	<td valign="top">&nbsp;<B>ADDRESS:</B></td>
	<td valign="top">&nbsp;<%=Replace(Rs("ClientAddr"),vbCrLf,"<br>")%></td>
	<td></td>
	<td></td>
	<td></td>
</tr>
<tr>
	<td colspan="9"><hr color="black" height="1" width ="100%"></hr></td>
</tr>

<tr>
<td colspan="6">
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
	<tr>
		<td width="100%" colspan="5">&nbsp;</td>
	</tr>
	
	<tr>
		<td width="35%"><font face="Arial"><b>&nbsp;Security</b></font></td>
		<td align="left" width="12%"><font face="Arial"><b>&nbsp;Quantity</b></font></td>
		<td align="left" width="12%"><font face="Arial"><b>&nbsp;Price</b></font></td>
		<!-- <td align="left" width="12%"><font face="Arial"><b>Amount</b></font></td> -->
		<td width="12%"><font face="Arial"><b>&nbsp;Valid&nbsp;Until&nbsp;&nbsp;</b></font></td>		
		<!-- <td>&nbsp;</td> -->
	</tr>
	<%	
	end if		
	k=0
	
	do until Rs.eof
	if(i<5) then
	k=k+1
		%>
		
		<tr>			
			<td width="35%" height="20"><%=Rs("SecurityName")%></td>						
			<td align="left" width="12%" height="20">&nbsp;<%=FormatNumEx(Rs("Alloted_Rights"),0)%>&nbsp;</td>
			<td align="left" width="12%" height="20">&nbsp;<%=FormatNum(Rs("Offering_Price"))%>&nbsp;</td>
			<!-- <td align="left" width="12%" height="20">&nbsp;</td> -->
			<% 
			end if
			%>			
			<td width="12%" height="20"><%=FormatDate(CDate("12/4/2006"))%></td>			
		</tr>
		
		<!-- <tr height="3">
			<td width="100%" colspan="6" height="20">&nbsp;</td>
		</tr> -->
		<%
		'end if
		Rs.movenext

		i = i + 1			
		
	loop
		
	for j = 0 to (2-i)
		%>
		 <tr>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
		</tr> 
		<%
	next
	
	if(Page1<>page2) then
	%>
	<tr>
		<td width="100%" colspan="6">&nbsp;</td>
	</tr>
</table>
	

</td>
</tr>
<tr><td colspan="6">&nbsp;</td></tr>
<tr><td colspan="6"> I hereby indemnify African Alliance and hold it harmless against any liability for all and any actual or contigent losses, liabilities, damages and costs (including, without limitation, legal costs on the scale as between attorney and own client and any additional costs) and  any expenses of any nature whatsoever which African Alliance may suffer or incur as a result of my altering this order instruction in any mode other than writing.</td></tr>
<tr><td colspan="6"><br>&nbsp;<br>&nbsp;<br>&nbsp;<br>&nbsp;</td></tr>

<tr><td colspan="2">__________________________________________<br>&nbsp;Client's signature</td>
<td>&nbsp;</td>
<td colspan="2">__________________________________________<br>&nbsp;Date</td></tr>

<tr><td colspan="6"><br>&nbsp;<br>&nbsp;<br>&nbsp;</td></tr>

<tr><td colspan="6">
	<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#CCCCCC" width="100%">
	<tr><td colspan="6"><B><U>FOR OFFICIAL USE ONLY</U></B></td></tr>
	<tr>
		<td width="20%">Forward taken by: </td>
		<td>&nbsp;</td>
		<td width="20%">Forward approved by:</td>
		<td>&nbsp;</td>

	</tr>
	</table>
</td></tr>
</table>




<%
response.end
rs.movefirst%>

<table border="" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
<tr>
<td width="20%">&nbsp;</td>
<td width="80%">

<br>&nbsp;
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
	<tr>
		<td width="100%" colspan="2" height="30"><font face="Arial"><b>Client:</b></font></td>
	</tr>
	<tr>
		<td width="40%" height="15">&nbsp;<font face="Arial"><%=Rs("ClientName")%></font></td>
		<td width="60%" height="15">&nbsp;<font face="Arial">[<%=Rs("Client_DPA_")%>]</font></td>
	</tr>
	<tr>
		<td width="40%" height="15"><font face="Arial">&nbsp;<%=Replace(Rs("ClientAddr"),vbCrLf,"<br>")%></font></td>
		<td width="60%" height="15">&nbsp;</td>
	</tr>
	<tr>
		<td width="40%" height="15">&nbsp;<font face="Arial">Tel&nbsp;:&nbsp;<%=Rs("ClientOfficeTel")%></font></td>
		<td width="60%" height="15">&nbsp;</td>
	</tr>
	<%
	
	if Len(Rs("ClientAddr")) = 0 then
		%>
		<tr>
			<td width="100%" colspan="2" height="30">&nbsp;</td>
		</tr>
		<%
	end if
	%>
	<tr>
		<td width="100%" colspan="2" height="30"><font face="Arial"><b>Agent</b></font></td>
	</tr>
	<tr>
		<td width="40%" height="15">&nbsp;<font face="Arial"><%=Rs("AgentName")%></font></td>
		<td width="60%" height="15">&nbsp;<font face="Arial">[<%=Rs("Agent_DPA_")%>]</font></td>
	</tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
<tr>
	<td width="100%" height="85">&nbsp;</td>
</tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
	<tr>
		<td width="100%" colspan="5">&nbsp;</td>
	</tr>
	
	<tr>
		<td width="35%"><font face="Arial"><b>&nbsp;Security</b></font></td>
		<td align="left" width="12%"><font face="Arial"><b>&nbsp;Quantity</b></font></td>
		<td align="left" width="12%"><font face="Arial"><b>&nbsp;Price</b></font></td>
		<td align="left" width="12%"><font face="Arial"><b>&nbsp;Amount</b></font></td>
		<td width="12%"><font face="Arial"><b>&nbsp;Valid&nbsp;Until&nbsp;&nbsp;</b></font></td>		
		<td>&nbsp;</td>
	</tr>
	<%	
	end if
		if isNull(Rs("OrdDetailValidity")) then
			if Len(Rs("OrdDetailValidity")) <> 0 then
				Validity = FormatDate(Rs("OrdDetailValidity"))
			else
				Validity = "Until Execution"
			end if
		else
			Validity = "Until Execution"
		end if
	k=0
	
	do until Rs.eof
	if(i<5) then
	k=k+1
		%>
		
		<tr>
			<%
			if(Cint(Rs("OrderSecType_DPA_"))=1) then
			%>
			<td width="35%" height="20"><%=Rs("SecurityName")%>&nbsp;<%=Rs("BondDescription")%></td>
			<%
			else
			%>
			<td width="35%" height="20"><%=Rs("SecurityName")%>&nbsp;</td>
			<%
			end if
			%>
			<%if(Rs("Best")=True) then
			AmtDep=1
				if(Rs("OrdDetailQty")<>0 and Trim(rs("OrderTypeDescription"))<>"Purchase") then
				%>
				<td align="left" width="12%" height="20"><%=FormatNumEx(Rs("OrdDetailQty"),0)%>&nbsp;</td>
				<td align="left" width="12%" height="20">Best</td>			
				<td align="left" width="12%" height="20">&nbsp;</td>
				<%
				else
				%>
				<td align="left" width="12%" height="20">AMT DEP*</td>
				<td align="left" width="12%" height="20">Best</td>			
				<td align="left" width="12%" height="20"><%=FormatNum(Rs("Amount"))%>&nbsp;</td>
				<%
			end if			
			else
			%>
			<td align="left" width="12%" height="20"><%=FormatNumEx(Rs("OrdDetailQty"),0)%>&nbsp;</td>
			<td align="left" width="12%" height="20"><%=FormatNum(Rs("OrdDetailPrice"))%>&nbsp;</td>
			<td align="left" width="12%" height="20">&nbsp;</td>
			<% 
			end if
			%>			
			<td width="12%" height="20"><%=Validity%>&nbsp;</td>
			<td width="10%" height="20"><%=Rs("OrdDetailCertNo")%>&nbsp;</td>			
			<td>&nbsp;</td>
		</tr>
		
		<tr height="3">
			<td width="100%" colspan="6" height="20">&nbsp;</td>
		</tr>
		<%
		end if
		Rs.movenext

		i = i + 1			
		
	loop
		
	for j = 0 to (9-i)
		%>
		<tr>
			<td width="100%" colspan="6" height="20">&nbsp;</td>
		</tr>
		<%
	next
	
	if(Page1<>page2) then
	%>
	<tr>
		<td width="100%" colspan="6">&nbsp;</td>
	</tr>
</table>

<br>&nbsp;
<br>&nbsp;

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="101%">
	
	<tr>
		<td width="10%"><font face="Arial"><b>Date</b></font></td>
		<td width="40%" valign="bottom">______________________________</td>
		<td width="10%">&nbsp;</td>
		<td width="40%">&nbsp;</td>
	</tr>
	<tr>
	<td colspan="4" height="20">&nbsp;</td>
	</tr>
	<tr>
		<td width="10%"><font face="Arial"><b>Signature</b></font></td>
		<td width="40%" valign="bottom">______________________________</td>
		<td width="10%">&nbsp;</td>
		<td width="40%">&nbsp;</td>
		
	</tr>
	<tr>
		<td width="101%" colspan="4">&nbsp;</td>
	</tr>
	<tr>
		<td width="10%"><font face="Arial"><b>Customer Service</b></font></td>
		<td width="40%" valign="bottom">______________________________</td>
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
			dividends/bonus due to delay <br> in the execution and delivery of documents 
			to our office in good time</font></li>
		</ol>
		</td>
	</tr>
	<tr>
		<td width="101%" colspan="4">&nbsp;</td>
	</tr>
	</tr>
	<tr>
		<td width="101%" colspan="4">&nbsp;</td>
	</tr>	
</table>
</td>
</tr>
</table>
<%
end if
page2=page1
	if(Cint(Rs.recordCount)>5) then
	pageNumber=pageNumber + 1
	m=m+i
	i=0

    set Rs = conn.execute("SELECT TOP 5 * FROM  OrderForm WHERE (Order_DPA_ NOT IN (SELECT TOP " & m & " Order_DPA_" & _
                          " FROM   OrderForm)) and (Order_DPA_ = " & OrderID & ")")
	
	%>
			<BR class="newpage">
	<%	
	end if
	
	
	
	p=p+1
	loop
%>
</body>

</html>