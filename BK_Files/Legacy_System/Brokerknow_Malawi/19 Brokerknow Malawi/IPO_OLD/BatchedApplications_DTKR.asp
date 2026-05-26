<html>

<head>
<title>Batch Control Form (DTKR)</title>
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
FirstDay=DateSerial(Year(Date), Month(Date)-1 + iOffset, 1)

genReport = trim(Request("genReport"))
selectedBatch = trim(Request("cboBatch"))
Offering = trim(Request("Offering"))
%>

<% DrawPageFunctions True, True, True %>

<%
	Set conn = GetActiveConnection("KBroker")
	Set Rs = CreateObject("ADODB.Recordset")						        
	
	If Offering = "" Then Offering = 1
	
	sqlStr = "SELECT * FROM OfferingsList WHERE Offering = "& Offering &" AND Batch_No = " & selectedBatch & _
		" ORDER BY cast(floor(cast(Offerings_Date as float)) as datetime), Rtrim(ltrim(ClientName)) "
	
	Rs.CursorLocation = adUseClient	
	Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
	
	If rs.EOF Or rs.BOF Then%>
		<Script Language="JavaScript">
			alert("No batched applications were found for the selected batch.")
			window.location.href='BatchedApplications_DTKR.asp';
		</Script>
		<%Set Rs = Nothing
		Set Conn = Nothing
		Response.End
	End If
	
%>	
<br><br>
<table align=center border="0" cellspacing="0" cellpadding="1" width="70%">
	<tr>
		<td width="100%" colspan=5 nowrap align="left">
			<table border="1" cellspacing="5" cellpadding="5" width="100%" style="border: 0 solid black;">
				<tr>
					<td rowspan=2 width="30%" align=left><b><img src='image001.png' width='243' height='135'></b></td>
					<td width="70%" align=left style="font-size: 14pt;"><b>DIAMOND TRUST BANK KENYA LIMITED</b></td>
				</tr>
				<tr>
					<td width="70%" align=left style="font-size: 11pt;"><b>Rights Issue 2006</b></td>
				</tr>
			</table>
		</td>      
	</tr>
	<tr>
		<td width="100%" colspan=5 nowrap height="0" align="LEFT" style="border-top:3 solid black;font-size: 11pt;"><b>AGENT BATCH CONTROL FOR PAL <%If Instr(1,Rs("SecurityName"),"II")>0 Then Response.Write "II" Else Response.Write "I"%></b></td>      
	</tr>
	<tr>
		<td width="100%" colspan=5 nowrap align="left">&nbsp;</td>      
	</tr>
	<tr>
		<td width="15%" style="border:1 solid black;" nowrap align="left"><b>AUTHORIZED AGENT:</b></td>      
		<td width="20%" style="border:1 solid black;" nowrap align="left">AFRICAN ALLIANCE</td>      
		<td width="30%" height=50 style="border:1 solid black;border-bottom:0;" nowrap align="left" valign=top>Agent Stamp:</td>      
		<td width="35%" style="background-color:gainsboro;border:1 solid black;font-size:14pt;" nowrap align="center" colspan=2><b>PAL I BATCH NO</b></td>      
	</tr>
	<tr>
		<td width="15%" style="border:1 solid black;" nowrap align="left"><b>AGENT CODE:</b></td>      
		<td width="20%" style="border:1 solid black;" nowrap align="left">B23B</td>      
		<td width="30%" height=50 style="border:1 solid black;border-top:0;" nowrap align="left">&nbsp;</td>      
		<td width="35%" style="border:1 solid black;font-size:14pt;" nowrap align="center" colspan=2><b>23-<%If Instr(1,Rs("SecurityName"),"II")>0 Then Response.Write "2" Else Response.Write "1"%>000<%=Rs("Batch_No")%></b></td>          
	</tr>
	<tr>
		<td width="100%" colspan=5 nowrap align="left">&nbsp;</td>      
	</tr>
	<tr>
		<td width="100%" colspan=5 nowrap align="left">
		
			<table border="0" cellspacing="0" cellpadding="5" width="100%">
				<tr>
					<td width="10%" style="border:1 solid black;" align="center"><b>NO</b></td>
					<td width="10%" style="border:1 solid black;" align="left"><b>SHAREHOLDER NAME</b></td>
					<td width="10%" style="border:1 solid black;" align="center"><b>PAL NO</b></td>      
					<td width="10%" style="border:1 solid black;" align="center"><b>NO OF SHARES ACCEPTED</b></td>
					<td width="10%" style="border:1 solid black;" align="center"><b>NO OF ADDITIONAL SHARES</b></td>
					<td width="10%" style="border:1 solid black;" align="center"><b>TOTAL SHARES PAID FOR</b></td>   
					<td width="10%" style="border:1 solid black;" align="center"><b>PAYMENT AMOUNT</b></td>   
					<td width="10%" style="border:1 solid black;" align="center"><b>BANK AND BRANCH CODE</b></td>   
					<td width="10%" style="border:1 solid black;" align="center"><b>CHEQUE NO</b></td>   
				</tr>

				<%
				totalquantity = 0
				totalpayable = 0
				num = 1
				
				do until Rs.eof
						%>
						<tr>
							<td width="10%" style="border:1 solid black;" nowrap align="center"><%=num%></td>
							<td width="10%" style="border:1 solid black;" nowrap align="left"><%=Rs("Client_DPA_") & " - " & Mid(RS("ClientName"),1,30)%></td>		
							<td width="10%" style="border:1 solid black;" nowrap align="center"><%=Rs("Pal_No")%></td>
							<td width="10%" style="border:1 solid black;" nowrap align="center"><%=FormatNumEx(Rs("Alloted_Rights")-Rs("Additional"),0)%></td>
							<td width="10%" style="border:1 solid black;" nowrap align="center"><%=FormatNumEx(Rs("Additional"),0)%></td>
							<td width="10%" style="border:1 solid black;" nowrap align="center"><%=FormatNumEx(Rs("Alloted_Rights"),0)%></td>
							<td width="10%" style="border:1 solid black;" nowrap align="center"><%=FormatNum(Rs("Alloted_Rights")*Rs("Offering_Price"))%></td>
							<td width="10%" style="border:1 solid black;" nowrap align="center">&nbsp;</td>
							<td width="10%" style="border:1 solid black;" nowrap align="center">&nbsp;</td>
						</tr>
						<%
						totalquantity = totalquantity + Rs("Alloted_Rights")
						adtotalquantity = adtotalquantity + Rs("Additional")
						totalpayable = totalpayable + (Rs("Alloted_Rights")*Rs("Offering_Price"))
						num = num + 1
					Rs.MoveNext
				loop
				%>
				<tr>
					<td width="30%" style="border:1 solid black;" colspan="3"><b>TOTALS</b></td>
					<td width="10%" style="border:1 solid black;" align="center"><b><%=FormatNumEx(totalquantity-adtotalquantity,0)%></b></td>
					<td width="10%" style="border:1 solid black;" align="center"><b><%=FormatNumEx(adtotalquantity,0)%></b></td>
					<td width="10%" style="border:1 solid black;" align="center"><b><%=FormatNumEx(totalquantity,0)%></b></td>
					<td width="10%" style="border:1 solid black;" align="center"><b><%=FormatNum(totalpayable)%></b></td>
					<td width="10%" style="border:1 solid black;" align="center">&nbsp;</td>
					<td width="10%" style="border:1 solid black;" align="center">&nbsp;</td>
				</tr>
			</table>
  
		</td>      
	</tr>
	<tr>
		<td width="100%" colspan=9 nowrap align="left">
			<table border="0" cellspacing="5" cellpadding="5" width="100%">
				<tr>
					<td width="50%" align=center nowrap valign=top>
						<table border="0" cellspacing="5" cellpadding="5" width="100%" style="border: 1 solid black;">
							<tr>
								<td width="100%" height=30 align=left nowrap><b>COMPLETED BY AGENT REPRESENTATIVE:</b></td>
							</tr>
							<tr>
								<td width="100%" height=30 align=left nowrap><b>NAME: ______________________</b></td>
							</tr>
							<tr>
								<td width="100%" height=30 align=left nowrap><b>SIGNATURE: _________________</b></td>
							</tr>
							<tr>
								<td width="100%" height=30 align=left nowrap><b>DATE: ______________________</b></td>
							</tr>
						</table>
					</td>
				
					<td width="50%" align=center nowrap valign=top>
						<table border="0" cellspacing="5" cellpadding="5" width="100%" style="border: 1 solid black;">
							<tr>
								<td width="100%" height=30 align=left nowrap><b>RECEIVED BY:</b></td>
							</tr>
							<tr>
								<td width="100%" height=30 align=left nowrap><b>NAME: __________________________</b></td>
							</tr>
							<tr>
								<td width="100%" height=30 align=left nowrap><b>SIGNATURE: _____________________</b></td>
							</tr>
							<tr>
								<td width="100%" height=30 align=left nowrap><b>DATE: __________________________</b></td>
							</tr>
							<tr>
								<td style="border: 1 solid black;" width="100%" height=100 align=left nowrap valign=top>Rights Processing Unit Stamp</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>      
	</tr>
	<tr>
		<td width="100%" colspan=9 nowrap align="left">NOTE: To be completed in quadriplicate.</td>      
	</tr>
</table>
<%
Set Rs = Nothing
Set Conn = Nothing
%>   
</body>

</html>