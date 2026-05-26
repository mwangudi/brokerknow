<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Batched applications</title>
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

If genReport <> "1" Or selectedBatch = ""  Then%>
	<Script Language="JavaScript">
		report_SetBodyClass();
		function validateForm(frm){			
			if (frm.cboBatch.selectedIndex < 0){
				alert("Select a batch.");
				frm.cboBatch.focus();
				return;
			}
			
			frm.target = '_self';			
			frm.submit();
		}
	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="BatchedApplications.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">
		<table>
			<tr>
				<td nowrap>Batch No: </td>
				<td><select name = 'cboBatch' id = 'cboBatch' size="1">					
					<%
					        Set conn = GetActiveConnection("KBroker")
					        
					        sqlStr = "SELECT distinct Batch_No FROM Offerings where Batch_NO is not null and Forward <>1 and deleted=0 ORDER BY Batch_No desc"
					        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
					        If Not (rs.EOF Or rs.BOF) Then
					                rs.MoveFirst
					                Do Until rs.EOF%>
					                        <option value = '<%=rs.Fields("Batch_No")%>'><%=rs.Fields("Batch_No")  %></option>
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

<% DrawPageFunctions True, True, True %>

<%
	Offering = Request("Offering")
	If Offering = "" Then Offering = 1
	
	Set conn = GetActiveConnection("KBroker")
	Set Rs = CreateObject("ADODB.Recordset")						        
	sqlStr = "SELECT * FROM OfferingsList where Offering = "& Offering &" and Batch_No = " & selectedBatch & _
	         " order by cast(floor(cast(Offerings_Date as float)) as datetime), Rtrim(ltrim(ClientName)) "
	
	Rs.CursorLocation = adUseClient	
	Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
	
	If rs.EOF Or rs.BOF Then%>
		<Script Language="JavaScript">
			alert("No batched applications were found for the selected batch.")
			window.location.href='BatchedApplications.asp';
		</Script>
		<%Set Rs = Nothing
		Set Conn = Nothing
		Response.End
	End If
	
%>	
<br><br>
<table align=center border="0" cellspacing="0" cellpadding="5" width="90%">
	<tr>
		<td width="100%" colspan=5 nowrap align="left">BATCHED APPLICATIONS</td>      
	</tr>
	<tr>
		<td width="100%" colspan=5 nowrap height="0" align="LEFT" style="border-top:3 solid black;font-size: 11pt;"><b>MUMIAS SUGAR CO LTD O.F.S BATCH CONTROL SCHEDULE</b></td>      
	</tr>
	<tr>
		<td width="100%" colspan=5 nowrap align="left">&nbsp;</td>      
	</tr>
	<tr>
		<td width="15%" style="border:1 solid black;" nowrap align="left"><b>AGENT NAME</b></td>      
		<td width="20%" style="border:1 solid black;" nowrap align="left">AFRICAN ALLIANCE</td>      
		<td width="30%" nowrap align="left">&nbsp;</td>      
		<td width="15%" style="border:1 solid black;" nowrap align="left"><b>AGENT CODE</b></td>      
		<td width="20%" style="border:1 solid black;" nowrap align="left">B23B</td>      
	</tr>
	<tr>
		<td width="15%" style="border:1 solid black;" nowrap align="left"><b>DATE</b></td>      
		<td width="20%" style="border:1 solid black;" nowrap align="left"><%= FormatDate(Date) %></td>      
		<td width="30%" nowrap align="left">&nbsp;</td>      
		<td width="15%" style="border:1 solid black;" nowrap align="left"><b>SCHEDULE NO.</b></td>      
		<td width="20%" style="border:1 solid black;" nowrap align="left">&nbsp;</td>      
	</tr>
	<tr>
		<td width="100%" colspan=5 nowrap align="left">&nbsp;</td>      
	</tr>
	<tr>
		<td width="100%" colspan=5 nowrap align="left">
		
			<table border="0" cellspacing="0" cellpadding="5" width="100%">
				<tr>
					<td width="10%" style="border:1 solid black;" nowrap align="center"><b>ITEM</b></td>
					<td width="10%" style="border:1 solid black;" nowrap align="center"><b>SERIAL NO</b></td>
					<td width="50%" style="border:1 solid black;" nowrap align="center"><b>NAME OF APPLICANT</b></td>      
					<td width="10%" style="border:1 solid black;" nowrap align="center"><b>NO. OF SHARES</b></td>
					<td width="10%" style="border:1 solid black;" nowrap align="center"><b>AMOUNT PAYABLE</b></td>   
					<td width="10%" style="border:1 solid black;" nowrap align="center"><b>CHEQUE NUMBER</b></td>
				</tr>

				<%
				totalquantity = 0
				totalpayable = 0
				num = 1
				
				do until Rs.eof
						%>
						<tr>
							<td width="10%" style="border:1 solid black;" nowrap align="center"><%=num%></td>
							<td width="10%" style="border:1 solid black;" nowrap align="center"><%=Rs("Pal_No")%></td>
							<td width="50%" style="border:1 solid black;" nowrap align="center"><%=Rs("Client_DPA_") & " - " & Mid(RS("ClientName"),1,30)%></td>		
							<td width="10%" style="border:1 solid black;" nowrap align="center"><%=FormatNumEx(Rs("Alloted_Rights"),0)%></td>
							<td width="10%" style="border:1 solid black;" nowrap align="center"><%=FormatNum(Rs("Alloted_Rights")*Rs("Offering_Price"))%></td>
							<td width="10%" style="border:1 solid black;" nowrap align="center">-</td>
						</tr>
						<%
						totalquantity = totalquantity + Rs("Alloted_Rights")
						totalpayable = totalpayable + (Rs("Alloted_Rights")*Rs("Offering_Price"))
						num = num + 1
					Rs.MoveNext
				loop
				%>
				<tr>
					<td width="70%" style="border:1 solid black;" colspan="3"><b>TOTALS</b></td>
					<td width="10%" style="border:1 solid black;" align="center"><b><%=FormatNumEx(totalquantity,0)%></b></td>
					<td width="10%" style="border:1 solid black;" align="center"><b><%=FormatNum(totalpayable)%></b></td>
					<td width="10%" style="border:1 solid black;" align="center">&nbsp;</td>
				</tr>
			</table>
  
		</td>      
	</tr>
	<tr>
		<td width="100%" colspan=5 nowrap align="left">
			<table border="0" cellspacing="0" cellpadding="2" width="100%">
				<tr>
					<td width="35%" height="100" valign=top align=center style="border: 1 solid black;"><b>SIGNATURE OF AGENT</b></td>
					<td width="30%">&nbsp;</td>
					<td width="35%" height="100" valign=top align=center style="border: 1 solid black;"><b>STAMP OF AGENT</b></td>
				</tr>
			</table>
		</td>      
	</tr>
	<tr>
		<td width="15%" nowrap align="left">&nbsp;</td>      
		<td width="20%" nowrap align="left">&nbsp;</td>      
		<td width="30%" nowrap align="left">&nbsp;</td>      
		<td width="15%" style="border:1 solid black;" nowrap align="left" colspan=2><b>PAYMENT MODE</b></td>      
	</tr>
	<tr>
		<td width="15%" style="border:1 solid black;" nowrap align="left"><b>Delivery by:</b></td>      
		<td width="20%" style="border:1 solid black;" nowrap align="left">&nbsp;</td>      
		<td width="30%" nowrap align="left">&nbsp;</td>      
		<td width="15%" style="border:1 solid black;" nowrap align="left"><b>Bankers' Cheque</b></td>      
		<td width="20%" style="border:1 solid black;" nowrap align="left">&nbsp;</td>      
	</tr>
	<tr>
		<td width="15%" style="border:1 solid black;" nowrap align="left"><b>Received by:</b></td>      
		<td width="20%" style="border:1 solid black;" nowrap align="left">&nbsp;</td>      
		<td width="30%" nowrap align="left">&nbsp;</td>      
		<td width="15%" style="border:1 solid black;" nowrap align="left"><b>Global Payment System</b></td>      
		<td width="20%" style="border:1 solid black;" nowrap align="left"><%="YES"%></td>      
	</tr>
	<tr>
		<td width="15%" style="border:1 solid black;" nowrap align="left"><b>Batch No.:</b></td>      
		<td width="20%" style="border:1 solid black;" nowrap align="left"><%=selectedBatch%></td>      
		<td width="30%" nowrap align="left">&nbsp;</td>      
		<td width="15%" style="border:1 solid black;" nowrap align="left"><b>EFT</b></td>      
		<td width="20%" style="border:1 solid black;" nowrap align="left">&nbsp;</td>      
	</tr>
	<tr>
		<td width="100%" colspan=5 nowrap align="left">NOTE: To be completed in triplicate and retain the third copy.</td>      
	</tr>
</table>

  
   
<%
Set Rs = Nothing
Set Conn = Nothing
%>   
</body>

</html>