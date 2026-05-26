<html>
<head>
<title>Batched Forwards</title>
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

If len(Request.QueryString) > 0 Then
	genReport = 1
	IDs = Trim(Request.QueryString("From"))
	ID = Split(IDs,"<->")
	selectedFrom = ID(0)
	offering = ID(1)
Else
	genReport = Request.Form("genReport")
	selectedFrom = Trim(Request.Form("txtFrom"))
	selectedTo = Trim(Request.Form("txtTo"))
	offering = Trim(Request.Form("cboOfferings"))

	If genReport <> "1" Or selectedFrom = "" Or selectedTo = "" Then%>
		<Script Language="JavaScript">
			report_SetBodyClass();
			function validateForm(frm){
				if (frm.cboOfferings[frm.cboOfferings.selectedIndex].SearchCode == ''){
					alert("Please select an Offering.");
					frm.cboOfferings.focus();
					return;
				}

				if (frm.txtFrom == ""){
					alert("Please specify a From Batch Number.");
					frm.txtFrom.focus();
					return;
				}

				if (frm.txtTo == ""){
					alert("Please specify a To Batch Number.");
					frm.txtTo.focus();
					return;
				}

				if (parseInt(frm.txtFrom.value) > parseInt(frm.txtTo.value)){
					alert("Invalid batch selection criteria.");
					frm.txtFrom.focus();
					return;
				}

				frm.target = '_self';
				frm.submit();

			}

		</Script>
		<form method="POST" action="BatchedForwards.asp" Name="frmMain" id="frmMain">
			<input type="hidden" value="1" name="genReport">
			<table>
				<tr>
					<td nowrap>Select an Offering&nbsp;</td>
					<td>
					<select name = 'cboOfferings' id = 'cboOfferings' size="1">
						<option selected SearchCode = '' value = ''></option>
						<%
						Set conn = GetActiveConnection("KBroker")
						sqlStr = "SELECT * FROM [SecurityListOfferings] WHERE Security_DPA_ = 157" & _
						" Order By DefaultSelection DESC, Security_DPA_ DESC"
						
						sqlStr = "SELECT DISTINCT  " & _
                                 "                       SecurityName, SecurityMktPrice, SecurityCode, SecurityAddr, Security_DPA_, BatchSize, OfferType, ParentSecurity_DPA_, DefaultSelection,  " & _
                                 "                       MinimumQty, Ratio, StepQty " & _
                                 " FROM         SecurityListOfferings " & _
                                 " ORDER BY DefaultSelection DESC, Security_DPA_ DESC"
						
						sqlStr = "SELECT SecurityCode,Security_DPA_,SecurityName,isnull(DefaultSelection,0) as DefaultSelection FROM [SecurityListOfferings] " & _
								 " Order By SecurityName ASC"
						
						Set rs = conn.Execute(SqlStr)

						If Not (rs.EOF Or rs.BOF) Then
							Do Until rs.EOF
								if(Cint(rs("DefaultSelection"))=1) then
								%>
								<option selected SearchCode = '<%=rs("SecurityCode")%>' value = '<%=rs("Security_DPA_")%>'><%=rs("SecurityName")%></option>
								<%
								else							
								%>
								<option SearchCode = '<%=rs("SecurityCode")%>' value = '<%=rs("Security_DPA_")%>'><%=rs("SecurityName")%></option>
								<%
								end if
							rs.MoveNext
							Loop
						End If%>
					</select>
					</td>
				</tr>
			</table>
			<table>
				<tr>
					<td nowrap>From Batch No:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
					<td><input type='text' id='txtFrom' name='txtFrom' size='10'></td>
					<td>&nbsp;</td>
					<td nowrap>To Batch No: </td>
					<td><input type='text' id='txtTo' name='txtTo' size='10'></td>
				</tr>
				<tr>
					<td colspan=2>Include Brokerknow Codes&nbsp;&nbsp;&nbsp;<input type="checkbox" name="brokerCodes" value="1"></td>
				</tr>
				<tr>
					<td colspan=2><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... " id=Button1 name=Button1>&nbsp;&nbsp;</td>
				</tr>
			</table>

		</form>

		<%Set rs = Nothing
		Set Conn = Nothing
		Response.End
	End If
End If

If len(Request.QueryString) > 0 Then
	DrawPageFunctions True, False, False
Else
	DrawPageFunctions True, True, True
End If
selectedBrokerCode = Trim (Request.Form("brokerCodes"))
'response.write selectedBrokerCode : response.end
If selectedBrokerCode ="" Then selectedBrokerCode="0"

Set conn = GetActiveConnection("KBroker")
Set Rs = CreateObject("ADODB.Recordset")

If Len(selectedFrom) > 0 And Len(selectedTo) > 0 Then
	sqlStr = "SELECT * FROM ForwardsList " & _
	" WHERE (Batch_No BETWEEN "& selectedFrom &" AND "& selectedTo &") AND (Offering = "& offering &")"
	'Response.Write sqlStr
	'Response.End
Else
	sqlStr = "SELECT * FROM ForwardsList where (Batch_No = "& selectedFrom &") AND (Offering = "& offering &")"
	
End If


Rs.CursorLocation = adUseClient
Rs.Open (sqlStr), conn.ConnectionString, adOpenKeyset, adLockOptimistic

If rs.EOF Or rs.BOF Then
	%>
	<Script Language="JavaScript">
		alert("The report did not find any values available.")
	</Script>
	<%Set Rs = Nothing
	Set Conn = Nothing
	Response.End
Else
	offeringName = Trim(Rs("SecurityName"))
	'paymentMode = Trim(Rs("BatchPaymentMode"))
End If

If len(Request.QueryString) > 0 Then
	'output single batch
	%>
	<table border="0" cellspacing="2" cellpadding="3" style="font-family: Arial Narrow" width="100%">
		<tr>
		   <table border ="0" width ="100%">
		        <tr>
			     <td><font face="Impact" size="3">BATCHED FORWARDS</font></td>
			      <td align="right" width="70%" nowrap align=right><font face="Impact" size="3"><%= UCase(Session("CompanyName")) %></font></td>
			    </tr>
				<tr><td align="center" width="100%" nowrap colspan='2'><font face="Impact" size="5"><%=UCase(offeringName)%></font></td></tr>
				<tr><td align="center" width="100%"colspan='2'><%= UCase(offeringName) & "&nbsp;BATCH CONTROL SCHEDULE" %></font></td></tr>
			 </table>
		</tr>
		<!--tr>
			<td align="center" width="100%" nowrap><font face="Impact" size="5"><%'=UCase(offeringName)%></font></td>
		</tr>
		<tr>
		  <td align="center" width="100%"><%'= UCase(offeringName) & "&nbsp;BATCH CONTROL SCHEDULE" %></font></td>
		</tr-->
	</table>
	</table>
	<br>
	<table border="0" cellspacing="0" cellpadding="3" style="font-family: Arial Narrow" width="100%">
		<tr>
		   <td width="5%">&nbsp;</td>
		   <td style="border-left:1px solid black;border-top:1px solid black;border-bottom:1px solid black" width="10%"><B>AGENT&nbsp;NAME</B></td>
		   <td style="border-left:1px solid black;border-top:1px solid black;border-bottom:1px solid black;border-right:1px solid black" width="28%"><%=UCase(Session("CompanyName"))%></td>
		   <td width="12%">&nbsp;</td>
		   <td width="13%">&nbsp;</td>
		   <td colspan="2" style="border-left:1px solid black;border-top:1px solid black;border-bottom:1px solid black" width="20%"><B>AGENT&nbsp;CODE</B></td>
		   <td style="border-left:1px solid black;border-top:1px solid black;border-bottom:1px solid black;border-right:1px solid black" width="12"><%="B" & Session("BrokerCode")%></td>
		</tr>
		<tr>
		   <td width="5%">&nbsp;</td>
		   <td style="border-left: 1px solid black; border-bottom: 1px solid black" width="10%"><B>DATE</B></td>
		   <td style="border-left: 1px solid black; border-bottom: 1px solid black;border-right:1px solid black" width="28%"><%=FormatDate(Now())%></td>
		   <td width="12%">&nbsp;</td>
		   <td width="13%">&nbsp;</td>
		   <td colspan="2" style="border-left: 1px solid black; border-bottom: 1px solid black" width="20%"><B>SCHEDULE&nbsp;NO:</B></td>
		   <td style="border-left: 1px solid black; border-bottom: 1px solid black;border-right:1px solid black" width="12%"><%=Session("BrokerCode") & Right("000000" & selectedFrom,6)%></td>
		</tr>

	</table>
	<BR>
	<table border="0" cellspacing="0" cellpadding="3"  style="font-family: Arial Narrow" width="100%">
		 <tr>
		   <td style="border-left: 1px solid black;border-top: 1px solid black;border-right: 1px solid black" width="5%"><B>ITEM</B></td>
		   <%
		   If selectedBrokerCode ="1" then
		   %>
				 <td style="border-right: 1px solid black;border-top: 1px solid black" width="10%"><B>Code&nbsp;</B></td>

		   <%
		   End if
		   %>
		   <td style="border-right: 1px solid black;border-top: 1px solid black" width="10%"><B>SERIAL&nbsp;NO</B></td>
		   <td style="border-right: 1px solid black;border-top: 1px solid black" width="28%"><B>NAME&nbsp;OF&nbsp;APPLICANT</B></td>
		   <td style="border-right: 1px solid black;border-top: 1px solid black" width="28%"><B>ID/PASSPORT NO.&nbsp;OF&nbsp;APPLICANT</B></td>
		   <td style="border-right: 1px solid black;border-top: 1px solid black" width="12%"><B>NO.&nbsp;OF&nbsp;SHARES</B></td>
		   <td style="border-right: 1px solid black;border-top: 1px solid black" width="13%"><B>AMOUNT&nbsp;PAYABLE</B></td>
		   <td style="border-right: 1px solid black;border-top: 1px solid black" width="10%"><B>BANK&nbsp;CODE</B></td>
		   <td style="border-right: 1px solid black;border-top: 1px solid black" width="10%"><B>CHEQUE NUMBER</B></td>
		   <td style="border-right: 1px solid black;border-top: 1px solid black" width="12%"><B>ACCOUNT NUMBER</B></td>
		</tr>
		<%
		totalquantity=0
		totalpayable=0
		num=1

		sqlStr = "SELECT * FROM ForwardsList where (Batch_No = "& selectedFrom &") AND (Offering = "& offering &")"
		set Rs = conn.execute(sqlStr)
		If Not(Rs.EOF Or Rs.BOF) Then
			Do Until Rs.EOF
			If Rs("Deleted") = True Then textStyle = "line-through" Else textStyle = "none"
				%>
				<tr style="text-decoration: <%=textStyle%>">
					<td style="border-left: 1px solid black;border-top: 1px solid black;border-right: 1px solid black" width="5%"><%=num%></td>

					<% 
					colSpans=1
		   If selectedBrokerCode ="1" Then
			colSpans=2
		   %>
				 <td style="border-right: 1px solid black;border-top: 1px solid black" width="10%"><%=Rs("Client_DPA_")%></td>
		   <%
		   End if
		   %>
					<td nowrap style="border-right: 1px solid black;border-top: 1px solid black" width="10%"><%=Rs("AgentCode") & " - " & Rs("Pal_No")%></td>
					<td style="border-right: 1px solid black;border-top: 1px solid black" width="28%"><%=RS("ClientName")%></td>
					<td style="border-right: 1px solid black;border-top: 1px solid black" width="12%" align="left"><%=Rs("ClientIDPass")%>&nbsp;</td>
					<td style="border-right: 1px solid black;border-top: 1px solid black" width="12%" align="left"><%=FormatNumEx(Rs("Alloted_Rights"),0)%>&nbsp;</td>
					<td style="border-right: 1px solid black;border-top: 1px solid black" width="13%" align="right"><%=FormatNum(Rs("Alloted_Rights")*Rs("Offering_Price"))+ Rs("CDSCharge") %>&nbsp;</td>
					<%If isNull(Rs("PaymentBankRef")) Then %><td style="border-right: 1px solid black;border-top: 1px solid black" width="10%" align="center"><%=("-")%><%Else%><td style="border-right: 1px solid black; border-top : 1px solid black" width="10%" align="left">&nbsp;<%=UCASE(Rs("PaymentBankRef"))%><%End If%></td>
					<td style="border-right: 1px solid black;border-top: 1px solid black" width="10%" align="left">&nbsp;<%=Rs("PaymentRef")%></td>
					<%If isNull(Rs("PaymentAccountNo")) Then %><td style="border-right: 1px solid black;border-top: 1px solid black" width="10%" align="center">&nbsp;<%=("-")%><%Else%><td style="border-right: 1px solid black; border-top: 1px solid black" width="10%" align="left">&nbsp;<%=Rs("PaymentAccountNo")%><%End If%></td>
				</tr>
				<%
				If Rs("Deleted") = True Then
					totalquantity = totalquantity 
					totalpayable = totalpayable  + Rs("CDSCharge")
					num=num
				Else
					totalquantity = totalquantity + Rs("Alloted_Rights")
					totalpayable = totalpayable + (Rs("Alloted_Rights")*Rs("Offering_Price")) + Rs("CDSCharge")
					num=num+1
				End If
				
				Rs.MoveNext
			loop
		End If
		%>
		<tr>
			<td style="border-left: 1px solid black;border-top: 1px solid black;border-right: 1px solid black;border-bottom: 1px solid black" width="5%">&nbsp</td>
			<td style="border-right: 1px solid black;border-top: 1px solid black;border-bottom: 1px solid black" width="10%">&nbsp</td>
			<td style="border-right: 1px solid black;border-top: 1px solid black;border-bottom: 1px solid black" width="10%">&nbsp</td>
			<td  colspan = <%=colSpan%> style="border-right: 1px solid black;border-top: 1px solid black;border-bottom: 1px solid black" width="28%"><b>TOTALS</b></td>
			<td style="border-right: 1px solid black;border-top: 1px solid black;border-bottom: 1px solid black" align="right" width="12%"><b><%=FormatNumEx(totalquantity,0)%></b>&nbsp;</td>
			<td style="border-right: 1px solid black;border-top: 1px solid black;border-bottom: 1px solid black" align="right" width="13%"><b><%=FormatNum(totalpayable)%></b>&nbsp;</td>
			<td style="border-right: 1px solid black;border-top: 1px solid black;border-bottom: 1px solid black" align="right" width="10%">&nbsp;</td>
			<td style="border-right: 1px solid black;border-top: 1px solid black;border-bottom: 1px solid black" align="right" width="10%">&nbsp;</td>
			<td style="border-right: 1px solid black;border-top: 1px solid black;border-bottom: 1px solid black" align="right" width="12%">&nbsp;</td>
		</tr>
	</table>
	<BR>
	<table border="0" cellspacing="0" cellpadding="3"  style="font-family: Arial Narrow" width="100%">
	<tr>
		<td width="5%">&nbsp;</td>
		<td colspan="2" style="border-style: solid; border-width: 1px; border-color: black" width="38%"><B>SIGNATURE OF AGENT</B></td>
		<td width="12%">&nbsp;</td>
		<td width="13%">&nbsp;</td>
		<td colspan="3" style="border-style: solid; border-width: 1px; border-color: black" width="32%"><B>STAMP OF AGENT</B></td>
	</tr>
	<tr>
		<td width="5%">&nbsp;</td>
		<td colspan="2" style="border-left-style: solid; border-left-width: 1; border-left-color: black; border-right-style: solid; border-right-width: 1; border-right-color: black" width="38%">&nbsp;</td>
		<td width="12%">&nbsp;</td>
		<td width="13%">&nbsp;</td>
		<td colspan="3" style="border-left-style: solid; border-left-width: 1; border-left-color: black; border-right-style: solid; border-right-width: 1; border-right-color: black" width="32%">&nbsp;</td>
	</tr>
	<tr>
		<td width="5%">&nbsp;</td>
		<td colspan="2" style="border-left-style: solid; border-left-width: 1; border-left-color: black; border-right-style: solid; border-right-width: 1; border-right-color: black" width="38%">&nbsp;</td>
		<td width="12%">&nbsp;</td>
		<td width="13%">&nbsp;</td>
		<td colspan="3" style="border-left-style: solid; border-left-width: 1; border-left-color: black; border-right-style: solid; border-right-width: 1; border-right-color: black" width="32%">&nbsp;</td>
	</tr>
	<tr>
		<td width="5%">&nbsp;</td>
		<td colspan="2" style="border-left-style: solid; border-left-width: 1; border-left-color: black; border-right-style: solid; border-right-width: 1; border-right-color: black; border-bottom-style: solid; border-bottom-width: 1; border-bottom-color: black" width="38%">&nbsp;</td>
		<td width="12%">&nbsp;</td>
		<td width="13%">&nbsp;</td>
		<td colspan="3" style="border-left-style: solid; border-left-width: 1; border-left-color: black; border-right-style: solid; border-right-width: 1; border-right-color: black; border-bottom-style: solid; border-bottom-width: 1; border-bottom-color: black" width="32%">&nbsp;</td>
	</tr>
	</table>
	<BR>
	<table border="0" cellspacing="0" cellpadding="3"  style="font-family: Arial Narrow" width="100%">
	<tr>
		<td width="5%">&nbsp;</td>
		<td colspan="2" style="border-left: 1px solid black;border-top: 1px solid black;border-right: 1px solid black;border-bottom: 1px solid black" width="38%"><B>FOR OFFICIAL USE ONLY</B></td>
		<td width="12%">&nbsp;</td>
		<td width="13%">&nbsp;</td>
		<td colspan="2" style="border-left: 1px solid black;border-top: 1px solid black;border-bottom: 1px solid black" width="20%"><B>PAYMENT&nbsp;MODE(TICK&nbsp;ONE&nbsp;ONLY)</B></td>
		<td style="border-top: 1px solid black; border-right: 1px solid black;border-bottom: 1px solid black" width="12%">&nbsp;</td>
	</tr>
	<tr>
		<td width="5%">&nbsp</td>
		<td style="border-left: 1px solid black;border-right: 1px solid black;border-bottom: 1px solid black" width="10%"><B>Delivery by:</B></td>
		<td style="border-right: 1px solid black;border-bottom: 1px solid black" width="28%">&nbsp;</td>
		<td width="12%">&nbsp;</td>
		<td width="13%">&nbsp;</td>
		<td colspan="2" style="border-left: 1px solid black;border-right: 1px solid black;border-bottom: 1px solid black" width="20%"><B>Banker's Cheque</B></td>
		<%if cint(paymentMode) = 3 then 'bankers chq %>
			<td style="border-right: 1px solid black;border-bottom: 1px solid black" width="25%"><input type="checkbox" name="chkPayment" id="chkPayment" style="border-style: none; color: white; background-color: white" disabled="true" checked></td>
		<%else%>
			<td style="border-right: 1px solid black;border-bottom: 1px solid black" width="25%">&nbsp;</td>
		<%end if%>
	</tr>
	<tr>
		<td width="5%">&nbsp</td>
		<td style="border-left: 1px solid black;border-right: 1px solid black;border-bottom: 1px solid black" width="10%"><B>Received&nbsp;by:</B></td>
		<td style="border-right: 1px solid black;border-bottom: 1px solid black" width="28%">&nbsp;</td>
		<td width="12%">&nbsp;</td>
		<td width="13%">&nbsp;</td>
		<td colspan="2" style="border-left: 1px solid black;border-right: 1px solid black;border-bottom: 1px solid black" width="20%"><B>Global Payment</B></td>
		<%if cint(paymentMode) = 1 then 'GPS %>
			<td style="border-right: 1px solid black;border-bottom: 1px solid black" width="25%"><input type="checkbox" name="chkPayment" id="chkPayment" style="border-style: none; color: white; background-color: white" disabled="true" checked></td>
		<%else%>
			<td style="border-right: 1px solid black;border-bottom: 1px solid black" width="25%">&nbsp;</td>
		<%end if%>
	</tr>
	<tr>
		<td width="5%">&nbsp</td>
		<td style="border-left: 1px solid black;border-right: 1px solid black;border-bottom: 1px solid black" width="10%"><B>Batch No:</B></td>
		<td style="border-right: 1px solid black;border-bottom: 1px solid black" width="28%"><%=Session("BrokerCode") & Right("000000" & selectedFrom,6)%></td>
		<td width="12%">&nbsp;</td>
		<td width="13%">&nbsp;</td>
		<td colspan="2" style="border-left: 1px solid black;border-right: 1px solid black;border-bottom: 1px solid black" width="20%"><B>EFT</B></td>
		<%if cint(paymentMode) = 2 then 'EFT %>
			<td style="border-right: 1px solid black;border-bottom: 1px solid black" width="25%"><input type="checkbox" name="chkPayment" id="chkPayment" style="border-style: none; color: white; background-color: white" disabled="true" checked></td>
		<%else%>
			<td style="border-right: 1px solid black;border-bottom: 1px solid black" width="25%">&nbsp;</td>
		<%end if%>
	</tr>
	</table>
	<BR>
	<table border="0" cellspacing="0" cellpadding="3"  style="font-family: Arial Narrow" width="100%">
	<!--<tr>
		<td colspan="6" width="100%">NOTE:&nbsp;&nbsp;&nbsp;To be completed in triplicate. Send all copies to the Receiving Agent's
		Processing Centre for stamping following which one stamped copy will be returned to the Agent.</td>
	</tr>-->
	</table>
<%Else
	'output several batches
	first = 1
	For i = selectedFrom To selectedTo
		If first = 0 then
				%><BR Class="newpage"><%
		end if

		sqlStr = "SELECT * FROM ForwardsList where (Batch_No = "& i &") AND (Offering = "& offering &")"
		set Rs = conn.execute(sqlStr)
		If Not(Rs.EOF Or Rs.BOF) Then
			'paymentMode = Trim(Rs("BatchPaymentMode"))%>
			<table border="0" cellspacing="0" cellpadding="3" style="font-family: Arial Narrow" width="100%">
			 <tr>
			   <table border ="0" width ="100%">
		        <tr>
			     <td nowrap><font face="Impact" size="3">BATCHED FORWARDS</font></td>
			      <td align="right" width="70%" nowrap align=right><font face="Impact" size="3"><%= UCase(Session("CompanyName")) %></font></td>
			    </tr>
				<tr><td align="center" width="100%" nowrap colspan='2'><font face="Impact" size="5"><%'=UCase(offeringName)%></font></td></tr>
				<tr><td align="center" width="100%"colspan='2'><%'= UCase(offeringName) & "&nbsp;BATCH CONTROL SCHEDULE" %></font></td></tr>
			   </table>
				</tr>
				<!--tr>
					<td align="center" width="100%" nowrap><font face="Impact" size="5"><%'=UCase(offeringName)%></font></td>
				</tr>
				<tr>
				  <td align="center" width="100%"><%'= UCase(offeringName) & "&nbsp;BATCH CONTROL SCHEDULE" %></font></td>
				</tr-->
			</table>
			<br>
			<table border="0" cellspacing="0" cellpadding="3" style="font-family: Arial Narrow" width="100%">
			    <tr>
				  <td width="5%">&nbsp;</td>
			      <td style="border-left:1px solid black;border-top:1px solid black;border-bottom:1px solid black" width="10%"><B>AGENT&nbsp;NAME</B></td>
			      <td style="border-left:1px solid black;border-top:1px solid black;border-bottom:1px solid black;border-right:1px solid black" width="28%"><%=UCase(Session("CompanyName"))%></td>
				  <td width="12%">&nbsp;</td>
				  <td width="13%">&nbsp;</td>
				  <td colspan="2" style="border-left:1px solid black;border-top:1px solid black;border-bottom:1px solid black" width="20%"><B>AGENT&nbsp;CODE</B></td>
			      <td style="border-left:1px solid black;border-top:1px solid black;border-bottom:1px solid black;border-right:1px solid black" width="12"><%="B" & Session("BrokerCode")%></td>
			   </tr>
			   <tr>
				  <td width="5%">&nbsp;</td>
			      <td style="border-left: 1px solid black; border-bottom: 1px solid black" width="10%"><B>DATE</B></td>
			      <td style="border-left: 1px solid black; border-bottom: 1px solid black;border-right:1px solid black" width="28%"><%=FormatDate(Now())%></td>
				  <td width="12%">&nbsp;</td>
				  <td width="13%">&nbsp;</td>
				  <td colspan="2" style="border-left: 1px solid black; border-bottom: 1px solid black" width="20%"><B>SCHEDULE&nbsp;NO:</B></td>
			      <td style="border-left: 1px solid black; border-bottom: 1px solid black;border-right:1px solid black" width="12%"><%=Session("BrokerCode") & Right("000000" & i,6)%></td>
			   </tr>

			</table>
			<BR>
			<table border="0" cellspacing="0" cellpadding="3"  style="font-family: Arial Narrow" width="100%">
				  <tr>
				   <td style="border-left: 1px solid black;border-top: 1px solid black;border-right: 1px solid black" width="5%"><B>ITEM</td>
				    <%
		   If selectedBrokerCode ="1" then
		   %>
				 <td style="border-right: 1px solid black;border-top: 1px solid black" width="10%"><B>Code&nbsp;</B></td>

		   <%
		   End if
		   %>
				   <td style="border-right: 1px solid black;border-top: 1px solid black" width="10%"><B>SERIAL&nbsp;NO</B></td>
				   <td style="border-right: 1px solid black;border-top: 1px solid black" width="28%"><B>NAME&nbsp;OF&nbsp;APPLICANT</B></td>
				   <td style="border-right: 1px solid black;border-top: 1px solid black" ><B>ID/PASSPORT NO.&nbsp;OF APPLICANT</B></td>
				   <td style="border-right: 1px solid black;border-top: 1px solid black" width="12%"><B>NO.&nbsp;OF&nbsp;SHARES</B></td>
				   <td style="border-right: 1px solid black;border-top: 1px solid black" width="13%"><B>AMOUNT&nbsp;PAYABLE</B></td>
				   <td style="border-right: 1px solid black;border-top: 1px solid black" width="10%"><B>BANK&nbsp;CODE</B></td>
				   <td style="border-right: 1px solid black;border-top: 1px solid black" width="10%"><B>CHEQUE NUMBER</B></td>
				   <td style="border-right: 1px solid black;border-top: 1px solid black" width="12%"><B>ACCOUNT NUMBER</B></td>
				</tr>
				<%
				totalquantity=0
				totalpayable=0
				num=1
				Do Until Rs.EOF
					If Rs("Deleted") = True Then textStyle = "line-through" Else textStyle = "none"
					%>
					<tr style="text-decoration: <%=textStyle%>">
						<td style="border-left: 1px solid black;border-top: 1px solid black;border-right: 1px solid black" width="5%"><%=num%></td>

						<% 
			colspan = 1
			If selectedBrokerCode ="1" Then
				colspan = 2
		   %>
				 <td style="border-right: 1px solid black;border-top: 1px solid black" width="10%"><%=Rs("Client_DPA_")%></td>

		   <%
		   End if
		   %>
						<td nowrap style="border-right: 1px solid black;border-top: 1px solid black" width="10%"><%=Rs("Pal_No")%></td>
						<td style="border-right: 1px solid black;border-top: 1px solid black" width="28%"><%=RS("ClientName")%></td>

						<td style="border-right: 1px solid black;border-top: 1px solid black" width="">&nbsp;<%=RS("ClientIDPASS")%></td>
						<td style="border-right: 1px solid black;border-top: 1px solid black" width="12%" align="right"><%=FormatNumEx(Rs("Alloted_Rights"),0)%>&nbsp;</td>
						<td style="border-right: 1px solid black;border-top: 1px solid black" width="13%" align="right"><%=FormatNum((Rs("Alloted_Rights")*Rs("Offering_Price")))%>&nbsp;</td>
						<%If isNull(Rs("PaymentBankRef")) Then %>
						<td style="border-right: 1px solid black;border-top: 1px solid black" width="10%" align="center"><%=("-")%>
						<%Else%>
						<td style="border-right: 1px solid black;border-top: 1px solid black" width="10%" align="left">&nbsp;
						<%=UCASE(Rs("PaymentBankRef"))%>
						<%End If%></td>
						<td style="border-right: 1px solid black;border-top: 1px solid black" width="10%" align="left"><%=Rs("PaymentRef")%></td>
						<%If isNull(Rs("PaymentAccountNo")) Then %><td style="border-right: 1px solid black;border-top: 1px solid black" width="10%" align="center">&nbsp;<%=("-")%><%Else%><td style="border-right: 1px solid black;border-top: 1px solid black" width="10%" align="left">&nbsp;<%=Rs("PaymentAccountNo")%><%End If%></td>
					</tr>
					<%
					If Rs("Deleted") = True Then
						totalquantity = totalquantity 
						totalpayable = totalpayable 
						num=num
					Else
						totalquantity = totalquantity + Rs("Alloted_Rights")
						totalpayable = totalpayable + (Rs("Alloted_Rights")*Rs("Offering_Price")) 
						num=num+1
					End If
					Rs.MoveNext

					'If Rs.EOF Then Exit For
				loop
				%>
				<tr>
					<td style="border-left: 1px solid black;border-top: 1px solid black;border-right: 1px solid black;border-bottom: 1px solid black" width="5%">&nbsp</td>
					<td style="border-right: 1px solid black;border-top: 1px solid black;border-bottom: 1px solid black" width="10%">&nbsp</td>
					<td style="border-left: 1px solid black;border-top: 1px solid black;border-right: 1px solid black;border-bottom: 1px solid black" width="5%">&nbsp</td>
					<td colspan = <%=colSpan%> style="border-right: 1px solid black;border-top: 1px solid black;border-bottom: 1px solid black" ><b>TOTALS</b></td>
					<td style="border-right: 1px solid black;border-top: 1px solid black;border-bottom: 1px solid black" align="right" width="12%"><b><%=FormatNumEx(totalquantity,0)%></b>&nbsp;</td>
					<td style="border-right: 1px solid black;border-top: 1px solid black;border-bottom: 1px solid black" align="right" width="13%"><b><%=FormatNum(totalpayable)%></b>&nbsp;</td>
					<td style="border-right: 1px solid black;border-top: 1px solid black;border-bottom: 1px solid black" align="right" width="10%">&nbsp;</td>
					<td style="border-right: 1px solid black;border-top: 1px solid black;border-bottom: 1px solid black" align="right" width="10%">&nbsp;</td>
					<td style="border-right: 1px solid black;border-top: 1px solid black;border-bottom: 1px solid black" align="right" width="12%">&nbsp;</td>
				</tr>
			</table>
			<BR>
			<table border="0" cellspacing="0" cellpadding="3"  style="font-family: Arial Narrow" width="100%">
			<tr>
				<td width="5%">&nbsp;</td>
				<td colspan="2" style="border-style: solid; border-width: 1px; border-color: black" width="38%"><B>SIGNATURE OF AGENT</B></td>
				<td width="12%">&nbsp;</td>
				<td width="13%">&nbsp;</td>
				<td colspan="3" style="border-style: solid; border-width: 1px; border-color: black" width="32%"><B>STAMP OF AGENT</B></td>
			</tr>
			<tr>
				<td width="5%">&nbsp;</td>
				<td colspan="2" style="border-left-style: solid; border-left-width: 1; border-left-color: black; border-right-style: solid; border-right-width: 1; border-right-color: black" width="38%">&nbsp;</td>
				<td width="12%">&nbsp;</td>
				<td width="13%">&nbsp;</td>
				<td colspan="3" style="border-left-style: solid; border-left-width: 1; border-left-color: black; border-right-style: solid; border-right-width: 1; border-right-color: black" width="32%">&nbsp;</td>
			</tr>
			<tr>
				<td width="5%">&nbsp;</td>
				<td colspan="2" style="border-left-style: solid; border-left-width: 1; border-left-color: black; border-right-style: solid; border-right-width: 1; border-right-color: black" width="38%">&nbsp;</td>
				<td width="12%">&nbsp;</td>
				<td width="13%">&nbsp;</td>
				<td colspan="3" style="border-left-style: solid; border-left-width: 1; border-left-color: black; border-right-style: solid; border-right-width: 1; border-right-color: black" width="32%">&nbsp;</td>
			</tr>
			<tr>
				<td width="5%">&nbsp;</td>
				<td colspan="2" style="border-left-style: solid; border-left-width: 1; border-left-color: black; border-right-style: solid; border-right-width: 1; border-right-color: black; border-bottom-style: solid; border-bottom-width: 1; border-bottom-color: black" width="38%">&nbsp;</td>
				<td width="12%">&nbsp;</td>
				<td width="13%">&nbsp;</td>
				<td colspan="3" style="border-left-style: solid; border-left-width: 1; border-left-color: black; border-right-style: solid; border-right-width: 1; border-right-color: black; border-bottom-style: solid; border-bottom-width: 1; border-bottom-color: black" width="32%">&nbsp;</td>
			</tr>
			</table>
			<BR>
			<table border="0" cellspacing="0" cellpadding="3"  style="font-family: Arial Narrow" width="100%">
			<tr>
				<td width="5%">&nbsp;</td>
				<td colspan="2" style="border-left: 1px solid black;border-top: 1px solid black;border-right: 1px solid black;border-bottom: 1px solid black" width="38%"><B>FOR OFFICIAL USE ONLY</B></td>
				<td width="12%">&nbsp;</td>
				<td width="13%">&nbsp;</td>
				<td colspan="2" style="border-left: 1px solid black;border-top: 1px solid black;border-bottom: 1px solid black" width="20%"><B>PAYMENT&nbsp;MODE(TICK&nbsp;ONE&nbsp;ONLY)</B></td>
				<td style="border-top: 1px solid black; border-right: 1px solid black;border-bottom: 1px solid black" width="12%">&nbsp;</td>
			</tr>
			<tr>
				<td width="5%">&nbsp</td>
				<td style="border-left: 1px solid black;border-right: 1px solid black;border-bottom: 1px solid black" width="10%"><B>Delivery by:</B></td>
				<td style="border-right: 1px solid black;border-bottom: 1px solid black" width="28%">&nbsp;</td>
				<td width="12%">&nbsp;</td>
				<td width="13%">&nbsp;</td>
				<td colspan="2" style="border-left: 1px solid black;border-right: 1px solid black;border-bottom: 1px solid black" width="20%"><B>Banker's Cheque</B></td>
				<%if cint(paymentMode) = 1 then 'bankers chq %>
					<td style="border-right: 1px solid black;border-bottom: 1px solid black" width="25%"><input type="checkbox" name="chkPayment" id="chkPayment" style="border-style: none; color: white; background-color: white" disabled="true" checked></td>
				<%else%>
					<td style="border-right: 1px solid black;border-bottom: 1px solid black" width="25%">&nbsp;</td>
				<%end if%>
			</tr>
			<tr>
				<td width="5%">&nbsp</td>
				<td style="border-left: 1px solid black;border-right: 1px solid black;border-bottom: 1px solid black" width="10%"><B>Received&nbsp;by:</B></td>
				<td style="border-right: 1px solid black;border-bottom: 1px solid black" width="28%">&nbsp;</td>
				<td width="12%">&nbsp;</td>
				<td width="13%">&nbsp;</td>
				<td colspan="2" style="border-left: 1px solid black;border-right: 1px solid black;border-bottom: 1px solid black" width="20%"><B>Global Payment</B></td>
				<%if cint(paymentMode) = 1 then 'GPS %>
					<td style="border-right: 1px solid black;border-bottom: 1px solid black" width="25%">&nbsp;<!--<input type="checkbox" name="chkPayment" id="chkPayment" style="border-style: none; color: white; background-color: white" disabled="true" checked>--></td>
				<%else%>
					<td style="border-right: 1px solid black;border-bottom: 1px solid black" width="25%">&nbsp;</td>
				<%end if%>
			</tr>
			<tr>
				<td width="5%">&nbsp</td>
				<td style="border-left: 1px solid black;border-right: 1px solid black;border-bottom: 1px solid black" width="10%"><B>Batch No:</B></td>
				<td style="border-right: 1px solid black;border-bottom: 1px solid black" width="28%"><%=Session("BrokerCode") & Right("000000" & i,6)%></td>
				<td width="12%">&nbsp;</td>
				<td width="13%">&nbsp;</td>
				<td colspan="2" style="border-left: 1px solid black;border-right: 1px solid black;border-bottom: 1px solid black" width="20%"><B>EFT</B></td>
				<%if cint(paymentMode) = 2 then 'EFT %>
					<td style="border-right: 1px solid black;border-bottom: 1px solid black" width="25%"><input type="checkbox" name="chkPayment" id="chkPayment" style="border-style: none; color: white; background-color: white" disabled="true" checked></td>
				<%else%>
					<td style="border-right: 1px solid black;border-bottom: 1px solid black" width="25%">&nbsp;</td>
				<%end if%>
			</tr>
			</table>
			<BR>
			<table border="0" cellspacing="0" cellpadding="3"  style="font-family: Arial Narrow" width="100%">
				<!--<tr>
					<td colspan="6" width="100%">NOTE:&nbsp;&nbsp;&nbsp;To be completed in triplicate. Send all copies to the Receiving Agent's
					Processing Centre for stamping following which one stamped copy will be returned to the Agent.</td>
				</tr>-->
			</table>
		<%End If
		first = 0
	Next
End If
Set Rs = Nothing
Set Conn = Nothing%>
</body>

</html>

