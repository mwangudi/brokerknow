<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Profit and Loss Account</title>

	<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
	<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
	<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
	<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
	<SCRIPT language=VBScript src="../scripts/reports.vbs"></SCRIPT>
	<SCRIPT language=Javascript src="../scripts/reports.js"></SCRIPT>
	<SCRIPT language=Javascript src="../scripts/accountList.js"></SCRIPT>



	<style media="print">
		@page {
			@top{font-family: Helvetica, Arial, sans-serif;
				font-size: 150%;
				font-weight: bolder;
				text-align: left;
				content: "<%= FormatDate(Date) %>";			
			}
			
			margin-left: 0cm;
			margin-right: 0cm;
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
'selectedBank = Request.Form("cboAccount")
selectedFromDate = Request.Form("transFromDate")
selectedToDate = Request.Form("txtToDate")
SelectedType=Request.Form("cboEntity")
FirstDay=DateSerial(Year(Date), Month(Date)-1 + iOffset, 1)
thistype=Request.Form("Selectedtype")

If genReport <> "1" Or Not IsDate(selectedFromDate) Then%>
	<Script Language="JavaScript">
		report_SetBodyClass();
		
		function validateForm(frm)
		{						
			frm.target = '_self';			
			frm.submit();
		}
		
		
		function evaluateEntity(Val, Entity)
		{      	
	  	FetchAccounts1(Entity)
		}
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "transFromDate","cmdDate","<%= FormatDate(FirstDay) %>",1);
		var cal1=new ctlSpiffyCalendarBox("cal1", "frmMain", "txtToDate","cmdDate","<%= FormatDate(Date) %>",1);

	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="ProfitAndLoss.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">
		<% currentEntityType=5 %>
		<table>
			<tr>
				<td colspan="2">Select date from:</td>
				<td>
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>				
			</tr>
			<tr>
				<td colspan="2">To date:</td>
				<td>
					<SCRIPT language="JavaScript">cal1.writeControl();</SCRIPT>	
				</td>
				
			</tr>			

			<tr>
				<td colspan="3"><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... ">&nbsp;&nbsp;</td>
			</tr>
		</table>
		
	</form>
	
	<%
	'Set rs = Nothing
	'Set Conn = Nothing
	Response.End
End If

%>
<% DrawPageFunctions True, True, True %>
<%
	Set conn = GetActiveConnection("KBroker")
	Set Rs = CreateObject("ADODB.Recordset")	
			 


sqlStr = "PnLProc '" & FormatDate(selectedFromDate) & "',' " & FormatDate(selectedToDate) & "' "

'Response.write(sqlStr)
'Response.end

	''Rs.CursorLocation = adUseClient	
	Rs.Open SQLServerFormat((sqlStr)), conn.ConnectionString, 0, 1

	'Rs.Filter = "Client_DPA_ LIKE '" & selectedClient & "' AND TransDate >= '" & FormatDate(selectedFromDate) & "'"	

	''Set Rs = Conn.Execute(sqlStr)

	'Conn.Execute("DROP TABLE #tableA")
	'Conn.Execute("DROP TABLE #tableB")
	
	'Set Rs=Conn.Execute(sqlStr)

	Dim i 
	i=0

	'for i = 1 to 9
	'	 Rs =  Rs.NextRecordset 
	 'next
	
	If Rs.EOF Or Rs.BOF Then%>
			<script Language="JavaScript">
				alert("There are no transactions in the system")
				window.history.go(-1);
			</script>
			<%Set Rs = Nothing
			Set Conn = Nothing
			Response.End
	End If%>

	<table border="0" cellspacing="2" cellpadding="2" style="font-family: Arial Narrow" width="100%">
		<tr>
		<td width="10%" nowrap><font face="Impact" size="4">PROFIT and LOSS
		    STATEMENT</font></td>
		<td width="60%" nowrap align="right"><font face="Impact" size="3"><%= Session("CompanyName") %>
		    </font></td>
		</tr>
	</table>
	<br>
	<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
		<tr>
		<td width="1%"><b>Date Between: </b></td>
		<td width="48%"><%= FormatDate(selectedFromDate) %>&nbsp;&nbsp;<%= FormatDate(selectedToDate) %>
		</td>
		</tr>
	</table>
	<br>
	<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow; LEFT-MARGIN:100PX" width="100%">
		<tr>
		<td style="border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" colspan="2"><b><font face="Arial Narrow" size="3">Income</font></b></td>
		<td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" align="right"><b><font face="Arial Narrow" size="3">&nbsp;</font></b></td>
		<td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" align="right"><b><font face="Arial Narrow" size="3">&nbsp;</font></b></td>
		</tr>
		<%	
	
    Do Until Rs.EOF 
		%>
		<tr>
		  <td colspan="2"><%= Rs.Fields("AccountName").Value %>
		  </td>
		  <td align="right">&nbsp;
		  </td>
		  <td align="right"><%=FormatNum(Rs.Fields("Balance").Value)%>
		  </td>
		</tr>
		<%	
				If trim(Rs.Fields("SubTotal").Value) <> "" Then%>
					<tr>
						<td colspan="2"><b>Sub Total</b>
						</td>
						<td align="right">&nbsp;
						</td>
						<td align="right"><b><%=FormatNum(Rs.Fields("SubTotal").Value)%></b></td>
					</tr>				
				<%
				'Response.write "here"
				exit do
				End If   	
		Rs.MoveNext
		
	Loop%>
	
	<tr>
		<td style="border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" colspan="2"><b><font face="Arial Narrow" size="3">Expenditure</font></b></td>
		<td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" align="right"><b><font face="Arial Narrow" size="3">&nbsp;</font></b></td>
		<td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" align="right"><b><font face="Arial Narrow" size="3">&nbsp;</font></b></td>
		</tr>
		<%
	
	Dim pnlSummary
	
	if Not(Rs.Eof or Rs.Bof) then
		If trim(Rs.Fields("GrandTotal").Value) = "" Then
				Rs.MoveNext
				Do Until Rs.EOF 
					%>
					<tr>
					  <td colspan="2"><%= Rs.Fields("AccountName").Value %>
					  </td>
					  <td align="right"><%=FormatNum(Rs.Fields("Balance").Value)%>
					  </td>
					  <td align="right">&nbsp;
					  </td>
					</tr>
					<%	
							If trim(Rs.Fields("SubTotal").Value) <> "" Then%>
								<tr>
									<td colspan="2"><b>Sub Total</b>
									</td>
									<td align="right"><b><%=FormatNum(Rs.Fields("SubTotal").Value)%></b>
									</td>
									<td align="right">&nbsp;
									</td>
								</tr>	
											
							<%
							exit do
							End If   	
					Rs.MoveNext
				Loop
		end if%>
	<tr>
		<%
		Dim summaryCaption
		pnlSummary = Rs.Fields("GrandTotal").Value
		
		if pnlSummary < 0 Then 
				summaryCaption = "Net Loss"
		else
				summaryCaption = "Net Profit"
		end if%>
		<td colspan="2"><b><%=ucase(summaryCaption)%></b>
		</td>
		<td align="right">&nbsp;
		</td>
		<td align="right"><b><%=FormatNum(pnlSummary)%></b>
		</td>
	</tr>
	</table>
	
	<%
	end if
Set Rs = Nothing
Set Conn = Nothing%>

</body>

</html>
