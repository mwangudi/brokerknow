<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Delivery Slip</title>
 
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css">		
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
	<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
	<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
	<!--CALENDAR -->	
	<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
	<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
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
selectedTradeDate = Request.Form("txtDate")
timeLimit = Request.Form("timeLimit")

If genReport <> "1" Or Not IsDate(selectedTradeDate) Then%>
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
	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="ActualDeliveries.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">	
		<table>
			<tr>
				<td>Select delivery date:</td>
				<td>
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>
			</tr>
			
			<tr>
				<td colspan=2>&nbsp;</td>
			</tr>
			<tr style="display: none">
				<td>Select Transaction Type:</td>
				<td>
					<SELECT NAME="TransType">
						<OPTION VALUE=0 SELECTED>Sale</OPTION>
						<OPTION VALUE=1>Purchase</OPTION>		
					</SELECT>
				</td>
			</tr>
			<tr>
				<td colspan=2><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... ">&nbsp;&nbsp;</td>
			</tr>
		</table>
		
	</form>
	
	<%
	Response.End
End If

%>
<% DrawPageFunctions True, True, True %>
<%
   Dim conn 
   Dim sqlStr
   Dim rs
   Dim deliveriesFound
   
   Set conn = GetActiveConnection("KBroker")
	  deliveriesFound = false
   For i = 0 To 1
			transType = i 
				 
				  
			sqlStr = "SELECT * FROM [DeliveredContractList] WHERE OrderTypeSale = " & transType & " AND ContractDeliveryDate = '" & FormatDate(selectedTradeDate) & "'"
				       
		    Set Rs = CreateObject("ADODB.Recordset")						       
			Rs.CursorLocation = adUseClient	
			Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
		        
		    If Not (rs.EOF Or rs.BOF) Then
		       deliveriesFound = true        
		        
		    rs.MoveFirst
		        
		        
		%>


			<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
			     <tr>
					     <td>
					        <b><font face="Arial Narrow" size="4">Delivery Slip</font></b></td>
						 <td nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>			        
					</tr>	
			     <tr>
					     <td colspan=2>
					        <font face="Arial" size="2">to be produced in duplicate</font></td>
					</tr>
			    <tr>
					     <td colspan=2>
					        <font face="Arial" size="2">&nbsp;</font></td>
					</tr>
			</table>

			<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
			    <tr>
			      <td  style="border-top-style: solid; border-top-width: 1"><b><font face="Arial Narrow" size="2">TO:</font></b></td>
			      <td style="border-top-style: solid; border-top-width: 1"><b><font face="Arial Narrow" size="2">NAIROBI STOCK EXCHANGE</font></b></td>
			    </tr>

			    <tr>
			      <td  style="border-bottom-style: solid; border-bottom-width: 1">
			        <p align="left"><b><b><font face="Arial Narrow" size="2">DELIVERY DATE:</font></b></td>

			      <td style="border-bottom-style: solid; border-bottom-width: 1">
			        <p align="left"><font face="Arial" size="2">
						<%= FormatDate(Rs.Fields("ContractDeliveryDate").Value) %>
					</font></td>
			    </tr>

			  </table>


			<BR>


		  <table border="0" cellspacing="0" cellpadding="4" style="font-family: Arial Narrow; LEFT-MARGIN:100PX">
		    <tr>
		      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Slip</font></b></td>
		      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Traded</font></b></td>
		      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Security</font></b></td>
		      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Quantity</font></b></td>
		      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Price</font></b></td>
		       <%
		      If transType = 1 Then
				detailInfo = "Certificate No"
		      Else
				detailInfo = "Transfer No"
		      End If
		      %>
		      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow" size="3"><%= detailInfo %></font></b></td>
		    </tr>
			<%Do Until Rs.EOF
					    %>
					<tr>
							<td align="left"><%= Rs.Fields("LotSlipNo").Value %>	</td>

							<td><%= FormatDate(Rs.Fields("LotTDate").Value) %></td>
							<td><%= Rs.Fields("SecurityCode").Value %></td>
							<td align="right"><%= FormatNum(Rs.Fields("LotQty").Value) %></td>
							<td align="right"><%= FormatNum(Rs.Fields("LotPrice").Value) %></td>
							<%
							If transType = 1 Then
								detailInfo = Rs.Fields("ContractNCertificate").Value
							Else
								detailInfo = Rs.Fields("ContractTransferNo").Value
							End If
							%>
							<td><%= detailInfo %></td>
					</tr>
			
					<%
						Rs.MoveNext
			Loop
			%>
		  </table>
		  
		  <BR class="newpage">
		 <% 
				
			End If
	Next
	
	if deliveriesFound = false then%>
			<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
			     <tr>
					     <td>
					        <b><font face="Arial Narrow" size="4">No deliveries found for this date</font></b></td>
				</tr>						     
			</table>
	<%end if
Set Rs = Nothing
Set Conn = Nothing
%>


</body>
</html>
