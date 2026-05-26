<html>
<%
Offering = Request("cboOfferings")
%>
<head>
<title>Batched Forwards By Date</title>
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

genReport = Request.Form("genReport")
selectedFromDate = Request.Form("transFromDate")
selectedToDate = Request.Form("transToDate")

If genReport <> "1" Or Not IsDate(selectedFromDate) Then%>
	<Script Language="JavaScript">
		report_SetBodyClass();
		function validateForm(frm){						
			
			frm.target = '_self';			
			frm.submit();
		}
		
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "transFromDate","cmdDate","<%= FormatDate(FirstDay) %>",1);
		var cal1=new ctlSpiffyCalendarBox("cal1", "frmMain", "transToDate","cmdDate","<%= FormatDate(Date) %>",1);

	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="BatchedForwardDate.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">
		<table>			
			<tr>
			<td width="20%" nowrap>Offering Name</td>
			<td width="80%" nowrap>
				<select name = 'cboOfferings' id = 'cboOfferings' size="1">
				<% 
				Set conn = GetActiveConnection("KBroker")
				
				sqlStr = "SELECT * FROM [SecurityListOfferings] " & _
				" WHERE cast(floor(cast(ClosingDate as float)) as datetime) >= cast(floor(cast(GetDate() as float)) as datetime)" & _
				" Order By SecurityName ASC"
				Set rs = conn.Execute(SqlStr)
				
				If Not (rs.EOF Or rs.BOF) Then
					Do Until rs.EOF
							if Offering = "" then
								if trim(rs.Fields("DefaultSelection")) = 1 Then
									%>
									<option selected ParentSecurity = "<%=Rs("ParentSecurity_DPA_")%>" Ratio = "<%=Rs("Ratio")%>" OfferType = "<%=Rs("OfferType_DPA_")%>" SearchPrice = "<%=rs.Fields("SecurityMktPrice")%>" value = '<%=rs.Fields("Security_DPA_")%>'><%=rs.Fields("SecurityName")%></option>
									<%
									price = rs("SecurityMktPrice")
									ratio = Rs("Ratio")
								else
									%>                   						
									<option ParentSecurity = "<%=Rs("ParentSecurity_DPA_")%>" Ratio = "<%=Rs("Ratio")%>" OfferType = "<%=Rs("OfferType_DPA_")%>" SearchPrice = "<%=rs.Fields("SecurityMktPrice")%>" value = '<%=rs.Fields("Security_DPA_")%>'><%=rs.Fields("SecurityName")%></option>
									<%
									'price = rs("SecurityMktPrice")
								end if 
							else
								if trim(rs.Fields("Security_DPA_")) = trim(Offering) Then
									%>
									<option selected ParentSecurity = "<%=Rs("ParentSecurity_DPA_")%>" Ratio = "<%=Rs("Ratio")%>" OfferType = "<%=Rs("OfferType_DPA_")%>" SearchPrice = "<%=rs.Fields("SecurityMktPrice")%>" value = '<%=rs.Fields("Security_DPA_")%>'><%=rs.Fields("SecurityName")%></option>
									<%
									price = rs("SecurityMktPrice")
									ratio = Rs("Ratio")
								else
									%>                   						
									<option ParentSecurity = "<%=Rs("ParentSecurity_DPA_")%>" Ratio = "<%=Rs("Ratio")%>" OfferType = "<%=Rs("OfferType_DPA_")%>" SearchPrice = "<%=rs.Fields("SecurityMktPrice")%>" value = '<%=rs.Fields("Security_DPA_")%>'><%=rs.Fields("SecurityName")%></option>
									<%
									'price = rs("SecurityMktPrice")
								end if 
							end if
						rs.MoveNext
					Loop
				End If
				%>
				</select>
			</td>     
		</tr>
			<tr>
				<td>Select date from:</td>
				<td>
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>
			</tr>
			<tr>
				<td>Select date To:</td>
				<td>
					<SCRIPT language="JavaScript">cal1.writeControl();</SCRIPT>	
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

Dim sqlstr, rs, rst, conn

Set Conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set rst = Server.CreateObject("ADODB.Recordset")
Set Conn = GetActiveConnection("KBroker")

%>
<%
	If Offering = "" Then Offering = 1
	
	'Get Details					        
	sqlStr = "SELECT * FROM ForwardsList where Offering = "& Offering &" AND (NOT (Batch_No IS NULL)) AND Cast(floor(cast(Offerings_Date as float)) as DateTime) Between '" & FormatDate(CDate(selectedFromDate)) & "' and '" & FormatDate(CDate(selectedToDate)) & "' " & _
	          " order by cast(floor(cast(Offerings_Date as float)) as datetime), Rtrim(ltrim(ClientName)) "
	
	Rs.CursorLocation = adUseClient	
	Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
	
	If rs.EOF Or rs.BOF Then%>
		<Script Language="JavaScript">
			alert("No batched applications were found for the specified criterion.")
			window.location.href='BatchedForwardDate.asp';
		</Script>
		<%Set Rs = Nothing
		Set Rst = Nothing
		Set Conn = Nothing
		Response.End
	End If
	
%>	
<table border="0" cellspacing="0" cellpadding="0" style="font-family: Arial Narrow" width="100%">
	<tr>
		<td width="100%" align="middle"><img src="../data/photos/aaprintlogo.jpg" border="0" width="482" height="178"></td>
	</tr>
	<br>
	<tr>
		<td width="100%" nowrap align="left"><font face="Impact" size="4">BATCHED FORWARDS SUMMARY</font></td>      
	</tr>
	<tr>
		<td width="100%" nowrap height="0" align="right">&nbsp;</td>      
	</tr>
</table>
<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
    <tr>
      <td width="10%"><b>DATE:</b></td>
      <td width="90%"><%= FormatDate(Date) %></td>
    </tr>   
	<tr>
      <td width="10%"><b>OFFER</b></td>
      <td width="90%"><font face="Arial Narrow" size="3"><%= ucase(trim(Rs("SecurityName"))) %></font></td>
    </tr>
	<tr>
      <td width="10%"><font face="Arial Narrow" size="2"><b>PAYMENT</b></font></td>
      <td width="90%"><font face="Arial Narrow" size="2">TO BE PAID BY BANKERS CHEQUE</font></td>
    </tr>
</table>
<br>

	<%
	'Prepare cover page
	 sqlstr = "SELECT Batch_No as BatchNo, SUM(Alloted_Rights) AS TotalQty,  " & _
		 "     SUM(ISNULL(Alloted_Rights, 0) * ISNULL(Offering_Price, 0)) AS TotalAmt, COUNT(Offering_DPA_)  " & _
		 "     AS AppCount " & _
		 " FROM  ForwardsList " & _
		 " where (NOT (Batch_No IS NULL)) AND Cast(floor(cast(Offerings_Date as float)) as DateTime) Between '" & FormatDate(CDate(selectedFromDate)) & "' and '" & FormatDate(CDate(selectedToDate)) & "' "  & _
		 " GROUP BY Batch_No "
    
     set rst = conn.execute(sqlstr)

	 intrscount =  rst.recordcount

	 if intrscount <= 0 then
		 %>
		 <script language="javascript">
		    alert("No batched applications were found for the specified criterion.")
			window.location.href='BatchedForwardDate.asp';
		 </script>
		 <%
		  set rst = nothing
		  set rs = nothing
		  set conn = nothing
		  response.end
	 else
       rst.movefirst
       batchData =  rst.getrows
	 end if

   '==========================================================================================
   'Prepare cover page
   '==========================================================================================
	%>
	<br>
	<table border="0" cellspacing="2" cellpadding="2" style="font-family: Arial Narrow" width="100%">
		 <tr>
			 <td style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;"><b>BATCH</b></td>
			 <td style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;"><b>PAYMENT</b></td>
			 <td align="right" style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;"><b>APPLICATIONS</b></td>
			 <td align="right" style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;"><b>TOTAL SHARES</b></td>
			 <td align="right" style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;"><b>TOTAL AMOUNT</b></td>
		 </tr>
	
	<%
	TotalAppNo = 0
	TotalShares = 0
    TotalAmount = 0

   for intcount=0 to intrscount-1
     BatchID = trim(batchData(0,intCount))
     AppNo = trim(batchData(3,intCount))
     TotalQty = trim(batchData(1,intCount))
     TotalAmt = trim(batchData(2,intCount))
     
	 %>
	 <tr>
		  <td ><%=BatchID%></td>
		  <td >Agent Cheque</td>
		  <td align="right">&nbsp;<%=AppNo%></td>
		  <td align="right">&nbsp;<%=FormatNumCommasOnly(TotalQty)%></td>
		  <td align="right">&nbsp;<%=FormatNum(TotalAmt)%></td>
	 </tr>
	 <%
     TotalAppNo = TotalAppNo + AppNo
	 TotalShares = TotalShares + TotalQty
     TotalAmount = TotalAmount + TotalAmt
   next
   %>
  
	<tr>
		  <td colspan="2"><b>TOTALS</b></td>
		  <td align="right">&nbsp;<b><%=TotalAppNo%></b></td>
		  <td align="right">&nbsp;<b><%=FormatNumCommasOnly(TotalShares)%></b></td>
		  <td align="right">&nbsp;<b><%=FormatNum(TotalAmount)%></<b></td>
	 </tr>
   </TABLE>
   <br><br>
  <table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
	 <tr>
		 <td width="5%" height = "100">&nbsp;</td>
		 <td height = "100" style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;">&nbsp;</td>
		 <td width="5%">&nbsp;</td>
		 <td height = "100" style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;">&nbsp;</td>
		 <td height = "100" width="5%">&nbsp;</td>
	 </tr>
	 <tr>
		 <td width="10%">&nbsp;</td>
		 <td width="35%" align="center">&nbsp;Agent stamp and signature</td>
		 <td width="10%">&nbsp;</td>
		 <td width="35%" align="center">&nbsp;Receiving bank and signature</td>
		 <td width="10%">&nbsp;</td>
	 </tr>
  </table>
   <BR class="newpage">
   

<table border="0" cellspacing="0" cellpadding="0" style="font-family: Arial Narrow" width="100%">
	<tr>
		<td width="100%" align="middle"><img src="../data/photos/aaprintlogo.jpg" border="0" width="482" height="178"></td>
	</tr>
	<br>
	<tr>
		<td width="100%" nowrap align="left"><font face="Impact" size="4">BATCHED FORWARDS</font></td>      
	</tr>
	<tr>
		<td width="100%" nowrap height="0" align="right">&nbsp;</td>      
	</tr>
</table>
<br>
<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
    <tr>
      <td width="10%"><b>DATE:</b></td>
      <td width="90%"><%= FormatDate(Date) %></td>
    </tr>   
	<tr>
      <td width="10%"><b>OFFER</b></td>
      <td width="90%"><font face="Arial Narrow" size="3"><%= ucase(trim(Rs("SecurityName"))) %></font></td>
    </tr>
	<tr>
      <td width="10%"><font face="Arial Narrow" size="2"><b>PAYMENT</b></font></td>
      <td width="90%"><font face="Arial Narrow" size="2">TO BE PAID BY BANKERS CHEQUE</font></td>
    </tr>
</table>
<BR>


  <table border="0" cellspacing="3" cellpadding="0" style="font-family: Arial Narrow; LEFT-MARGIN:100PX"  width="100%">
    <tr>   
	  <td style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;"  nowrap><b><font face="Arial Narrow" size="2">Date</font></b></td>
      <td style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;"  nowrap><b><font face="Arial Narrow" size="2">Serial No</font></b></td>
      <td style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;"><b><font face="Arial Narrow" size="2">Code</font></b></td>
      <td style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;"  nowrap><b><font face="Arial Narrow" size="2">Client Name</font></b></td>      
	  <td style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;" nowrap><b><font face="Arial Narrow" size="2">Bank</font></b></td>
      <td style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;" nowrap><b><font face="Arial Narrow" size="2">Cheque No</font></b></td>         
      <td align="right" style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;"  align=right nowrap><b><font face="Arial Narrow" size="2">Quantity</font></b></td>
	  <td align="right" style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;"  nowrap><b><font face="Arial Narrow" size="2">Price</font></b></td>
	  <td align="right" style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;"  nowrap><b><font face="Arial Narrow" size="2">Payable</font></b></td>
    </tr>

   <%
   totalquantity=0
   totalpayable=0

   do until Rs.eof
	%>
	<tr>	
        <td nowrap><%=FormatDate(Rs("Offerings_Date"))%></td>
		<td><%=Rs("Pal_No")%></td>
		<td><%=Rs("Client_DPA_")%></td>
		<td nowrap><%=Mid(RS("ClientName"),1,30)%></td>		
		<td nowrap><%=Mid(RS("OfferBank"),1,30)%></td>
		<td nowrap><%=Mid(RS("OfferCheque"),1,30)%></td>		
		<td align="right"><%=FormatNumEx(Rs("Alloted_Rights"),0)%></td>
		<td align="right"><%=FormatNum(Rs("Offering_Price"))%></td>
		<td align="right"><%=FormatNum(Rs("Alloted_Rights")*Rs("Offering_Price"))%></td>
	</tr>
	<%

	totalquantity = totalquantity + Rs("Alloted_Rights")
	totalpayable = totalpayable + (Rs("Alloted_Rights")*Rs("Offering_Price"))
	
	Rs.MoveNext
   loop

   %>
   <tr>
		<td colspan="5">&nbsp;</td>
		<td align="left"><b>Totals</b></td>
		<td align="right"><b><%=FormatNumEx(totalquantity,0)%></b></td>
		<td align="right">&nbsp;</td>
		<td align="right"><b><%=FormatNum(totalpayable)%></b></td>
   </tr>
  </table>
  
  <p>
<p>
  <table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
	 <tr>
		 <td width="5%" height = "100">&nbsp;</td>
		 <td height = "100" style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;">&nbsp;</td>
		 <td width="5%">&nbsp;</td>
		 <td height = "100" style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;">&nbsp;</td>
		 <td height = "100" width="5%">&nbsp;</td>
	 </tr>
	 <tr>
		 <td width="10%">&nbsp;</td>
		 <td width="35%" align="center">&nbsp;Agent stamp and signature</td>
		 <td width="10%">&nbsp;</td>
		 <td width="35%" align="center">&nbsp;Receiving bank and signature</td>
		 <td width="10%">&nbsp;</td>
	 </tr>
  </table>
   
<%Set Rs = Nothing
Set Conn = Nothing%>   
</body>

</html>