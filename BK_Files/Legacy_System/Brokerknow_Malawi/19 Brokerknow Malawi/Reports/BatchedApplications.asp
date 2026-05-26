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

genReport = trim(Request.Form("genReport"))
selectedBatch = trim(Request.Form("cboBatch"))

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
					        
					        sqlStr = "SELECT distinct Batch_No FROM Offerings where Batch_NO is not null and Forward <>1 and deleted=0 ORDER BY Batch_No"
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
	Set conn = GetActiveConnection("KBroker")
	Set Rs = CreateObject("ADODB.Recordset")						        
	sqlStr = "SELECT * FROM OfferingsList where Batch_No = " & selectedBatch & _
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

<table border="0" cellspacing="0" cellpadding="0" style="font-family: Arial Narrow" width="100%">
	<tr>
		<td width="100%" align="middle"><img src="../data/photos/aaprintlogo.jpg" border="0" width="482" height="178"></td>
	</tr>
	<br>
	<tr>
		<td width="100%" nowrap align="left"><font face="Impact" size="4">BATCHED APPLICATIONS</font></td>      
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
      <td width="90%"><font face="Arial Narrow" size="2">TO BE PAID BY AGENT</font></td>
    </tr>
	<tr>
      <td width="10%"><font face="Arial Narrow" size="2"><b>BATCH NO</b></font></td>
      <td width="90%"><font face="Arial Narrow" size="2"><%=selectedBatch%></font></td>
    </tr>
<!--    
<%
     'filename = trim(rs("BatchFileName"))
     
     'if filename <> "" then
    %>
    <tr>
		  <td width="10%"><font face="Arial Narrow" size="2"><b>FILE NAME</b></font></td>
		  <td width="90%"><font face="Arial Narrow" size="2"><%=filename%></font></td>
		</tr>
	<%'end if%>-->
</table>
<BR>


  <table border="0" cellspacing="3" cellpadding="0" style="font-family: Arial Narrow; LEFT-MARGIN:100PX"  width="100%">
    <tr>
	  <td><b><font face="Arial Narrow" size="2">#</font></b></td>
      <td style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;"><b><font face="Arial Narrow" size="2">Date</font></b></td>
      <td style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;"  nowrap><b><font face="Arial Narrow" size="2">Serial No</font></b></td>
      <td style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;"><b><font face="Arial Narrow" size="2">Code</font></b></td>
      <td style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;"  nowrap><b><font face="Arial Narrow" size="2">Client Name</font></b></td>
      <td align="right" style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;"  align=right nowrap><b><font face="Arial Narrow" size="2">Quantity</font></b></td>
	  <td align="right" style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;"  nowrap><b><font face="Arial Narrow" size="2">Price</font></b></td>
	  <td align="right" style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;"  nowrap><b><font face="Arial Narrow" size="2">Payable</font></b></td>
    </tr>
   <%
   totalquantity=0
   totalpayable=0
   num=1
   do until Rs.eof
	%>
	<tr>
		<td><b><%=num%></b></td>
		<td nowrap><%=FormatDate(Rs("Offerings_Date"))%></td>
		<td><%=Rs("Pal_No")%></td>
		<td><%=Rs("Client_DPA_")%></td>
		<td nowrap><%=Mid(RS("ClientName"),1,30)%></td>		
		<td align="right"><%=FormatNumEx(Rs("Alloted_Rights"),0)%></td>
		<td align="right"><%=FormatNum(Rs("Offering_Price"))%></td>
		<td align="right"><%=FormatNum(Rs("Alloted_Rights")*Rs("Offering_Price"))%></td>
	</tr>
	<%
	totalquantity = totalquantity + Rs("Alloted_Rights")
	totalpayable = totalpayable + (Rs("Alloted_Rights")*Rs("Offering_Price"))
	num=num+1
	Rs.MoveNext
   loop

   %>
   <tr>
		<td colspan="5" align="right"><b>Totals</b></td>
		<td align="right"><b><%=FormatNumEx(totalquantity,0)%></b></td>
		<td align="right">&nbsp;</td>
		<td align="right"><b><%=FormatNum(totalpayable)%></b></td>
   </tr>
  </table>

  <P> <P>

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