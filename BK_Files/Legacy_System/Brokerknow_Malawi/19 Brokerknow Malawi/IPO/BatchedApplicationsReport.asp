<html>
<%
Offering = Request("Offering")
%>
<head>
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
	<form method="POST" action="BatchedApplicationsReport.asp?Offering=<%=Offering%>" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">
		<table>
			<tr>
			<td width="20%" nowrap>Offering Name</td>
			<td width="80%" nowrap>
				<select name = 'cboOfferings' id = 'cboOfferings' size="1" onChange="window.location.href='BatchedApplicationsReport.asp?Offering='+this.value">
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
				<td nowrap>Batch No: </td>
				<td><select name = 'cboBatch' id = 'cboBatch' size="1">					
					<%
							If Offering = "" Then Offering = 1
					        
					        sqlStr = "SELECT distinct Batch_No FROM Offerings where Offering= "& Offering &" AND Batch_NO is not null and Forward <>1 and deleted=0 ORDER BY Batch_No desc"
					        Set rs = conn.Execute(SqlStr)
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
		<%If Request("Offering") = "" Then%>
		<script language=javascript>window.location.href='BatchedApplicationsReport.asp?Offering='+document.all.item("cboOfferings").value;</script>
		<%End If%>
	</form>
	
	<%Set rs = Nothing
	Set Conn = Nothing
	Response.End
End If

%>

<% DrawPageFunctions True, True, True %>

<%
If Offering = "" Then Offering = 1
	
Set conn = GetActiveConnection("KBroker")
Set Rs = CreateObject("ADODB.Recordset")						        

SqlStr = "SELECT OfferingsList.*, ISNULL(Client.ClientAddr, '') AS ClientAddr, ISNULL(Residency.ResidencyDescription, '') AS Residence, ISNULL(Class.ClassDescription, " & _
	" '') AS Category, '' AS Citizenship, ISNULL(Bank.BankName, '') AS BankName, ISNULL(BankAcc.BankAccNumber, '') AS BankAccount,  " & _
	" ISNULL(BankAcc.BnkBranch, '') AS BankBranch " & _
	" FROM OfferingsList INNER JOIN " & _
	" Client ON OfferingsList.Client_DPA_ = Client.Client_DPA_ LEFT OUTER JOIN " & _
	" Class ON Client.Class_DPA_ = Class.Class_DPA_ LEFT OUTER JOIN " & _
	" Residency ON Client.Residency_DPA_ = Residency.Residency_DPA_ LEFT OUTER JOIN " & _
	" BankAcc ON Client.Client_DPA_ = BankAcc.Client_DPA_ LEFT OUTER JOIN " & _
	" Bank ON BankAcc.Bank_DPA_ = Bank.Bank_DPA_ " & _
	" WHERE (OfferingsList.Offering = "& Offering &") AND (OfferingsList.Batch_No = "& selectedBatch &") " & _
	" ORDER BY CAST(FLOOR(CAST(OfferingsList.Offerings_Date AS float)) AS datetime), RTRIM(LTRIM(OfferingsList.ClientName))"	

Rs.CursorLocation = adUseClient	
Rs.Open SqlStr, conn.ConnectionString, adOpenKeyset, adLockOptimistic
	
If rs.EOF Or rs.BOF Then%>
	<Script Language="JavaScript">
		alert("No batched applications were found for the selected batch.")
		window.location.href='BatchedApplicationsReport.asp';
	</Script>
	<%Set Rs = Nothing
	Set Conn = Nothing
	Response.End
End If
%>	

<table border="0" cellspacing="2" cellpadding="2" style="font-family: Arial;" width="95%" align=center>
	<tr>
		<td width="100%" align="middle"><img src="../data/photos/aaprintlogo.jpg" border="0" width="482" height="178"></td>
	</tr>
	<br>
	<tr>
		<td width="100%" nowrap align="left"><font face="Impact" size="4">BATCHED APPLICATIONS</td>      
	</tr>
	<tr>
		<td width="100%" nowrap height="0" align="right">&nbsp;</td>      
	</tr>
</table>

<br>

<table border="0" cellspacing="2" cellpadding="2" style="font-family: Arial;" width="95%" align=center>
    <tr>
      <td width="10%"><b>DATE:</b></td>
      <td width="90%"><%= FormatDate(Date) %></td>
    </tr>   
	<tr>
      <td width="10%"><b>OFFER</b></td>
      <td width="90%"><font face="Arial Narrow" size="3"><%= ucase(trim(Rs("SecurityName"))) %></td>
    </tr>
	<tr>
      <td width="10%"><b>PAYMENT</b></td>
      <td width="90%">TO BE PAID BY AGENT</td>
    </tr>
	<tr>
      <td width="10%"><b>BATCH NO</b></td>
      <td width="90%"><%=selectedBatch%></td>
    </tr>
</table>

<BR>

<table border="0" cellspacing="2" cellpadding="2" style="font-family: Arial;" width="95%" align=center>
    <tr>
	  <td><b>#</b></td>
      <td style="border: 1 solid black;"><b>Serial No</b></td>
      <td style="border: 1 solid black;"><b>Receiving Agent</b></td>
      <td style="border: 1 solid black;"><b>Applicant's Name</b></td>
      <td style="border: 1 solid black;"><b>Address</b></td>
      <td align="right" style="border: 1 solid black;"><b><font face="Arial Narrow" size="2">Applied</b></td>
	  <td align="right" style="border: 1 solid black;"><b>Payable</b></td>
	  <td align="right" style="border: 1 solid black;"><b>Category</b></td>
	  <td align="right" style="border: 1 solid black;"><b>Residence</b></td>
	  <td align="right" style="border: 1 solid black;"><b>Citizenship</b></td>
	  <td align="right" style="border: 1 solid black;"><b>Bank Name</b></td>
	  <td align="right" style="border: 1 solid black;"><b>Bank Branch</b></td>
	  <td align="right" style="border: 1 solid black;"><b>Bank Account</b></td>
    </tr>
   <%
   totalquantity=0
   totalpayable=0
   num=1
   do until Rs.eof
	%>
	<tr>
		<td><%=num%></td>
		<td><%=Rs("Pal_No")%></td>
		<td><%=Mid(RS("BrokerName"),1,30)%></td>
		<td><%=Rs("Client_DPA_") & " - " & Mid(RS("ClientName"),1,30)%></td>
		<td><%=Rs("ClientAddr")%></td>
		<td align="right"><%=FormatNumEx(Rs("Alloted_Rights"),0)%></td>
		<td align="right"><%=FormatNum(Rs("Alloted_Rights")*Rs("Offering_Price"))%></td>
		<td><%=FormatNum(Rs("Category"))%></td>
		<td><%=FormatNum(Rs("Citizenship"))%></td>
		<td><%=FormatNum(Rs("Residence"))%></td>
		<td><%=FormatNum(Rs("BankName"))%></td>
		<td><%=FormatNum(Rs("BankBranch"))%></td>
		<td><%=FormatNum(Rs("BankAccount"))%></td>
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
		<td align="right"><b><%=FormatNum(totalpayable)%></b></td>
		<td colspan=6 align="right">&nbsp;</td>
   </tr>
  </table>

  <P> <P>

  <table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
	 <tr>
		 <td width="5%" height = "100">&nbsp;</td>
		 <td height = "100" style="border: 1 solid black;">&nbsp;</td>
		 <td width="5%">&nbsp;</td>
		 <td height = "100" style="border: 1 solid black;">&nbsp;</td>
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