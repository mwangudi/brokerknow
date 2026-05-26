<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Batched Applications By Date</title>
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
offering = Trim(Request.Form("cboOfferings"))
selectedFromDate = Request.Form("transFromDate")
selectedToDate = Request.Form("transToDate")
filterDebits = Request.Form("txtDebit")
filtereIPO = Request.Form("txteipo")
filterAmount = Request.Form("txtAmount")

If genReport <> "1" Or Not IsDate(selectedFromDate) Then%>
	<Script Language="JavaScript">
		report_SetBodyClass();
		function validateForm(frm){						
			if (frm.cboOfferings[frm.cboOfferings.selectedIndex].SearchCode == ''){
				alert("Please select an Offering.");
				frm.cboOfferings.focus();
				return;
			}

			frm.target = '_self';			
			frm.submit();
		}

		function updateChk(theChk){
		    if(theChk.name.toLowerCase()=='chkdebit'){
		        if(theChk.checked){
		            document.frmMain.elements("txtDebit").value="1";
		        }
		        else{
		            document.frmMain.elements("txtDebit").value="0";
		        }
		    }

			else if(theChk.name.toLowerCase()=='chkeipo'){
			    if(theChk.checked){
		            document.frmMain.elements("txteipo").value="1";
		        }
			    else{
		                document.frmMain.elements("txteipo").value="0";
		        }
		    }
		}
		
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "transFromDate","cmdDate","<%= FormatDate(FirstDay) %>",1);
		var cal1=new ctlSpiffyCalendarBox("cal1", "frmMain", "transToDate","cmdDate","<%= FormatDate(Date) %>",1);

	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="BatchedApplicationsDate.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">
		<table>
			<tr>
				<td nowrap>Select an Offering&nbsp;</td>
				<td>
				<select name = 'cboOfferings' id = 'cboOfferings' size="1">
					<%
					Set conn = GetActiveConnection("KBroker")
					sqlStr = "SELECT * FROM [SecurityListOfferings] " & _

					" ORDER BY DefaultSelection DESC, Security_DPA_ DESC"
					Set rs = conn.Execute(SqlStr)

					If Not (rs.EOF Or rs.BOF) Then
					    if DefaultSelection = 1 Then%>
					        <option selected SearchCode = '<%=rs("SecurityCode")%>' value = '<%=rs("Security_DPA_")%>'><%=rs("SecurityName")%></option>
					        <%
					    else    
						    Do Until rs.EOF%>
						    <option SearchCode = '<%=rs("SecurityCode")%>' value = '<%=rs("Security_DPA_")%>'><%=rs("SecurityName")%></option>
						    <%
							    rs.MoveNext
						    Loop
						end if
					End If%>
				</select>
				</td>
			</tr>
			<tr>
				<td>Show Debits Only:</td>
				<td>
					<input type="checkbox" name="chkDebit" id="chkDebit" OnClick="JavaScript: updateChk(this)">
				</td>
			</tr>
			<tr style="Display:none">
				<td>Show eIPO Only:</td>
				<td>
					<input type="checkbox" name="chkeIPO" id="chkeIPO" OnClick="JavaScript: updateChk(this)">
				</td>
			</tr>
			<tr>
			    <td>Filter By Amount</td>
				<td width="80%" nowrap><input type = 'text' name ='txtAmount' id = 'txtAmount' size="25" value=0></td>
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
				<input type="hidden" name="txtDebit" id="txtDebit" value="0">
				<input type="hidden" name="txteipo" id="txteipo" value="0">
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
	if Not(isnumeric(filterAmount)) Or filterAmount < 0 then filterAmount = 0
	filterAmount = Replace(filterAmount,",","")
	
	if filterDebits = 1 and filtereIPO = 1 Then 'show both debit and eipo applications
		sqlStr = "SELECT     Offerings.Offerings_Date, Offerings.PAL_No, Offerings.Batch_No, Offerings.Client_DPA_, Client.ClientName, Offerings.Alloted_Rights, Offerings.Offering_Price,  " & _
				 "                       Security.SecurityName, ISNULL(ClientBalances.CurrentBal,0) as CurrentBalance " & _
				 " FROM         Offerings INNER JOIN " & _
				 "                       Security ON Offerings.Offering = Security.Security_DPA_ INNER JOIN " & _
				 "                       Client ON Offerings.Client_DPA_ = Client.Client_DPA_ LEFT OUTER JOIN " & _
				 "                       ClientBalances ON Client.Client_DPA_ = ClientBalances.client_DPA_ " & _
				 " WHERE     (CAST(FLOOR(CAST(Offerings.Offerings_Date AS float)) AS DateTime) BETWEEN '" & FormatDate(CDate(selectedFromDate)) & "' and '" & FormatDate(CDate(selectedToDate)) & "') AND (Offerings.Offering = "& offering &") AND  " & _
				 "                       (Offerings.Deleted <> 1) AND (Offerings.Forward <> 1) AND (ROUND(ISNULL(ClientBalances.CurrentBal,0),2)<0) AND (NOT (Offerings.CitiAccepted IS NULL)) AND (Offerings.Alloted_Rights*Offerings.Offering_Price>="& filteramount &")" & _
				 " ORDER BY cast(floor(cast(Offerings.Offerings_Date as float)) as datetime), Offerings.Batch_No"
				 
	elseif filterDebits = 1 and filtereIPO = 0 Then 'show debit applications only
		sqlStr = "SELECT     Offerings.Offerings_Date, Offerings.PAL_No, Offerings.Batch_No, Offerings.Client_DPA_, Client.ClientName, Offerings.Alloted_Rights, Offerings.Offering_Price,  " & _
				 "                       Security.SecurityName,ISNULL(ClientBalances.CurrentBal,0) as CurrentBalance " & _
				 " FROM         Offerings INNER JOIN " & _
				 "                       Security ON Offerings.Offering = Security.Security_DPA_ INNER JOIN " & _
				 "                       Client ON Offerings.Client_DPA_ = Client.Client_DPA_ LEFT OUTER JOIN " & _
				 "                       ClientBalances ON Client.Client_DPA_ = ClientBalances.client_DPA_ " & _
				 " WHERE     (CAST(FLOOR(CAST(Offerings.Offerings_Date AS float)) AS DateTime) BETWEEN '" & FormatDate(CDate(selectedFromDate)) & "' and '" & FormatDate(CDate(selectedToDate)) & "') AND (Offerings.Offering = "& offering &") AND  " & _
				 "                       (Offerings.Deleted <> 1) AND (Offerings.Forward <> 1) AND (ROUND(ISNULL(ClientBalances.CurrentBal,0),2)<0) AND (Offerings.Alloted_Rights*Offerings.Offering_Price>="& filteramount &")" & _
				 " ORDER BY cast(floor(cast(Offerings.Offerings_Date as float)) as datetime), Offerings.Batch_No"
				 
	elseif filterDebits = 0 and filtereIPO = 1 Then 'show both debit and eipo
	    sqlStr = "SELECT     Offerings.Offerings_Date, Offerings.PAL_No, Offerings.Batch_No, Offerings.Client_DPA_, Client.ClientName, Offerings.Alloted_Rights, Offerings.Offering_Price,  " & _
				 "                       Security.SecurityName,  ISNULL(ClientBalances.CurrentBal,0) as CurrentBalance " & _
				 " FROM         Offerings INNER JOIN " & _
				 "                       Security ON Offerings.Offering = Security.Security_DPA_ INNER JOIN " & _
				 "                       Client ON Offerings.Client_DPA_ = Client.Client_DPA_ LEFT OUTER JOIN " & _
				 "                       ClientBalances ON Client.Client_DPA_ = ClientBalances.client_DPA_ " & _
				 " WHERE     (CAST(FLOOR(CAST(Offerings.Offerings_Date AS float)) AS DateTime) BETWEEN '" & FormatDate(CDate(selectedFromDate)) & "' and '" & FormatDate(CDate(selectedToDate)) & "') AND (Offerings.Offering = "& offering &") AND  " & _
				 "                       (Offerings.Deleted <> 1) AND (Offerings.Forward <> 1) AND (NOT (Offerings.CitiAccepted IS NULL)) AND (Offerings.Alloted_Rights*Offerings.Offering_Price>="& filteramount &")" & _
				 " ORDER BY cast(floor(cast(Offerings.Offerings_Date as float)) as datetime), Offerings.Batch_No"
				 
	else 'show all
	    sqlStr = "SELECT     Offerings.Offerings_Date, Offerings.PAL_No, Offerings.Batch_No, Offerings.Client_DPA_, Client.ClientName, Offerings.Alloted_Rights, Offerings.Offering_Price,  " & _
				 "                       Security.SecurityName,ISNULL(ClientBalances.CurrentBal,0) as CurrentBalance " & _
				 " FROM         Offerings INNER JOIN " & _
				 "                       Security ON Offerings.Offering = Security.Security_DPA_ INNER JOIN " & _
				 "                       Client ON Offerings.Client_DPA_ = Client.Client_DPA_ LEFT OUTER JOIN " & _
				 "                       ClientBalances ON Client.Client_DPA_ = ClientBalances.client_DPA_ " & _
				 " WHERE     (CAST(FLOOR(CAST(Offerings.Offerings_Date AS float)) AS DateTime) BETWEEN '" & FormatDate(CDate(selectedFromDate)) & "' and '" & FormatDate(CDate(selectedToDate)) & "') AND (Offerings.Offering = "& offering &") AND  " & _
				 "                       (Offerings.Deleted <> 1) AND (Offerings.Forward <> 1) AND (Offerings.Alloted_Rights*Offerings.Offering_Price>="& filteramount &")" & _
				 " ORDER BY cast(floor(cast(Offerings.Offerings_Date as float)) as datetime), Offerings.Batch_No"		 
	end if
	
	'Response.write sqlStr : Response.end

	set rs = conn.execute(sqlStr)

	If rs.EOF Or rs.BOF Then%>
		<Script Language="JavaScript">
			alert("No data found")
			window.history.go(-1);
		</Script>
		<%Set Rs = Nothing
		Set Conn = Nothing
		Response.End
	End If
	
%>	

<table border="0" cellspacing="2" cellpadding="2" style="font-family: Arial Narrow" width="100%">
    <tr>
		<td width="10%" nowrap><font face="Impact" size="4">Applications By Date</font></td>
      <td width="60%" nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>
      
    </tr>

  </table>
<br>
<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
    <tr>
      <td width="1%"><b>Date:</b></td>
      <td width="48%"><%= FormatDate(Date) %></td>
    </tr>
	<tr>
      <td colspan="2" width="100%"><font face="Arial Narrow" size="3"><b><%= Rs("SecurityName") %></b></font></td>
    </tr>
</table>
<BR>


  <table border="0" cellspacing="3" cellpadding="0" style="font-family: Arial Narrow; LEFT-MARGIN:100PX"  width="100%">
    <tr>
      <td style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"><b><font face="Arial Narrow" size="2">Date</font></b></td>
      <td style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"  nowrap><b><font face="Arial Narrow" size="2">Serial No</font></b></td>
	  <td style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"  nowrap><b><font face="Arial Narrow" size="2">Batch No</font></b></td>
	  <td style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;display:none;"  nowrap><b><font face="Arial Narrow" size="2">eIPO Serial No</font></b></td>
	  <td style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;display:none;"  nowrap><b><font face="Arial Narrow" size="2">eIPO Batch No</font></b></td>
      <td style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"><b><font face="Arial Narrow" size="2">Code</font></b></td>
      <td style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"  nowrap><b><font face="Arial Narrow" size="2">Client Name:</font></b></td>
      <td style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"  nowrap><b><font face="Arial Narrow" size="2">Current Balance</font></b></td>
	  <td align="right" style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"  align=right nowrap><b><font face="Arial Narrow" size="2">Quantity</font></b></td>
	  <td align="right" style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"  nowrap><b><font face="Arial Narrow" size="2">Price</font></b></td>
	  <td align="right" style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"  nowrap><b><font face="Arial Narrow" size="2">Payable</font></b></td>
	  <td style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;display:none;"  nowrap><b><font face="Arial Narrow" size="2">&nbsp;Status</font></b></td>
    </tr>

   <%
   totalquantity=0
   totalpayable=0

   do until Rs.eof
	%>
	<tr>
		<td nowrap><%=FormatDate(Rs("Offerings_Date"))%></td>
		<td><%=Rs("Pal_No")%></td>
		<td><%=Rs("Batch_No")%></td>
		<td style="display:none;"><%'=Rs("CitiSerialNo")%></td>
		<td style="display:none;"><%'=Rs("CitiBatchNo")%></td>
		<td><%=Rs("Client_DPA_")%></td>
		<td nowrap><%=Mid(RS("ClientName"),1,30)%></td>	
		<td align="right"><%=FormatNumEx(Rs("CurrentBalance"),0)%></td>
		<td align="right"><%=FormatNumEx(Rs("Alloted_Rights"),0)%></td>
		<td align="right"><%=FormatNum(Rs("Offering_Price"))%></td>
		<td align="right"><%=FormatNum(Rs("Alloted_Rights")*Rs("Offering_Price"))%></td>
		<td align="left" style="display:none;">&nbsp;<%=Rs("Status")%></td>
	</tr>
	<%

	totalquantity = totalquantity + Rs("Alloted_Rights")
	totalpayable = totalpayable + (Rs("Alloted_Rights")*Rs("Offering_Price"))
	
	Rs.MoveNext
   loop

   %>
   <tr>
		<td colspan="6" align="right"><b>Totals</b></td>
		
		<td align="right"><b><%=FormatNumEx(totalquantity,0)%></b></td>
		<td align="right">&nbsp;</td>
		<td align="right"><b><%=FormatNumEx(totalpayable,2)%></b></td>
   </tr>
  </table>
   
<%
Set Rs = Nothing
Conn.close
Set Conn = Nothing

%>   
</body>

</html>