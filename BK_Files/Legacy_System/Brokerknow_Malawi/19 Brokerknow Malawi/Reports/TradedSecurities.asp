<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Debtors and Creditors</title>
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
Set conn = GetActiveConnection("KBroker")
genReport = Request.Form("genReport")

selectedFromDate = Request.Form("transFromDate")
secType = Request.Form("secType")
TDate = Request.Form("TDate")

If genReport <> "1" Or Not IsDate(selectedFromDate) Then%>
	<Script Language="JavaScript">
		report_SetBodyClass();
		
		function validateForm(frm){			
			frm.target = '_self';			
			frm.submit();
		}
		
		
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "transFromDate","cmdDate","<%= FormatDate(Date) %>",1);
	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="TradedSecurities.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">
		&nbsp;<p>
		<table>
		
			<tr>
				<td><input type="checkbox" name="Tdate" value="1">  Trade Date:</td>
				<td>
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>
			</tr>
			<tr>
				<td>Security</td>
				<td>
				<select id="secType" name="secType">
					<option value=''>Select Security</option>
				<%set rsSec = server.CreateObject("adodb.recordset")
					secSql= "Select security_DPA_ ,securityCode from security"
					
				  rsSec.Open SecSql, conn,0,1
					  
				  if not rsSec.EOF or not rsSec.BOF then
					do until rsSec.eof
				%>
			
					<option value='<%=rsSec("Security_DPA_")%>'><%=rsSec("SecurityCode")%></option>
				<%
						rsSec.MoveNext 
					loop
				%>
					</select>	
				<%end if%>
				</td>
			</tr>
			<tr>
				<td colspan=2><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... ">&nbsp;&nbsp;</td>
			</tr>
		</table>
		
	</form>
	
	<%Set rsSec = Nothing
	
	Response.End
End If

%>

<% DrawPageFunctions True, True, True %>

<%
   Dim conn 
   Dim sqlStr
   Dim rs
	Set Rs = CreateObject("ADODB.Recordset")						        
	Rs.CursorLocation = adUseClient	
	'select the securities traded here
		
		If Len(secType) = 0 then
			%>
			<script>
				alert("Please select a security.");
				window.history.back();
			</script>
			<%
			Response.End
		End If

		sqlstr= "Select * from LotList where Security_DPA_= "& secType  
		
		if trim(TDate)<>"" then
			sqlstr= sqlstr &" and LotTDate = '" & FormatDate(selectedFromDate) &"'"
		end if
		sqlstr = sqlstr &" order by LotTDate"
	
		rs.Open sqlStr,conn, 0,1
			
		If rs.EOF Or rs.BOF Then
                %>
                <Script Language="JavaScript">
					alert("No records were found using the search criteria");
					window.parent.history.go(-1);					
                </Script>
                <% Set Rs = Nothing
                Set Conn = Nothing
                Response.End
        End If
        
        rs.MoveFirst
        
%>



<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
    <tr>
      <td width="10%" nowrap><b><font face="Arial" size="4"><b>Traded Securities</b></font></b></td>
      <td nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>
      
    </tr>

    <tr>
      <td nowrap colspan=2><font size="2" face="Arial">as at: <%= FormatDate(selectedFromDate) %></font></td>
    </tr>

  </table>
<br>
 <table border="0" cellspacing="0" cellpadding="3" style="font-family: Arial Narrow; LEFT-MARGIN:100PX" width="100%">
    <tr>
      <td bgcolor="#000000" width="10%"><b><font color="#FFFFFF" face="Arial Narrow" size="3" >Security </font></b></td>
      <td bgcolor="#000000" width="20%"><b><font color="#FFFFFF" face="Arial Narrow" size="3" >Contract</font></b></td>
	  <td bgcolor="#000000" width="20%" ><b><font color="#FFFFFF" face="Arial Narrow" size="3" >Lot Qty</font></b></td>
      <td bgcolor="#000000" width="20%"><b><font color="#FFFFFF" face="Arial Narrow" size="3" >Trade Date</font></b></td>
      <td bgcolor="#000000" width="20%"><b><font color="#FFFFFF" face="Arial Narrow" size="3" >REF</font></b></td>
      <td align="left" bgcolor="#000000" width="50%"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Client</font></b></td>
      
      
    </tr>
    
    <%
    totalBal = 0
    Do Until Rs.EOF
		%>
			<tr>
			  <td width="20%" nowrap><font face="Arial Narrow" size="2"><%= rs("SecurityCode") %></font></td>
			  <td width="10%" nowrap><font face="Arial Narrow" size="2"><%= rs("ContractNumber")%></font></td>
			  <td width="10%" nowrap ><font face="Arial Narrow" size="2"><%= formatnumber(rs("LotQty"),0)%></font></td>
			  <td width="15%" nowrap><font face="Arial Narrow" size="2"><%= rs("LotTDate") %></font></td>
			  <td width="15%" nowrap><font  face="Arial Narrow" size="2"><%= rs("ClientCdsNo")%></font></td>
			  <td width="40%" nowrap><font  face="Arial Narrow" size="2"><%= rs("ordDetailClient") %></font></td>
			  
			</tr>
    
    <%		
		Rs.MoveNext
    Loop%>

    
  </table>

</body>

</html>
