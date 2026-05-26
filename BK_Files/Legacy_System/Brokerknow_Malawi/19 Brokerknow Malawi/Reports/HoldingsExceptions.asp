<!--#include file="../libroutinesTEST.asp"-->

<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Holdings Exceptions</title>  
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
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
			content: "<%=FormatDate(Date)%>";			
		}
			
		margin-left: 2cm;
		margin-right: 5cm;
		margin-top: 1cm;    
		margin-bottom: 2cm;
		size: landscape;
		br.newpage{
			page-break-before:always;
		}
	}
</style>

</head>

<body Class="Reports">
<%
genReport = Request.Form("genReport")
selectedSecurity = Request.Form("cboSecurity")
HoldExp = Request.Form("chkExp")
selectedclienttype = Request.Form("clienttype")
selectedclientclass = Request.Form("clientclass")


If genReport <> "1" Then%>
	<Script Language="JavaScript">
		report_SetBodyClass();
		function validateForm(frm){			
			frm.target = '_self';			
			frm.submit();
		}
	</Script>
	
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="HoldingsExceptions.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">	
		<br>
		
  <table align="center" cellspacing=3 cellpadding=3 width="90%">
    <tr> 
      <td>Security: </td>
      <td><input type = 'text' name ='txtSecCode' id = 'txtSecCode' size="10" onBlur="txtval = this.value; selectItem(cboSecurity);"></td>
      <td><select name = 'cboSecurity' id = "cboSecurity" size="1" 
    				onchange='UpdateCode(true,cboSecurity,txtSecCode)'
					onKeypress="return (dodefaultaction()==''); " 
					onKeydown="return (dodefaultaction()==''); " 
					onKeyup="return (FilterData(this,1,UpdateCode(change(cboSecurity,0),cboSecurity,txtSecCode)));" 
					onfocus="txtval = '';inputIsItemCode = 1;" 
					onblur="txtval = '';inputIsItemCode = 1;">
          <option selected SearchCode = "0" SearchText = "All Securities"  value = '0'>All Securities</option>
        <%
					dim SecurityName
					dim NameSecurity
					
					Set conn = GetActiveConnection("KBroker")
					
					 sqlStr = "SELECT * FROM Security order by SecurityName"
						Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
						If Not (rs.EOF Or rs.BOF) Then
							rs.MoveFirst
							Do Until rs.EOF
								SecurityName=rs.Fields("SecurityName")
								NameSecurity=SecurityName
		%>
          <option SearchCode = "<%=rs.Fields("SecurityCode")%>" SearchText = "<%=NameSecurity%>" value = '<%=rs.Fields("Security_DPA_")%>'><%=NameSecurity%></option>
        <%
						            rs.MoveNext
					            Loop
					    End If
					%>
        </select> </td>
    </tr>
    <tr>
      <td>Client Type :</td>
      <td valign=center colspan=2><select name="clienttype" id="clienttype">
          <option value="0" selected>Both Nominees &amp; Non Nominee</option>
          <option value="1">Nominees Only</option>
          <option value="2">Non Nominees Only</option>
        </select></td>
    </tr>
	<tr>
      <td>Client Class :</td>
      <td valign=center colspan=2><select name="clientclass" id="clientclass">
          <option value="0" selected>All Clients</option>
          <option value="1">Non-Custodian Client</option>
          <option value="2">Custodian Client</option>
        </select></td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td valign=center colspan=2> <input type="checkbox" name="chkExp" id="chkExp" value="1" checked>
        &nbsp;<b>Show Exceptions Only</b> </td>
    </tr>
    <tr> 
      <td colspan=3><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... " id=Button1 name=Button1></td>
    </tr>
  </table>
		
	</form>
	
	<%
	Response.End
End If

%>

<%  DrawPageFunctions True, True, True, True %>

<%

'response.Write(cbool(HoldExp))
'response.End()

If selectedSecurity="" Then
	%>
	<Script Language="JavaScript">
		alert("Please select The Security")
		window.history.go(-1);
	</Script>
	<%
	Response.End
End If

Select Case selectedclienttype
	Case 0
		if Cstr(selectedSecurity) = "0" then
			NomCliStr =  " (IsNominee = 1 Or IsNominee = 0) "
		else
			NomCliStr =  " "
		end if
	Case 1
		'NomCliStr =  " IsNominee = 1 "
		NomCliStr =   " ClientName Like N'%Nominee%'"
	Case 2
		NomCliStr =  " IsNominee = 0 "
End Select 

Select Case selectedclientclass
	Case 0
		if Cstr(selectedSecurity) = "0" then
			CusCliStr =  " (IsNominee = 1 Or IsNominee = 0) "
		else
			CusCliStr =  " "
		end if
	Case 1
		CusCliStr = " IsCustodian = 0 "
	Case 2
		CusCliStr = " IsCustodian = 1 " 
End Select

Set conn = GetActiveConnection("KBroker")
conn.CommandTimeOut = 0

if Cstr(selectedSecurity) = "0" then
	NomCliStr = Replace(NomCliStr," AND "," ")
	if trim(NomCliStr) <> "" then CusCliStr = Replace(CusCliStr," AND "," ")
	sqlStr = "SELECT * FROM [HoldingExceptions] WHERE " & NomCliStr & " AND " & CusCliStr & " ORDER BY SecurityName,Code" 
else
	sqlStr = "SELECT * FROM [HoldingExceptions] WHERE Security_DPA_ = " & selectedSecurity & " " & NomCliStr & " " & CusCliStr & " ORDER BY Code" 
end if 

'Response.Write sqlStr
'Response.End

'Conn.execute("ClientTotalsDelete")		 
'Conn.execute("ClientBalancesDelete")		 
'Conn.execute("ClientTotalsProcedure")		 
'Conn.execute("ClientBalancesProcedure")		 

set rs = conn.Execute(sqlStr)

if Cstr(selectedSecurity) = "0" then

Dim HArray
	
If rs.EOF Or rs.BOF Then%>
	<Script Language="JavaScript">
		alert("There are no records based on the specified criterion.")
		window.parent.history.go(-1);			
	</Script>
	<%Set Rs = Nothing
	Set Conn = Nothing
	Response.End
End If

If Not (rs.EOF Or rs.BOF) Then		
	HoldArray = rs.GetRows()
	RowNum = rs.RecordCount - 1
End If

function roundrobin ()
	psecname = ""
		for k = 0 to RowNum
					
			SecurityName = HoldArray(2,k)
					
			if SecurityName <> psecname then
				HArray = HoldArray
				SecCount = GetSecurityCount(k+1,SecurityName)
				Endat = k + SecCount
				HArray = HoldArray
				call DoCreatePage(k,Endat,SecurityName)
			end if
			psecname = SecurityName
		next
end function
		
function GetSecurityCount(startat, SecurityName)
		
	counter = 0
		
	for k = startat to RowNum
		if SecurityName = HArray(2,k) then
			counter = counter + 1
		else
			exit for	
		end if
	next
			
	GetSecurityCount = counter
	Set HArray = Nothing
		
end function
		
call roundrobin()
		
		function DoCreatePage(startat, endat, security)	
		%>
		<BR class="newpage">
		<p id="toPDFOrient" name="toPDFOrient" value="" style="display:none;">&nbsp;
		<p id="toPDF" name="toPDF">
		<table border="0" cellspacing="2" cellpadding="2" style="font-family: Arial Narrow" width="639">
			<tr>
				<td bgcolor="#000000" width="80%" nowrap align="left"><font color="#FFFFFF" face="Impact" size="2">HOLDINGS EXCEPTIONS</font></td>
				<td bgcolor="#000000" width="20%" nowrap align=right><font color="#FFFFFF" face="Impact" size="2"><%= Session("CompanyName") %></font></td>
			</tr>
		</table>

		<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" width="635" > 
			<tr>      
				<td width="627"><b>Date:&nbsp;<%= FormatDate(Date) %></b></td>      			
			</tr>
			<tr>
				<td style="border: 1px solid #000000" valign="top">
					<table cellspacing=0 cellpadding=0 border=0>		
						<tr>
							<td nowrap>SECURITY:&nbsp;<%= security %> </td>					
						</tr>					
					</table>
				</td>
			</tr>
			<tr>
				<td align="right" height="8">&nbsp;</td>		
			</tr>
		</table>

		<table border="0" cellspacing=2 cellpadding=2 width="635">
		  <tr> 
		    <td align="left" colspan="2" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="3"><b>&nbsp;</b></font></td>
		    <td align="left" colspan="4" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="3"><b>HOLDINGS</b></font></td>
		  </tr>
		  <tr> 
		    <td width="69" align="left" valign="top" style="border: 1px solid #000000"><font face="Arial Narrow" size="2"><b>Code&nbsp;</b></font></td>
		    <td width="207" align="left" valign="top" style="border: 1px solid #000000"><font face="Arial Narrow" size="2"><b>Client&nbsp;</b></font></td>
		    <td width="92" align="Right" valign="top" style="border: 1px solid #000000"><font face="Arial Narrow" size="2"><b>System</b></font></td>
		    <td width="91" align="Right" valign="top" style="border: 1px solid #000000"><font face="Arial Narrow" size="2"><b>CDS</b></font></td>
		    <td width="134" align="Right" valign="top" style="border: 1px solid #000000"><font face="Arial Narrow" size="2"><b>Difference</b></font></td>
		  </tr>
		  <%
			Dim clicode, cliname, sysqty, cdsqty, cdssysdiff
			Dim totalsysqty, totalcdsqty, totalcdssysdiff
			
			totalsysqty = 0
			totalcdsqty = 0
			totalcdssysdiff = 0
			
			for i = startat to endat	
				clicode = HArray(0,i)
				cliname = HArray(1,i)
				sysqty = clng(HArray(4,i))
				cdsqty = clng(HArray(3,i))
				cdssysdiff = clng(HArray(5,i))
				
				if (cbool(HoldExp) = true) and (cdssysdiff = 0) then
				  Response.Write("")
				else  
				 %>
				  <tr> 
					<td nowrap><font face="Arial Narrow" size="2"><%=clicode%></font></td>
					<td nowrap><font face="Arial Narrow" size="2"><%=cliname%></font></td>
					<td align="Right" nowrap><font face="Arial Narrow" size="2"><%=FormatNumber(sysqty,0)%></font></td>
					<td align="Right" nowrap><font face="Arial Narrow" size="2"><%=FormatNumEx(cdsqty,0)%></font></td>
					<td align="Right" nowrap><font face="Arial Narrow" size="2"><%=FormatNumEx((cdssysdiff),0)%></font></td>
				  </tr>
				  
			<%
			end if
						i = i + 1
						totalsysqty = sysqty + totalsysqty
						totalcdsqty = cdsqty + totalcdsqty
						totalcdssysdiff = cdssysdiff + totalcdssysdiff
			 next
			 Set HArray = Nothing
			%>

		  <tr> 
		    <td align="right" colspan="9"><font face="Arial Narrow" size="3">&nbsp;</font></td>
		  </tr>
		   <tr> 
		    <td align="right" colspan="2"><font face="Arial Narrow" size="2"><b>TOTAL</b></font></td>
		    <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b><%=FormatNumEx(totalsysqty,0)%></b></font></td>
		    <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b><%=FormatNumEx(totalcdsqty,0)%></b></font></td>
		    <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b><%=FormatNumEx(totalcdssysdiff,0)%></b></font></td>
		    <td width="4" align="Right"  valign="top"><font face="Arial Narrow" size="2">&nbsp;</font></td>
		  </tr>
		</table>
		<%end function%>
<%
Else

		If rs.EOF Or rs.BOF Then%>
			<Script Language="JavaScript">
				alert("There are no records based on the specified criterion.")
				window.parent.history.go(-1);			
			</Script>
			<%Set Rs = Nothing
			Set Conn = Nothing
			Response.End
		End If
			
		If Not (rs.EOF Or rs.BOF) Then		
			SecurityName = rs.fields("SecurityName")
		End If

		%>
		<p id="toPDFOrient" name="toPDFOrient" value="" style="display:none;">&nbsp;
		<p id="toPDF" name="toPDF">
		<table border="0" cellspacing="2" cellpadding="2" style="font-family: Arial Narrow" width="635">
			<tr>
				<td bgcolor="#000000" width="80%" nowrap align="left"><font color="#FFFFFF" face="Impact" size="2">HOLDINGS EXCEPTIONS</font></td>
				<td bgcolor="#000000" width="20%" nowrap align=right><font color="#FFFFFF" face="Impact" size="2"><%= Session("CompanyName") %></font></td>
			</tr>
		</table>

		<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" width="634" > 
			<tr>      
				<td width="626"><b>Date:&nbsp;<%= FormatDate(Date) %></b></td>      			
			</tr>
			<tr>
				<td style="border: 1px solid #000000" valign="top">
					<table cellspacing=0 cellpadding=0 border=0>		
						<tr>
							<td nowrap>SECURITY:&nbsp;<%= SecurityName %> </td>					
						</tr>					
					</table>
				</td>
			</tr>
			<tr>
				<td align="right" height="8">&nbsp;</td>		
			</tr>
		</table>

		
<table border="0" cellspacing=2 cellpadding=2 width="634">
  <tr> 
    <td align="left" colspan="2" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="3"><b>&nbsp;</b></font></td>
    <td align="left" colspan="4" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="3"><b>HOLDINGS</b></font></td>
  </tr>
  <tr> 
    <td width="75" align="left" valign="top" style="border: 1px solid #000000"><font face="Arial Narrow" size="2"><b>Code&nbsp;</b></font></td>
    <td width="197" align="left" valign="top" style="border: 1px solid #000000"><font face="Arial Narrow" size="2"><b>Client&nbsp;</b></font></td>
    <td width="102" align="Right" valign="top" style="border: 1px solid #000000"><font face="Arial Narrow" size="2"><b>System</b></font></td>
    <td width="107" align="Right" valign="top" style="border: 1px solid #000000"><font face="Arial Narrow" size="2"><b>CDS</b></font></td>
    <td width="112" align="Right" valign="top" style="border: 1px solid #000000"><font face="Arial Narrow" size="2"><b>Difference</b></font></td>
  </tr>
  <%
			Dim clicode, cliname, sysqty, cdsqty, cdssysdiff
			Dim totalsysqty, totalcdsqty, totalcdssysdiff
			Dim HoldArray 
			
			If Not (rs.EOF Or rs.BOF) Then		
				HoldArray = rs.GetRows()
				RowNum = rs.RecordCount - 1
			End If

			totalsysqty = 0
			totalcdsqty = 0
			totalcdssysdiff = 0

			for i = 0 to RowNum

				clicode = HoldArray(0,i)
				cliname = HoldArray(1,i)
				sysqty = clng(HoldArray(4,i))
				cdsqty = clng(HoldArray(3,i))
				'cdssysdiff = clng(HoldArray(5,i))
				cdssysdiff = cdsqty - sysqty
				if (cbool(HoldExp) = true) and (cdssysdiff = 0) then
				  	Response.Write("")
				else  
				 %>
					  <tr> 
						<td nowrap><font face="Arial Narrow" size="2"><%=clicode%></font></td>
						<td nowrap><font face="Arial Narrow" size="2"><%=cliname%></font></td>
						<td align="Right" nowrap><font face="Arial Narrow" size="2"><%=FormatNumber(sysqty,0)%></font></td>
						<td align="Right" nowrap><font face="Arial Narrow" size="2"><%=FormatNumEx(cdsqty,0)%></font></td>
						<td align="Right" nowrap><font face="Arial Narrow" size="2"><%=FormatNumEx((cdssysdiff),0)%></font></td>
					  </tr>
				 <%
						totalsysqty = sysqty + totalsysqty
						totalcdsqty = cdsqty + totalcdsqty
						totalcdssysdiff = cdssysdiff + totalcdssysdiff
				 end if
						i = i + 1
			 next
			%>
  <tr> 
    <td align="right" colspan="9"><font face="Arial Narrow" size="3">&nbsp;</font></td>
  </tr>
  <tr> 
    <td align="right" colspan="2"><font face="Arial Narrow" size="2"><b>TOTAL</b></font></td>
    <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b><%=FormatNumEx(totalsysqty,0)%></b></font></td>
    <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b><%=FormatNumEx(totalcdsqty,0)%></b></font></td>
    <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b><%=FormatNumEx(totalcdssysdiff,0)%></b></font></td>
    <td width="3" align="Right"  valign="top"><font face="Arial Narrow" size="2">&nbsp;</font></td>
  </tr>
</table>
<%End if%>
</body>
<%Set Rs = Nothing%>
<%Set Conn = Nothing%>
<%Set HArray = Nothing%>
<%Set HoldArray = Nothing%>
</html>

