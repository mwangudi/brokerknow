<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Securities Volme By Client</title>
	<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
	<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
	<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
	<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
	
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

If genReport <> "1" Then%>
	<Script Language="JavaScript">
		function validateForm(frm){					
			frm.target = '_self';			
			frm.submit();
		}
		document.body.className = 'dialog';
		
	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="ClientVolumeTotals.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">
		<table>						
			<tr>
				<td colspan=2><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... ">&nbsp;&nbsp; <input type="Button" class="Buttons" Value=" Close " OnClick="JavaScript: window.parent.self.close();"></td>
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
	sqlStr = "SELECT * FROM ClientVolumeTotals"
	Rs.CursorLocation = adUseClient	
	Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
	'Rs.Filter = "Client_DPA_ LIKE '" & selectedClient & "' AND TransDate >= '" & FormatDate(selectedFromDate) & "'"
	
	
	If rs.EOF Or rs.BOF Then%>
		<Script Language="JavaScript">
			alert("There are no transactions in the system")
			window.history.go(-1);
		</Script>
		<%Set Rs = Nothing
		Set Conn = Nothing
		Response.End
	End If
	
	
	Dim SecurityName
	'OpeningBalance =  Rs.Fields("Balance").Value	
	SecurityName = Rs.Fields("SecurityName").Value  	
		
Do Until Rs.EOF 
%> 
<table border="0" cellspacing="2" cellpadding="2" style="font-family: Arial Narrow" width="100%">
    <tr>
		<td width="10%" nowrap><font face="Impact" size="4">Security Volumes By Client</font></td>
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
      <td width="1%"><b>Security:</b></td>
      <td width="48%"><b><%= SecurityName %></b></td>
    </tr>

    <tr>
      <td width="1%"><b><font size="2" face="Arial">&nbsp;</font></b></td>
      <td width="48%"><%= accountAddress %></td>
    </tr>

</table>
<BR>


  <table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow; LEFT-MARGIN:100PX"  width="100%">
    <tr>
      <td style="border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"><b><font face="Arial Narrow" size="3">Client Code:</font></b></td>
      <td style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"><b><font face="Arial Narrow" size="3">Client Name:</font></b></td>
      <td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"><b><font face="Arial Narrow" size="3">Volume:</font></b></td>
      <td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"  align=right><b><font face="Arial Narrow" size="3">Gross:</font></b></td>
    </tr>
    
<%
    Dim totalVolume
    Dim totalGross
    
    totalVolume = 0
    totalGross = 0
    nextEntityFound = False    
    IsOpeningBalance = True
    
    Do Until nextEntityFound
		totalVolume = totalVolume + Rs.Fields("Volume").Value 
		totalGross = totalGross + Rs.Fields("Gross").Value%>
		<tr>	
		  <td><%= Rs.Fields("ClientCode").Value %></td>
		  <td><%= Rs.Fields("ClientName").Value %></td>
		  <td align="right"><%= FormatNum(Rs.Fields("Volume").Value) %></td>
		  <td align="right"><%= FormatNum(Rs.Fields("Gross").Value) %></td>
		</tr>
	
	<%	Rs.MoveNext
		If Not Rs.EOF  Then
			If StrComp(SecurityName, Rs.Fields("SecurityName").Value) <> 0 Then
				SecurityName = Rs.Fields("SecurityName").Value
				nextEntityFound = True				
			End If   
		Else
			nextEntityFound = True	
		End If
		
	Loop
	%>
	
		<tr>	
		  <td>&nbsp;</td>
		  <td>&nbsp;</td>
		  <td align="right">&nbsp;</td>
		  <td align="right">&nbsp; </td>
		  <td align="right">&nbsp;</td>
		</tr>
		<tr>	
		  <td>Total </td>
		  <td>&nbsp;</td>
		  <td align="right"><%=FormatNum(totalVolume)%></td>
		  <td align="right"><%=FormatNum(totalGross)%> </td>
		  <td align="right">&nbsp;</td>
		</tr>  
	
    <tr>
      <td colspan="6" align="right" style="border-bottom-style: solid; border-bottom-width: 1">
        &nbsp;&nbsp;&nbsp; </td>

    </tr>

   

  </table>
  
  <%If Not Rs.EOF Then %>
			<BR class="newpage">
	<%	End If
   

Loop

Set Rs = Nothing
Set Conn = Nothing%>   
</body>

</html>
