<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Expired Orders</title>
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
	 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
	 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
	<!--CALENDAR -->
	<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
	<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>

	<style media="print">	
		@page {
				size: landscape;
				margin-left: 2cm;
				margin-right: 5cm;
				margin-top: 1cm;    
				margin-bottom: 2cm;
				writing-mode: tb-rl;
				height: 80%;
				margin: 10% 0%;						
				br.newpage{
					page-break-before:always;
				}		
			}		 
	</style>


</head>

<body Class="Reports">
<Script Language="JavaScript">
	function HideRemindSelectLandscape(){
		try{			
			document.getElementById('landRem').style.display = 'none';
		}	
		catch(e){}	
	}
	function ShowRemindSelectLandscape(){
		try{			
			document.getElementById('landRem').style.display = '';
		}	
		catch(e){}	
	}
	window.onload = HideRemindSelectLandscape;
	//window.onbeforeprint = HideRemindSelectLandscape;
	//window.onafterprint = ShowRemindSelectLandscape;
</Script>

<!--#include file="../libroutines.asp"-->

<%
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim rstotal
	Dim rsCurrent
	Dim sqlStr1
	Dim sqlStr2
	Dim Excess	
	Dim Security1
	Dim Security2
	
		 sqlStr = "SELECT distinct OrdDetailSecurity,Order_DPA_, OrdDetailSecType,Code,OrdDetailType,Client FROM [FineTradingSchedule] WHERE(Validity < { fn CURDATE() }) AND " & _
                  "(Validity IS NOT NULL) ORDER BY Client"
				 
		 Set conn = GetActiveConnection("KBroker")
		 Set rstotal = Server.CreateObject("ADODB.Recordset")
		 Set rsCurrent = Server.CreateObject("ADODB.Recordset")		
		 
		 rstotal.CursorLocation = adUseClient
		 rsCurrent.CursorLocation = adUseClient
			       
	    Set groupRs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
	    If groupRs.EOF Or groupRs.BOF Then %>
				<Script Language="JavaScript">
					alert("No orders available");
					window.parent.close();					
	            </Script>
	            <% Set groupRs = Nothing
	            Set Conn = Nothing
	            Response.End
	    End If
		        
	    groupRs.MoveFirst
		        
selectedTradeDate = FormatDate(Date)

DrawPageFunctions True, True, False

headerDescription = FormatDateFull(selectedTradeDate) %>
<i id="landRem">Remember to select landscape settings while printing.</i>
<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
     <tr>
		<td nowrap><b><font face="Arial Narrow" size="4">Expired Orders</font></b></td>
		<td nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>
	</tr>	
    <tr>
	   <td COLSPAN=2><font face="Arial" size="2">As Of:  <%= headerDescription %></font></td>
	</tr>
    <tr>
		  <td COLSPAN=2><font face="Arial" size="2">&nbsp;</font></td>
	</tr>
</table>				

  <table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">	 
    <tr>
	  <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></b></td>
      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow">Security&nbsp;Name&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></b></td>            
      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow">Order&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></b></td>
      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow">Type&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></b></td>
      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow">Sec&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></b></td>
      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow">Security&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></b></td>
      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow">Order&nbsp;Date&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></b></td>
      <td bgcolor="#000000" align="Right"><b><font color="#FFFFFF" face="Arial Narrow">&nbsp;&nbsp;&nbsp;&nbsp;Quantity</font></b></td>
      <td bgcolor="#000000" align="right"><b><font color="#FFFFFF" face="Arial Narrow">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Price</font></b></td>
      <td bgcolor="#000000" align="left"><b><font color="#FFFFFF" face="Arial Narrow">Expired&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></b></td>
    </tr>
    <tr>
      <td colspan="8">&nbsp; </td>
    </tr>
        
    <%
    Security1=""
    Security2=""
    
    Do Until groupRs.EOF   
    
	OrdDetailSecurity = groupRs.Fields("Code").Value 
	Security1=OrdDetailSecurity
    
    if Trim(Security1) <> Trim(Security2) then	
	%>		
    <tr>
      <td align="Right"><b><%= OrdDetailSecurity %></b></td>
      <td colspan="5"><b><%= groupRs.Fields("Client").Value %></b></td>
    </tr>
    
    <%
    sqlStr = "SELECT * FROM FineTradingSchedule WHERE Code = '" & OrdDetailSecurity & "' AND ((Validity < { fn CURDATE() }) AND " & _
             "(Validity IS NOT NULL)) ORDER BY OrdDetailType DESC, OrderDate"
    	
    Set Rs = Conn.Execute (SQLServerFormat(HandleQuote(sqlStr)))
    
    If Not (Rs.EOF Or Rs.BOF) Then
    'Rs.Fields("OrdDetailSecType").Value
		Do Until Rs.EOF%>
			<tr>
			  <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
			  <td align="left"><%= Rs.Fields("OrdDetailSecurity").Value %></td>			  				
			  <td align="left"><%= Rs.Fields("Order_DPA_").Value %></td>
			  <td><%= Rs.Fields("OrdDetailType").Value %></td>
			  <td><%= Rs.Fields("OrdDetailSecType").Value %></td>
			  <td><%= Rs.Fields("SecurityCode").Value %></td>
			  <td><%= FormatDate(Rs.Fields("OrderDate").Value) %></td>
			  <td align="right"><%= Rs.Fields("BalanceQty").Value %> </td>
			  <td align="right"><%= FormatNum(Rs.Fields("OrdDetailPrice").Value) %> </td>
			  <td align="left"><%= FormatDate(Rs.Fields("Validity").Value) %> </td>
			</tr>
    
    <%		Rs.MoveNext
		Loop
	End If
	Set Rs = Nothing
	End if	
	Security2=Security1
		groupRs.MoveNext
    Loop
    
    Set groupRs = Nothing
    Set Conn = Nothing
    %>

    <tr>
      <td colspan="6">&nbsp;</td>
    </tr>
    <tr>
      <td colspan="6">&nbsp;</td>
    </tr>
  </table>
  

</body>

</html>
