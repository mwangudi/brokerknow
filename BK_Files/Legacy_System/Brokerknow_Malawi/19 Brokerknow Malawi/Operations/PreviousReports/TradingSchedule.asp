<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Trading Schedule</title>
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
	 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
	 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
	<!--CALENDAR -->
	<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
	<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>

	<style media="print">	
		@page domes {
			size: portrait;
			margin-left: 2cm;
			margin-right: 5cm;
			margin-top: 1cm;    
			margin-bottom: 2cm;
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
	
		 'sqlStr = "SELECT distinct OrdDetailSecurity,Order_DPA_, OrdDetailSecType,Code,OrdDetailType,Best,Excess FROM [FineTradingSchedule] WHERE(Validity >= { fn CURDATE() }) OR " & _
          '        "(Validity IS NULL) ORDER BY OrdDetailSecurity, OrdDetailSecType "
				 
		 sqlStr = "SELECT * FROM DB_FineTradingSchedule WHERE ((Validity >= { fn CURDATE() }) OR " & _
             "(Validity IS NULL)) ORDER BY ordDetailSecurity, OrdDetailSecType ,OrdDetailType DESC, OrderDate"
    		
		 'sqlStr = "SELECT * FROM Wanjaus_TradingSchedule WHERE ((Validity >= { fn CURDATE() }) OR " & _
             '"(Validity IS NULL)) ORDER BY ordDetailSecurity, OrdDetailSecType ,OrdDetailType DESC, OrderDate"
    		
'Wanjaus_TradingSchedule
		 'Response.write(sqlStr)
		 'Response.end
		 Set conn = GetActiveConnection("KBroker")
		 
		 Conn.execute("ClientTotalsDelete")		 
		 Conn.execute("ClientBalancesDelete")		 
			
		 Conn.execute("ClientTotalsProcedure")		 
		 Conn.execute("ClientBalancesProcedure")		 
	
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
i=0
headerDescription = FormatDateFull(selectedTradeDate) %>
<i id="landRem">Remember to select landscape settings while printing.</i>
<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
     <tr>
		<td nowrap><b><font face="Arial Narrow" size="4">Trading Schedule</font></b></td>
		<td nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>
	</tr>	
    <tr>
	   <td COLSPAN=2><font face="Arial" size="2">for Deals traded on:  <%= headerDescription %></font></td>
	</tr>
    <tr>
		  <td COLSPAN=2><font face="Arial" size="2">&nbsp;</font></td>
	</tr>
</table>				

  <table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="946">
    <tr>
	  <td align="right" width="80"><b><font  face="Arial Narrow">CODE</font></b></td>
      <td  width="272"><b><font  face="Arial Narrow">CLIENT</font></b></td>      
      <td  width="52" align="right"><b><font  face="Arial Narrow">ORDER</font></b></td>
      <td  width="55"><b><font  face="Arial Narrow">TYPE</font></b></td>
      <td  width="100"><b><font  face="Arial Narrow">AGENT</font></b></td>
      <td  width="90"><b><font  face="Arial Narrow">ORDER&nbsp;DATE</font></b></td>
      <td  align="right" width="59"><b><font  face="Arial Narrow">QUANTITY</font></b></td>
      <td  align="right" width="52"><b><font  face="Arial Narrow">PRICE</font></b></td>
    </tr>
    <tr>
      <td colspan="8" width="737">&nbsp; </td>
    </tr>
        
    <%    
    Security1=""
    Security2=""
    
    Do Until groupRs.EOF       	
		
    if (Cdbl(groupRs("Excess"))>0) and groupRs.Fields("OrdDetailType")="Purchase" then    
    
    else
    'Response.write(Security1)
    i=i+1
    OrdDetailSecurity = groupRs.Fields("OrdDetailSecurity").Value 
	Security1=OrdDetailSecurity    
		if Trim(Security1) <> Trim(Security2) then		
		%>		
    	<tr>
      	<td width="80"><b><%= groupRs.Fields("OrdDetailSecType").Value %></b></td>
      	<td colspan="5" width="533"><b><%= OrdDetailSecurity %></b></td>
    	</tr>
    	<%
    	end if
    	%>
    		<tr>
			  <td align="right" width="80"><%= groupRs.Fields("ClientCDSNO").Value %></td>
			  <td width="272"><%= groupRs.Fields("Client").Value %></td>			  				
			  <td align="right" width="52"><%= groupRs.Fields("Order_DPA_").Value %></td>			  
			  <td width="55"><%= groupRs.Fields("OrdDetailType").Value %></td>
			  <td width="100"><%= groupRs.Fields("AgentName").Value %></td>
			  <td width="90"><%= FormatDate(groupRs.Fields("OrderDate").Value) %></td>
			  <td align="right" width="59"><%= FormatNumEx(groupRs.Fields("BalanceQty").Value,0) %> </td>
			  <% if (groupRs.Fields("Best") = true) then %>
			  <td align="right" width="52">BEST</td>
				<% else %>
			  <td align="right" width="52"><%= FormatNum(groupRs.Fields("OrdDetailPrice").Value) %> </td>
			</tr>
    			<%	
				   End if
	End if 'End of Excess
	Security2=Security1
		groupRs.MoveNext
    Loop
    
    Set groupRs = Nothing
    Set Conn = Nothing
    %>

    <tr>
      <td colspan="6" width="617">&nbsp;</td>
    </tr>
    <tr>
      <td colspan="6" width="617">&nbsp;</td>
    </tr>
  </table>
  

</body>

</html>