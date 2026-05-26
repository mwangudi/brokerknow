<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Sale Orders</title>
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
	
		 sqlStr = "SELECT * FROM TotalTradingSchedule WHERE ((Validity >= { fn CURDATE() }) OR " & _
             "(Validity IS NULL)) ORDER BY Code,Client,Order_DPA_"
    
    	Set conn = GetActiveConnection("KBroker")
		 Set rstotal = Server.CreateObject("ADODB.Recordset")
		 Set rsCurrent = Server.CreateObject("ADODB.Recordset")		
		 
		 rstotal.CursorLocation = adUseClient
		 rsCurrent.CursorLocation = adUseClient
		
		 'Response.write(sqlStr)
		 'Response.end
		 Set conn = GetActiveConnection("KBroker")
		 
		 Conn.execute("ClientTotalsDelete")		 
		 Conn.execute("ClientBalancesDelete")		 
			
		 Conn.execute("ClientTotalsProcedure")		 
		 Conn.execute("ClientBalancesProcedure")		 			 
		 	       
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
		<td nowrap><b><font face="Arial Narrow" size="4">Sale Orders</font></b></td>
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
	  <td bgcolor="#000000" width="30"><b><font color="#FFFFFF" face="Arial Narrow">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></b></td>
      <td bgcolor="#000000" width="50"><b><font color="#FFFFFF" face="Arial Narrow">Security</font></b></td>            
      <td bgcolor="#000000" width="50"><b><font color="#FFFFFF" face="Arial Narrow">Order</font></b></td>
      <td bgcolor="#000000" width="70"><b><font color="#FFFFFF" face="Arial Narrow">Order&nbsp;Date</font></b></td>
      <td bgcolor="#000000" align="right" width="50"><b><font color="#FFFFFF" face="Arial Narrow">Quantity</font></b></td>
      <td bgcolor="#000000" align="right" width="50"><b><font color="#FFFFFF" face="Arial Narrow">Total</font></b></td>
      <td bgcolor="#000000" align="right" width="50"><b><font color="#FFFFFF" face="Arial Narrow">Balance</font></b></td>
      <td bgcolor="#000000" align="right" width="50"><b><font color="#FFFFFF" face="Arial Narrow">Excess</font></b></td>
    </tr>
    <tr>
      <td colspan=14 >&nbsp; </td>
    </tr>
        
    <%
    Security1=""
    Security2=""
    
    Do Until groupRs.EOF   
    
	OrdDetailSecurity = groupRs.Fields("code").Value 
	
	if groupRs.Fields("OrdDetailType")="Sale" then    	
	Security1=OrdDetailSecurity 
			
	if Trim(Security1) <> Trim(Security2) then
	
	%>		
    <tr>
      <td align="right" width="30"><b><%= OrdDetailSecurity %></b></td>
      <td colspan="13" width="970"><b><%= groupRs.Fields("Client") %></b></td>
    </tr>    
    <%
    end if
    %>
			<tr>
			  <td width="30">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
			  <td width="44"><%= groupRs.Fields("SecurityCode").Value %></td>			  				
			  <td width="60"><%= groupRs.Fields("Order_DPA_").Value %></td>
			  <td width="100"><%= FormatDate(groupRs.Fields("OrderDate").Value) %></td>
			  <td align="right" width="50"><%= FormatNumber(groupRs.Fields("BalanceQty").Value,0)%> </td>
			  
			  <td align="right" width="50"><%= FormatNumber(groupRs("Total"),0) %> </td>
			  <% if(groupRs("BalanceFree")="Y") then%>
			  <td align="right" width="50"><%= FormatNumber(groupRs("Balance"),0) %> </td>
			  <% else
				%>
			  <td align="right" width="50"><%= FormatNumber(groupRs("Balance"),0) %><b>*</b></td>
			  <%
			  end if
			  %>
			   <td align="right" width="50"><%= FormatNumber(groupRs("Excess"),0) %> </td>
			</tr>
    <%
    End if 'End of Excess
	Security2=Security1
		groupRs.MoveNext
    Loop
    
    Set groupRs = Nothing
    Set Conn = Nothing
    %>

    <tr>
      <td width="30">&nbsp;</td>
    </tr>
    <tr>
      <td width="30">&nbsp;</td>
    </tr>
  </table> 

</body>

</html>