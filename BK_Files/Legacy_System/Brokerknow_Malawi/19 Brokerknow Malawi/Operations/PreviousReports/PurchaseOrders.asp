<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Purchase Orders</title>
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
	
		 sqlStr = "SELECT * FROM Db_FineTradingSchedule WHERE ((Validity >= { fn CURDATE() }) OR " & _
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
		<td nowrap><b><font face="Arial Narrow" size="4">Purchase Orders</font></b></td>
		<td nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>
	</tr>	
    <tr>
	   <td COLSPAN=2><font face="Arial" size="2">As Of:  <%= headerDescription %></font></td>
	</tr>
    <tr>
		  <td COLSPAN=2><font face="Arial" size="2">&nbsp;</font></td>
	</tr>
</table>				

  <table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="1000">        
    <tr>
	  <td bgcolor="#000000" width="30"><b><font color="#FFFFFF" face="Arial Narrow">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></b></td>
      <td bgcolor="#000000" width="44"><b><font color="#FFFFFF" face="Arial Narrow">Security</font></b></td>            
      <td bgcolor="#000000" width="60"><b><font color="#FFFFFF" face="Arial Narrow">Order</font></b></td>
      <td bgcolor="#000000" width="26"><b><font color="#FFFFFF" face="Arial Narrow">Type</font></b></td>
      <td bgcolor="#000000" width="20"><b><font color="#FFFFFF" face="Arial Narrow">Sec</font></b></td>      
      <td bgcolor="#000000" width="100"><b><font color="#FFFFFF" face="Arial Narrow">Order&nbsp;Date</font></b></td>
      <td bgcolor="#000000" align="right" width="50"><b><font color="#FFFFFF" face="Arial Narrow">Quantity</font></b></td>
      <td bgcolor="#000000" align="right" width="74"><b><font color="#FFFFFF" face="Arial Narrow">Price</font></b></td>
      <td bgcolor="#000000" align="right" width="74"><b><font color="#FFFFFF" face="Arial Narrow">Cost</font></b></td>
      <td bgcolor="#000000" align="right" width="75"><b><font color="#FFFFFF" face="Arial Narrow">Total</font></b></td>
      <td bgcolor="#000000" align="right" width="75"><b><font color="#FFFFFF" face="Arial Narrow">Current&nbsp;Balance</font></b></td>
      <td bgcolor="#000000" align="right" width="75"><b><font color="#FFFFFF" face="Arial Narrow">Working&nbsp;Balance</font></b></td>
      <td bgcolor="#000000" align="right" width="75"><b><font color="#FFFFFF" face="Arial Narrow">Credit&nbsp;Limit</font></b></td>
      <td bgcolor="#000000" align="right" width="75"><b><font color="#FFFFFF" face="Arial Narrow">Excess</font></b></td>
    </tr>
    <tr>
      <td colspan=14 width="1000">&nbsp; </td>
    </tr>
        
    <%
    Security1=""
    Security2=""
    
    Do Until groupRs.EOF   
    
	OrdDetailSecurity = groupRs.Fields("code").Value 
	
	if groupRs.Fields("OrdDetailType")="Purchase" then    	
	Security1=OrdDetailSecurity 
			
	if Trim(Security1) <> Trim(Security2) then
	
	%>		
    <tr>
      <td align="right" width="30"><b><%= OrdDetailSecurity %></b></td>
      <td colspan="13" width="970"><b><%= groupRs.Fields("Client") %></b></td>
    </tr>    
    <%
    end if
    	if(Trim(groupRs.Fields("OrdDetailType").Value)="Purchase") then
    	DetailType="P"
    	else
    	DetailType="S"
    	end if
    	
    	if(Trim(groupRs.Fields("OrdDetailSecType").Value)="Security") then
    	DetailSecType="S"
    	else
    	DetailSecType="F"
    	end if
    %>
			<tr>
			  <td width="30">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
			  <td width="44"><%= groupRs.Fields("SecurityCode").Value %></td>			  				
			  <td width="60"><%= groupRs.Fields("Order_DPA_").Value %></td>
			  <td width="26"><%= DetailType %></td>
			  <td width="20"><%= DetailSecType %></td>			  
			  <td width="100"><%= FormatDate(groupRs.Fields("OrderDate").Value) %></td>
			  <td align="right" width="50"><%= groupRs.Fields("BalanceQty").Value %> </td>
			  	<% if(groupRs("Best")=true or groupRs.Fields("OrdDetailPrice")="BEST") then
			  	%>
			  	<td align="right" width="74"><%= FormatNum(groupRs.Fields("Price").Value * 1.105) %> </td>
			  	<td align="right" width="74"><%= FormatNum(groupRs.Fields("BalanceQty")*groupRs.Fields("price").Value * 1.105) %> </td>			  	
			  	<%
			  	else
			  	%>
			  	<td align="right" width="74"><%= FormatNum(groupRs.Fields("OrdDetailPrice").Value) %> </td>
			  	<td align="right" width="74"><%= FormatNum(groupRs.Fields("BalanceQty")*groupRs.Fields("OrdDetailPrice")) %> </td>
			  	<% end if%>
			  <td align="right" width="75"><%= FormatNum(groupRs("Total")) %> </td>
			  <td align="right" width="75"><%= FormatNum(groupRs("CurrentBal")) %> </td>
			  <td align="right" width="75"><%= FormatNum(groupRs("CurrentBal")-groupRs("Total")) %> </td>
			  <td align="right" width="75"><%= FormatNum(-(groupRs("CreditLimit"))) %> </td>
			  <td align="right" width="75"><%= FormatNum(groupRs("Excess")) %> </td>
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