<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Allocation Schedule</title>
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
	
		 sqlStr = "SELECT * FROM DB_FineTradingSchedule WHERE ( (Validity >= { fn CURDATE() }) OR " & _
             "(Validity IS NULL)) ORDER BY ordDetailSecurity, OrdDetailSecType ,OrdDetailType DESC, OrderDate"
    	
				 
		 Set conn = GetActiveConnection("KBroker")
		 Conn.execute("ClientTotalsDelete")		 
		 Conn.execute("ClientBalancesDelete")		 
			
		 Conn.execute("ClientTotalsProcedure")		 
		 Conn.execute("ClientBalancesProcedure")		 
	
		 'Set rstotal = Server.CreateObject("ADODB.Recordset")
		 'Set rsCurrent = Server.CreateObject("ADODB.Recordset")				 
		 
		 'rstotal.CursorLocation = adUseClient
		 'rsCurrent.CursorLocation = adUseClient
		       
	    Set Rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
	    If Rs.EOF Or Rs.BOF Then %>
				<Script Language="JavaScript">
					alert("No Allocation available");
					window.parent.close();					
	            </Script>
	            <% Set groupRs = Nothing
	            Set Conn = Nothing
	            Response.End
	    End If
		        
	    Rs.MoveFirst
		        
selectedTradeDate = FormatDate(Date)

DrawPageFunctions True, True, False

headerDescription = FormatDateFull(selectedTradeDate) %>
<i id="landRem">Remember to select landscape settings while printing.</i>
<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
     <tr>
		<td nowrap><b><font face="Arial Narrow" size="4">Allocation Schedule</font></b></td>
		<td nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>
	</tr>	
    <tr>
	   <td COLSPAN=2><font face="Arial" size="2">AS Of:  <%= headerDescription %></font></td>
	</tr>
    <tr>
		  <td COLSPAN=2><font face="Arial" size="2">&nbsp;</font></td>
	</tr>
</table>
<%
    Security1=0
    Security2=0
    saletype1=0
    saletype2=0
    first=1

    If Not (Rs.EOF Or Rs.BOF) Then
		Do Until Rs.EOF
		Security1=Cint(Rs("Security_DPA_"))
    		if(Rs("OrderTypeSale")=true) then 
			saletype1=1
			else 
			saletype1=0
		end if
if (Cdbl(Rs("Excess"))>0) and Rs.Fields("OrdDetailType")="Purchase" then    
    
    else
    
		if(Security1<>Security2) then
            %> 
    
    <table> 
    <tr><td colpan="10">&nbsp;</td></tr> </table>
    <tr><td colpan="10"><b><font size=2><%=rs("ordDetailSecurity")%></font></b></td></tr>    
    <table border="1" cellspacing="0" cellpadding="1" style="font-family: Arial Narrow" width="900">					
    <tr>
	<td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow">Client</font></b></td>
      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow">Agent</font></b></td>      
      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow">Order</font></b></td>
      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow">Security</font></b></td>
      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow">Price</font></b></td>
      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow">Slip</font></b></td>
      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow">Time</font></b></td>
      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow">Price</font></b></td>
      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow">Trade Qty</font></b></td>
      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow">Broker</font></b></td>
    </tr> 
	<% end if
		if(saletype1<>saletype2 and first=0 and security1=security2) then %>
			</table>
			<table>
			<tr><td colpan="10">&nbsp;</td></tr>
			</table>
			<table border="1" cellspacing="0" cellpadding="1" style="font-family: Arial Narrow" width="900">		

		<%
		end if
		%>  
		     
    		<tr>
			<td width="200"><table border="0" cellspacing="0" cellpadding="0" >			
			<tr><td><%=rs("Client")%></td></tr>
			<tr><td><%=rs("Code")%></td></tr>
			<tr><td><%=rs("ClientCDSNo")%></td></tr>			
			</table></td>
			
			<td width="118"><table>
			<tr><td><%=rs("AgentName")%></td></tr>
			<tr><td><%=rs("AgentCode")%></td></tr>
			<tr><td>&nbsp;&nbsp;</td></tr>						
			</table></td>			  
			
			<td width="80"><table>
			<tr><td><%=rs("Order_DPA_")%></td></tr>
			<tr><td><%=FormatDate(rs("OrderDate"))%></td></tr>
			<tr><td><%=FormatNum(rs("BalanceQty"))%></td></tr>									
			</table></td>			  
			<%		
			
			if(Rs("OrderTypeSale")=true) then 
			OrderType="SALE"
			else 
			OrderType="BUY"
			end if
			%>
			<td width="50"><table>
			<tr><td><%=rs("SecurityCode")%></td></tr>
			<tr><td><%=rs("BalanceQty")%></td></tr>
			<tr><td><%=OrderType%></td></tr>												
			</table></td>			  
			
			<td width="36"><table>
			<% if(rs("Best")=true) then
			%>
			<tr><td width="36">BEST</td></tr>
			<%
			else
			%>
			<tr><td><%=rs("OrdDetailPrice")%></td></tr>
			<% end if%>
			<tr><td>&nbsp;&nbsp;</td></tr>
			<tr><td>&nbsp;&nbsp;</td></tr>												
			
			</table></td>			  
			
			<td width="65"><table border="0" cellspacing="0" cellpadding="2">
			<tr><td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>
			<tr><td>&nbsp;&nbsp;</td></tr>												
			<tr><td>&nbsp;&nbsp;</td></tr>															
			</table></td>			  
						
			<td width="65"><table border="0" cellspacing="0" cellpadding="2">
			<tr><td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>
			<tr><td>&nbsp;&nbsp;</td></tr>												
			<tr><td>&nbsp;&nbsp;</td></tr>																		
			</table></td>			  
			
			<td width="65"><table border="0" cellspacing="0" cellpadding="2">
			<tr><td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>
			<tr><td>&nbsp;&nbsp;</td></tr>												
			<tr><td>&nbsp;&nbsp;</td></tr>																		
			</table></td>			  
			
			<td width="65"><table border="0" cellspacing="0" cellpadding="2">
			<tr><td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>
			<tr><td>&nbsp;&nbsp;</td></tr>												
			<tr><td>&nbsp;&nbsp;</td></tr>																		
			</table></td>			  
			
			<td width="65"><table border="0" cellspacing="0" cellpadding="2">
			<tr><td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>
			<tr><td>&nbsp;&nbsp;</td></tr>												
			<tr><td>&nbsp;&nbsp;</td></tr>																		
			</table></td>			  
									  						  
			</tr>
    			
    <%	
		
		Security2=Security1
		first=0
            saletype2=saletype1	
		
		end if
		Rs.MoveNext
		Loop
	End If
	Set Rs = Nothing   
    Set Conn = Nothing
    %>
  </table>
  

</body>

</html>