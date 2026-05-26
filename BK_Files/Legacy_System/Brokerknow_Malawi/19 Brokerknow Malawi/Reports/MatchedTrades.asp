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
	
		 sqlStr = "SELECT * From CDSMatchedTradesList"
				 
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
<i id="landRem"></i>
<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
     <tr>
		<td nowrap><b><font face="Arial Narrow" size="4">Matched Trades</font></b></td>
		<td nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>
	</tr>	    
    <tr>
		  <td COLSPAN=2><font face="Arial" size="2">&nbsp;</font></td>
	</tr>
</table>				

  <table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="946">
    <tr>
    
	 <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow">Order&nbsp;No</font></b></td>
     <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow">Date</font></b></td>      
     <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow">Type</font></b></td>
     <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow">Client</font></b></td>
     <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow">Security</font></b></td>
     <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow">Limit</font></b></td>
     <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow">Balance</font></b></td>
     <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow">Trade&nbsp;Date</font></b></td>
     <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow">Ref</font></b></td>
     <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow">Quantity</font></b></td>      
     <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow">Price</font></b></td>
     <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow">Broker</font></b></td>
			
	</tr>
    <tr>
      <td colspan="12" width="737">&nbsp; </td>
    </tr>        
    <%    
    Do Until groupRs.EOF       			    
		%>
	<tr>
	  <td><font face="Arial Narrow"><%=grouprs.Fields("Order_DPA_")%></font></td>
      <td><font face="Arial Narrow"><%=FormatDate(grouprs.Fields("OrderDate"))%></font></td>      
      <td><font face="Arial Narrow"><%=IIf(CBool(grouprs.Fields("OrderTypeSale")) = True, "S", "P")%></font></b></td>
      <td><font face="Arial Narrow"><%=grouprs.Fields("OrdDetailClient")%></font></td>
      <td><font face="Arial Narrow"><%=grouprs.Fields("SecurityCode")%></font></td>
      <td><font face="Arial Narrow"><%=grouprs.Fields("OrdDetailQty")%></font></td>
      <td><font face="Arial Narrow"><%=grouprs.Fields("BalanceQty")%></font></td>
      <td><font face="Arial Narrow"><%=grouprs.Fields("TradeDate")%></font></td>
      <td><font face="Arial Narrow"><%=grouprs.Fields("CDSRef")%></font></td>
      <td><font face="Arial Narrow"><%=grouprs.Fields("Quantity")%></font></td>      
      <td><font face="Arial Narrow"><%=grouprs.Fields("Price")%></font></td>
      <td><font face="Arial Narrow"><%=grouprs.Fields("BrokerCode")%></font></td>  
    </tr>
    		<%		    		
		groupRs.MoveNext
    Loop
    
    Set groupRs = Nothing
    Set Conn = Nothing
    %>

    <tr>
      <td colspan="12" width="737">&nbsp;</td>
    </tr>
    <tr>
      <td colspan="12" width="737">&nbsp;</td>
    </tr>
  </table>
  

</body>

</html>