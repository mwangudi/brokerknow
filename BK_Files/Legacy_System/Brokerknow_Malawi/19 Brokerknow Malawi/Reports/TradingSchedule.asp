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
			
		'=====================================================
        'Filter Parameter 1. Purchase Orders with excess <= 0
        '                 2. Order validity  
        '                 3. Include Bond & Custodian Orders
        '                 4. Sales with Balance Free = 'N'
        '====================================================
        
        sqlstr = " SELECT OrdDetailSecurity, OrdDetailSecType, Security_DPA_ FROM DB_FineTradingSchedule " & _
				"  WHERE ((Validity >= { fn CURDATE() }) OR (Validity IS NULL))  AND  " & _
				" ((OrderTypeSale = 0 AND Excess <= 0)  OR   " & _
				" (OrderTypeSale =1 AND (BalanceFree <> 'N' OR  Excess < = 0))OR (ordDetailSecType = 'Fixed') OR (Iscustodian = 1)) " & _
				" Group By OrdDetailSecurity, OrdDetailSecType, Security_DPA_   " & _
				" ORDER BY OrdDetailSecurity ASC" 		           
				
		sqlstr = "SELECT     ordDetailSecurity, OrdDetailSecType, Security_DPA_ " & _
			" FROM         DB_FineTradingSchedule " & _
			" WHERE     (Validity >= { fn CURDATE() } OR " & _
			"                       Validity IS NULL) AND (OrderTypeSale IN (0)) AND (Excess <= 0) " & _
			" UNION ALL " & _
			" SELECT     OrdDetailSecurity, OrdDetailSecType, Security_DPA_ " & _
			" FROM         DB_FineTradingSchedule " & _
			" WHERE     (Validity >= { fn CURDATE() } OR " & _
			"                       Validity IS NULL) AND (OrderTypeSale IN (1, 2)) AND (BalanceFree <> 'N') " & _
			" UNION ALL " & _
			" SELECT     OrdDetailSecurity, OrdDetailSecType, Security_DPA_ " & _
			" FROM         DB_FineTradingSchedule " & _
			" WHERE     (Validity >= { fn CURDATE() } OR " & _
			"                       Validity IS NULL) AND (OrderTypeSale IN (1, 2)) AND (Excess <= 0) " & _
			" UNION ALL " & _
			" SELECT     OrdDetailSecurity, OrdDetailSecType, Security_DPA_ " & _
			" FROM         DB_FineTradingSchedule " & _
			" WHERE     (Validity >= { fn CURDATE() } OR " & _
			"                       Validity IS NULL) AND (ordDetailSecType = 'Fixed') " & _
			" UNION ALL " & _
			" SELECT     OrdDetailSecurity, OrdDetailSecType, Security_DPA_ " & _
			" FROM         DB_FineTradingSchedule " & _
			" WHERE     (Validity >= { fn CURDATE() } OR " & _
			"                       Validity IS NULL) AND (Iscustodian = 1)" & _
			" GROUP BY OrdDetailSecurity, OrdDetailSecType, Security_DPA_" & _
			" ORDER BY OrdDetailSecurity ASC"
          
          'response.write sqlstr
		  'response.end
		 Set conn = GetActiveConnection("KBroker")
		 
		 'Run Stored procedures to update client Balances( Current Balance i.e monies) and Totals( Value of ordered Shares)
		   
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
i=0
headerDescription = FormatDateFull(selectedTradeDate)%>
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

  <table border="0" cellspacing="0" cellpadding="5" style="font-family: Arial Narrow" width="100%">
    <tr>
      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Order</font></b></td>
      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Type</font></b></td>
      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Security</font></b></td>
      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Date of Order</font></b></td>
      <td bgcolor="#000000" align="right"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Quantity</font></b></td>
      <td bgcolor="#000000" align="right"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Price</font></b></td>
    </tr>
    <tr>
      <td colspan="6">&nbsp; </td>
    </tr>
    <%
    Do Until groupRs.EOF
			
	  Security_DPA_ = groupRs.Fields("Security_DPA_").Value 
	  OrdDetailSecurity = groupRs.Fields("OrdDetailSecurity").Value
	  
			%>   
						<tr>
						<td><%= groupRs.Fields("OrdDetailSecType").Value %></td>
						<td colspan="5"><b><%= OrdDetailSecurity %></b></td>
						</tr>   
			<%
				'sqlStr = "SELECT * FROM TradingSchedule WHERE OrdDetailSecurity = '" & OrdDetailSecurity & "' ORDER BY OrdDetailType DESC, OrderDate"
																			                        
               sqlstr =  " SELECT * FROM DB_FineTradingSchedule   " & _
						 "  WHERE ((Validity >= { fn CURDATE() }) OR (Validity IS NULL))  AND  " & _
						 "    ((OrderTypeSale = 0 AND Excess <= 0)  " & _
						 "       OR  " & _
						 "     (OrderTypeSale =1 AND (BalanceFree <> 'N' OR  Excess < = 0)) OR (ordDetailSecType = 'Fixed') OR (Iscustodian = 1)) " & _
						 " AND Security_DPA_ = " & Security_DPA_ & " " & _
						 " ORDER BY OrdDetailType DESC, OrderDate"      
				
				sqlstr = "SELECT     * " & _
					" FROM         DB_FineTradingSchedule " & _
					" WHERE     (Validity >= { fn CURDATE() } OR " & _
					"                       Validity IS NULL) AND (OrderTypeSale IN (0)) AND (Excess <= 0) " & _
					" AND Security_DPA_ = " & Security_DPA_ & " " & _
					" UNION ALL " & _
					" SELECT     * " & _
					" FROM         DB_FineTradingSchedule " & _
					" WHERE     (Validity >= { fn CURDATE() } OR " & _
					"                       Validity IS NULL) AND (OrderTypeSale IN (1, 2)) AND (BalanceFree <> 'N') " & _
					" AND Security_DPA_ = " & Security_DPA_ & " " & _
					" UNION ALL " & _
					" SELECT     * " & _
					" FROM         DB_FineTradingSchedule " & _
					" WHERE     (Validity >= { fn CURDATE() } OR " & _
					"                       Validity IS NULL) AND (OrderTypeSale IN (1, 2)) AND (Excess <= 0) " & _
					" AND Security_DPA_ = " & Security_DPA_ & " " & _
					" UNION ALL " & _
					" SELECT     * " & _
					" FROM         DB_FineTradingSchedule " & _
					" WHERE     (Validity >= { fn CURDATE() } OR " & _
					"                       Validity IS NULL) AND (ordDetailSecType = 'Fixed') " & _
					" AND Security_DPA_ = " & Security_DPA_ & " " & _
					" UNION ALL " & _
					" SELECT     * " & _
					" FROM         DB_FineTradingSchedule " & _
					" WHERE     (Validity >= { fn CURDATE() } OR " & _
					"                       Validity IS NULL) AND (Iscustodian = 1) AND Security_DPA_ = " & Security_DPA_ & _
					" ORDER BY OrdDetailType DESC, OrderDate"    
					
				Set Rs = Conn.Execute (SQLServerFormat(HandleQuote(sqlStr)))
																				    
				If Not (Rs.EOF Or Rs.BOF) Then
					Do Until Rs.EOF%>
						<tr>			
						<td align="right"><%= Rs.Fields("Order_DPA_").Value %></td>
						<td><%= Rs.Fields("OrdDetailType").Value %></td>
						<%
						 if Ucase(trim(Rs.Fields("ordDetailSecType"))) = "FIXED" then
							'Display Bond Issue Number instead of security code
							%>
						 <td><%= Rs.Fields("BondIssue").Value %></td>
							<%
						 else
							%>
							 <td><%= Rs.Fields("SecurityCode").Value %></td>
							<%
						 end if
							%>
						
						<td><%= FormatDate(Rs.Fields("OrderDate").Value) %></td>

						<%
						 Limit = Cdbl(Rs("Limit"))
						 if (Cdbl(rs("BalanceQty")) > Limit AND Limit > 0) then
							OrderQty = Limit
						  Else
							OrderQty = Cdbl(rs("BalanceQty"))
						  end if
						%>
						<!-- <td align="right"><%'= Rs.Fields("BalanceQty").Value %> </td> -->
						<td align="right"><%=OrderQty%> </td>
						<%
						 'Indicate Best Order Price as BEST
						 
						 if cbool(Rs.Fields("Best").Value) then
							%>
								<td align="right">BEST</td>
							<%
						 else
							%>
							<td align="right"><%= Rs.Fields("OrdDetailPrice").Value %> </td>
							<%
						end if
						%>
						</tr>
				       <%Rs.MoveNext
				    Loop
			   End If
		
			   Set Rs = Nothing
	
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
