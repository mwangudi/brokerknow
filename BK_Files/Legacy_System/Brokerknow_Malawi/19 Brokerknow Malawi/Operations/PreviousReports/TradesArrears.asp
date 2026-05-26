<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Trades Arrears</title>

	<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
	<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
	<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
	<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
	<SCRIPT language=VBScript src="../scripts/reports.vbs"></SCRIPT>
	<SCRIPT language=Javascript src="../scripts/reports.js"></SCRIPT>
	<SCRIPT language=Javascript src="../scripts/accountList.js"></SCRIPT>

	<style media="print">
		@page {
			@top{font-family: Helvetica, Arial, sans-serif;
				font-size: 150%;
				font-weight: bolder;
				text-align: left;
				content: "<%= FormatDate(Date) %>";			
			}
			
			margin-left: 0cm;
			margin-right: 0cm;
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
'selectedBank = Request.Form("cboAccount")
selectedFromDate = Request.Form("transFromDate")
selectedToDate = Request.Form("txtToDate")
SelectedType=Request.Form("cboEntity")
FirstDay=DateSerial(Year(Date), Month(Date)-1 + iOffset, 1)
thistype=Request.Form("Selectedtype")

If genReport <> "1" Or Not IsDate(selectedFromDate) Then%>
	<Script Language="JavaScript">
		report_SetBodyClass();
		
		function validateForm(frm)
		{						
			frm.target = '_self';			
			frm.submit();
		}
		
		
		function evaluateEntity(Val, Entity)
		{      	
	  	FetchAccounts1(Entity)
		}
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "transFromDate","cmdDate","<%= FormatDate(FirstDay) %>",1);
		var cal1=new ctlSpiffyCalendarBox("cal1", "frmMain", "txtToDate","cmdDate","<%= FormatDate(Date) %>",1);

	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="TradesArrears.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">
		<% currentEntityType=5 %>
		<table>
			<tr>
				<td colspan="2">Select date:</td>
				<td>
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>				
			</tr>		

			<tr>
				<td colspan="3"><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... ">&nbsp;&nbsp;</td>
			</tr>
		</table>
		
	</form>
	
	<%Set rs = Nothing
	Set Conn = Nothing
	Response.End
End If
%>
<%DrawPageFunctions True, True, True%>
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
	
		 'sqlStr = "SELECT * FROM Db_FineTradingSchedule WHERE ((Validity >= { 'fn CURDATE() }) OR " & _
          '   "(Validity IS NULL)) ORDER BY Code,Client,Order_DPA_"
		
		sqlStr="SELECT DISTINCT  " & _
				 "                       TOP 100 PERCENT dbo.OrderType.OrderTypeDescription AS OrdDetailType, dbo.OrdDetail.OrdDetailPrice, dbo.OrdDetail.Order_DPA_,  " & _
				 "                       dbo.tbOrder.OrderRef, dbo.tbOrder.OrderDate, dbo.Security.SecurityCode, dbo.Client.Client_DPA_ AS Code, dbo.Client.ClientName AS Client,  " & _
				 "                       dbo.OrdDetail.Best, dbo.OrdDetail.OrdDetailValidity AS Validity, dbo.OrdDetail.Amount, DB_DataStreamPriceList.Price AS Price,  " & _
				 "                       dbo.DB_OrdDetailContractedQtyList.BalanceQty, ISNULL(- dbo.Client.CreditLimit - (ClientBalances.CurrentBal - ClientTotal.Total), 0) AS Excess,  " & _
				 "                       ClientBalances.CurrentBal, ISNULL(ClientTotal.Total, 0) AS Total, dbo.OrderSecType.OrderSecTypeDisplayName AS OrdDetailSecType,  " & _
				 "                       dbo.Security.SecurityName AS ordDetailSecurity, dbo.Client.CreditLimit, dbo.OrderType.OrderTypeSale, dbo.OrdDetail.Security_DPA_,  " & _
				 "                       dbo.Client.ClientCDSNo, CASE WHEN len(dbo.Agent.AgentName) > 10 THEN LEFT(dbo.Agent.AgentName, 10)  " & _
				 "                       + '...' ELSE dbo.Agent.AgentName END AS AgentName, dbo.Agent.Agent_DPA_ AS AgentCode, Lots.ContractNumber , " & _
                 "     Cast(Floor(Cast(Lots.LotTDate AS Float)) AS DateTime) AS LotTDate" & _
				 " FROM         dbo.OrdDetail INNER JOIN " & _
				 "                       dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN " & _
				 "                       dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN " & _
				 "                       dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN " & _
				 "                       dbo.OrderSecType ON dbo.tbOrder.OrderSecType_DPA_ = dbo.OrderSecType.OrderSecType_DPA_ INNER JOIN " & _
				 "                       dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN " & _
				 "                       dbo.DB_OrdDetailContractedQtyList ON dbo.OrdDetail.OrdDetail_DPA_ = dbo.DB_OrdDetailContractedQtyList.OrdDetail_DPA_ LEFT OUTER JOIN " & _
				 "                       dbo.Agent ON dbo.Client.Agent_DPA_ = dbo.Agent.Agent_DPA_ LEFT OUTER JOIN " & _
				 "                       ( Select Sum(Balance) as CurrentBal,Client_DPA_ From StatementList where TransDate<=#" & CDate(SelectedFromDate) & "# Group By Client_DPA_) ClientBalances ON dbo.Client.Client_DPA_ = ClientBalances.client_DPA_ LEFT OUTER JOIN " & _
				 "(SELECT     SUM(ISNULL(BalanceQty * OrdDetailPrice, 0)) AS Total, Client_DPA_ " & _
				 " FROM         (SELECT DISTINCT  " & _
				 "                                               dbo.LotList.BalanceQty, CASE (dbo.OrdDetail.Best) WHEN 1 THEN datastream_SecurityPriceList.Price * 1.105 ELSE CONVERT(float,  " & _
				 "                                               dbo.LotList.OrdDetailPrice) END AS OrdDetailPrice, dbo.LotList.Order_DPA_, dbo.LotList.Client_DPA_ " & _
				 "                        FROM          dbo.LotList INNER JOIN " & _
				 "                                               dbo.OrdDetail ON dbo.LotList.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN " & _
				 "                                                   (SELECT DISTINCT  " & _
				 "                                                                            TOP 100 PERCENT dsM.MktDate AS [Date], ISNULL(dsM.MktClose, 0) AS Price, dbo.Security.Security_DPA_,  " & _
				 "                                                                            dbo.Security.SecurityCode, datastream_Securities.SecCode " & _
				 "                                                     FROM          dbo.datastream_Market dsM INNER JOIN " & _
				 "                                                                            dbo.datastream_Securities ON dsM.MktCode = dbo.datastream_Securities.SecCode INNER JOIN " & _
				 "                                                                                (SELECT DISTINCT TOP 100 PERCENT MAX(FLOOR(CAST(MktDate AS float))) AS MktUnique, LTRIM(MktCode) AS Code " & _
				 "                                                                                  FROM          dbo.datastream_Market " & _
				 "                                                                                  WHERE      (NOT (MktClose IS NULL)) AND (MktClose <> 0) AND (Cast(Floor(Cast(MktDate as Float)) as Datetime) <= #" & CDate(SelectedFromDate) & "#) " & _
				 "                                                                                  GROUP BY MktCode " & _
				 "                                                                                  ORDER BY LTRIM(MktCode)) MarketView ON FLOOR(CAST(dsM.MktDate AS float)) = MarketView.MktUnique AND  " & _
				 "                                                                            dsM.MktCode = MarketView.Code RIGHT OUTER JOIN " & _
				 "                                                                            dbo.Security ON dbo.datastream_Securities.SecKnow_DPA = dbo.Security.Security_DPA_ " & _
				 "                                                     ORDER BY dbo.Security.Security_DPA_) datastream_SecurityPriceList ON  " & _
				 "                                               dbo.OrdDetail.Security_DPA_ = datastream_SecurityPriceList.Security_DPA_ " & _
				 "                        WHERE      (RTRIM(dbo.LotList.OrdDetailType) LIKE '%Purchase%')) a " & _
				 " GROUP BY Client_DPA_) ClientTotal ON dbo.Client.Client_DPA_ = ClientTotal.Client_DPA_ LEFT OUTER JOIN " & _
								 "                       (SELECT DISTINCT  " & _
				 "                       TOP 100 PERCENT dsM.MktDate AS [Date], ISNULL(dsM.MktClose, 0) AS Price, dbo.Security.Security_DPA_, dbo.Security.SecurityCode,  " & _
				 "                       dbo.datastream_Securities.SecCode " & _
				 " FROM         dbo.datastream_Market dsM INNER JOIN " & _
				 "                       dbo.datastream_Securities ON dsM.MktCode = dbo.datastream_Securities.SecCode INNER JOIN " & _
				 "                           (SELECT DISTINCT TOP 100 PERCENT MAX(FLOOR(CAST(MktDate AS float))) AS MktUnique, LTRIM(MktCode) AS Code " & _
				 "                             FROM          dbo.datastream_Market " & _
				 "                             WHERE      (NOT (MktClose IS NULL)) AND (MktClose <> 0) AND (MktDate <= #" & CDate(SelectedFromDate) & "#)" & _
				 "                             GROUP BY MktCode " & _
				 "                             ORDER BY LTRIM(MktCode)) MarketView ON FLOOR(CAST(dsM.MktDate AS float)) = MarketView.MktUnique AND  " & _
				 "                       dsM.MktCode = MarketView.Code RIGHT OUTER JOIN " & _
				 "                       dbo.Security ON dbo.datastream_Securities.SecKnow_DPA = dbo.Security.Security_DPA_ " & _
				 " ORDER BY dbo.Security.Security_DPA_) DB_DataStreamPriceList ON dbo.Security.Security_DPA_ = DB_DataStreamPriceList.Security_DPA_ INNER JOIN " & _
				 "                       lots ON ordDetail.ordDetail_DPA_ = lots.OrdDetail_DPA_ " & _
				 " WHERE  (dbo.tbOrder.OrderHold = 0) AND (dbo.tbOrder.OrderCanceled = 0) AND  " & _
				 "                       (dbo.tbOrder.Deleted = 0) AND (dbo.OrdDetail.Deleted = 0) and Cast(Floor(Cast(Lots.LotTDate AS Float)) AS DateTime)=#" & CDate(SelectedFromDate) & "# and dbo.OrderSecType.OrderSecType_DPA_=2 ORDER BY Cast(Floor(Cast(Lots.LotTDate AS Float)) AS DateTime)"

    	Set conn = GetActiveConnection("KBroker")
		 Set rstotal = Server.CreateObject("ADODB.Recordset")
		 Set rsCurrent = Server.CreateObject("ADODB.Recordset")		
		 
		 rstotal.CursorLocation = adUseClient
		 rsCurrent.CursorLocation = adUseClient
		
		' Response.write(SQLServerFormat(HandleQuote(sqlStr)))
		' Response.end

		 Set conn = GetActiveConnection("KBroker")
		 
		 Conn.execute("ClientTotalsDelete")		 
		 Conn.execute("ClientBalancesDelete")		 
			
		 Conn.execute("ClientTotalsProcedure")		 
		 Conn.execute("ClientBalancesProcedure")		 			 
		 	       
	    Set groupRs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
	    If groupRs.EOF Or groupRs.BOF Then %>
				<Script Language="JavaScript">
					alert("No Contracts available");
					window.parent.close();					
	            </Script>
	            <% Set groupRs = Nothing
	            Set Conn = Nothing
	            Response.End
	    End If
		        
	    groupRs.MoveFirst
		        
selectedTradeDate = FormatDate(Date)

headerDescription = FormatDateFull(selectedFromDate) %>
<i id="landRem">Remember to select landscape settings while printing.</i>
<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
     <tr>
		<td nowrap><b><font face="Arial Narrow" size="4">Trades In Arrears</font></b></td>
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
      <td bgcolor="#000000" width="60"><b><font color="#FFFFFF" face="Arial Narrow">Contract</font></b></td>
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
			  <td width="60"><%= groupRs.Fields("ContractNumber").Value %></td>
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