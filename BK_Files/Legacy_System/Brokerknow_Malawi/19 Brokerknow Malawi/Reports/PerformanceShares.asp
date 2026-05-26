<!--#include file="../libroutines.asp"-->

<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Performance Shares</title>  
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



<%
'FirstDay=DateSerial(Year(Date), Month(Date) + iOffset, 1)
FirstDay=DateSerial(Year(Date), Month(Date)-1 + iOffset, 1)
LastDay=DateSerial(Year(Date), Month(Date) + iOffset, 1)-1

genReport = Request.Form("genReport")
selectedFromDate = Request.Form("txtFromDate")
selectedToDate = Request.Form("txtToDate")
If genReport <> "1" Then%>
		<Script Language="JavaScript">
		report_SetBodyClass();
		function validateForm(frm){			
			//if (frm.txtDate.value==''){
			//	alert("Select a date");
			//	frm.txtDate.focus();
			//	return;
			//}
			
			frm.target = '_self';			
			frm.submit();
		}
		
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "txtFromDate","cmdDate","<%= FormatDate(FirstDay) %>",1);
		var cal1=new ctlSpiffyCalendarBox("cal1", "frmMain", "txtToDate","cmdDate","<%= FormatDate(LastDay) %>",1);
	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="PerformanceShares.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">	
		<table>			
			<tr>
				<td>From Date:</td>
				<td>
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>
			</tr>
			<tr>
				<td>To date:</td>
				<td>
					<SCRIPT language="JavaScript">cal1.writeControl();</SCRIPT>	
				</td>
			</tr>
			
			<tr>
				<td colspan=2><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... ">&nbsp;&nbsp;</td>
			</tr>
		</table>
		
	</form>
	
	<%
	Response.End
End If

%>

<% DrawPageFunctions True, True, True %>


<%
	Set conn = GetActiveConnection("KBroker")
	Set Rs = CreateObject("ADODB.Recordset")						        
	
	Dim pageNumber
	
	pageNumber = 0	
	
		'Rs.Filter = "ContractNumber = '" & Rs.Fields("ContractNumber").Value & "'"
	
		
	pageNumber = pageNumber + 1
%>
<table border="0" cellspacing="2" cellpadding="2" style="font-family: Arial Narrow" width="100%">
    <tr>
	  <td bgcolor="#000000" width="80%" nowrap align="left"><font color="#FFFFFF" face="Impact" size="2">PERFORMANCE RESULTS IN SHARE MARKET</font></td>
      <td bgcolor="#000000" width="20%" nowrap align=right><font color="#FFFFFF" face="Impact" size="2"><%= Session("CompanyName") %></font></td>
      
    </tr>

  </table>
<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" width="100%" > 
	<tr>      
      			<td ><b>Date:&nbsp;&nbsp;<%= FormatDate(Date) %></b></td>      			
    </tr>
	<tr>
		<td align="right" height="8">
            &nbsp;			
		</td>		
	</tr>
  <tr>      
      			<td ><b>From :&nbsp;&nbsp;<%= FormatDate(selectedFromDate) %>&nbsp;&nbsp;&nbsp;&nbsp;To :&nbsp;&nbsp;<%= FormatDate(selectedToDate) %></b></td>      			
    </tr>
</table>
<table border="0" cellspacing=2 cellpadding=2 width="600">		    	    
	<tr>      
      <td align="left" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>TITLE&nbsp;&nbsp;</b></font></td>      
      <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>TURNOVER</b></font></td>                        
      <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>GROSS COMM</b></font></td>      
      <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>RET COMM</b></font></td>
      <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>NET COMM</b></font></td>            
    </tr>
	<%	
    sqlStr = "SELECT     SUM(Gross) AS Gross, SUM(Commission) AS Commission FROM performanceShares WHERE TransDate between '" & CDate(selectedFromDate) & "' and '" & CDate(selectedToDate) & "'"	 
	
	Rs.CursorLocation = adUseClient	
	Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic	
	
	
	TotalGross=Rs("Gross")
	TotalComm=Rs("Commission")	
	

	Runtotal=0
	RunComm=0
	RetComm=0
	NetComm=0
	TotalRet=0
	TotalNet=0

	Set Rs = Nothing

	%>
	<tr>      
	  <td align="left"><font face="Arial Narrow" size="2"><b>TOTAL TURNOVER</b></font></td>      
	  <td align="Right"><font face="Arial Narrow" size="2"><b><%=FormatNum(TotalGross)%></b></font></td>                        
	  <td align="Right"><font face="Arial Narrow" size="2"><b><%=FormatNum(TotalComm)%></b></font></td>      
	  <td align="Right"><font face="Arial Narrow" size="2">&nbsp;</font></td>
	  <td align="Right"><font face="Arial Narrow" size="2">&nbsp;</font></td>            
	</tr>	
	<%
	
	Set TempRs = CreateObject("ADODB.Recordset")						        
	
	'Response.end

	sqlStr="SELECT SUM(performanceShares.Gross) AS Gross, SUM(performanceShares.Commission) AS Commission, 		performanceShares.Owner_DPA_,  " & _
			 "                       OwnerList.OwnerName, OwnerList.CommissionRate " & _
			 " FROM         performanceShares INNER JOIN " & _
			 "                       OwnerList ON performanceShares.Owner_DPA_ = OwnerList.Owner_DPA_ " & _
			 " WHERE     (performanceShares.TransDate between '" & CDate(selectedFromDate) & "' and '" & CDate(selectedToDate) & "')"	 & _
			 " GROUP BY performanceShares.Owner_DPA_, OwnerList.OwnerName, OwnerList.CommissionRate " & _
			 " HAVING (NOT (performanceShares.Owner_DPA_ IS NULL))"
	
	'Response.write(sqlStr)
	'Response.end
	

	TempRs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic	
	
	Do Until TempRs.EOF	
	    'if(TempRs("Owner_DPA_")=4 OR TempRs("Owner_DPA_")=5 OR TempRs("Owner_DPA_")=8) then
			RetComm=TempRs("Commission")*TempRs("CommissionRate")/100
			NetComm=TempRs("Commission")-(TempRs("Commission")*TempRs("CommissionRate")/100)
		'	else
		'	RetComm=TempRs("Commission")*15/100
		'	NetComm=TempRs("Commission")-(TempRs("Commission")*15/100)			
		'	end if
		%>
		<tr>      
		  <td align="left"><font face="Arial Narrow" size="2"><%=Ucase(TempRs("OwnerName"))%></font></td>      
		  <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(TempRs("Gross"))%></font></td>                        
		  <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(TempRs("Commission"))%></font></td>      
		  <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(RetComm)%></font></td>
		  <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(NetComm)%></font></td>            
		</tr>
		<%
		Runtotal=Runtotal+TempRs("Gross")
		RunComm=RunComm+TempRs("Commission")
		
		TotalRet=TotalRet + RetComm
		TotalNet=TotalNet + NetComm

		TempRs.movenext
	Loop
	set TempRs=nothing

	Set Rs = CreateObject("ADODB.Recordset")						        

	sqlStr = "SELECT SUM(LevyAmount) AS Commission, SUM(LotGrossAmount) AS Gross FROM OwnerCommissions WHERE LotTDate between '" & CDate(selectedFromDate) & "' and '" & CDate(selectedToDate) & "' AND (Owner_DPA_ = - 1)"	 
	
	'Response.write(sqlStr)
	'Response.end

	Rs.CursorLocation = adUseClient	
	Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
	
	
	'TotalGross=Rs("Gross")
	'TotalComm=Rs("Commission")	
		%>
		<tr>      
			  <td align="left"><font face="Arial Narrow" size="2">WALK IN CLIENTS (CUSTODIAL)</font></td>      
			  <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(Rs("Gross"))%></font></td>                        
			  <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(Rs("Commission"))%></font></td>      
			  <td align="Right"><font face="Arial Narrow" size="2">&nbsp;</font></td>
			  <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(Rs("Commission"))%></font></td>            
			</tr>
			
	<%
		Runtotal=Runtotal+Rs("Gross")
		RunComm=RunComm+Rs("Commission")
		
		'TotalRet=TotalRet + RetComm
		TotalNet=TotalNet + Rs("Commission")
		
	Set Rs = Nothing
	set TempRs=nothing

	Set Rs = CreateObject("ADODB.Recordset")						        

	sqlStr = "SELECT SUM(LevyAmount) AS Commission, SUM(LotGrossAmount) AS Gross FROM OwnerCommissions WHERE LotTDate between '" & CDate(selectedFromDate) & "' and '" & CDate(selectedToDate) & "' AND (Owner_DPA_ = 0)"	 
	
	'Response.write(sqlStr)
	'Response.end

	Rs.CursorLocation = adUseClient	
	Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
	
	
	'TotalGross=Rs("Gross")
	'TotalComm=Rs("Commission")
	%>
	<tr>      
		  <td align="left"><font face="Arial Narrow" size="2">WALK IN CLIENTS (NON CUSTODIAL)</font></td>      
		  <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(Rs("Gross"))%></font></td>                        
		  <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(Rs("Commission"))%></font></td>      
		  <td align="Right"><font face="Arial Narrow" size="2">&nbsp;</font></td>
		  <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(Rs("Commission"))%></font></td>            
		</tr>
		
<%
	Runtotal=Runtotal+Rs("Gross")
	RunComm=RunComm+Rs("Commission")
	'TotalRet=TotalRet + RetComm
	TotalNet=TotalNet + Rs("Commission")
	
	Set Rs = Nothing
	Set Rs = CreateObject("ADODB.Recordset")						        

	sqlStr = "SELECT     SUM(LevyAmount) AS Commission, SUM(LotGrossAmount) AS Gross, AgentName,Agent_DPA_ " & _
			 " FROM         DB_AgentCommissions " & _
			 " WHERE     (LotTDate between #" & CDate(selectedFromDate) & "# and #" & CDate(selectedToDate) & "#)" & _
			 " GROUP BY Agent_DPA_, AgentName " & _
			 " HAVING (Agent_DPA_ = 10024)"	 
	
	sqlstr = SQLServerFormat(HandleQuote(sqlStr))		 

	Rs.CursorLocation = adUseClient	
	Rs.Open sqlstr, conn.ConnectionString, adOpenKeyset, adLockOptimistic
    
    if not (Rs.Bof or Rs.EOF) then						        
	
	sqlstr = " SELECT     SUM(dbo.LevyContract.LevyAmount) AS Commission, dbo.Client.Agent_DPA_ " & _
			 " FROM         dbo.LevyContract INNER JOIN " & _
			 "                       dbo.Lots ON dbo.LevyContract.Contract_DPA_ = dbo.Lots.Contract_DPA_ INNER JOIN " & _
			 "                       dbo.OrdDetail ON dbo.Lots.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN " & _
			 "                       dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN " & _
			 "                       dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ " & _
			 " WHERE     (dbo.LevyContract.SystemMaintained = 12) AND (dbo.tbOrder.OrderSecType_DPA_ = 2) AND (dbo.Client.Owner_DPA_ IS NULL) AND  " & _
			 "                       (CAST(FLOOR(CAST(dbo.Lots.LotTDate AS Float)) AS DateTime) between '" & CDate(selectedFromDate) & "' and '" & CDate(selectedToDate) & "')" & _
			 " GROUP BY dbo.Client.Agent_DPA_ " & _
			 " HAVING      (dbo.Client.Agent_DPA_ = "& Rs("Agent_DPA_") &")"
	
	sqlstr = SQLServerFormat(HandleQuote(sqlStr))		 
	
	
	
	Set TempRs = Nothing
	Set TempRs = server.CreateObject("ADODB.Recordset")
	
	TempRs.CursorLocation = adUseClient	
	TempRs.Open sqlstr, conn.ConnectionString, adOpenKeyset, adLockOptimistic
	
	%>
	<tr>      
		  <td align="left"><font face="Arial Narrow" size="2"><%=ucase(Rs("AgentName"))%></font></td>      
		  <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(Rs("Gross"))%></font></td>                        
		  <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(Rs("Commission"))%></font></td>      
		  <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(TempRs("Commission"))%></font></td>
		  <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(Rs("Commission")-TempRs("Commission"))%></font></td>            
		</tr>
		
<%
		Runtotal=Runtotal+Rs("Gross")
		RunComm=RunComm+Rs("Commission")
		
		TotalRet=TotalRet + TempRs("Commission")
		TotalNet=TotalNet + (Rs("Commission")-TempRs("Commission"))
	
	end if
	
	Set Rs = Nothing
	Set Rs = CreateObject("ADODB.Recordset")						        

	sqlStr = "SELECT     SUM(LevyAmount) AS Commission, SUM(LotGrossAmount) AS Gross" & _
			 " FROM         DB_AgentCommissions " & _
			 " WHERE     (LotTDate between '" & CDate(selectedFromDate) & "' and '" & CDate(selectedToDate) & "')" & _
			 " " & _
			 " AND (Agent_DPA_ <> 10024)"	 
	
	'Response.write(sqlStr)
	'Response.end

	Rs.CursorLocation = adUseClient	
	Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
	
	Set TempRs = Nothing
	Set TempRs = CreateObject("ADODB.Recordset")						        
	
	sqlstr="SELECT     SUM(dbo.LevyContract.LevyAmount) AS Commission" & _
			 " FROM         dbo.LevyContract INNER JOIN " & _
			 "                       dbo.Lots ON dbo.LevyContract.Contract_DPA_ = dbo.Lots.Contract_DPA_ INNER JOIN " & _
			 "                       dbo.OrdDetail ON dbo.Lots.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN " & _
			 "                       dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN " & _
			 "                       dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ " & _
			 " WHERE     (dbo.LevyContract.SystemMaintained = 12) AND (dbo.tbOrder.OrderSecType_DPA_ = 2) AND (dbo.Client.Owner_DPA_ IS NULL) AND  " & _
			 "                       (CAST(FLOOR(CAST(dbo.Lots.LotTDate AS Float)) AS DateTime) between '" & CDate(selectedFromDate) & "' and '" & CDate(selectedToDate) & "')" & _
			 "AND      (dbo.Client.Agent_DPA_ <> 10024) " & _
			 " "
	
	'Response.write(sqlStr)
	'Response.end

	TempRs.CursorLocation = adUseClient	
	TempRs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
	
	Runtotal=Runtotal+Rs("Gross")
	RunComm=RunComm+Rs("Commission")
	TotalRet=TotalRet + TempRs("Commission")
	TotalNet=TotalNet + (Rs("Commission")-TempRs("Commission"))

	
	%>
		<tr>      
		  <td align="left"><font face="Arial Narrow" size="2">AGENT RELATED CLIENTS</font></td>      
		  <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(Rs("Gross"))%></font></td>                        
		  <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(Rs("Commission"))%></font></td>      
		  <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(TempRs("Commission"))%></font></td>
		  <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(Rs("Commission")-TempRs("Commission"))%></font></td>            
		</tr>
		
		<tr>      
	  <td align="left"><font face="Arial Narrow" size="2"><b>TOTAL TURNOVER</b></font></td>      
	  <td align="Right"><font face="Arial Narrow" size="2"><b><%=FormatNum(Runtotal)%></b></font></td>                        
	  <td align="Right"><font face="Arial Narrow" size="2"><b><%=FormatNum(RunComm)%></b></font></td>      
	  <td align="Right"><font face="Arial Narrow" size="2"><b><%=FormatNum(TotalRet)%></b></font></td>
	  <td align="Right"><font face="Arial Narrow" size="2"><b><%=FormatNum(TotalNet)%></b></font></td>            
		</tr>
	
</table>	
</body>

</html>

<%
function convertSign(balance)
	dim sign
	sign = sgn(balance) '1 indicates positive, -1 indicates negative
	if sign = 1 then
		convertSign = "-" & balance	'// make this positive number a negative number
	elseif sign = -1 then
		convertSign = Abs(balance)	'// make this negative number a positive number. Abs removes the - element.
	end if
end function
%>

