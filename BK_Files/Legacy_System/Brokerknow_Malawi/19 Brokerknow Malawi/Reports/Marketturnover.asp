<!--#include file="../libroutines.asp"-->

<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Market Turnover</title>  
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
FirstDay=DateSerial(Year(Date), Month(Date) + iOffset, 1)

genReport = Request.Form("genReport")
selectedFromDate = Request.Form("txtFromDate")
selectedToDate = Request.Form("txtToDate")

If genReport <> "1" Then%>
		<Script Language="JavaScript">
		report_SetBodyClass();
		function validateForm(frm){					
			frm.target = '_self';			
			frm.submit();
		}
		
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "txtFromDate","cmdDate","<%= FormatDate(FirstDay) %>",1);
		var cal1=new ctlSpiffyCalendarBox("cal1", "frmMain", "txtToDate","cmdDate","<%= FormatDate(Date) %>",1);
	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="MarketTurnover.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">	
		<table>			
			
			<tr>
				<td>Start Date:&nbsp;</td>
				<td>
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>
			</tr>
			<tr>
				<td>End date:&nbsp;</td>
				<td>
					<SCRIPT language="JavaScript">cal1.writeControl();</SCRIPT>	
				</td>
			</tr>
			
			<tr>
				<td colspan=2 align="Center"><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... ">&nbsp;&nbsp;</td>
			</tr>
		</table>
		
	</form>
	
	<%
	Response.End
End If

DrawPageFunctions True, True, True 
	
	
	Set conn = GetActiveConnection("KBroker")
	Set Rs = CreateObject("ADODB.Recordset")						        

	sqlStr = "SELECT     Turnover_DPA_, TradeDate, MarketTurnOver " & _
			 " FROM         tblTurnOver " & _
			 " WHERE     (Deleted = 0) AND (CAST(FLOOR(CAST(TradeDate AS float)) AS datetime) BETWEEN '"& formatdate(selectedFromDate) &"' AND '"& formatdate(selectedToDate) &"') " & _
			 " ORDER BY TradeDate DESC "
	
	'Response.write(sqlStr):Response.end
	
	Rs.CursorLocation = adUseClient	
	Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
	if rs.eof or rs.bof then
		%>
		<Script Language="JavaScript">
			alert("No data was found using the provided search criteria")
			window.history.go(-1);
		</Script>
		<%Set Rs = Nothing
		Set Conn = Nothing
		Response.End
	end if
%>	

<table border="0" cellspacing="2" cellpadding="2" style="font-family: Arial Narrow" width="100%">
    <tr>
		<td width="10%" nowrap><font face="Impact" size="4">MARKET TURNOVER</font></td>
      <td width="60%" nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>
      
    </tr>

  </table>
<br>
<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
    <tr>
      <td width="1%"><b>Date:</b></td>
      <td width="48%"><b>From:&nbsp;</b><%= FormatDate(selectedFromDate) %>&nbsp;<b>To:&nbsp;</b><%= FormatDate(selectedToDate) %></td>
    </tr>

    
</table>
<BR>


  <table border="0" cellspacing="1" cellpadding="2" style="font-family: Arial Narrow; LEFT-MARGIN:100PX" >
    <tr>
          	<td  nowrap align="center"><b><font face="Arial Narrow" size="2">#</font></b></td>
            <td  nowrap align="center"><b><font face="Arial Narrow" size="2">Traded</font></b></td>
        	<td  nowrap align="center"><b><font face="Arial Narrow" size="2">Market Turnover</font></b></td>
      
               
    </tr>
    <%
	rs.movefirst
	iCounter=0
    Do until rs.eof
		iCounter=iCounter+1
		%>   
		<tr>	
			  <td nowrap><font face="Arial Narrow" size="1"><%=iCounter%></font></td>
			  <td nowrap align="right" width="100"><font face="Arial Narrow" size="1"><%=FormatDate(Rs("TradeDate"))%></font></td>
			  <td nowrap align="right" width="100"><font face="Arial Narrow" size="1"><%=formatnum(Rs("MarketTurnover"))%></font></td>		  
		</tr>    
		<%	
		Rs.MoveNext
	Loop
	%>
	
  </table>
   
<%Set Rs = Nothing
Set Conn = Nothing%>   
</body>

</html>