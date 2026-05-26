<html>

<head>
<title>Settlement Slip</title>
  
	<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
	<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
	<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
	<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
	<SCRIPT language=VBScript src="../scripts/reports.vbs"></SCRIPT>
	<SCRIPT language=Javascript src="../scripts/reports.js"></SCRIPT>

	<style media="print">
		@page {
			@top{font-family: Helvetica, Arial, sans-serif;
				font-size: 150%;
				font-weight: bolder;
				text-align: left;
				content: "<%= FormatDate(Date) %>";			
			}
			
			margin-left: 2cm;
			margin-right: 5cm;
			margin-top: 1cm;    
			margin-bottom: 2cm;
			size: portrait;
			
			br.newpage{
				page-break-before:always;
			}
			
			tr.pageNumbering{
				display:none;
			}
		}

	</style>
</head>

<body Class="Reports">

<!--#include file="../libroutinesTEST.asp"-->


<%

genReport = Request.Form("genReport")
selectedTradeDate = Request.Form("txtDate")
timeLimit = Request.Form("timeLimit")

If genReport <> "1" Or Not IsDate(selectedTradeDate) Then%>
	<Script Language="JavaScript">
		report_SetBodyClass();
		
		function validateForm(frm){			
			if (frm.txtDate.value==''){
				alert("Select a date");
				frm.txtDate.focus();
				return;
			}
			
			frm.target = '_self';			
			frm.submit();
		}
		
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "txtDate","cmdDate","<%= FormatDate(Date) %>",1);
	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="SettlementSlips.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">	
		<table>
			<tr>
				<td>Select any day of month</td>
				<td>
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>
			</tr>
			<tr>
				<td>&nbsp;</td>
				<td>
					<p><input type="radio" class="BorderLess" checked Name="timeLimit" value="0" id="NSETime"><label for="NSETime" style="cursor: hand">Use MSE time limit</label></p>
					<p><input type="radio" class="BorderLess" Name="timeLimit" value="1" id="InternalTime"><label for="InternalTime" style="cursor: hand">Use internal time limit</label></p>
					
				</td>
			</tr>
			<tr>
				<td colspan=2>&nbsp;</td>
			</tr>
			<tr>
				<td colspan=2><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... ">&nbsp;&nbsp;</td>
			</tr>
		</table>
		
	</form>
	
	<%Set rs = Nothing
	Set Conn = Nothing
	Response.End
End If

%>



<% DrawPageFunctions True, True, True, True %>


<%
   Dim conn 
   Dim sqlStr
   Dim rs
   Dim isDifferentBroker
   Dim currBrokerCode
   
   Set conn = GetActiveConnection("KBroker")
    sqlStr = "SELECT * FROM [TimeLimitList] WHERE TimeLimit_DPA_ = 1" 'Settlement (Brokers)
    Set tempRs = Conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        
    If Not (tempRs.EOF Or tempRs.BOF) Then
		If timeLimit = "0" Then
			'use NSE
			timeRange = tempRs.Fields("TimeLimitNSE").Value
		Else
			'use internal
			timeRange = tempRs.Fields("TimeLimitInternal").Value
		End If		
			
    Else%>
		<Script Language="JavaScript">
			alert("The time limits of type Settlement (Brokers) have not been set");
			window.parent.self.close();				
        </Script>
		<%
		Set tempRs = Nothing
		Set Conn = Nothing
		Response.End
    End If
        
    Set tempRs = Nothing
    
     sqlStr = "SELECT * FROM [Holidays]"
        Set holidayRs = CreateObject("ADODB.Recordset")
        holidayRs.CursorLocation = adUseClient
        holidayRs.Open SQLServerFormat(HandleQuote(sqlStr)), Conn.ConnectionString, 1, 3
                
        
        'get upper filter date
        originalselectedTradeDate = selectedTradeDate
        settlementDate = FormatDate(selectedTradeDate)
        theSelDay = WeekDay(selectedTradeDate)
        
        'get to a monday if selected date is on weekend
        settlementDate = JumpUpFromWeekendToWeek (settlementDate)
        
		
		'check upwards until upper date is not a holiday
               
        upperLimitDate = settlementDate
        
        
        
        For i = 1 To timeRange
			upperLimitDate = DateAdd("d", 1, upperLimitDate)
			upperLimitDate = JumpUpFromWeekendToWeek(upperLimitDate)
			
			'count only working days
			'check upwards until is not a holiday
			isProperDate = False
		
			Do Until isProperDate = True
				holidayRs.Cancel		
				holidayRs.Filter = "Holiday = '" & upperLimitDate & "'"		
				If Not (holidayRs.EOF Or holidayRs.BOF) Then
					upperLimitDate = DateAdd("d", 1, upperLimitDate)
				Else
					formerSelectedTradeDate = upperLimitDate
					upperLimitDate = JumpBackFromWeekendToWeek (upperLimitDate)
					If DateDiff("d", FormatDate(upperLimitDate), FormatDate(formerSelectedTradeDate)) = 0 Then
						isProperDate = True
						holidayRs.Cancel
					End If	
				End If

			Loop
   
        Next
        
        Set holidayRs = Nothing
	
		sqlStr = "SELECT * FROM [SettlementSlips] WHERE SettlementDate = '" & FormatDate(selectedTradeDate) & "' ORDER BY BrokerName"
		 
		
        
        Set rs = CreateObject("ADODB.Recordset")
        rs.CursorLocation = adUseClient
        Rs.Open SQLServerFormat(HandleQuote(sqlStr)), Conn.ConnectionString, 1, 3
        If rs.EOF Or rs.BOF Then
               %>
                <Script Language="JavaScript">
					alert("No settlement slips available");
					window.parent.history.back();				
                </Script>
                <% Set Rs = Nothing
                Set Conn = Nothing
                Response.End
        End If
        
        rs.MoveFirst
        
	Dim pageNumber
	
	pageNumber = 0

%>
<p id="toPDFOrient" name="toPDFOrient" value="P" style="display:none;">P
<p id="toPDF" name="toPDF">
<%

Do Until rs.EOF
	pageNumber = pageNumber + 1
	
	currBrokerCode = Rs.Fields("BrokerCode").Value %>


<table width=100%>
     <tr class="pageNumbering">
		<td align="left" colspan=2>
			<FONT FACE=ARIAL SIZE=2><B>Page <%=pageNumber%>	</B></FONT>	
		</td>		
	</tr>
     <tr>
		     <td>
		        <b><font face="Arial Narrow" size="4">Settlement Slip</font></b></td>
				
			<td align=right>
				<b><font face="Arial Narrow" size="4"><%= Session("CompanyName") %></font></b></td>
					
			</td>
	</tr>		
     <tr>
		     <td colspan=2>
		        <font face="Arial" size="2">printed for each individual broker&nbsp;</font></td>
		</tr>
	<tr>
		     <td colspan=2>
		        <font face="Arial" size="2">&nbsp;</font></td>
		</tr>	
  
</table>

<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
    <tr>
      <td width="20%" style="border-top-style: solid; border-top-width: 1"><b><font face="Arial Narrow" size="2">TO:</font></b></td>
      <td style="border-top-style: solid; border-top-width: 1"><b><font face="Arial Narrow" size="2">MALAWI STOCK EXCHANGE</font></b></td>
    </tr>
	<tr>
      <td width="20%"><b><font face="Arial Narrow" size="2">BROKER:</font></b></td>
      <td><b><font face="Arial Narrow" size="2"><%= Rs.Fields("BrokerName").Value %></font></b></td>
    </tr>
    <tr>
      <td width="20%" style="border-bottom-style: solid; border-bottom-width: 1">
        <p align="left"><b><font size="2"  face="Arial Narrow">SETTLEMENT DATE:</font></b></td>

      <td style="border-bottom-style: solid; border-bottom-width: 1">
        <p align="left"><font face="Arial" size="2"><%= FormatDateFull(upperLimitDate) %></font></td>
    </tr>


<BR>

<center>
<table border="0" cellspacing=0 cellpadding=3>
<tr>
	<td colspan=7>	
<PRE>
<font face="Arial" size="2">
			<b>Payment to Brokers</b></font>			
<b><font face="Arial" size="2">		
Please prepare payment payable to brokers as per the details shown on this voucher.
Kindly note that all contracts on the list were traded on the same date shown below:</font></b></PRE>

	</td>
</tr>
<tr bgcolor="#000000">
	<td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial" size="2">Slip</font></b></td>
	<td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial" size="2">Traded</font></b></td>	
	<td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial" size="2">Security</font></b></td>	
	<td align=right bgcolor="#000000"><b><font color="#FFFFFF" face="Arial" size="2">Quantity</font></b></td>
	<td align=right bgcolor="#000000"><b><font color="#FFFFFF" face="Arial" size="2">Price</font></b></td>
	<td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial" size="2">Contract</font></b></td>
	<td align=right bgcolor="#000000"><b><font color="#FFFFFF" face="Arial" size="2">Gross</font></b></td>
	
  </tr>
<%		
	totalLevyAmount = 0
	isDifferentBroker = False
	
		 Do Until isDifferentBroker
			totalLevyAmount = totalLevyAmount + FormatNum(rs.Fields("LevyAmount")) %>
                <tr>
                        <td><%=rs.Fields("LotSlipNo")%></td>
                        <td><%= FormatDate(rs.Fields("LotTDate")) %></td>
                        <td><%=rs.Fields("OrdDetailSecurity")%></td>
                        <td align=right><%= FormatNum(rs.Fields("LotQty")) %></td>
                        <td align=right><%= FormatNum(rs.Fields("LotPrice")) %></td>
                        <td style="BORDER-RIGHT: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent"><%=rs.Fields("ContractNumber")%></td>                   
                        <td align=right><%= FormatNum(rs.Fields("LevyAmount")) %></td>
                        
                </tr>
                <%
             Rs.MoveNext   
             If Not (Rs.EOF Or Rs.BOF) Then
				If Rs.Fields("BrokerCode").Value <> currBrokerCode Then					
					Rs.Move -1
					isDifferentBroker = True
				End If
             Else
				isDifferentBroker = True
				Rs.Move -1
             End If   
         Loop      
                %>
                
           <tr>
						<td colspan=7 align=right>&nbsp;</td>
           </tr>               
        
         <tr>
						<td colspan=6 align=right><b>Broker Totals:</b></td>
                        <td align=right style="border-style: solid; border-color: #000000; border-width: 1" height="30px"><%= FormatNum(totalLevyAmount) %></td>
                        
         </tr>        
</table>
</center>
<%

			rs.MoveNext
			'important!		
			If Not Rs.EOF Then %>
				<BR class="newpage">
		<%	End If
        Loop
        
        conn.Close
        Set conn = Nothing
        Set Rs = Nothing%>


</body>

</html>
