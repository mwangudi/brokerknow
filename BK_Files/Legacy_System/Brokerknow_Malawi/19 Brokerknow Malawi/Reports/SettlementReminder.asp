<%'server.scripttimeout=10000%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Settlement Reminder</title>
  
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
	 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
	 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
	  <!--CALENDAR -->
	<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
	<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
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
	<form method="POST" action="SettlementReminder.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">	
		<table>
			<tr>
				<td>Select date</td>
				<td>
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>
			</tr>
			<tr>
				<td>&nbsp;</td>
				<td>
					<p><input type="radio" class="BorderLess" checked Name="timeLimit" value="0" id="USETime"><label for="USETime" style="cursor: hand">Use MSE time limit</label></p>
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


<% DrawPageFunctions True, True, True, True%>

<%
   Dim conn 
   Dim sqlStr
   Dim rs
   Dim timeRange
	
		Set conn = GetActiveConnection("KBroker")
        
        sqlStr = "SELECT * FROM [TimeLimitList] WHERE TimeLimit_DPA_ = 1" 'Settlement (Brokers)
        Set tempRs = Conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        
        If Not (tempRs.EOF Or tempRs.BOF) Then
			If timeLimit = "0" Then
				'use USE
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
					upperLimitDate = JumpUpFromWeekendToWeek(upperLimitDate)
					isProperDate = True
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
        
        
        sqlStr = "SELECT * FROM [SettlementReminder] WHERE TRADED = '" & FormatDate(originalselectedTradeDate) & "' ORDER BY Traded DESC"
		
	
        Set rs = conn.execute(sqlstr)'CreateObject("ADODB.Recordset")
        'rs.CursorLocation = adUseClient
        'Rs.Open SQLServerFormat(HandleQuote(sqlStr)), Conn.ConnectionString, 1, 3
        If rs.EOF Or rs.BOF Then
               %>
                <Script Language="JavaScript">
					alert("No settlement reminders available");
					window.parent.history.go(-1);
                </Script>
                <% Set Rs = Nothing
                Set Conn = Nothing
                Response.End
        End If
        
        
        rs.MoveFirst
        

%>
<p id="toPDFOrient" name="toPDFOrient" value="P" style="display:none;">P
<p id="toPDF" name="toPDF">
<table width="100%">
     <tr>
		     <td>
		        <b><font face="Arial Narrow" size="4">Settlement Reminder</font></b></td>
		      <td align=right>
				<b><font face="Arial Narrow" size="4"><%= Session("CompanyName") %> </font></b></td>
					
			</td>  
		</tr>	
     <tr>
		     <td colspan=2> 
		        <font face="Arial" size="2">for funds from or to brokers on: <%= FormatDateFull(upperLimitDate)%></font></td>
		</tr>
	<tr>
		     <td  colspan=2>
		        <font face="Arial" size="2">&nbsp;</font></td>
		</tr>	
  
</table>


  <table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
    <tr>
      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow" size="4">Traded</font></b></td>
      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow" size="4">Broker</font></b></td>

      <td bgcolor="#000000">
        <p align="right"><b><font color="#FFFFFF" face="Arial Narrow" size="4">Gross</font></b></p>
    </td>
    </tr>
      <tr>
      <td colspan="3">&nbsp;</td>
    </tr>
    
    <%
    dailyTotals = 0
    myTotals = 0
    
    Rs.Filter = "OrderTypeSale = 0"
    
    If Not (Rs.EOF Or Rs.BOF) Then%>
    <tr>
      <td colspan="3"><font size="3" face="Impact">Amounts to be paid to brokers</font></td>
    </tr>
    <%	
		
		Do Until Rs.EOF
			myTotals = myTotals + Rs.Fields("GrossAmt").Value%>
		<tr>
		  <td><%= FormatDate(Rs.Fields("Traded").Value) %></td>
		  <td><%= Rs.Fields("Broker").Value  %></td>
		     <td>
		    <p align="right"><%= FormatNum(Rs.Fields("GrossAmt").Value)%></td>
		</tr>
		
		
	<%		Rs.MoveNext
		Loop%>
		
		<tr>
		  <td>&nbsp;</td>
		  <td>&nbsp;</td>	
		  <td style="border-top: solid; border-width: 1" align=right valign=bottom height="25px"><%= FormatNum(myTotals) %></td>
		</tr>
		  
		<%
		
	End If
    
    Rs.Cancel
    
    dailyTotals = dailyTotals + myTotals
    myTotals = 0
    
    Rs.Filter = "OrderTypeSale = 1"
    
    If Not (Rs.EOF Or Rs.BOF) Then%>    
    <tr>
      <td colspan="3">&nbsp; </td>
    </tr>
    
    <tr>
      <td colspan="3"><font size="3" face="Impact">Amounts expected from brokers</font></td>
    </tr>
    
    
     <%	
		
		Do Until Rs.EOF
			myTotals = myTotals + Rs.Fields("GrossAmt").Value%>
		<tr>
		  <td><%= FormatDate(Rs.Fields("Traded").Value) %></td>
		  <td><%= Rs.Fields("Broker").Value  %></td>
		     <td>
		    <p align="right"><%= FormatNum(Rs.Fields("GrossAmt").Value)%></td>
		</tr>
		
		
	<%		Rs.MoveNext
		Loop%>
		
		<tr>
		  <td>&nbsp;</td>
		  <td>&nbsp;</td>		
		  <td style="border-top: solid; border-width: 1; border-color: #000000;" align=right valign=bottom height="25px"><%= FormatNum(myTotals) %></td>
		</tr>
		  
		<%
		
	End If
	
	dailyTotals = dailyTotals + myTotals%>
	
	 <tr>
      <td colspan="3">&nbsp; </td>
    </tr>
     <tr>
      <td colspan="3">&nbsp; </td>
    </tr>   
    <tr>
      <td></td>
  
    <td>
      <p align="right">Daily Totals:&nbsp;&nbsp; </td>
    <td style="border-style: solid; border-width: 1">
			<p align="right"><%= FormatNum(dailyTotals) %></p>
		</td>
    </tr>
  </table>


<%
Set Rs = Nothing
Set Conn = Nothing
%>
</body>

</html>
