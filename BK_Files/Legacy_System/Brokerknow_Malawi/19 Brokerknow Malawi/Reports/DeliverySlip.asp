<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Delivery Slip</title>
 
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

<!--#include file="../libroutines.asp"-->

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
	<form method="POST" action="DeliverySlip.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">	
		<table>
			<tr>
				<td>Select trade date:</td>
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
			<tr style="display: none">
				<td>Select Transaction Type:</td>
				<td>
					<SELECT NAME="TransType">
						<OPTION VALUE=0 SELECTED>Sale</OPTION>
						<OPTION VALUE=1>Purchase</OPTION>		
					</SELECT>
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
   Dim conn 
   Dim sqlStr
   Dim rs
   
   Set conn = GetActiveConnection("KBroker")
   sqlStr = "SELECT * FROM [TimeLimitList] WHERE TimeLimit_DPA_ = 2" 'Delivery (Shares)
   Set tempRs = Conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        
    If Not (tempRs.EOF Or tempRs.BOF) Then
		If timeLimit = "0" Then
			'use NSE
			sharesTimeRange = tempRs.Fields("TimeLimitNSE").Value
		Else
			'use internal
			sharesTimeRange = tempRs.Fields("TimeLimitInternal").Value
		End If		
			
    Else%>
		<Script Language="JavaScript">
			alert("The time limits of type Delivery (Shares) have not been set");
			window.parent.self.close();				
        </Script>
		<%
		Set tempRs = Nothing
		Set Conn = Nothing
		Response.End
    End If
    
    
    sqlStr = "SELECT * FROM [TimeLimitList] WHERE TimeLimit_DPA_ = 3" 'Delivery (Bonds)
    Set tempRs = Conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        
    If Not (tempRs.EOF Or tempRs.BOF) Then
		If timeLimit = "0" Then
			'use NSE
			fixedTimeRange = tempRs.Fields("TimeLimitNSE").Value
		Else
			'use internal
			fixedTimeRange = tempRs.Fields("TimeLimitInternal").Value
		End If		
			
    Else%>
		<Script Language="JavaScript">
			alert("The time limits of type Delivery (Bonds) have not been set");
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
        deliveryDate = FormatDate(selectedTradeDate)
        theSelDay = WeekDay(selectedTradeDate)
        
        'get to a monday if selected date is on weekend
        deliveryDate = JumpUpFromWeekendToWeek (deliveryDate)
        
		
		'check upwards until upper date is not a holiday
		For k = 1 To 2               
				upperLimitDate = deliveryDate
				
				If k = 1 Then
					timeRange = sharesTimeRange
				Else
					timeRange = fixedTimeRange
				End If	
        
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
				
				If k = 1 Then
					sharesDeliveryDate = upperLimitDate
				Else
					fixedDeliveryDate = upperLimitDate
				End If	
        
				
        
        Next
        
        
        Set holidayRs = Nothing
        
      Dim RecordsFound 
      RecordsFound = false  
        
   
   For i = 0 To 1
		transType = i 
		 
		sqlStr = "SELECT * FROM [DeliverySlips] WHERE OrderTypeSale = " & transType & " AND LotTDate = '" & FormatDate(originalselectedTradeDate) & "'"
		       
        Set Rs = CreateObject("ADODB.Recordset")						       
		Rs.CursorLocation = adUseClient	
		Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
        
        If Not (rs.EOF Or rs.BOF) Then
          RecordsFound = true     
        
        rs.MoveFirst
        
        
%>


<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
     <tr>
		     <td>
		        <b><font face="Arial Narrow" size="4">Delivery Slip</font></b></td>
			 <td nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>			        
		</tr>	
     <tr>
		     <td colspan=2>
		        <font face="Arial" size="2">to be produced in duplicate</font></td>
		</tr>
    <tr>
		     <td colspan=2>
		        <font face="Arial" size="2">&nbsp;</font></td>
		</tr>
</table>

<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
    <tr>
      <td  style="border-top-style: solid; border-top-width: 1"><b><font face="Arial Narrow" size="2">TO:</font></b></td>
      <td style="border-top-style: solid; border-top-width: 1"><b><font face="Arial Narrow" size="2">MALAWI STOCK EXCHANGE</font></b></td>
    </tr>

    <tr>
      <td  style="border-bottom-style: solid; border-bottom-width: 1">
        <p align="left"><b><b><font face="Arial Narrow" size="2">EXPECTED DELIVERY DATE:</font></b></td>

      <td style="border-bottom-style: solid; border-bottom-width: 1">
        <p align="left"><font face="Arial" size="2">
			<%
			If LCase(Rs.Fields("OrderSecTypeDisplayName").Value) = "security" Then
				Response.Write FormatDateFull(sharesDeliveryDate) 
			Else
				Response.Write FormatDateFull(fixedDeliveryDate) 
			End If%>
		</font></td>
    </tr>

  </table>


<BR>


  <table border="0" cellspacing="0" cellpadding="4" style="font-family: Arial Narrow; LEFT-MARGIN:100PX">
    <tr>
      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Slip</font></b></td>
      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Traded</font></b></td>
      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Security</font></b></td>
      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Quantity</font></b></td>
      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Price</font></b></td>
       <%
      If transType = 1 Then
		detailInfo = "Certificate No"
      Else
		detailInfo = "Transfer No"
      End If
      %>
      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow" size="3"><%= detailInfo %></font></b></td>
    </tr>
	<%Do Until Rs.EOF
        %>
    <tr>
      <td align="left"><%= Rs.Fields("LotSlipNo").Value %>	</td>

      <td><%= FormatDate(Rs.Fields("LotTDate").Value) %></td>
      <td><%= Rs.Fields("SecurityCode").Value %></td>
      <td align="right"><%= FormatNum(Rs.Fields("LotQty").Value) %></td>
      <td align="right"><%= FormatNum(Rs.Fields("LotPrice").Value) %></td>
      <%
      If transType = 1 Then
		detailInfo = Rs.Fields("ContractNCertificate").Value
      Else
		detailInfo = Rs.Fields("ContractTransferNo").Value
      End If
      %>
      <td><%= detailInfo %></td>
    </tr>
	
	<%
		Rs.MoveNext
	Loop
	%>
  </table>
  
  <BR class="newpage">
   

<% 
		
	End If
Next


Set Rs = Nothing
Set Conn = Nothing
'Give message if no records found
if not RecordsFound then
 %>
                <script language = 'javascript'>
                		alert ("No Match Found");
                		window.parent.history.go(-1);          		
                </script>
                
                <% Response.End  
end if 
%>


</body>
</html>
