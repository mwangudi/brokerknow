<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Contract Schedule</title>
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
			
			margin-left: 0cm;
			margin-right: 0cm;
			margin-top: 0cm;    
			margin-bottom: 0cm;			
			
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
	window.onbeforeprint = HideRemindSelectLandscape;
	window.onafterprint = ShowRemindSelectLandscape;
</Script>

<!--#include file="../libroutines.asp"-->

<%

const beginLeviesCol = 15

genReport = Request.Form("genReport")
selectedTradeDate = Request.Form("txtDate")

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
	<form method="POST" action="ContractSchedule.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">	
		<table>
			<tr>
				<td>Select Date</td>
				<td>
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>
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


<% DrawPageFunctions True, True, True
		Dim fld
		Dim conn 
		Dim sqlStr
		Dim rs
		Dim i
		Dim dailyTotalsArray()
		Dim TotalDailyTotals()
		Dim LevyName
		Dim PageNumbers
		Dim PageNumber
		Dim intPageCount		' The number of pages in the recordset.
		Dim intRecordCount		' The number of records in the recordset.
		Dim intPage			' The current page that we are on.
		Dim intRecord			' Counter used to iterate through the recordset.
		Dim intStart			' The record that we are starting on.
		Dim intFinish			' The record that we are finishing on.
		Dim first
		
		Set conn = GetActiveConnection("KBroker")
 		sqlStr = "ContractLeviesCrossTab"	
		Set Rs = CreateObject("ADODB.Recordset")		
		Rs.CursorLocation = adUseClient		
		Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
		
		'Confirm that the stored procedure returned information anticipated
		 if rs.EOF or rs.BOf then
		  %>
                <script language = 'javascript'>
                		alert ("No Contract Schedules Found");
                		window.parent.history.go(-1);          		
                </script>
                
                <% Response.End   
		 end if
		
		Rs.Filter = "LotTDate = '" & FormatDate(selectedTradeDate) & "'"
		Rs.Sort="Contract"
		Rs.PageSize=30
		
		Conn.Execute("ShortName")

		Rs.CacheSize = Rs.PageSize
		intPageCount = Rs.PageCount 
		intRecordCount = Rs.RecordCount 
	
		first=0
		
		'Response.write(intPagecount)
		
		'Rs.Getrows(10)
		'="Select Top 10"
		
		
		
	' Now you must double check to make sure that you are not before the start
	' or beyond end of the recordset.  If you are beyond the end, set 
	' the current page equal to the last page of the recordset.  If you are
	' before the start, set the current page equal to the start of the recordset.	

	Rs.CacheSize = Rs.PageSize
	intPageCount = Rs.PageCount 
	intRecordCount = Rs.RecordCount 
	

	
		headerDescription = FormatDateFull(selectedTradeDate) 
        PageNumber1=PageNumber1 + 1
        
        intPage=0
        
         intPageCount=Cint(intPageCount)
        
      t=0  
    'i=Cint(intPageCount)
     'Do until t = intPageCount 
    
	
	Do while Cint(intPage) < intPageCount
	if(Cint(first)=1) then
	%>
             <BR class="newpage">
    <%
	end if

	first=1
	intPage=intPage + 1
	'Response.write(intpage)
	
	If CInt(intPage) > CInt(intPageCount) Then intPage = intPageCount
	If CInt(intPage) <= 0 Then intPage = 1
	
	 'Make sure that the recordset is not empty.  If it is not, then set the 
	 'AbsolutePage property and populate the intStart and the intFinish variables.
	If intRecordCount > 0 Then
		Rs.AbsolutePage = intPage
		intStart = Rs.AbsolutePosition
		'Response.write(intStart)
		
		If CInt(intPage) = CInt(intPageCount) Then
			intFinish = intRecordCount
		Else
			intFinish = intStart + (Rs.PageSize - 1)
		End if
	End If	
        %>
<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">	 
     <tr>		
		<td colspan=2 nowrap align=right><font face="Arial Narrow" size="5">Page&nbsp;<%= intPage %></font></td>
	</tr>	    
     <tr>
		<td nowrap><font face="Arial Narrow" size="5">Contracts</font><font face="Arial Narrow" size="6"> </font>
        <font face="Arial Narrow" size="5">Schedule</font></td>
		<td nowrap align=right><font face="Arial Narrow" size="5"><%= Session("CompanyName") %></font>&nbsp;</td>
	</tr>	
    <tr>
	   <td COLSPAN=2><font face="Arial Narrow"  size="5">for</font><font face="Arial Narrow"  size="5"> </font>
       <font face="Arial Narrow"  size="5">Deals</font><font face="Arial Narrow"  size="5"> </font>
       <font face="Arial Narrow"  size="5">traded</font><font face="Arial Narrow"  size="5"> </font>
       <font face="Arial Narrow"  size="5">on</font><font face="Arial Narrow"  size="5">:  <%= headerDescription %></font></td>
	</tr>
    <tr>
		  <td COLSPAN=2><font face="Arial Narrow"  size="5">&nbsp;</font></td>
	</tr>
</table>				

    <table border="0" width="100%" cellPadding="2" cellSpacing=3>
    <tr>
			
	  <td  nowrap align="left"><font face="Arial Narrow"  size="4" ><u>CODE</u></font></td>
      <td  nowrap align="left" ><font face="Arial Narrow"  size="4" ><u>CLIENT</u></font></td>
      <td  nowrap align="left"><u><font face="Arial Narrow"  size="4" >BR</font><font face="Arial Narrow"  size="5" >&nbsp;</font></u></td>
      <td  nowrap align="left"><font face="Arial Narrow"  size="4" ><u>CONT</u></font></td>
      <td  nowrap align="left"><font face="Arial Narrow"  size="4" ><u>SECUR</u></font></td>            
      <td  nowrap align="left"><font face="Arial Narrow"  size="4" ><u>SLIP</u></font></td>
      <td  nowrap align="Right"><font face="Arial Narrow"  size="4" ><u>PRICE</u></font></td>
      <td  nowrap align="Right"><font face="Arial Narrow"  size="4" ><u>QTY</u></font></td>
      <td  nowrap>&nbsp;</td>
      <%
		
		'i = 0
		'fldCount = fldCount + 2 'this is hard coding just to skip some unwanted columns 
		for i = beginLeviesCol to rs.fields.count - 2
				Redim Preserve dailyTotalsArray(i - beginLeviesCol)
				dailyTotalsArray(i - beginLeviesCol) = 0
								
				%>
					<td nowrap  align="right" width="10%"><font face="Arial Narrow" size="4" ><u><%=Ucase(mid(rs.fields(i).name,2,4))%></u></font>&nbsp;</td>		
				<%				
		next
		
		'add gross, and net amount
		Redim Preserve dailyTotalsArray((i) - beginLeviesCol)
		dailyTotalsArray((i) - beginLeviesCol) = 0
		
		Redim Preserve dailyTotalsArray((i + 1) - beginLeviesCol)
		dailyTotalsArray((i + 1) - beginLeviesCol) = 0
		
		'Redim Preserve dailyTotalsArray((i + 2) - beginLeviesCol)
		'dailyTotalsArray((i + 2) - beginLeviesCol) = 0
      %>
		
		<td nowrap align="right"><font  face="Arial Narrow" size="4" ><u>&nbsp;&nbsp;</font><font  face="Arial Narrow" size="4" ><span lang="en-us">GROSS</span></font><font  face="Arial Narrow" size="4" >&nbsp;&nbsp;</font></u></td>
		<td nowrap align="right"><u><font  face="Arial Narrow" size="4" ><span lang="en-us">NET AMOUNT</span></font></u></td>
    </tr>        
 <%
 i=0
 If Not(rs.EOF Or rs.BOF) Then
        'rs.MoveFirst
        'Do Until rs.EOF
        For intRecord = 1 to Rs.PageSize        
%>
     
        		<tr>
      <td nowrap><font face="Arial Narrow" size="4"><%= rs("Client_DPA_")%></font>&nbsp;</td>
      <td nowrap><font face="Arial Narrow" size="4"><% If Len(rs.Fields("ClientName")) > 25 Then 
									Response.Write Mid(rs.Fields("ClientName"), 1, 25)
								   Else
									Response.Write rs.Fields("ClientName")	
								   End If	 %></font>&nbsp;</td>
	  <td nowrap align="center"><font face="Arial Narrow" size="4"><%=rs.Fields("BrokerCode")%></font>&nbsp;</td>
	  <td nowrap><font face="Arial Narrow" size="4"><%=rs.Fields("ContractNumber")%></font>&nbsp;</td>
      <td nowrap><font face="Arial Narrow" size="4"><%=rs.Fields("SecurityCode")%></font>&nbsp;</td>            
      <td nowrap><font face="Arial Narrow" size="4"><%=rs.Fields("LotSlipNo")%></font>&nbsp;</td>
      <td nowrap align="right"><font face="Arial Narrow" size="4"><%=FormatNum(rs.Fields("LotPrice"))%></font>&nbsp;</td>
      <td nowrap align="right"><font face="Arial Narrow" size="4"><%=FormatNumCommasOnly(rs.Fields("LotQty"))%></font>&nbsp;</td>       
      <td style="BORDER-RIGHT: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent">&nbsp;</td>
           <%	levyTotals = 0
				for i = beginLeviesCol to rs.fields.count - 2
						If i = rs.Fields.Count - 2 Then
							myStyle =  "BORDER-RIGHT: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent"
						Else
							myStyle = ""	
						End If
						%>
							<td  nowrap align=right  Style="<%= myStyle %>"><font face="Arial Narrow" size="4"><%=FormatNum(rs.fields(i).value)%></font>&nbsp;</td>		
						<%
						'i = i + 1						
						levyTotals = levyTotals + rs.fields(i).value					
						dailyTotalsArray(i - beginLeviesCol) = dailyTotalsArray(i - beginLeviesCol) + rs.fields(i).value
				next
				
				'check whether this contract has the
				'levy of agent commission type, which is not
				'really a levy, but a rate of "broker commission" levy
				'if so, minus this value from the levyTotals variable
				sqlStr = "SELECT * FROM LevyContract WHERE Contract_DPA_ = " & rs.Fields("Contract_DPA_").Value & " AND SystemMaintained = 12"  
				Set tmpRs = Conn.Execute (SQLServerFormat(HandleQuote(sqlStr)))
				If Not (tmpRs.EOF OR tmpRs.BOF) Then
					'agent commission exists
					levyTotals = levyTotals - tmpRs.Fields("LevyAmount").Value
				End If
				Set tmpRs = Nothing		
				
				
				grossAmt = rs.Fields("LotGrossAmount") 'rs.Fields("LotPrice") * rs.Fields("LotQty") 
				
				If rs.Fields("OrderTypeSale").Value = 0 Then 
					netAmt = grossAmt + levyTotals 
				Else
					netAmt = grossAmt - levyTotals 
				End If
				
				dailyTotalsArray((i) - beginLeviesCol) = dailyTotalsArray((i) - beginLeviesCol) + grossAmt
				dailyTotalsArray((i + 1) - beginLeviesCol) = dailyTotalsArray((i + 1) - beginLeviesCol) + netAmt

				%>
             
             
             <td nowrap align="right"><font face="Arial Narrow" size="4"><%= FormatNum(grossAmt) %></font>&nbsp;</td>       
             <td nowrap align="right"><font face="Arial Narrow" size="4"><%= FormatNum(netAmt) %></font>&nbsp;</td> 
             </tr>  
                             
             <%
             rs.MoveNext
        'Loop
        'Rs.MoveNext
		If Rs.EOF Then Exit for

        Next
        %>        
        
     <%else%>
                <script language = 'javascript'>
                		alert ("No contracts found using the specified criteria");
                		window.parent.history.go(-1);          		
                </script>
                
                <%                  
             
               		Set Rs = Nothing
					Set Conn = Nothing
                Response.end
     
   End if    
   			
              'TotalDailyTotals
              Redim Preserve TotalDailyTotals(UBound(dailyTotalsArray))
				
			  for i = 0 to UBound(dailyTotalsArray)
			  TotalDailyTotals(i)=TotalDailyTotals(i)+dailyTotalsArray(i)
			  next

			if(intPage=intPageCount) then
				%>
				<tr>
			<td colspan="<%= 9 + UBound(dailyTotalsArray)%>">&nbsp;</td>
        </tr>
      
		<tr height="30px">
			<td colspan=9 align=right>
			<font face="Arial Narrow" size="4">	
					DAILY</font><font face="Arial Narrow" size="4">	
					</font>		
			<font face="Arial Narrow" size="4">	
					TOTALS</font><font face="Arial Narrow" size="4">:
			</font>		
			</td>	
			 <%	
				for i = 0 to UBound(TotalDailyTotals)
						If i = 0 Then 
							myStyle = "BORDER-LEFT: #C0C0C0 1px inset; BORDER-TOP: #C0C0C0 1px inset; BORDER-BOTTOM: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent; "
						ElseIf (i <> 0 And i <> UBound(TotalDailyTotals)) Then
							myStyle = "BORDER-TOP: #C0C0C0 1px inset; BORDER-BOTTOM: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent;"
						Else
							myStyle = "BORDER-RIGHT: #C0C0C0 1px inset;BORDER-TOP: #C0C0C0 1px inset; BORDER-BOTTOM: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent;"
						End If
						
					
						%>
							<td  nowrap align=right style="<%= myStyle %>"><font face="Arial Narrow" size="4"><%= FormatNum(Trim(TotalDailyTotals(i))) %></font>&nbsp;</td>		
						<%
				next
				
				%>
		</tr>
  </table>
  <%
  else
  %>
  </table>
  <%
  end if
             loop
	%>	

</body>

</html>