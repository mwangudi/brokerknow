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
			
			margin-left: 1cm;
			margin-right: 1cm;
			margin-top: 1cm;    
			margin-bottom: 1cm;
			writing-mode: tb-rl;
			height: 90%;
			margin: 10% 0%;			
			
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
	window.onbeforeprint = HideRemindSelectLandscape;
	window.onafterprint = ShowRemindSelectLandscape;
</Script>

<!--#include file="../libroutinesTEST.asp"-->

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
	<form method="POST" action="down_ContractSchedule.asp" Name="frmMain" id="frmMain">
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


<% headerDescription = FormatDateFull(selectedTradeDate) %>


	<%

		Dim fso, txtLine, txtStream, txtStreamOut, filename
			
			filename = "ContractSchedule_" & FormatDate(selectedTradeDate) & ".csv"

			Set fso = Server.CreateObject("Scripting.FileSystemObject")
			Set txtStreamOut = fso.CreateTextFile("D:\18 - Brokerknow Botswana - Test\Downloads\Downloaded\" & filename,true)

	%>

<p id="toPDFOrient" name="toPDFOrient" value="L" style="display:none;">L
<p id="toPDF" name="toPDF">

<i id="landRem">Remember to select landscape settings while printing.</i>
<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
     <tr>
		<td nowrap><b><font face="Arial Narrow" size="4">Contracts Schedule</font></b></td>
			<%txtLine="Contracts Schedule" & chr(13) & chr(10) %>
		<td nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>
	</tr>	
    <tr>
	   <td COLSPAN=2><font face="Arial" size="2">for Deals traded on:  <%= headerDescription %></font></td>
			<%txtLine=txtLine & "for Deals traded on: " & replace(headerDescription, ",", " ") & chr(13) & chr(10) %>
	</tr>
    <tr>
		  <td COLSPAN=2><font face="Arial" size="2">&nbsp;</font></td>
	</tr>
</table>				

    <table border="0" width="100%" cellPadding="2" cellSpacing=0>
    <tr bgColor="#000000">

			
	  <td bgColor="#000000" nowrap align="center"><b><font color="#FFFFFF">Traded</font></b></td>
			<%txtLine=txtLine & "Traded"%>
      <td bgColor="#000000" nowrap align="center"><b><font color="#FFFFFF">Client</font></b></td>
			<%txtLine=txtLine & "," & "Client"%>
      <td bgColor="#000000" nowrap align="center"><b><font color="#FFFFFF">Security</font></b></td>
			<%txtLine=txtLine & "," & "Security"%>
      <td bgColor="#000000" nowrap align="center"><b><font color="#FFFFFF">Br</font></b></td>
			<%txtLine=txtLine & "," & "Br"%>
      <td bgColor="#000000" nowrap align="center"><b><font color="#FFFFFF">Contr</font></b></td>
			<%txtLine=txtLine & "," & "Contr"%>
      <td bgColor="#000000" nowrap align="center"><b><font color="#FFFFFF">REF</font></b></td>
			<%txtLine=txtLine & "," & "REF"%>
      <td bgColor="#000000" nowrap align="right"><b><font color="#FFFFFF">Price</font></b></td>
			<%txtLine=txtLine & "," & "Price"%>
      <td bgColor="#000000" nowrap align="right"><b><font color="#FFFFFF">Quantity</font></b></td>
			<%txtLine=txtLine & "," & "Quantity"%>
      <td bgColor="#000000" nowrap>&nbsp;</td>
			<%'txtLine=txtLine & "," & " "%>
      <%


	


		Dim fld
		Dim conn 
		Dim sqlStr
		Dim rs
		Dim i
		Dim dailyTotalsArray()
			
		Set conn = GetActiveConnection("KBroker")
 		sqlStr = "ContractLeviesCrossTab"	
		Set Rs = CreateObject("ADODB.Recordset")		
		Rs.CursorLocation = adUseClient		
		Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
		Rs.Filter = "LotTDate = '" & FormatDate(selectedTradeDate) & "'"
		
		'i = 0
		'fldCount = fldCount + 2 'this is hard coding just to skip some unwanted columns 
		for i = beginLeviesCol to rs.fields.count - 2
				Redim Preserve dailyTotalsArray(i - beginLeviesCol)
				dailyTotalsArray(i - beginLeviesCol) = 0
				%>
					<td bgColor="#000000" nowrap  align="center"><b><font color="#FFFFFF"><%=rs.fields(i).name%></font></b></td>		
						<%txtLine=txtLine & "," & rs.fields(i).name%>
				<%
				'i = i + 1
		next
		
		'add gross, and net amount
		Redim Preserve dailyTotalsArray((i) - beginLeviesCol)
		dailyTotalsArray((i) - beginLeviesCol) = 0
		
		Redim Preserve dailyTotalsArray((i + 1) - beginLeviesCol)
		dailyTotalsArray((i + 1) - beginLeviesCol) = 0
		
		'Redim Preserve dailyTotalsArray((i + 2) - beginLeviesCol)
		'dailyTotalsArray((i + 2) - beginLeviesCol) = 0

				

	  %>
		<td bgColor="#000000" nowrap align="center"><b><font color="#FFFFFF">&nbsp;&nbsp;Gross&nbsp;&nbsp;</font></b></td>
			<%txtLine=txtLine & "," & "Gross"%>
		<td bgColor="#000000" nowrap align="center"><b><font color="#FFFFFF">Net Amount</font></b></td>
			<%txtLine=txtLine & "," & "Net Amount"%>
    </tr>
			<%txtLine=txtLine & chr(13) & chr(10) %>
 <%
     If Not(rs.EOF Or rs.BOF) Then
        'rs.MoveFirst
        Do Until rs.EOF%>
        		<tr>
      <td nowrap><font size="1"><%= Day(rs.Fields("LotTDate")) & " " & MonthName(Month(rs.Fields("LotTDate")), True) %></font></td>
			<%txtLine=txtLine & Day(rs.Fields("LotTDate")) & " " & MonthName(Month(rs.Fields("LotTDate")), True)%>
      <td nowrap><font size="1"><% If Len(rs.Fields("ClientName")) > 25 Then 
									Response.Write Mid(rs.Fields("ClientName"), 1, 25)
								   Else
									Response.Write rs.Fields("ClientName")	
								   End If	 %></font></td>
			<%txtLine=txtLine & "," & Mid(rs.Fields("ClientName"), 1, 25)%>
      <td nowrap><font size="1"><%=rs.Fields("SecurityCode")%></font></td>
			<%txtLine=txtLine & "," & Mid(rs.Fields("SecurityCode"), 1, 25)%>
      <td nowrap align="center"><font size="1"><%=rs.Fields("BrokerCode")%></font></td>
			<%txtLine=txtLine & "," & Mid(rs.Fields("BrokerCode"), 1, 25)%>
      <td nowrap><font size="1"><%=rs.Fields("ContractNumber")%></font></td>
			<%txtLine=txtLine & "," & Mid(rs.Fields("ContractNumber"), 1, 25)%>
      <td nowrap><font size="1"><%=rs.Fields("LotSlipNo")%></font></td>
			<%txtLine=txtLine & "," & Mid(rs.Fields("LotSlipNo"), 1, 25)%>
      <td nowrap align="right"><font size="1"><%=FormatNum(rs.Fields("LotPrice"))%></font></td>
			<%txtLine=txtLine & "," & Mid(rs.Fields("LotPrice"), 1, 25)%>
      <td nowrap align="right"><font size="1"><%=FormatNumCommasOnly(rs.Fields("LotQty"))%></font></td>       
			<%txtLine=txtLine & "," & Mid(rs.Fields("LotQty"), 1, 25)%>
      <td style="BORDER-RIGHT: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent">&nbsp;</td>
			<%'txtLine=txtLine & "," & Mid(rs.Fields("ClientName"), 1, 25)%>
           <%	levyTotals = 0
				for i = beginLeviesCol to rs.fields.count - 2
						If i = rs.Fields.Count - 2 Then
							myStyle =  "BORDER-RIGHT: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent"
						Else
							myStyle = ""	
						End If
						%>
							<td  nowrap align=right  Style="<%= myStyle %>"><font size="1"><%=FormatNum(rs.fields(i).value)%></font></td>		
								<%txtLine=txtLine & "," & rs.fields(i).value %>
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
             
             
             <td nowrap align="right"><font size="1"><%= FormatNum(grossAmt) %></font></td>       
				<%txtLine=txtLine & "," & grossAmt %>
             <td nowrap align="right"><font size="1"><%= FormatNum(netAmt) %></font></td> 
				<%txtLine=txtLine & "," & netAmt %>
             </tr>      
				<%txtLine=txtLine & chr(13) & chr(10) %>
             <%rs.MoveNext
        Loop%>
        
        <tr>
			<td colspan="<%= 9 + UBound(dailyTotalsArray)%>">&nbsp;</td>
				<%txtLine=txtLine & ",,,,,," %>
        </tr>
      
		<tr height="30px">
			<td colspan=9 align=right>
				<b>
					Daily totals:
				</b>	
			</td>	
				<%txtLine=txtLine & "," & "Daily totals: " %>
			 <%	
				for i = 0 to UBound(dailyTotalsArray)
						If i = 0 Then 
							myStyle = "BORDER-LEFT: #C0C0C0 1px inset; BORDER-TOP: #C0C0C0 1px inset; BORDER-BOTTOM: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent; "
						ElseIf (i <> 0 And i <> UBound(dailyTotalsArray)) Then
							myStyle = "BORDER-TOP: #C0C0C0 1px inset; BORDER-BOTTOM: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent;"
						Else
							myStyle = "BORDER-RIGHT: #C0C0C0 1px inset;BORDER-TOP: #C0C0C0 1px inset; BORDER-BOTTOM: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent;"
						End If
						
					
						%>
							<td  nowrap align=right style="<%= myStyle %>"><font size="1"><%= FormatNum(Trim(dailyTotalsArray(i))) %></font></td>		
								<%txtLine=txtLine & "," & Trim(dailyTotalsArray(i)) %>
						<%
				next
				
				%>
		</tr>
			<%txtLine=txtLine & chr(13) & chr(10) %>
        
     <%else%>
                <script language = 'javascript'>
					alert ("No contracts found using the specified criteria");
					window.location.replace("down_List.asp");
                </script>
                
     <%  
		Set Rs = Nothing
		Set Conn = Nothing
		Set fso = Nothing

		Response.end
     
   End if


		txtStreamOut.WriteLine (txtLine)
		
		conn.execute ("DELETE FROM down_File WHERE     (Filename = '" & filename & "') or (cast(floor(cast(TimeCreated as float)) as datetime) < cast(floor(cast(getdate() as float)) as datetime))")

		sqlStr=" INSERT INTO down_File (Filename, Report, CreatedBy, CreatedByDesc) "
		sqlStr=sqlStr & " SELECT '" & filename & "' AS Filename, '" & "Contract Schedule" & "' AS Report, " & Session("UserID") & " AS CreatedBy,  "
		sqlStr=sqlStr & " (SELECT LEFT(LTRIM(RTRIM(OtherNames)) + ' ' + LTRIM(RTRIM(Surname)), 100) AS CreatedByDesc  "
		sqlStr=sqlStr & " FROM Users  "
		sqlStr=sqlStr & " WHERE   (UserID = " & Session("UserID") & ")) AS CreatedByDesc  "

		conn.execute(sqlStr)
		'response.write sqlstr
		'response.end

 		Set Rs = Nothing
		Set Conn = Nothing
		Set fso = Nothing


 
 %>
	
	
	
  </table>
 

</body>

</html>

<SCRIPT Language="JavaScript">
	alert ("Schedule generated!");
	window.location.replace("down_List.asp");
</SCRIPT>

