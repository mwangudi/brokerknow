<html>

<head>

<title>Batch Report [LandScape]</title>
  
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<LINK rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
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
				page-break-before: always;
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
selectedContractDate = Request.Form("txtDate")

If genReport <> "1" Or selectedContractDate = "" Then
	%>
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
	
	<form method="POST" action="BatchReportsLandscape.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">	
		<table>
			<tr>
				<td>Select date:</td>
				<td>
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>
			</tr>
				
			<tr>
				<td colspan=2><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... " id=Button1 name=Button1>&nbsp;&nbsp;</td>
			</tr>
		</table>
	</form>
	<%
	Response.End
End If
%>

<% DrawPageFunctions True, True, True, True %>

<%
''CONTRACT SCHEDULE

response.write "CONTRACT SCHEDULE"

headerDescription = FormatDateFull(selectedContractDate)
%>

<i id="landRem">Remember to select landscape settings while printing.</i>

<p id="toPDFOrient" name="toPDFOrient" value="L" style="display:none;">L
<p id="toPDF" name="toPDF">

<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
     <tr>
		<td nowrap><b><font face="Arial Narrow" size="4">Contracts Schedule</font></b></td>
		<td nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>
	</tr>	
    <tr>
	   <td COLSPAN=2><font face="Arial" size="2">for Deals traded on:  <%= headerDescription %></font></td>
	</tr>
    <tr>
		  <td COLSPAN=2><font face="Arial" size="2">&nbsp;</font></td>
	</tr>
</table>				

<table border="0" width="100%" cellPadding="2" cellSpacing=0>
	<tr bgColor="#000000">
		<td bgColor="#000000" nowrap align="center"><b><font color="#FFFFFF">Traded</font></b></td>
		<td bgColor="#000000" nowrap align="center"><b><font color="#FFFFFF">Client</font></b></td>
		<td bgColor="#000000" nowrap align="center"><b><font color="#FFFFFF">Security</font></b></td>
		<td bgColor="#000000" nowrap align="center"><b><font color="#FFFFFF">Br</font></b></td>
		<td bgColor="#000000" nowrap align="center"><b><font color="#FFFFFF">Contr</font></b></td>
		<td bgColor="#000000" nowrap align="center"><b><font color="#FFFFFF">CDS Ref</font></b></td>
		<td bgColor="#000000" nowrap align="right"><b><font color="#FFFFFF">Price</font></b></td>
		<td bgColor="#000000" nowrap align="right"><b><font color="#FFFFFF">Quantity</font></b></td>
		<td bgColor="#000000" nowrap>&nbsp;</td>
		<%
		Dim fld
		Dim conn 
		Dim sqlStr
		Dim rs
		Dim i
		Dim dailyTotalsArray1()
		
		beginLeviesCol1=15	
		
		Set conn = GetActiveConnection("KBroker")
 		sqlStr = "ContractLeviesCrossTab"	
		Set Rs = CreateObject("ADODB.Recordset")		
		Rs.CursorLocation = adUseClient		
		Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
		Rs.Filter = "LotTDate = '" & FormatDate(selectedContractDate) & "'"
		
		'i = 0
		'fldCount = fldCount + 2 'this is hard coding just to skip some unwanted columns 
		for i = beginLeviesCol1 to rs.fields.count - 2
			Redim Preserve dailyTotalsArray1(i - beginLeviesCol1)
			dailyTotalsArray1(i - beginLeviesCol1) = 0
			%>
				<td bgColor="#000000" nowrap  align="center"><b><font color="#FFFFFF"><%=rs.fields(i).name%></font></b></td>		
			<%
			'i = i + 1
		next
		
		'add gross, and net amount
		Redim Preserve dailyTotalsArray1((i) - beginLeviesCol1)
		dailyTotalsArray1((i) - beginLeviesCol1) = 0
		
		Redim Preserve dailyTotalsArray1((i + 1) - beginLeviesCol1)
		dailyTotalsArray1((i + 1) - beginLeviesCol1) = 0
		
		'Redim Preserve dailyTotalsArray1((i + 2) - beginLeviesCol1)
		'dailyTotalsArray1((i + 2) - beginLeviesCol1) = 0
		%>
		<td bgColor="#000000" nowrap align="center"><b><font color="#FFFFFF">&nbsp;&nbsp;Gross&nbsp;&nbsp;</font></b></td>
		<td bgColor="#000000" nowrap align="center"><b><font color="#FFFFFF">Net Amount</font></b></td>
    </tr>
    
	<%
    If Not(rs.EOF Or rs.BOF) Then
		'rs.MoveFirst
		Do Until rs.EOF
				%>
        		<tr>
					<td nowrap><font size="1"><%= Day(rs.Fields("LotTDate")) & " " & MonthName(Month(rs.Fields("LotTDate")), True) %></font></td>
					<td nowrap><font size="1">
						<%
						If Len(rs.Fields("ClientName")) > 25 Then 
							Response.Write Mid(rs.Fields("ClientName"), 1, 25)
						Else
							Response.Write rs.Fields("ClientName")	
						End If
						%>
					</font></td>
					<td nowrap><font size="1"><%=rs.Fields("SecurityCode")%></font></td>
					<td nowrap align="center"><font size="1"><%=rs.Fields("BrokerCode")%></font></td>
					<td nowrap><font size="1"><%=rs.Fields("ContractNumber")%></font></td>
					<td nowrap><font size="1"><%=rs.Fields("LotSlipNo")%></font></td>
					<td nowrap align="right"><font size="1"><%=FormatNum(rs.Fields("LotPrice"))%></font></td>
					<td nowrap align="right"><font size="1"><%=FormatNumCommasOnly(rs.Fields("LotQty"))%></font></td>       
					<td style="BORDER-RIGHT: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent">&nbsp;</td>
					<%	
					levyTotals = 0
				
					for i = beginLeviesCol1 to rs.fields.count - 2
						If i = rs.Fields.Count - 2 Then
							myStyle =  "BORDER-RIGHT: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent"
						Else
							myStyle = ""	
						End If
						%>
							<td  nowrap align=right  Style="<%= myStyle %>"><font size="1"><%=FormatNum(rs.fields(i).value)%></font></td>		
						<%
						'i = i + 1						
						levyTotals = levyTotals + rs.fields(i).value					
						dailyTotalsArray1(i - beginLeviesCol1) = dailyTotalsArray1(i - beginLeviesCol1) + rs.fields(i).value
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
				
					dailyTotalsArray1((i) - beginLeviesCol1) = dailyTotalsArray1((i) - beginLeviesCol1) + grossAmt
					dailyTotalsArray1((i + 1) - beginLeviesCol1) = dailyTotalsArray1((i + 1) - beginLeviesCol1) + netAmt
					%>
             		<td nowrap align="right"><font size="1"><%= FormatNum(grossAmt) %></font></td>       
					<td nowrap align="right"><font size="1"><%= FormatNum(netAmt) %></font></td> 
				</tr>      
				<%
             rs.MoveNext
		Loop
        %>
        
        <tr>
			<td colspan="<%= 9 + UBound(dailyTotalsArray1)%>">&nbsp;</td>
        </tr>
      
		<tr height="30px">
			<td colspan=9 align=right><b>Daily totals:</b></td>	
			<%
			for i = 0 to UBound(dailyTotalsArray1)
				If i = 0 Then 
					myStyle = "BORDER-LEFT: #C0C0C0 1px inset; BORDER-TOP: #C0C0C0 1px inset; BORDER-BOTTOM: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent; "
				ElseIf (i <> 0 And i <> UBound(dailyTotalsArray1)) Then
					myStyle = "BORDER-TOP: #C0C0C0 1px inset; BORDER-BOTTOM: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent;"
				Else
					myStyle = "BORDER-RIGHT: #C0C0C0 1px inset;BORDER-TOP: #C0C0C0 1px inset; BORDER-BOTTOM: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent;"
				End If
				%>
				<td  nowrap align=right style="<%= myStyle %>"><font size="1"><%= FormatNum(Trim(dailyTotalsArray1(i))) %></font></td>		
				<%
			next
			%>
		</tr>
		<%
	Else
		%>
	    <script language = 'javascript'>
	    	alert ("No contracts found using the specified criteria");
	    	window.parent.history.go(-1);          		
	    </script>
	    <%
		Set Rs = Nothing
		Set Conn = Nothing

	    Response.end
	End if
	%>
</body>
</html>



