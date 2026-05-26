<%
server.ScriptTimeout = 60000000
Response.Expires = 0
Response.Buffer = false
%>

<html>
<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Traded Levies (Non Custodians)</title>
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

<!--#include file="../libroutines.asp"-->

<%

const beginLeviesCol = 13

genReport = Request.Form("genReport")
selectedFromDate = Request.Form("transFromDate")
selectedToDate = Request.Form("transToDate")

If genReport <> "1" Or Not IsDate(selectedFromDate)  Or Not IsDate(selectedToDate)Then%>
	<Script Language="JavaScript">
		report_SetBodyClass();
		
		function validateForm(frm){			
			if (frm.transFromDate.value=='' ){
				alert("Select a date");
				frm.transFromDate.focus();
				return;
			}if (frm.transToDate.value=='' ){
				alert("Select a date");
				frm.transToDate.focus();
				return;
			}
			
			frm.target = '_self';			
			frm.submit();
		}
		
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "transFromDate","cmdDate","<%= FormatDate(date-30) %>",1);
		var cal1=new ctlSpiffyCalendarBox("cal1", "frmMain", "transToDate","cmdDate","<%= FormatDate(Date) %>",1);
	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="TradedLeviesNonCust.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">	
		<table>
			<tr>
				<td>Select date from:</td>
				<td>
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>
			</tr>
			<tr>
				<td>Select date To:</td>
				<td>
					<SCRIPT language="JavaScript">cal1.writeControl();</SCRIPT>	
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

selectedTradeDate = FormatDate(selectedFromDate)
headerDescription = formatdate(selectedFromDate) & " to " & formatdate(selectedToDate) %>

<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
     <tr>
		<td nowrap><b><font face="Arial Narrow" size="4">Traded Levies (Non Custodians)</font></b></td>
		<td nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>
	</tr>	
    <tr>
	   <td COLSPAN=2><font face="Arial" size="2">for the period <%= headerDescription %></font></td>
	</tr>
    <tr>
		  <td COLSPAN=2><font face="Arial" size="2">&nbsp;</font></td>
	</tr>
</table>				

    <table border="0" width="100%" cellPadding="2" cellSpacing=0>
    <tr bgColor="#000000">
			
	  <td nowrap><b><font color="#FFFFFF">No. </font></b></td>
	  <td nowrap><b><font color="#FFFFFF">Traded</font></b></td>
      <td nowrap><b><font color="#FFFFFF">Type</font></b></td>
      <td nowrap><b><font color="#FFFFFF">Security</font></b></td>
      <!--<td nowrap align="center"><b><font color="#FFFFFF">Broker</font></b></td>-->
      <td nowrap><b><font color="#FFFFFF">Contract</font></b></td>
      <td nowrap><b><font color="#FFFFFF">Slip</font></b></td>
      <td nowrap align="right"><b><font color="#FFFFFF">Price</font></b></td>
      <td nowrap align="right"><b><font color="#FFFFFF">Quantity</font></b></td>
      <td nowrap>&nbsp;</td>
      <%
		Dim fld
		Dim conn 
		Dim sqlStr
		Dim rs
		Dim i
		Dim TrNox
		
		'Dim upperDate, lowerDate
		Dim dailyTotalsArray()
		
		Set conn = GetActiveConnection("KBroker")
		conn.commandtimeout = 0
		TrNox = 0 
		procstr= "ContractLeviesCrossTabForNonCustodians '"& selectedFromDate & "', '" & selectedToDate & "'" 
		'response.write procstr:response.end
 		sqlStr =  procstr
		Set Rs = CreateObject("ADODB.Recordset")		
		Rs.CursorLocation = adUseClient		
		Rs.Open HandleQuote(sqlStr), conn, adOpenKeyset, adLockOptimistic
		'Rs.Sort = "LotTDate"
		
		'i = 0
		'fldCount = fldCount + 2 'this is hard coding just to skip some unwanted columns 
		
		dim skipTransfer
		dim colTransfer
		
		'response.write selectedTradeDate :response.end
		
		if cdate(selectedTradeDate) > #28 Feb 2005# then
				skipTransfer = true
				colTransfer = 0 'Default
		else
				skipTransfer = false
		end if
		
		Icount = 0
		for i = beginLeviesCol to rs.fields.count - 2
				

				IF (trim(lcase(rs.fields(i).name)) = "commission") or (trim(lcase(rs.fields(i).name)) = "Agent") then
					'ignore the above levies
				else
					IF (trim(lcase(rs.fields(i).name)) = "transfer") AND (skipTransfer = true) then
							'do nonthing
							'not needed anymore

							colTransfer = -1 'Note the column position so as to skip it when totaling.
					else
							
							Redim Preserve dailyTotalsArray(Icount)
							dailyTotalsArray(Icount) = 0
				
							%>
								<td width="60px" nowrap align="right"><b><font color="#FFFFFF"><%=rs.fields(i).name%></font></b></td>		
							<%
							'i = i + 1
							Icount =Icount+1
					end if
					
				end if
		next
		
		'add gross
		Redim Preserve dailyTotalsArray(Icount)
		dailyTotalsArray(Icount) = 0
		%>
		<td nowrap align="right"><b><font color="#FFFFFF">&nbsp;&nbsp;Net&nbsp;&nbsp;</font></b></td>
		
    </tr>
 <%
     If Not(rs.EOF Or rs.BOF) Then
        'rs.MoveFirst
        Do Until rs.EOF
			thisDate = FormatDate(rs.Fields("LotTDate"))
			
			'If Month(thisDate) = Month(selectedTradeDate) And Year(thisDate) = Year(selectedTradeDate) Then%>
			<%TrNox = TrNox + 1%>
        		<tr>
      <td nowrap><%= TrNox %></td>
	  <td nowrap><%= thisDate %></td>
      <td nowrap><%If instr(1,rs.Fields("ContractNumber"),"S") > 0 Then
						Response.Write "Sale"
				   Else
						Response.Write "Purchase"
				   End If %>
	  </td>
      <td nowrap><%=rs.Fields("SecurityCode")%></td>
     <!-- <td nowrap align="center"><></td>-->
      <td nowrap><%=rs.Fields("ContractNumber")%></td>
      <td nowrap><%=rs.Fields("LotSlipNo")%></td>
      <td nowrap align="right"><%=FormatNum(rs.Fields("LotPrice"))%></td>
      <td nowrap align="right"><%=FormatNum(rs.Fields("LotQty"))%></td>       
      <td style="BORDER-RIGHT: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent">&nbsp;</td>
           <%	levyTotals = 0
				ICount = 0
				for i = beginLeviesCol to rs.fields.count - 2
						If i = rs.Fields.Count - 2 Then
							myStyle =  "BORDER-RIGHT: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent"
						Else
							myStyle = ""	
						End If
						IF (trim(lcase(rs.fields(i).name)) = "commission") or (trim(lcase(rs.fields(i).name)) = "agent") then
							'ignore the above levies
						else
							IF (trim(lcase(rs.fields(i).name)) = "transfer") AND (skipTransfer = true) then
									'do nonthing
							else
									
									
									%>
										<td  nowrap align=right width="60px" Style="<%= myStyle %>"><%=FormatNum(rs.fields(i).value)%></td>		
									<%
									
									levyTotals = levyTotals + rs.fields(i).value
												
									dailyTotalsArray(ICount) = dailyTotalsArray(ICount) + rs.fields(i).value
									ICount = ICount + 1	
							end if
						end if
						
						'i = i + 1
						
				next
				
				'check whether this contract has the
				'levy of agent commission type, which is not
				'really a levy, but a rate of "broker commission" levy
				'if so, minus this value from the levyTotals variable
				'sqlStr = "SELECT * FROM LevyContract WHERE Contract_DPA_ = " & rs.Fields("Contract_DPA_").Value & " AND SystemMaintained = 12"  
				'Set tmpRs = Conn.Execute (SQLServerFormat(HandleQuote(sqlStr)))
				'If Not (tmpRs.EOF OR tmpRs.BOF) Then
					'agent commission exists
					'levyTotals = levyTotals - tmpRs.Fields("LevyAmount").Value
				'End If
				'Set tmpRs = Nothing		
				
				
				
				'grossAmt = rs.Fields("LotGrossAmount")
				
				'If rs.Fields("OrderTypeSale").Value = 0 Then 
					'netAmt = grossAmt + levyTotals 
				'Else
				'	netAmt = grossAmt - levyTotals 
				'End If
				
				'dailyTotalsArray(ICount) = dailyTotalsArray(ICount) + grossAmt
		
				'response.write UBound(dailyTotalsArray) :response.end
				%>
             
             
            <!-- <td nowrap align="right"><%= FormatNum(grossAmt) %></td>   -->
			 <td nowrap align="right"><%
						netAmt = levyTotals 						
				   
				   Response.Write formatnum(netAmt)
				   dailyTotalsArray(ICount) = dailyTotalsArray(ICount) + netAmt%>
		<td nowrap align="right">
             
             </tr>      
             <%'End If

			 rs.MoveNext
        Loop%>
        
        <tr>
			<td colspan="<%= 7 + UBound(dailyTotalsArray)%>">&nbsp;</td>
        </tr>
      
		<tr height="30px">
			<td colspan="9" align=right>
				<b>
					Totals:
				</b>	
			</td>	
			 <%	
				for i = 0 to UBound(dailyTotalsArray)
						'response.write i& ": - "
						If i = 0 Then 
							myStyle = "BORDER-LEFT: #C0C0C0 1px inset; BORDER-TOP: #C0C0C0 1px inset; BORDER-BOTTOM: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent; "
						ElseIf (i <> 0 And i <> UBound(dailyTotalsArray)) Then
							myStyle = "BORDER-TOP: #C0C0C0 1px inset; BORDER-BOTTOM: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent;"
						Else
							myStyle = "BORDER-RIGHT: #C0C0C0 1px inset;BORDER-TOP: #C0C0C0 1px inset; BORDER-BOTTOM: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent;"
						End If
						'IF (trim(lcase(rs.fields(i).name)) = "commission") or (trim(lcase(rs.fields(i).name)) = "cdsc") then
							'ignore the above levies
						'else
							if (skipTransfer = true AND colTransfer = i ) then
							 'do nothing
							else
						
								%>
									<td  nowrap align=right style="<%= myStyle %>"><%= FormatNum(Trim(dailyTotalsArray(i))) %></td>		
								<%
							end if
						'end if
				next
				
				%>
		</tr>
        
     <%else%>
                <script language = 'javascript'>
                		alert ("No contracts found using the specified criteria");
                		window.parent.history.go(-1);          		
                </script>
                
                <%  Set Rs = Nothing
					Set Conn = Nothing
                response.end
     
   End if
 %>
	
	
	
  </table>
 

</body>

</html>
