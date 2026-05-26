<html>
<%PAGENUMBERINGENABLED = True%>
<head>

<title>Client Statement</title>  

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
				size: landscape;
				margin-left: 0cm;
				margin-right: 0cm;
				margin-top: 0cm;    
				margin-bottom: 0cm;
				writing-mode: tb-rl;
				p.newpage{
					page-break-before:always;
				}		
			}		 
	</style>
</head>

<body Class="Reports">
<!--#include file="../libroutinesVB.asp"-->

<form name=frm1 id=frm1>
<%
	selectedClient = Request.QueryString("a")
	selectedFromDate = Request.QueryString("b")
	selectedToDate = Request.QueryString("c")
	
	Set conn = GetActiveConnection("KBroker")
 	Set Rs = CreateObject("ADODB.Recordset")						        
	
	'sqlStr = "SELECT * FROM StatementList WHERE Client_DPA_ = '" & selectedClient & "' AND TransDate BETWEEN '" & FormatDate(selectedFromDate) & "' AND '" & FormatDate(DateAdd("d", 1, selectedToDate)) & "'"
	sqlStr = "SELECT * FROM StatementList WHERE Client_DPA_ = '" & selectedClient & "' AND TransDate BETWEEN '" & selectedFromDate & "' AND '" & selectedToDate & "'"
	'sqlStr = "SELECT * FROM StatementList WHERE Client_DPA_ = '" & selectedClient & "' AND TransDate <= '" & selectedToDate & "'"
	
	Rs.CursorLocation = adUseClient	
	Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
	
	If rs.EOF Or rs.BOF Then
		''NO CLIENT RECORDS
		Set Rs = Nothing
		Set Conn = Nothing
		%>
			<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
				<tr>
					<td width="100%" nowrap align="left"><font face="Impact" size="4">Nothing found!</font></td>      
				</tr>
			</table>
		<%
		Response.End
	Else

		Dim pageNumber
		pageNumber = 0	
		Rs.PageSize = 30
		Rs.CacheSize = Rs.PageSize
		intPageCount = Rs.PageCount 
		intRecordCount = Rs.RecordCount 
		first=0
		' Now you must double check to make sure that you are not before the start
		' or beyond end of the recordset.  If you are beyond the end, set 
		' the current page equal to the last page of the recordset.  If you are
		' before the start, set the current page equal to the start of the recordset.	
		if Not (rs.eof or rs.bof) then
			intPageCount = Rs.PageCount 
			intRecordCount = Rs.RecordCount 
		else
			intPageCount=1
			intRecordCount=0
		end if	
		intPage=0
		intPageCount=Cint(intPageCount)
        
        
		''CLIENT RECORDS FOUND
		Set rsClient = Conn.Execute ("SELECT * FROM Client WHERE Client_DPA_ = " & selectedClient)
	
		If Not (rsClient.EOF Or rsClient.BOF) Then
			accountDesc = rsClient.Fields("ClientName").Value & " [  " & rsClient.Fields("Client_DPA_").Value & " ]"
			accountAddress = rsClient.Fields("ClientAddr").Value
		End If
	
		Set rsClient = Nothing

		isOpeningBalance = CBool(Rs.Fields("IsOpeningBalance").Value)

		If Not IsOpeningBalance Then
			''GET THE LATEST PREVIOUS BALANCE
			sqlStr = "SELECT Client_DPA_, '" & FormatDate(selectedFromDate) & "' AS TransDate, '' AS REF, 'Opening Balance' AS Particulars, " & _
					" 0 AS Debit, 0 AS Credit, 1 AS IsOpeningBalance, SUM(Credit - Debit) " & _
					" AS Balance FROM  ClientStatement " & _
					" WHERE     (Client_DPA_ = " & selectedClient & ") AND (TransDate < '" & FormatDate(selectedFromDate) & "')" & _
					" GROUP BY Client_DPA_"
			
			sqlStr = "SELECT Client_DPA_, '" & FormatDate(selectedFromDate) & "' AS TransDate, '' AS REF, 'Opening Balance' AS Particulars, " & _
					" 0 AS Debit, 0 AS Credit, 1 AS IsOpeningBalance, SUM(Credit - Debit) " & _
					" AS Balance FROM  StatementList " & _
					" WHERE     (Client_DPA_ = " & selectedClient & ") AND (TransDate < '" & FormatDate(selectedFromDate) & "')" & _
					" GROUP BY Client_DPA_"
			Set cloneRs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
			
			If cloneRs.EOF Or cloneRs.BOF Then
				OpeningBalance =  Rs.Fields("Balance").Value
			End If
			OpeningBalance = cloneRs.Fields("Balance").Value
		Else
			OpeningBalance =  Rs.Fields("Balance").Value				
		End If
		%>
		<div id="emailDoc" name="emailDoc">
		<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
			<tr>
				<td width="100%" nowrap align="left"><font face="Impact" size="4">CLIENTS STATEMENT</font></td>      
			</tr>
			
			<tr>
				<td width="100%" nowrap align="right">&nbsp;Page&nbsp;<%=1%>&nbsp;of&nbsp;<%=intPageCount%></td>      
			</tr>
		</table>

		<br>
		
		<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
			<tr>
				<td width="1%"><b>Date:</b></td>
				<td width="48%"><%= FormatDate(Date) %></td>
			</tr>

			<tr>
				<td width="1%"><b>Account:</b></td>
				<td width="48%"><%= accountDesc %> </td>
			</tr>

			<tr>
				<td width="1%"><b><font size="2" face="Arial">&nbsp;</font></b></td>
			<td width="48%"><%= accountAddress %></td>
			</tr>
		</table>
		
			
			<BR>
  <table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow; LEFT-MARGIN:100PX"  width="100%">
			<tr>
				<td style="border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"><b><font face="Arial Narrow" size="3">Date</font></b></td>
				<td style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"><b><font face="Arial Narrow" size="3">Ref:</font></b></td>
				<td style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"><b><font face="Arial Narrow" size="3">Particulars:</font></b></td>
				<td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"  align=right><b><font face="Arial Narrow" size="3">Debit:</font></b></td>
				<td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" align=right><b><font face="Arial Narrow" size="3">Credit:</font></b></td>
				<td align="right" style="border-right-style: solid; border-right-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"  align=right><b><font face="Arial Narrow" size="3">Balance</font></b></td>
			</tr>

			<%
			totalDebits = 0
			totalCredits = 0
    
			kwanza=0
			'Do Until Rs.EOF
			Do while Cint(intPage) < intPageCount	
				intPage=intPage + 1
				If CInt(intPage) > CInt(intPageCount) Then intPage = intPageCount
				If CInt(intPage) <= 0 Then intPage = 1
				'Make sure that the recordset is not empty.  If it is not, then set the 
				'AbsolutePage property and populate the intStart and the intFinish variables.
				If intRecordCount > 0 Then 'and Not(Rs.eof and Rs.bof)
					Rs.AbsolutePage = intPage
					intStart = Rs.AbsolutePosition
							
					If CInt(intPage) = CInt(intPageCount) Then
						intFinish = intRecordCount
					Else
						intFinish = intStart + (Rs.PageSize - 1)
					End if
				End If	  
				
				If Not IsOpeningBalance AND (intPage = 1) Then
					%>
					<tr>	
						<td><font size="1"><%= Day(selectedFromDate) & " " & MonthName(Month(selectedFromDate), True) & " " & Right(Year(selectedFromDate),2) %></font></td>
						<td><font size="1">&nbsp;</font></td>
						<td><font size="1">Opening Balance</font></td>
						<td align="right"><font size="1">&nbsp;</font></td>
						<td align="right"><font size="1">&nbsp;</font></td>
						<td align="right"><font size="1"><%= CreditDebitValue(FormatNum(OpeningBalance)) %></font></td>
					</tr>    
					<%
					runningBal = CreditDebitValueRev(OpeningBalance)
					'OpeningBalance = cloneRs.Fields("Balance").Value
					'Set cloneRs = Nothing
				Else
					runningBal = 0	
				End If
									
				totalDebits = totalDebits + Rs.Fields("Debit").Value 
				totalCredits = totalCredits + Rs.Fields("Credit").Value
				
				if Not (rs.eof or rs.bof) then
					For intRecord = 1 to Rs.PageSize
						
						IF PAGENUMBERINGENABLED THEN
							if intRecord = 1 then
								If intPage <> 1 Then
									%>
									<tr>
								<td colspan="6">
								<p class="newpage">
								</td>
							</tr>
										
							<tr>
								<td colspan=6 align=center><img src="Include/aaprintlogo.jpg" width="482" height="178"></td>
							</tr>
										
							<tr>
								<td colspan=6 align=center>
									<table>
										<tr>
											<td colspan="2" width="100%" nowrap align="left"><font face="Impact" size="4">CLIENTS STATEMENT</font></td>      
										</tr>
													
										<tr>
											<td colspan="2" width="100%" nowrap align="right">&nbsp;Page&nbsp;<%=intPage%>&nbsp;of&nbsp;<%=intPageCount%></td>      
										</tr>
													
										<tr>
											<td width="1%"><b>Date:</b></td>
											<td width="48%"><%= FormatDate(Date) %></td>
										</tr>

										<tr>
											<td width="1%"><b>Account:</b></td>
											<td width="48%"><%= accountDesc %> </td>
										</tr>

										<tr>
											<td width="1%"><b><font size="2" face="Arial">&nbsp;</font></b></td>
											<td width="48%"><%= accountAddress %></td>
										</tr>
													
										<tr>
											<td colspan="2" width="100%" nowrap align="left">&nbsp;</td>      
										</tr>
									</table>
											
								</td>
							</tr>
								<tr>
							<td style="border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"><b><font face="Arial Narrow" size="3">Date</font></b></td>
							<td style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"><b><font face="Arial Narrow" size="3">Ref:</font></b></td>
							<td style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"><b><font face="Arial Narrow" size="3">Particulars:</font></b></td>
							<td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"  align=right><b><font face="Arial Narrow" size="3">Debit:</font></b></td>
							<td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" align=right><b><font face="Arial Narrow" size="3">Credit:</font></b></td>
							<td align="right" style="border-right-style: solid; border-right-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"  align=right><b><font face="Arial Narrow" size="3">Balance</font></b></td>
						</tr>	
									<tr>	
						<td><font size="1"><%= Day(lastDate) & " " & MonthName(Month(lastDate), True) & " " & Right(Year(lastDate),2) %></font></td>
						<td><font size="1">&nbsp;</font></td>
						<td><font size="1">BALANCE B/F</font></td>
						<td align="right"><font size="1">&nbsp;</font></td>
						<td align="right"><font size="1">&nbsp;</font></td>
						<td align="right"><font size="1"><%= FormatNum(BalOnNextPage) %></font></td>
					</tr>  
									<%
								End If
								
								%>
								
								<%
								
							end if
							kwanza=1
						END IF
						
						
						%>
						<tr>	
							<td><font size="1"><%= Day(rs.Fields("TransDate")) & " " & MonthName(Month(rs.Fields("TransDate")), True) & " " & Right(Year(rs.Fields("TransDate")),2) %></font></td>
							<td><font size="1"><%= Rs.Fields("Ref").Value %></font></td>
							<%
							if(Rs("receipttype")=1) then
								%>
								<td><font size="1">RECEIPT:<%= Mid(Ucase(Rs.Fields("Particulars").Value),1,47) %></font></td>
								<%
							else
								if(Rs("receipttype")=3) then
									%>
									<td><font size="1">PAYMENT:<%= Mid(Ucase(Rs.Fields("Particulars").Value),1,47) %></font></td>
									<%			
								else
									if(rs("IsOpeningBalance")=1) then				
										If Not IsOpeningBalance Then
										else
											%>
											<td><font size="1"><%= Mid(Ucase(Rs.Fields("Particulars").Value),1,55) %></font>&nbsp;</td>
											<%
										end if
									else
										%>
										<td><font size="1"><%= Mid(Ucase(Rs.Fields("Particulars").Value),1,55) %></font>&nbsp;</td>
										<%					
									end if
								end if
							end if
							%>
							<td align="right"><font size="1">
								<%
								If Rs.Fields("Debit").Value <> "0" Then 
									Response.Write FormatNum(Rs.Fields("Debit").Value) 
								End If		
								%>
							</font></td>
							<td align="right"><font size="1">
								<%
								If Rs.Fields("Credit").Value <> "0" Then
									Response.Write FormatNum(Rs.Fields("Credit").Value) 
								End If			
								%>
							</font></td>
							<td align="right"><font size="1">
							<%
							If (rs("IsOpeningBalance") <> 1 ) Then
								runningBal = runningBal + (Rs.Fields("Credit").Value - Rs.Fields("Debit").Value)					
								Response.Write  FormatNum(CreditDebitValue(runningBal)) 
							Else	
								If Not IsOpeningBalance Then
								else
									runningBal = runningBal + Rs.Fields("Balance").Value					
									Response.Write FormatNum(Rs.Fields("Balance").Value)
								end if
							End If 
							lastDate = Rs.Fields("TransDate").Value
							%>
							</font></td>
						</tr>
						<%
						
						IF PAGENUMBERINGENABLED THEN
							if intRecord = Rs.PageSize then
								%>
								<tr>
									<td style="border-top: 1 solid gray;border-bottom: 1 solid gray;">&nbsp;</td>
									<td style="border-top: 1 solid gray;border-bottom: 1 solid gray;">&nbsp;</td>
									<td style="border-top: 1 solid gray;border-bottom: 1 solid gray;"><b>Running Balance</b></td>
									<td style="border-top: 1 solid gray;border-bottom: 1 solid gray;">&nbsp;</td>
									<td style="border-top: 1 solid gray;border-bottom: 1 solid gray;">&nbsp;</td>
									<td style="border-top: 1 solid gray;border-bottom: 1 solid gray;" align="right"><b><%=FormatNum(CreditDebitValue(runningbal))%></b></td>
								</tr>
								<%
								BalOnNextPage = CreditDebitValue(runningbal)
							end if
						END IF
						
						Rs.MoveNext
				
						If Rs.EOF Then Exit for
					next
				end if
			Loop
			%>
			
			<tr>	
				<td><font size="1"><%= Day(lastDate) & " " & MonthName(Month(lastDate), True) & " " & Right(Year(lastDate),2) %></font></td>
				<td><font size="1">&nbsp;</font></td>
				<td><font size="1">Closing Balance</font></td>
				<td align="right"><font size="1">&nbsp;</font></td>
				<td align="right"><font size="1">&nbsp;</font></td>
				<td align="right"><font size="1"><%= FormatNum(CreditDebitValue(runningBal)) %></font></td>
			</tr>  
				
			<tr>
				<td colspan="6" align="right" style="border-bottom-style: solid; border-bottom-width: 1">&nbsp;&nbsp;&nbsp; </td>
			</tr>

			<tr>
				<td colspan="5" align="right"><font size="1">Opening Balance:</font></td>
				<td align="right"><font size="1"><%= FormatNum(CreditDebitValueRev(OpeningBalance)) %></font></td>
			</tr>

			<tr>
				<td colspan="5" align="right"><font size="1">less Total Debits:</font></td>
				<td align="right"><font size="1"><%= FormatNum(0 - totalDebits) %></font></td>
			</tr>

			<tr>
				<td colspan="5" align="right"><font size="1">add Total Credits:</font></td>
				<td align="right"> <font size="1"> <%= FormatNum(totalCredits) %></font></td>
			</tr>

			<tr>
				<td colspan="6" align="right" style="border-bottom-style: solid; border-bottom-width: 1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>
			</tr>

			<tr>
				<td colspan="2"><font size="1">Current: <%= FormatNum(CreditDebitValue(runningBal)) %></font></td>
				<td><font size="1">30-60 Days:</font> C </td>
				<td><font size="1">Over 60 Days:</font> </td>
				<td align="right"><b><font size="1">Total Balance:</font></b></td>
				<td align="right"><b><font size="1"><%= FormatNum(CreditDebitValue(runningBal)) %></font></b></td>
			</tr>
		</table>
		</div>
		<%
		
		
		Set Rs = Nothing 
		Set Conn = Nothing
	End If
	%>
	<input type="hidden" value="" id="hidData" name="hidData">
	<input type="hidden" value="<%=selectedClient%>" id="hidClient" name="hidClient">
	<input type="hidden" value="Statement" id="hidCategory" name="hidCategory">


</form>
</body>
</html>

<%'Response.End%>

<script language="vbscript">
	theData = emailDoc.innerHTML 
	frm1.hidData.value = theData
	frm1.method = "post"
	frm1.action = "PDFMail.asp"
	frm1.submit 
</script>

