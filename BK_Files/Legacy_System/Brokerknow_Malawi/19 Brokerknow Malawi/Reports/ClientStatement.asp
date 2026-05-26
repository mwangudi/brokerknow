<html>
<%PAGENUMBERINGENABLED = True%>
<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
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

genReport = Request("genReport")
selectedClient = Request("cboClient")
selectedFromDate = Request("transFromDate")
selectedToDate = Request("transToDate") 

If genReport <> "1" Or selectedClient = "" Or Not IsDate(selectedFromDate) Then%>
	<Script Language="JavaScript">
		report_SetBodyClass();
		function validateForm(frm){			
			if (frm.cboClient.selectedIndex < 0){
				alert("Select a client");
				frm.cboClient.focus();
				return;
			}
			
			frm.target = '_self';			
			frm.submit();
		}
		
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "transFromDate","cmdDate","<%= FormatDate(DateAdd("d", -90, Date)) %>",1);
		var cal2=new ctlSpiffyCalendarBox("cal2", "frmMain", "transToDate","cmdDate2","<%= FormatDate(Date) %>",1);
	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="ClientStatement.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">
		<table border=0>
			<tr>
				<td>Client: </td>
				<td><input type = 'text' name ='txtClientCode' id = 'txtClientCode' size="10" onBlur="txtval = this.value; selectItem(cboClient);"> 
				&nbsp;&nbsp;&nbsp;
				<select name = "cboClient" id = "cboClient" size="1" 
    				onchange = "UpdateCode(true,cboClient,txtClientCode)"
					onKeypress = "return (dodefaultaction()==''); "  
					onKeydown = "return (dodefaultaction()==''); " 
					onKeyup = "change(cboClient,0);"
					<!--onKeyup = "return (change(cboClient,0));"-->
				  onfocus = "txtval = '';inputIsItemCode = 0;" 
					onblur = "txtval = '';inputIsItemCode = 0;">
					<option selected SearchCode = "" SearchText = ""  value = ""></option>
					<%
					        Set conn = GetActiveConnection("KBroker")
					        
					        sqlStr = "SELECT * FROM Client ORDER BY ClientName"
					        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
					        If Not (rs.EOF Or rs.BOF) Then
					                rs.MoveFirst
					                Do Until rs.EOF%>
					                        <option SearchCode = "<%=rs.Fields("Client_DPA_")%>" SearchText = "<%=rs.Fields("ClientName")%>"  value = '<%=rs.Fields("Client_DPA_")%>'><%=mid(rs.Fields("ClientName"),1,30)%></option>
					                        <%rs.MoveNext
					                Loop
					        End If
					%>

					    </select>
				</td>
			</tr>
			<tr>
				<td>Select date from:</td>
				<td>
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>
			</tr>
			<tr>
				<td>Select date to:</td>
				<td>
					<SCRIPT language="JavaScript">cal2.writeControl();</SCRIPT>	
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

<% DrawPageFunctions True, True, True, True %>

<%
	Set conn = GetActiveConnection("KBroker")
 	Set Rs = CreateObject("ADODB.Recordset")						        
	sqlStr = "SELECT * FROM StatementList WHERE Client_DPA_ = '" & selectedClient & "' AND TransDate BETWEEN '" & FormatDate(selectedFromDate) & "' AND '" & FormatDate(DateAdd("d", 1, selectedToDate)) & "'"
	
	'Response.Write sqlStr
	Rs.CursorLocation = adUseClient	
	conn.begintrans
		conn.execute("ClientTotalProcedure " & selectedClient)
		conn.execute("ClientBalanceProcedure " & selectedClient)
	conn.committrans
	Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
	
	If rs.EOF Or rs.BOF Then%>
		<Script Language="JavaScript">
			alert("The specified client does not have any transaction using the specified date criterion")
			window.history.go(-1);
		</Script>
		<%Set Rs = Nothing
		Set Conn = Nothing
		Response.End
	End If
	
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
		
	Set rsClient = Conn.Execute ("SELECT * FROM Client WHERE Client_DPA_ = " & selectedClient)
	
	If Not (rsClient.EOF Or rsClient.BOF) Then
		'accountDesc = rsClient.Fields("ClientName").Value & " [  " & rsClient.Fields("Client_DPA_").Value & " ] "
		If Len(rsClient("ClientContact")) > 0 Then	
			accountDesc = rsClient.Fields("ClientName").Value & " [  " & rsClient.Fields("Client_DPA_").Value & " ] " 
		Else
			accountDesc = rsClient.Fields("ClientName").Value & " [  " & rsClient.Fields("Client_DPA_").Value & " ] " 
		End If
		accountAddress = rsClient.Fields("ClientAddr").Value
	End If
	
	Set rsClient = Nothing

	isOpeningBalance = CBool(Rs.Fields("IsOpeningBalance").Value)

	If Not IsOpeningBalance Then
		'get latest prev balance
		
		sqlStr = "SELECT Client_DPA_, '" & FormatDate(selectedFromDate) & "' AS TransDate, '' AS REF, 'Opening Balance' AS Particulars, " & _
				" 0 AS Debit, 0 AS Credit, 1 AS IsOpeningBalance, SUM(isnull(Credit,0) - isnull(Debit,0)) " & _
				" AS Balance FROM  ClientStatement " & _
				" WHERE     (Client_DPA_ = " & selectedClient & ") AND (TransDate < '" & FormatDate(selectedFromDate) & "')" & _
				" GROUP BY Client_DPA_"
		
		sqlstr =" SELECT     Client_DPA_, TransDate, REF, Particulars, Debit, Credit, IsOpeningBalance, SUM(Balance) AS Balance " & _
				" FROM         (SELECT     Client_DPA_, '" & FormatDate(selectedFromDate) & "' AS TransDate, '' AS REF, 'Opening Balance' AS Particulars, 0 AS Debit, 0 AS Credit, 1 AS IsOpeningBalance,  " & _
				"                                               SUM(isnull(Credit, 0) - isnull(Debit, 0)) AS Balance " & _
				"                        FROM          StatementList " & _
				"                        WHERE      (Client_DPA_ = " & selectedClient & ") AND (TransDate < '" & FormatDate(selectedFromDate) & "') " & _
				"                        GROUP BY Client_DPA_ " & _
				"                        UNION " & _
				"                        SELECT     " & selectedClient & " AS Client_DPA_, '" & FormatDate(selectedFromDate) & "' AS TransDate, '' AS REF, 'Opening Balance' AS Particulars, 0 AS Debit, 0 AS Credit,  " & _
				"                                              1 AS IsOpeningBalance, 0 AS Balance " & _
				"                        FROM         StatementList) Derivedtbl " & _
				" GROUP BY Client_DPA_, TransDate, REF, Particulars, Debit, Credit, IsOpeningBalance "
		
		Set cloneRs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		If cloneRs.EOF Or cloneRs.BOF Then%>
				<Script Language="JavaScript">
					alert("A problem occured when calculating the opening balance");
					window.history.go(-1);
				</Script>
			<%	Set cloneRs = Nothing
				Set Rs = Nothing
				Set Conn = Nothing
				Response.End
		End If
		OpeningBalance = cloneRs.Fields("Balance").Value

	Else
		OpeningBalance =  Rs.Fields("Balance").Value				
	End If
%>	
<p id="toPDFOrient" name="toPDFOrient" value="P" style="display:none;">P
<p id="toPDF" name="toPDF">
<table border="0" cellspacing="0" cellpadding="0" style="font-family: Arial Narrow" width="100%">
	<tr>
		<td width="100%" align="middle" valign="top">
			<!--#include file="Header.asp"--></td>
	</tr>
	<!--<tr>
		<td width="100%" height="60" align=center>&nbsp;</td>
	</tr>-->
	<tr>
		<td width="100%" nowrap align="left"><font face="Impact" size="4">CLIENTS STATEMENT</font></td>      
	</tr>
	<tr>
		<td width="100%" nowrap height="0" align="right">&nbsp;Page&nbsp;<%=1%>&nbsp;of&nbsp;<%=intPageCount%></td>      
	</tr>
</table>

<table border="0" cellspacing="0" cellpadding="0" style="font-family: Arial Narrow" width="100%">
    <tr>
      <td width="1%"><b>Date:&nbsp;</b></td>
      <td width="48%"><%= FormatDate(Date) %></td>
    </tr>

    <tr>
      <td width="1%"><b>Account:&nbsp;</b></td>
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

    <%If Not IsOpeningBalance Then%>
		<tr>	
		  <td><font size="1"><%= Day(selectedFromDate) & " " & MonthName(Month(selectedFromDate), True) & " " & Right(Year(selectedFromDate),2) %></font></td>
		  <td><font size="1">&nbsp;</font></td>
		  <td><font size="1">Opening Balance</font></td>
		  <td align="right"><font size="1">&nbsp;</font></td>
		  <td align="right"><font size="1">&nbsp;</font></td>
		  <td align="right"><font size="1"><%= CreditDebitValue(FormatNum(cloneRs.Fields("Balance").Value)) %></font></td>
		</tr>    
    <%	runningBal = CreditDebitValueRev(cloneRs.Fields("Balance").Value)
		OpeningBalance = cloneRs.Fields("Balance").Value
		Set cloneRs = Nothing
		
	Else
	 runningBal = 0	
    End If
    
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
		
		if Not (rs.eof or rs.bof) then
			For intRecord = 1 to Rs.PageSize
				
				totalDebits = totalDebits + Rs.Fields("Debit").Value 
				totalCredits = totalCredits + Rs.Fields("Credit").Value
							
				IF PAGENUMBERINGENABLED THEN
					if intRecord = 1 then
						If intPage <> 1 Then
							%>
							</table>
<BR class="newpage">
							<!--<tr>
								<td colspan="6">
								<p class="newpage">
								</td>
							</tr>-->
							<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow; LEFT-MARGIN:100PX"  width="100%">			
							
							<tr>
								<td colspan=6 height="60" align=center>&nbsp;</td>
							</tr>
							
							<tr>
								<td colspan=6 align=center><img src="../data/photos/aaprintlogofooter.jpg"></td>
							</tr>
										
							<!--<tr>
								<td colspan=6 height="60" align=center>&nbsp;</td>
							</tr>-->
										
							<tr>
								<td colspan=6 height="60" align=center>
									<table border="0" cellspacing="0" cellpadding="0">
										<tr>
											<td colspan="2" width="100%" nowrap align="left"><font face="Impact" size="4">CLIENTS STATEMENT</font></td>      
										</tr>
													
										<tr>
											<td colspan="2" width="100%" nowrap height="0" align="right">&nbsp;Page&nbsp;<%=intPage%>&nbsp;of&nbsp;<%=intPageCount%></td>      
										</tr>
													
										<tr>
											<td width="1%"><b>Date:&nbsp;</b></td>
											<td width="48%"><%= FormatDate(Date) %></td>
										</tr>

										<tr>
											<td width="1%"><b>Account:&nbsp;</b></td>
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
      <td colspan="5" align="right">
        <font size="1">
        Opening Balance:</font></td>

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
      <td colspan="6" align="right" style="border-bottom-style: solid; border-bottom-width: 1">
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>

    </tr>

    <tr>
      <td colspan="2"><font size="1">Current: <%= FormatNum(CreditDebitValue(runningBal)) %></font></td>
      <td><font size="1">30-60 Days:</font> C
		
      </td>
      <td><font size="1">Over 60 Days:</font> </td>
      <td align="right"><b><font size="1">Total Balance:</font></b></td>
      <td align="right"><b><font size="1"><%= FormatNum(CreditDebitValue(runningBal)) %></font></b></td>
    </tr>
 
 	<td colspan=2>&nbsp;</td>
	</tr>
	<tr>
		<td>
		<i>For and on behalf of</i>
		<br> Cedar Capital Limited
		<br><br>Sign..........................................
		</td>
		
		<td>&nbsp;</td>
	</tr>
		<tr>
		<td valign=top colspan=2>
	

		</tr>	
	<tr>
						<td colspan="6"   align = "center" nowrap valign="top">
						<br>
						<small><font face="ARAIAL" color='black'>“Please notify Cedar Capital Limited in writing if any confirmation, contract note or statement is incorrect immediately <br>following your receipt of the confirmation, contract note or statement and, in any event, <br> by no later than 16h00 on the first business day following the trade date, <br> failing which (in the absence of manifest error) such confirmation, contract note or <br> statement will be conclusive and binding on you.” <Font>
						</small> </td>
					</tr>
 
  </table>
   
<%Set Rs = Nothing 
Set Conn = Nothing%>    
</body>
</html> 
 