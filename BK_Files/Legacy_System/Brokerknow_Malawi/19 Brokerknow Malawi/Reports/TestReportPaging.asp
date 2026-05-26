<html>

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
			
			margin-left: 1cm;
			margin-right: 1cm;
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
selectedClient = Request.Form("cboClient")
selectedFromDate = Request.Form("transFromDate")
thedate="3/1/2005"

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
		
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "transFromDate","cmdDate","<%= FormatDate(thedate) %>",1);
	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="ClientStatement.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">
		<table>
			<tr>
				<td>Client: </td>
				<td><input type = 'text' name ='txtClientCode' id = 'txtClientCode' size="10" onBlur="txtval = this.value; selectItem(cboClient);"></td>
				<td><select name = 'cboClient' id = "cboClient" size="1" 
    				onchange='UpdateCode(true,cboClient,txtClientCode)'
					onKeypress="return (dodefaultaction()==''); " 
					onKeydown="return (dodefaultaction()==''); " 
					onKeyup="return (FilterData(this,1,UpdateCode(change(cboClient,0),cboClient,txtClientCode)));" 
					onfocus="txtval = '';inputIsItemCode = 1;" 
					onblur="txtval = '';inputIsItemCode = 1;">
					<option selected SearchCode = "" SearchText = ""  value = ''></option>
					<%
					dim ClientName
					dim NameClient
					        Set conn = GetActiveConnection("KBroker")
					        
					        sqlStr = "SELECT * FROM Client where deleted=0 order by Ltrim(RTrim(ClientName))"
					        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
					        If Not (rs.EOF Or rs.BOF) Then
					                rs.MoveFirst
					                Do Until rs.EOF
					                ClientName=rs.Fields("ClientName")
					                NameClient=Mid(ClientName,1,30)
					                %>					                        
					                        <option SearchCode = "<%=rs.Fields("Client_DPA_")%>" SearchText = "<%=NameClient%>" value = '<%=rs.Fields("Client_DPA_")%>'><%=NameClient%></option>

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
				<td colspan=2><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... ">&nbsp;&nbsp;</td>
			</tr>
		</table>
		
	</form>
	
	<%Set rs = Nothing
	Set Conn = Nothing
	Response.End
End If

%>

<% DrawPageFunctions True, True, True %>

<%
	Set conn = GetActiveConnection("KBroker")
	Set Rs = CreateObject("ADODB.Recordset")
	Rs.CursorLocation = adUseClient		
						        
	sqlStr = "SELECT * FROM StatementList WHERE Client_DPA_ = " & selectedClient & " AND TransDate >= '" & FormatDate(selectedFromDate) & "'"	
	
	Rs.CursorLocation = adUseClient	
	Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
	
	Set rsClient = Conn.Execute ("SELECT * FROM Client WHERE Client_DPA_ = " & selectedClient)
	
	If Not (rsClient.EOF Or rsClient.BOF) Then
		accountDesc = rsClient.Fields("ClientName").Value & "&nbsp;&nbsp;&nbsp;" & "[" & rsClient.Fields("Client_DPA_").Value & "]"
		accountAddress = rsClient.Fields("ClientAddr").Value
	End If
	
	Set rsClient = Nothing
	
	Rs.PageSize = 50

	Rs.CacheSize = Rs.PageSize
		intPageCount = Rs.PageCount 
		intRecordCount = Rs.RecordCount 
	
		first=0
		
	' Now you must double check to make sure that you are not before the start
	' or beyond end of the recordset.  If you are beyond the end, set 
	' the current page equal to the last page of the recordset.  If you are
	' before the start, set the current page equal to the start of the recordset.	

	Rs.CacheSize = Rs.PageSize

	if Not (rs.eof or rs.bof) then
	intPageCount = Rs.PageCount 
	intRecordCount = Rs.RecordCount 
	else
	intPageCount=1
	intRecordCount=0
	end if	
		
        PageNumber1=PageNumber1 + 1
        
        intPage=0
        
         intPageCount=Cint(intPageCount)
        
      t=0      
	
	dim Valuewrite
	dim runningBal
	totalDebits = 0
    totalCredits = 0
    runningBal=0
    
	if Not (rs.eof or rs.bof) then
		if (rs("IsOpeningBalance")<>1) then
		sqlStr = "SELECT SUM(ISNULL(dbo.StatementList.Credit - dbo.StatementList.Debit, 0)) + dbo.Client.ClientOpeningBal AS CurrentBal, dbo.StatementList.Client_DPA_" & _
					 " FROM  dbo.StatementList INNER JOIN dbo.Client ON dbo.StatementList.Client_DPA_ = dbo.Client.Client_DPA_" & _
					 " WHERE (dbo.Client.Deleted = 0) AND (dbo.StatementList.TransDate < '" & FormatDate(selectedFromDate) & "') GROUP BY dbo.StatementList.Client_DPA_, dbo.Client.ClientOpeningBal having dbo.StatementList.Client_DPA_=" & selectedClient 

		Set rsClient = Conn.Execute(sqlStr)
	
			If Not (rsClient.EOF Or rsClient.BOF) Then
				OpeningBalance=rsClient("CurrentBal")
			end if
		end if
	else
		sqlStr = "SELECT SUM(ISNULL(dbo.StatementList.Credit - dbo.StatementList.Debit, 0)) + dbo.Client.ClientOpeningBal AS CurrentBal, dbo.StatementList.Client_DPA_" & _
					 " FROM  dbo.StatementList INNER JOIN dbo.Client ON dbo.StatementList.Client_DPA_ = dbo.Client.Client_DPA_" & _
					 " WHERE (dbo.Client.Deleted = 0) AND (dbo.StatementList.TransDate < '" & FormatDate(selectedFromDate) & "') GROUP BY dbo.StatementList.Client_DPA_, dbo.Client.ClientOpeningBal having dbo.StatementList.Client_DPA_=" & selectedClient 
	

	
	Set rsClient = Conn.Execute(sqlStr)
	
		If Not (rsClient.EOF Or rsClient.BOF) Then
			OpeningBalance=rsClient("CurrentBal")
		end if
	end if	
	
	kwanza=0

	Do while Cint(intPage) < intPageCount	
	intPage=intPage + 1
	'Response.write(intpage)

	if(Cint(first)=1) then
	%>
             <BR class="newpage">
    <%
	end if

	first = 1

	If CInt(intPage) > CInt(intPageCount) Then intPage = intPageCount
	If CInt(intPage) <= 0 Then intPage = 1
	
	 'Make sure that the recordset is not empty.  If it is not, then set the 
	 'AbsolutePage property and populate the intStart and the intFinish variables.
	
	'if Not(Rs.eof and Rs.bof) Then

	If intRecordCount > 0 Then 'and Not(Rs.eof and Rs.bof)
		
		Rs.AbsolutePage = intPage
		intStart = Rs.AbsolutePosition
		'Response.write(intStart)
		
		If CInt(intPage) = CInt(intPageCount) Then
			intFinish = intRecordCount
		Else
			intFinish = intStart + (Rs.PageSize - 1)
		End if
	End If	  
	
	if Not (rs.eof or rs.bof) then
	isOpeningBalance = CBool(Rs.Fields("IsOpeningBalance").Value)
	else
	isOpeningBalance=false
	end if
	
	

	If Not IsOpeningBalance Then		
		OpeningBalance = OpeningBalance		
			Else
		OpeningBalance =  Rs.Fields("Balance").Value		
		
	End If	

%>	

<table border="0" cellspacing="2" cellpadding="2" style="font-family: Arial Narrow" width="650">
    <tr class="pageNumbering">
    <td width="10%" nowrap><font face="Impact" size="2">&nbsp;</font></td>
		<td align="right" >
			<FONT FACE=ARIAL SIZE=2><B>Page <%=intPage%>/<%=intPageCount%></B></FONT>	
		</td>		
	</tr>
    <tr>    	
		<td width="10%" nowrap><font face="Impact" size="4">CLIENTS STATEMENT</font></td>
      <td width="40%" nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>
      
    </tr>

  </table>
<br>
<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="650">
    <tr>      
      <td width="1%"><b>Date:</b></td>
      <td width="48%"><%= FormatDate(Date) %></td>
    </tr>

    <tr>    
      <td width="1%"><b>Account:</b></td>
      <td width="48%"><%= accountDesc %></td>
    </tr>

    <tr>      
      <td width="1%"><b><font size="2" face="Arial">&nbsp;</font></b></td>
      <td width="48%"><%= accountAddress %></td>
    </tr>

</table>

  <table border="0" cellspacing="1" cellpadding="0" style="font-family: Arial Narrow; LEFT-MARGIN:100PX"  width="650">
    <tr>
      <td colspan="6" align="right" style="border-bottom-style: solid; border-bottom-width: 1">
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>
    </tr>  
    <tr>      
      <td width="80"><b><font face="Arial Narrow" size="3">Date:</font></b></td>
      <td><b><font face="Arial Narrow" size="3">Ref:</font></b></td>
      <td><b><font face="Arial Narrow" size="3">Particulars:</font></b></td>
      <td align="right"><b><font face="Arial Narrow" size="3">Debit:</font></b></td>
      <td align="right"><b><font face="Arial Narrow" size="3">Credit:</font></b></td>
      <td align="right"><b><font face="Arial Narrow" size="3">Balance:&nbsp;&nbsp;&nbsp;</font></b></td>
    </tr> 
    <tr>
      <td colspan="6" align="right" style="border-bottom-style: solid; border-bottom-width: 1">
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>

    </tr>    
    <%    
    'Do Until Rs.EOF       
 if(Cint(kwanza=0)) then   
	If Not IsOpeningBalance Then

	if(abs(FormatNum(OpeningBalance))=0) then
	OpeningBalance=0
	end if
	%>
	<tr>      
      <td width="80"><b><font size="1"><%= Day(selectedFromDate) & " " & MonthName(Month(selectedFromDate), True) & " " & Right(Year(selectedFromDate),2) %></font></b></td>
      <td><b><font size="1">&nbsp;</font></b></td>
      <td><b><font size="1">Opening Balance</font></b></td>
      <td align="right"><b><font size="1">&nbsp;</font></b></td>
      <td align="right"><b><font size="1">&nbsp;</font></b></td>
      <td align="right"><b><font size="1"><%=CreditDebitValue(FormatNum(OpeningBalance))%>&nbsp;&nbsp;</font></b></td>
    </tr> 
	<%

	runningBal = runningBal + OpeningBalance
	end if
end if
kwanza=1
if Not (rs.eof or rs.bof) then
    For intRecord = 1 to Rs.PageSize
    
		totalDebits = totalDebits + Rs.Fields("Debit").Value 
		totalCredits = totalCredits + Rs.Fields("Credit").Value
		
		
		%>
		<tr>		
		  <td><font size="1"><%= Day(rs.Fields("TransDate")) & " " & MonthName(Month(rs.Fields("TransDate")), True) & " " & Right(Year(rs.Fields("TransDate")),2) %></font>&nbsp;</td>		  
		  <td><font size="1"><%= Rs.Fields("PaymentReceiptNo").Value %></font>&nbsp;</td>
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
		  <td align="right"><font size="1"><% If Rs.Fields("Debit").Value <> "0" Then 
									Response.Write FormatNum(Rs.Fields("Debit").Value) 
							   End If			%>
            </font>
		  &nbsp;</td>
		  <td align="right"><font size="1"><% If Rs.Fields("Credit").Value <> "0" Then
									Response.Write FormatNum(Rs.Fields("Credit").Value) 
							   End If			
							%>
            </font>
		  &nbsp;</td>
		  <td align="right">
            <font size="1">
			<% 					

			If Not IsOpeningBalance Then
					Valuewrite=(Rs.Fields("Credit").Value - Rs.Fields("Debit").Value)										
					Response.write(FormatNum(CreditDebitValue(runningBal+ValueWrite)))											
					runningBal = runningBal + (Rs.Fields("Credit").Value - Rs.Fields("Debit").Value)										
				Else											
					if(rs("IsOpeningBalance")=1) then						
						Response.Write FormatNum(CreditDebitValue(Rs.Fields("Balance").Value))						
						runningBal = rs("CreditBal")
					else
					    
						Valuewrite=(Cdbl(Rs.Fields("Credit").Value) - CDbl(Rs.Fields("Debit").Value))
						'Response.write("yeah")
						if(abs(FormatNum(runningBal+Cdbl(ValueWrite)))=0) then
						Response.write(CreditDebitValue(abs(FormatNum(runningBal+Cdbl(ValueWrite)))))
						else
						Response.write(CreditDebitValue(FormatNum(runningBal+Cdbl(ValueWrite))))
						end if
						runningBal = runningBal + Valuewrite
						
						'Response.Write FormatNum(Rs.Fields("Balance").Value)
						
						'Response.write(runningBal)
						  if(abs(FormatNum(runningBal))=0) then
							runningBal=abs(FormatNum(runningBal))
						  end if
						  'Response.write(runningBal)

					end if
				End If 
				lastDate = Rs.Fields("TransDate").Value%>
            </font>
		  &nbsp;</td>
		</tr>
	
	<%	'Rs.MoveNext	
	rs.MoveNext
        
		If Rs.EOF Then Exit for

        Next
  else
  lastDate=Date
  end if      
	if Cint(intPage) <> intPageCount then 	
  	%>
  		<tr>
      <td colspan="6" align="right" style="border-bottom-style: solid; border-bottom-width: 1">
        &nbsp;&nbsp;&nbsp; </td>

    </tr>
	<%
	if(abs(FormatNum(OpeningBalance))=0) then
	OpeningBalance=0
	end if
	
	%>
    <tr>
      <td colspan="5" align="right">
        <font size="1">
        Opening&nbsp;Balance:</font></td>

      <td align="right"><font size="1"><%= FormatNum(CreditDebitValueRev(OpeningBalance)) %></font>&nbsp;</td>
    </tr>

    <tr>
      <td colspan="5" align="right"><font size="1">less&nbsp;Running&nbsp;Debits:</font></td>

      <td align="right"><font size="1"><%= FormatNum(0 - totalDebits) %></font>&nbsp;</td>
    </tr>

    <tr>
      <td colspan="5" align="right"><font size="1">add&nbsp;Running&nbsp;Credits:</font></td>

      <td align="right"> <font size="1"> <%= FormatNum(totalCredits) %></font>&nbsp;</td>
    </tr>

    <tr>
      <td colspan="6" align="right" style="border-bottom-style: solid; border-bottom-width: 1">
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>

    </tr>
	<%
	if(abs(FormatNum(runningBal))=0) then
		runningBal=abs(FormatNum(runningBal))
	end if
	%>

    <tr>
      <td colspan="4">&nbsp;</td>
      <td align="right"><b><font size="1">Running&nbsp;Balance:</font></b></td>
      <td align="right"><b><font size="1"><%= CreditDebitValue(FormatNum(runningBal)) %></font></b>&nbsp;</td>
    </tr>
   	<%
   	end if
   	%>	
	</table>
	<%	
	loop
	%>
	<table border="0" cellspacing="1" cellpadding="0" style="font-family: Arial Narrow; LEFT-MARGIN:100PX"  width="650">     
    <tr>		  
		  <td colspan="3"><font size="1"><%= Day(lastDate) & " " & MonthName(Month(lastDate), True) & " " & Right(Year(lastDate),2) %>		  
		  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;CLOSING BALANCE</font></td>
		  <td align="right" colspan="5"><font size="1"><%= CreditDebitValue(FormatNum(CDbl(runningBal))) %></font>&nbsp;&nbsp;</td>
		</tr>
	<tr>
      <td colspan="7" align="right" style="border-bottom-style: solid; border-bottom-width: 1">
        &nbsp;&nbsp;&nbsp; </td>

    </tr>

    <tr>
      <td colspan="6" align="right">
        <font size="1">
        Opening Balance:</font></td>

      <td align="right"><font size="1"><%= FormatNum(CreditDebitValueRev(OpeningBalance)) %></font>&nbsp;</td>
    </tr>

    <tr>
      <td colspan="6" align="right"><font size="1">less Total Debits:</font></td>

      <td align="right"><font size="1"><%= FormatNum(0 - totalDebits) %></font>&nbsp;</td>
    </tr>

    <tr>
      <td colspan="6" align="right"><font size="1">add Total Credits:</font></td>

      <td align="right"> <font size="1"> <%= FormatNum(totalCredits) %></font>&nbsp;</td>
    </tr>

    <tr>
      <td colspan="7" align="right" style="border-bottom-style: solid; border-bottom-width: 1">
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>

    </tr>
    <%
	Dim BeforeBal
	Dim LaterBal
		
	BeforeBal =	ClientBalance(30)
	
	if(abs(FormatNum(BeforeBal))=0) then
	BeforeBal=0
	end if

	LaterBal = ClientBalance(60)

	if(abs(FormatNum(LaterBal))=0) then
	LaterBal=0
	end if

	%>

    <tr>
      <td colspan="5"><font size="1">Current: <%= CreditDebitValue(FormatNum(runningBal)) %>
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;30-60 Days:&nbsp;<%=CreditDebitValue(FormatNum(BeforeBal))%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Over 60 Days:&nbsp;<%=CreditDebitValue(FormatNum(LaterBal))%></font> </td>
      <td align="right"><b><font size="1">Total Balance:</font></b></td>
      <td align="right"><b><font size="1"><%= CreditDebitValue(FormatNum(runningBal)) %></font></b>&nbsp;</td>
    </tr>
      
  </table>
   
<%
'Set Rs = Nothing

function ClientBalance(Days)
sqlStr="SELECT SUM(ISNULL(dbo.StatementList.Credit - dbo.StatementList.Debit, 0)) + dbo.Client.ClientOpeningBal AS CurrentBal," & _ 
       " dbo.StatementList.Client_DPA_ FROM  dbo.StatementList INNER JOIN dbo.Client ON dbo.StatementList.Client_DPA_ = dbo.Client.Client_DPA_" & _
	   " WHERE (dbo.StatementList.TransDate < GETDATE() - " & Days & ") and (StatementList.Client_DPA_=" & selectedClient & ")" & _
	   " GROUP BY dbo.StatementList.Client_DPA_, dbo.Client.ClientOpeningBal"

Set Rs = Conn.Execute(sqlStr)
 if not(Rs.eof and Rs.Bof) then
 ClientBalance=Rs("CurrentBal")
 else
 ClientBalance=0
 end if
end function 

Set Rs = Nothing
Set Conn = Nothing
%>   
</body>

</html>