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

selectedClient = Session("Client_DPA_")
SelectedFromDate=DateSerial(Year(Date), Month(Date)-3 + iOffset, 1)

if(SelectedFromDate<CDate(FormatDate("03/01/2005"))) then
	SelectedFromDate=CDate(FormatDate("03/01/2005"))
end if

SelectedToDate=Date


 DrawPageFunctions True, False, True 
 
	Set conn = GetActiveConnection("KBroker")
	Set Rs = CreateObject("ADODB.Recordset")	
	

	sqlStr = "SELECT * FROM StatementList WHERE Client_DPA_ = "  & selectedClient &_ 
    		 " and TransDate  between '" & CDate(selectedFromDate) & "'" & _
			 " and '" & CDate(selectedToDate) & "' order by TransDate"
	
	'sqlStr = "ClientStatement"
	Rs.CursorLocation = adUseClient	
	Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
	'Rs.Filter = "Client_DPA_ = '" & selectedClient & "' AND TransDate >= '" & FormatDate(selectedFromDate) & "'"
	
	
	If rs.EOF Or rs.BOF Then%>
		<Script Language="JavaScript">
			alert("The specified client does not have any transaction using the specified date criterion")
			window.history.go(-1);
		</Script>
		<%Set Rs = Nothing
		Set Conn = Nothing
		Response.End
	End If
	
	Set rsClient = Conn.Execute ("SELECT * FROM Client WHERE Client_DPA_ = " & selectedClient)
	
	If Not (rsClient.EOF Or rsClient.BOF) Then
		accountDesc = rsClient.Fields("ClientName").Value & "&nbsp;&nbsp;&nbsp;" & "[" & rsClient.Fields("Client_DPA_").Value & "]"
		accountAddress = rsClient.Fields("ClientAddr").Value
	End If
	
	Set rsClient = Nothing
	
	Rs.PageSize=50

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
		
        PageNumber1=PageNumber1 + 1
        
        intPage=0
        
         intPageCount=Cint(intPageCount)
        
      t=0      
	
	dim Valuewrite
	dim runningBal
	totalDebits = 0
    totalCredits = 0
    runningBal=0
    
	Do while Cint(intPage) < intPageCount	
	intPage=intPage + 1
	'Response.write(intpage)
	
	if(Cint(first)=1) then
	%>
             <BR class="newpage">
    <%
	end if

	first=1

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

	isOpeningBalance = CBool(Rs.Fields("IsOpeningBalance").Value)
	
	

	If Not IsOpeningBalance Then		
		OpeningBalance = OpeningBalance		
			Else
		OpeningBalance =  Rs.Fields("Balance").Value		
		
	End If

%>	

<table border="0" cellspacing="2" cellpadding="2" style="font-family: Arial Narrow" width="100%">
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
<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
    <tr>      
      <td width="1%"><b>Date:</b></td>
      <td width="48%"><b>From</b>&nbsp;<%= FormatDate(SelectedFromDate)%>&nbsp;<b>To</b>&nbsp;<%= FormatDate(SelectedToDate)%></td>
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

  <table border="0" cellspacing="1" cellpadding="0" style="font-family: Arial Narrow; LEFT-MARGIN:100PX"  width="670">
    <tr>
      <td colspan="6" align="right" style="border-bottom-style: solid; border-bottom-width: 1">
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>
    </tr>  
    <tr>      
      <td><b><font face="Arial Narrow" size="3">Date:</font></b></td>
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
    
    For intRecord = 1 to Rs.PageSize
    
		totalDebits = totalDebits + Rs.Fields("Debit").Value 
		totalCredits = totalCredits + Rs.Fields("Credit").Value%>
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
					'Response.Write  FormatNum(CreditDebitValue(runningBal))					
				'Response.write(runningBal)
				'Response.end							
				Else											
					if(rs("IsOpeningBalance")=1) then						
						Response.Write FormatNum(CreditDebitValue(Rs.Fields("Balance").Value))						
						runningBal = rs("CreditBal")
					else
						Valuewrite=(Rs.Fields("Credit").Value - Rs.Fields("Debit").Value)					
						Response.write(FormatNum(CreditDebitValue(runningBal+ValueWrite)))											
						runningBal = runningBal + Valuewrite
						'Response.Write FormatNum(Rs.Fields("Balance").Value)										
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
        
	if Cint(intPage) <> intPageCount then 	
  	%>
  		<tr>
      <td colspan="6" align="right" style="border-bottom-style: solid; border-bottom-width: 1">
        &nbsp;&nbsp;&nbsp; </td>

    </tr>

    <tr>
      <td colspan="5" align="right">
        <font size="1">
        Opening Balance:</font></td>

      <td align="right"><font size="1"><%= FormatNum(CreditDebitValueRev(OpeningBalance)) %></font>&nbsp;</td>
    </tr>

    <tr>
      <td colspan="5" align="right"><font size="1">less Running Debits:</font></td>

      <td align="right"><font size="1"><%= FormatNum(0 - totalDebits) %></font>&nbsp;</td>
    </tr>

    <tr>
      <td colspan="5" align="right"><font size="1">add Running Credits:</font></td>

      <td align="right"> <font size="1"> <%= FormatNum(totalCredits) %></font>&nbsp;</td>
    </tr>

    <tr>
      <td colspan="6" align="right" style="border-bottom-style: solid; border-bottom-width: 1">
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>

    </tr>

    <tr>
      <td colspan="4">&nbsp;</td>
      <td align="right"><b><font size="1">Running Balance:</font></b></td>
      <td align="right"><b><font size="1"><%= CreditDebitValue(FormatNum(runningBal)) %></font></b>&nbsp;</td>
    </tr>
   	<%
   	end if
   	%>	
	</table>
	<%	
	loop
	%>
	<table border="0" cellspacing="1" cellpadding="0" style="font-family: Arial Narrow; LEFT-MARGIN:100PX"  width="670">     
    <tr>		  
		  <td colspan="3"><font size="1"><%= Day(lastDate) & " " & MonthName(Month(lastDate), True) & " " & Right(Year(lastDate),2) %>		  
		  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;CLOSING BALANCE</font></td>
		  <td align="right" colspan="5"><font size="1"><%= FormatNum(CreditDebitValue(runningBal)) %></font>&nbsp;</td>
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

    <tr>
      <td colspan="5"><font size="1">Current: <%= CreditDebitValue(FormatNum(runningBal)) %>
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;30-60 Days:&nbsp;<%=FormatNum(CreditDebitValue(ClientBalance(30)))%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Over 60 Days:&nbsp;<%=FormatNum(CreditDebitValue(ClientBalance(60)))%></font> </td>
      <td align="right"><b><font size="1">Total Balance:</font></b></td>
      <td align="right"><b><font size="1"><%= CreditDebitValue(FormatNum(runningBal)) %></font></b>&nbsp;</td>
    </tr>
      
  </table>
   
<%
'Set Rs = Nothing

function ClientBalance(Days)
sqlStr="SELECT SUM(ISNULL(dbo.ClientStatement.Credit - dbo.ClientStatement.Debit, 0)) + dbo.Client.ClientOpeningBal AS CurrentBal," & _ 
       " dbo.ClientStatement.Client_DPA_ FROM  dbo.ClientStatement INNER JOIN dbo.Client ON dbo.ClientStatement.Client_DPA_ = dbo.Client.Client_DPA_" & _
	   " WHERE (dbo.ClientStatement.TransDate < GETDATE() - " & Days & ") and (ClientStatement.Client_DPA_=" & selectedClient & ")" & _
	   " GROUP BY dbo.ClientStatement.Client_DPA_, dbo.Client.ClientOpeningBal"

'Response.write(sqlStr)
'Response.end

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