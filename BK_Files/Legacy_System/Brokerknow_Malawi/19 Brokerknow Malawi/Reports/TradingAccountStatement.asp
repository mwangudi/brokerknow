<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Profit and Loss Account</title>

	<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
	<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
	<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
	<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
	<SCRIPT language=VBScript src="../scripts/reports.vbs"></SCRIPT>
	<SCRIPT language=Javascript src="../scripts/reports.js"></SCRIPT>
	<SCRIPT language=Javascript src="../scripts/accountList.js"></SCRIPT>



	<style media="print">
		@page {
			@top{font-family: Helvetica, Arial, sans-serif;
				font-size: 150%;
				font-weight: bolder;
				text-align: left;
				content: "<%= FormatDate(Date) %>";			
			}
			
			margin-left: 0cm;
			margin-right: 0cm;
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
'selectedBank = Request.Form("cboAccount")
selectedFromDate = Request.Form("transFromDate")
selectedToDate = Request.Form("txtToDate")
SelectedType=Request.Form("cboEntity")
FirstDay=DateSerial(Year(Date), Month(Date)-1 + iOffset, 1)
thistype=Request.Form("Selectedtype")

If genReport <> "1" Or Not IsDate(selectedFromDate) Then%>
	<Script Language="JavaScript">
		report_SetBodyClass();
		
		function validateForm(frm)
		{						
			frm.target = '_self';			
			frm.submit();
		}
		
		
		function evaluateEntity(Val, Entity)
		{      	
	  	FetchAccounts1(Entity)
		}
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "transFromDate","cmdDate","<%= FormatDate(FirstDay) %>",1);
		var cal1=new ctlSpiffyCalendarBox("cal1", "frmMain", "txtToDate","cmdDate","<%= FormatDate(Date) %>",1);

	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="TradingAccountStatement.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">
		<% currentEntityType=5 %>
		<table>
			<tr>
				<td colspan="2">Select date from:</td>
				<td>
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>				
			</tr>
			<tr>
				<td colspan="2">To date:</td>
				<td>
					<SCRIPT language="JavaScript">cal1.writeControl();</SCRIPT>	
				</td>
				
			</tr>			

			<tr>
				<td colspan="3"><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... ">&nbsp;&nbsp;</td>
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
	
	Dim Subtotal
	Dim total
	
	Subtotal=0
	total=0
	sqlStr="SELECT Distinct TOP 100 PERCENT SettlementDate AS TransDate,'Equity trades: '+ Convert(char(12),Min(LotTDate)) as Particulars,SUM(CASE (OrderType_DPA_) " & _
           " WHEN 1 THEN LotGrossAmount ELSE 0 END) - SUM(CASE (OrderType_DPA_) WHEN 2 THEN LotGrossAmount ELSE 0 END) AS DepositeAmount FROM SettlementSchedule WHERE OrderSecType_DPA_ = 2 AND SettlementDate between '" & FormatDate(CDate(selectedFromDate)) & "' and '" & FormatDate(CDate(selectedToDate)) & "' GROUP BY SettlementDate"
    
    sqlStr=sqlStr & " Union SELECT Distinct TOP 100 PERCENT SettlementDate AS TransDate,'CDS Settlement: ' + Convert(Char(12),Min(LotTDate)) as Particulars,SUM(CASE (OrderType_DPA_) " & _
                    " WHEN 2 THEN LotGrossAmount ELSE 0 END) - SUM(CASE (OrderType_DPA_) WHEN 1 THEN LotGrossAmount ELSE 0 END) AS DepositeAmount" & _
                    " FROM SettlementSchedule WHERE OrderSecType_DPA_ = 2 AND IsCustodian = 0 AND SettlementDate between '" & FormatDate(CDate(selectedFromDate)) & "' and '" & FormatDate(CDate(selectedToDate)) & "' GROUP BY SettlementDate"       
	
	sqlStr=sqlStr & " Union SELECT Distinct TOP 100 PERCENT SettlementDate AS TransDate,'CDA Settlement: ' + Convert(Char(12),Min(LotTDate)) as Particulars,SUM(CASE (OrderType_DPA_) " & _
                    " WHEN 2 THEN LotGrossAmount ELSE 0 END) - SUM(CASE (OrderType_DPA_) WHEN 1 THEN LotGrossAmount ELSE 0 END) AS DepositeAmount" & _
                    " FROM SettlementSchedule WHERE OrderSecType_DPA_ = 2 AND IsCustodian = 1 AND SettlementDate between '" & FormatDate(CDate(selectedFromDate)) & "' and '" & FormatDate(CDate(selectedToDate)) & "' GROUP BY SettlementDate"       
	
	sqlStr=sqlStr & " Union SELECT Distinct TOP 100 PERCENT SettlementDate AS TransDate,'Bonds trades:' as Particulars,SUM(CASE (OrderType_DPA_) " & _
           " WHEN 1 THEN LotGrossAmount ELSE 0 END) - SUM(CASE (OrderType_DPA_) WHEN 2 THEN LotGrossAmount ELSE 0 END) AS DepositeAmount FROM SettlementSchedule WHERE OrderSecType_DPA_ <> 2 AND SettlementDate between '" & FormatDate(CDate(selectedFromDate)) & "' and '" & FormatDate(CDate(selectedToDate)) & "' GROUP BY SettlementDate"
    
    sqlStr=sqlStr & " Union SELECT Distinct TOP 100 PERCENT SettlementDate AS TransDate,'Bonds Settlement: ' as Particulars,SUM(CASE (OrderType_DPA_) " & _
                    " WHEN 2 THEN LotGrossAmount ELSE 0 END) - SUM(CASE (OrderType_DPA_) WHEN 1 THEN LotGrossAmount ELSE 0 END) AS DepositeAmount " & _
                    " FROM SettlementSchedule WHERE OrderSecType_DPA_ <> 2 AND SettlementDate between '" & FormatDate(CDate(selectedFromDate)) & "' and '" & FormatDate(CDate(selectedToDate)) & "' and Not (Contract_DPA_ is null)  GROUP BY SettlementDate"
	
	Set Rs = Conn.Execute(sqlStr) 
	
	'Confirm that records are returned
		 if rs.EOF or rs.BOf then
		  %>
                <script language = 'javascript'>
                		alert ("No Orders Found");
                		window.parent.history.go(-1);          		
                </script>
                
                <% Response.End   
		 end if
		 
	
	Rs.PageSize=40

	Rs.CacheSize = Rs.PageSize
		intPageCount = Rs.PageCount 
		intRecordCount = Rs.RecordCount 
	
		first=0
		
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
      m=0  
      t=0   
	totalDebits = 0
    	totalCredits = 0
    Balance=0
	Credits=0
	Debits=0
	   
	Do while Cint(intPage) < intPageCount	
	intPage=intPage + 1	
	'Response.write(sqlstr)
	
	if(Cint(first)=1) then
	%>
             <BR class="newpage">
    <%
	end if
	

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

	%>	
	
	<table border="0" cellspacing="2" cellpadding="2" style="font-family: Arial Narrow" width="500">
		<tr>
		<td width="10%" nowrap><font face="Impact" size="3">TRADING ACCOUNT STATEMENT</font></td>
		<td width="60%" nowrap align="right"><font face="Impact" size="3"><%= Session("CompanyName") %>
		    </font></td>
		</tr>
	</table>	
	<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="500">		
		<tr>
		<td width="1%"><b>Between: </b></td>
		<td width="48%"><b><%=FormatDate(selectedFromDate)%> &nbsp; And &nbsp;<%=FormatDate(selectedToDate)%></b>
		</td>
		</tr>
	</table>
	
	<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow; LEFT-MARGIN:100PX" width="500">		
		<tr>
		  <td width="80"><b>Date</b></td>
		  <td width="200"><b>Particulars</b></td>
		  <td align="right" width="60"><b>Debit</b></td>
		  <td align="right" width="60"><b>Credit</b></td>
		  <td align="right" width="100"><b>Balance</b></td>
		</tr>
		
	<%	
    For intRecord = 1 to Rs.PageSize
    if(rs("DepositeAmount")=0) then
    else 
       Balance=Balance + rs("DepositeAmount")
		%>
		<tr>
		  <td><%= FormatDate(Rs.Fields("TransDate").Value) %>
		  </td>
		  <td><%= Rs.Fields("Particulars").Value %>
		  </td>
		  <% if(rs("DepositeAmount")<0) then 
		  Debits=Debits+rs("DepositeAmount")
		  %>		  
		  <td align="right"><%=FormatNum(Abs(Rs.Fields("DepositeAmount").Value))%>
		  </td>
		  <td>&nbsp;</td>
		  <% else 
		  Credits=Credits + rs("DepositeAmount")
		  %>
		  <td>&nbsp;</td>
		  <td align="right"><%=FormatNum(Abs(Rs.Fields("DepositeAmount").Value))%>
		  <% end if%>
		  <td align="right"><%=CreditDebitValue(FormatNum(Balance))%>
		</tr>
		<%
	end if		
		rs.MoveNext
        
		If Rs.EOF Then Exit for

        Next
        
        %>	
	</table>
	<%	
	loop
	%>
	<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow; LEFT-MARGIN:100PX" width="500">		
	<tr><td colspan="5">&nbsp;</td></tr>
	<tr>		  
		  <td align="right" colspan="2" width="280">Balances</td>
		  <td align="right" width="60"><b><%=FormatNum(Debits)%></b></td>
		  <td align="right" width="60"><b><%=FormatNum(Credits)%></b></td>
		  <td align="right" width="100"><b><%=CreditDebitValue(FormatNum(Balance))%></b></td>
		</tr>
		
		</table>
	
<%	
Set Rs = Nothing
Set Conn = Nothing%>

</body>

</html>
