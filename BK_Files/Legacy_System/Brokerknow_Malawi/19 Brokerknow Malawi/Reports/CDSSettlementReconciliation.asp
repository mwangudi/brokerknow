<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Settlement Reconciliation (CDS) </title>
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
'FirstDay=DateSerial(Year(Date), Month(Date)-1 + iOffset, 1)

genReport = Request.Form("genReport")

selectedFromDate = Request.Form("transFromDate")

If genReport <> "1" Or Not IsDate(selectedFromDate) Then%>
	<Script Language="JavaScript">
		report_SetBodyClass();
		
		function validateForm(frm){			
			frm.target = '_self';			
			frm.submit();
		}
		
		
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "transFromDate","cmdDate","<%= FormatDate(Date) %>",1);
	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="CDSSettlementReconciliation.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">
		<table>
		
			<tr>
				<td>Select filter date:</td>
				<td>
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>
			</tr>
			<tr>
				<td colspan=2><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... " id=Button1 name=Button1>&nbsp;&nbsp;</td>
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
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim TimeLimitRs
	Dim NoOfDays
	Dim SettlementDate   
	Dim TotalAmount
		
	TotalAmount=0

	Set TimeLimitRs = CreateObject("ADODB.Recordset")						        
	TimeLimitRs.CursorLocation = adUseClient	

	Set conn = GetActiveConnection("KBroker")
	    	
	sqlStr="SELECT TimeLimitLimDaysNSE From TimeLimit where TimeLimit_DPA_=1"
	set TimeLimitRs=Conn.Execute(sqlStr) 
	   
	if not(TimeLimitRS.eof and TimeLimitRs.bof) then
		NoOfDays=TimeLimitRS("TimeLimitLimDaysNSE")
	end if
			
	SettlementDate=LTdate(CDate(selectedFromDate),5)	
		
	sqlStr = "SELECT * FROM CDASettlement WHERE (Day(LotTDate) = Day(" & "#" & SelectedFromDate & "#" & ")) " & _
		 " and (Month(LotTDate) = Month(" & "#" & SelectedFromDate & "#" & ")) and " & _
	           " (Year(LotTDate) = Year(" & "#" & SelectedFromDate & "#" & "))"
			
	sqlStr = "SELECT * FROM SettlementSchedule  WHERE (Day(SettlementDate) = Day(" & "#" & SelectedFromDate & "#" & ")) " & _
		 " and (Month(SettlementDate) = Month(" & "#" & SelectedFromDate & "#" & ")) and " & _
	           " (Year(SettlementDate) = Year(" & "#" & SelectedFromDate & "#" & ")) and (OrderSecType_DPA_ = 2) and IsCustodian=0 order by LotSlipNo"		
			
	'sqlStr="SELECT * FROM  ReceiptSchedule WHERE (DAY(Payment_Date) = DAY('01-Mar-2005')) OR (DAY(Payment_Date) = DAY('9-Mar-2005'))"    
	     	
	sqlStr = "SELECT LotTDate, MIN(LotSlipNo) AS LotSlipNo, MIN(ContractNumber) AS ContractNumber, ClientCDSNo, SecurityCode, ClientName, SUM(LotGrossAmount)  " & _
		" / SUM(LotQty) AS Price, SUM(LotQty) AS LotQty, SUM(LotGrossAmount) AS LotGrossAmount, MAX(OrderType_DPA_) AS OrderType_DPA_ " & _
		" FROM SettlementSchedule " & _
		" WHERE (Day(SettlementDate) = Day(" & "#" & SelectedFromDate & "#" & ")) " & _
		" and (Month(SettlementDate) = Month(" & "#" & SelectedFromDate & "#" & ")) and " & _
		" (Year(SettlementDate) = Year(" & "#" & SelectedFromDate & "#" & ")) and (OrderSecType_DPA_ = 2) and IsCustodian=0" & _
		" GROUP BY LotTDate, ClientCDSNo, SecurityCode, ClientName ORDER BY LotSlipNo"
	     	
	'Response.write(sqlStr)
	'Response.end	

	Set Rs = Conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
	If rs.EOF Or rs.BOF Then
	        %>
	        <Script Language="JavaScript">
				alert("The report did not find any values available");
				window.parent.history.go(-1);					
	        </Script>
	        <% Set Rs = Nothing
	        Set Conn = Nothing
	        Response.End
	End If
	        
	rs.MoveFirst

	Rs.PageSize=30

	SaleType=2
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
	SubAmount=0   
    totalBal = 0
    NetAmount=0
	Cds1=0
    Cds2=0
    Cda1=0
    Cda2=0
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

%>

<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
	<tr class="pageNumbering">
    <td width="10%" nowrap><font face="Impact" size="2">&nbsp;</font></td>
		<td align="right" >
			<FONT FACE=ARIAL SIZE=2><B>Page <%=intPage%>/<%=intPageCount%></B></FONT>	
		</td>		
	</tr>    
    <tr>
      <td width="10%" nowrap><b><font face="Arial" size="4"><b>SETTLEMENT RECONCILIATION (CDS)</b></font></b></td>
      <td nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>
      
    </tr>

    <tr>
      <td nowrap colspan=2><font size="2" face="Arial">Settlement Date:&nbsp;<%=FormatDate(SelectedFromDate)%></font></td>
    </tr>

  </table>
<br>
 <table border="0" cellspacing="0" cellpadding="3" style="font-family: Arial Narrow; LEFT-MARGIN:100PX" width="900">
 
	<tr>
		<td bgcolor="#000000" width="9"><b><font color="#FFFFFF" face="Arial Narrow" size="3">&nbsp;</font></b></td>
		<td bgcolor="#000000" width="90"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Trade&nbsp;Date:</font></b></td>                  
		<td bgcolor="#000000" width="63"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Cds Ref:</font></b></td>      
		<td align="left" bgcolor="#000000" width="80"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Contract&nbsp;No:</font></b></td>
		<td align="left" bgcolor="#000000" width="50"><b><font color="#FFFFFF" face="Arial Narrow" size="3">CDS&nbsp;NO:</font></b></td>
		<td align="left" bgcolor="#000000" width="41"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Sec:</font></b></td>    
		<td align="left" bgcolor="#000000" width="300"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Client:</font></b></td>       
		<td align="right" bgcolor="#000000" width="30"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Price</font></b></td>
		<td align="right" bgcolor="#000000" width="60"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Quantity</font></b></td>      
		<td align="right" bgcolor="#000000" width="60"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Amount</font></b></td>      
	</tr>
    <%	
    first=0   
    
    'Do Until Rs.EOF
	For intRecord = 1 to Rs.PageSize
	
	SaleType=rs("OrderType_DPA_")
	
	if(SaleType=2) then
	TotalAmount=TotalAmount + rs("LotGrossAmount")
	else
	TotalAmount=TotalAmount - rs("LotGrossAmount")
	end if

	'if (Cda1<>Cda2) and (first<>0) then 
	first=0
	

    		Cds1=Rs.Fields("OrderType_DPA_")						
		first=1
		
		ClientName=Mid(rs("ClientName"),1,30)		 	
		'CounterParty=Mid(rs("OwnerName"),1,30)
			If Rs.Fields("LotGrossAmount").Value <> 0 Then
			
			%>

			<tr>
			  <td align="left" width="9">&nbsp;</td>
			  <td align="left" width="90"><font  face="Arial Narrow" size="2"><%= FormatDate(rs("LotTDate")) %></font></td>							  			  			  
			  <td align="left" width="63"><font  face="Arial Narrow" size="2"><%= rs("LotSlipNo") %></font></td>							  
			  <td align="left" width="60"><font face="Arial Narrow" size="2"><%= rs("ContractNumber") %></font></td>
			  <td align="left" width="50"><font face="Arial Narrow" size="2"><%= rs("ClientCDSNo") %></font></td>
			  <td align="left" width="41"><font  face="Arial Narrow" size="2"><%= rs("SecurityCode")%></font></td>
			  <td align="left" width="300"><font face="Arial Narrow" size="2"><%= ClientName %></font></td>			  			  
			  <td align="right"><font  face="Arial Narrow" size="2"><%= FormatNum(rs("Price")) %></font></td>
			  <td align="right" width="60"><font  face="Arial Narrow" size="2"><%= FormatNumEx(rs("LotQty"),0) %></font></td>
			  <% if(SaleType=2) then%>
			  <td align="right"><font  face="Arial Narrow" size="2"><%= FormatNum(rs("LotGrossAmount")) %></font></td>
			  <% else%>
			  <td align="right"><font  face="Arial Narrow" size="2"><%= FormatNum(- rs("LotGrossAmount")) %></font></td>
			  <% end if%>
			</tr>
    
    <%		
		End If
    	
	rs.MoveNext
        
		If Rs.EOF Then Exit for

        Next
	loop
	%>
	<tr>
      <td align="right" COLSPAN=9><b>NET SETTLEMENT:</b></td>
      <td align="right"><b><font  face="Arial Narrow" size="2"><%= FormatNum(TotalAmount) %></font></b></td>     
      </td>
    </tr>  	
  </table>

</body>

</html>