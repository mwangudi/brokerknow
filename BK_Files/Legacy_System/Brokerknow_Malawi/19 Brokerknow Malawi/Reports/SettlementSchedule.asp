<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Debtors and Creditors</title>
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

genReport = Request.Form("genReport")
FirstDay=DateSerial(Year(Date), Month(Date)-1 + iOffset, 1)

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
	<form method="POST" action="SettlementSchedule.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">
		<table>
		
			<tr>
				<td>Select filter date:</td>
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

<% DrawPageFunctions True, True, True, True %>

<%
   Dim conn 
   Dim sqlStr
   Dim rs
   Dim TimeLimitRs
   Dim NoOfDays
   Dim SettlementDate
   
   Set TimeLimitRs = CreateObject("ADODB.Recordset")						        
   TimeLimitRs.CursorLocation = adUseClient	

   Set conn = GetActiveConnection("KBroker")
    	
   sqlStr="SELECT TimeLimitLimDaysNSE From TimeLimit where TimeLimit_DPA_=1"
   set TimeLimitRs=Conn.Execute(sqlStr) 
   
	if not(TimeLimitRS.eof and TimeLimitRs.bof) then
	 NoOfDays=TimeLimitRS("TimeLimitLimDaysNSE")
	end if
		
	SettlementDate=LTdate(CDate(selectedFromDate),5)	
		
	    sqlStr = "SELECT * FROM SettlementSchedule  WHERE (Day(LotTDate) = Day(" & "#" & SelectedFromDate & "#" & ")) " & _
			 " and (Month(LotTDate) = Month(" & "#" & SelectedFromDate & "#" & ")) and " & _
                   " (Year(LotTDate) = Year(" & "#" & SelectedFromDate & "#" & "))"
		
		'sqlStr="SELECT * FROM  ReceiptSchedule WHERE (DAY(Payment_Date) = DAY('01-Mar-2005')) OR (DAY(Payment_Date) = DAY('9-Mar-2005'))"    
     	
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
        
%>
<p id="toPDFOrient" name="toPDFOrient" value="P" style="display:none;">P
<p id="toPDF" name="toPDF">
<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
    <tr>
      <td width="10%" nowrap><b><font face="Arial" size="4"><b>SETTLEMENT SCHEDULE</b></font></b></td>
      <td nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>
      
    </tr>

    <tr>
      <td nowrap colspan=2><font size="2" face="Arial">For Contracts Traded On: <%= FormatDate(selectedFromDate) %>&nbsp;&nbsp;(Settlement Date:&nbsp;<%=FormatDate(SettlementDate)%>)</font></td>
    </tr>

  </table>
<br>
 <table border="0" cellspacing="0" cellpadding="3" style="font-family: Arial Narrow; LEFT-MARGIN:100PX" width="800">
    <tr>
      <td colspan="9" width="948"><b><font face="Arial Narrow" size="3">RECEIPTS:</font></b></td>      
    </tr>    
    <tr>
	  <td bgcolor="#000000" width="9"><b><font color="#FFFFFF" face="Arial Narrow" size="3">&nbsp;</font></b></td>	
      <td bgcolor="#000000" width="63"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Ref:</font></b></td>      
      <td align="left" bgcolor="#000000" width="66"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Contract</font></b></td>
      <td align="left" bgcolor="#000000" width="50"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Code:</font></b></td>
      <td align="left" bgcolor="#000000" width="300"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Client:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></b></td>
      <td align="left" bgcolor="#000000" width="41"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Sec:</font></b></td>      
      <td align="right" bgcolor="#000000" width="30"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Price</font></b></td>
      <td align="right" bgcolor="#000000" width="60"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Quantity</font></b></td>      
	  <td align="right" bgcolor="#000000" width="60"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Amount</font></b></td>      
    </tr>
    <%
    Cds1=0
    Cds2=0
    first=0
    
    SubAmount=0   
    totalBal = 0
    Do Until Rs.EOF
    Cds1=Rs.Fields("OrderType_DPA_")				
		
	if (Cds1<>Cds2) and (first<>0) then %>
		<tr>
        <td colspan="9" width="814">
        &nbsp;</td>
        </tr> 
		<tr>
      	<td colspan="4" width="144"><b>Total </b></td>
      	<td colspan="5" align="right" width="732"><b><%= FormatNum(SubAmount) %></b>&nbsp;</td>     
      	</td>
    	</tr>
    	<tr>
        <td colspan="9" width="948">
        &nbsp;</td>
        </tr>         
		<tr>
		<td colspan="9" width="948"><b><font face="Arial Narrow" size="3">PAYMENTS:</font></b></td>      
		</tr>
		<tr>
	  <td bgcolor="#000000" width="9"><b><font color="#FFFFFF" face="Arial Narrow" size="3">&nbsp;</font></b></td>	
      <td bgcolor="#000000" width="63"><b><font color="#FFFFFF" face="Arial Narrow" size="3">REF Ref:</font></b></td>      
      <td align="left" bgcolor="#000000" width="60"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Contract</font></b></td>
      <td align="left" bgcolor="#000000" width="50"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Code:</font></b></td>
      <td align="left" bgcolor="#000000" width="300"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Client:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></b></td>
      <td align="left" bgcolor="#000000" width="41"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Sec:</font></b></td>      
      <td align="right" bgcolor="#000000" width="30"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Price</font></b></td>
      <td align="right" bgcolor="#000000" width="60"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Quantity</font></b></td>      
	  <td align="right" bgcolor="#000000" width="60"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Amount</font></b></td>      
    </tr>
    
		<%		
		totalBal = SubAmount
		SubAmount=0 		
		end if
	SubAmount=SubAmount + Rs.Fields("LotGrossAmount").Value 
		first=1
		
		ClientName=Mid(rs("ClientName"),1,30)		 	
		'CounterParty=Mid(rs("OwnerName"),1,30)
			If Rs.Fields("LotGrossAmount").Value <> 0 Then
			
			%>

			<tr>
			  <td align="left" width="9">&nbsp;</td>							  	
			  <td align="left" width="63"><font  face="Arial Narrow" size="3"><%= rs("LotSlipNo") %></font></td>							  
			  <td align="left" width="60"><font face="Arial Narrow" size="3"><%= rs("ContractNumber") %></font></td>
			  <td align="left" width="50"><font face="Arial Narrow" size="3"><%= rs("Client_DPA_") %></font></td>
			  <td align="left" width="300"><font face="Arial Narrow" size="3"><%= ClientName %></font></td>			  
			  <td align="left" width="41"><font  face="Arial Narrow" size="3"><%= rs("SecurityCode")%></font></td>
			  <td align="right"><font  face="Arial Narrow" size="3"><%= FormatNum(rs("LotPrice")) %></font></td>
			  <td align="right" width="60"><font  face="Arial Narrow" size="3"><%= FormatNumEx(rs("LotQty"),0) %></font></td>
			  <td align="right"><font  face="Arial Narrow" size="3"><%= FormatNum(rs("LotGrossAmount")) %></font></td>
			</tr>
    
    <%		End If
    	Cds2=Cds1
		Rs.MoveNext
    Loop%>
	
    <tr>
      <td colspan="8" width="882">
        &nbsp;</td>
    </tr> 
       
	<tr>
      	<td colspan="4" width="144"><b>Total </b></td>
      	<td colspan="5" align="right" width="732"><b><%= FormatNum(SubAmount) %></b>&nbsp;</td>     
      	</td>
    	</tr>
    	<tr><td colspan="9" width="882"></td>
    </tr>
	
	<tr>
      <td colspan="9" width="882">
        &nbsp;</td>
    </tr> 

    <tr>
      <td colspan="4" width="144"><b>Net Amount: </b></td>
      <td colspan="5" align="right" width="732"><b><%= FormatNum(totalBal-SubAmount) %></b>&nbsp;</td>     
      </td>
    </tr>

  </table>

</body>

</html>