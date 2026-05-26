<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Payment Schedule</title>
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
	<form method="POST" action="PaymentSchedule.asp" Name="frmMain" id="frmMain">
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

<% DrawPageFunctions True, True, True %>

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
		
	sqlStr = "SELECT * FROM PaymentSchedule  WHERE (Day(Payment_Date) = "& Day(CDate(SelectedFromDate))  & ") AND     (Month(Payment_Date) = " &  Month(CDate(SelectedFromDate)) & ") AND (Year(Payment_Date) = " & Year(CDate(SelectedFromDate)) & ")and CdsGroup<>1"
     	
		
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



<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
    <tr>
      <td width="10%" nowrap><b><font face="Arial" size="4"><b>PAYMENT SCHEDULE</b></font></b></td>
      <td nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>
      
    </tr>

    <tr>
      <td nowrap colspan=2><font size="2" face="Arial">as at: <%= FormatDate(selectedFromDate) %></font></td>
    </tr>

  </table>
<br>
 <table border="0" cellspacing="0" cellpadding="3" style="font-family: Arial Narrow; LEFT-MARGIN:100PX">
    <tr>
      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Entity&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></b></td>
      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Recipient&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></b></td>
      <td align="left" bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Account</font></b></td>
      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Receipt</font></b></td>
      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Reference</font></b></td>
      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Payment&nbsp;Date</font></b></td>
      <td align="right" bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow" size="3">Amount</font></b></td>
    </tr>
    
    <%
    Cds1=0
    Cds2=0
    first=0
    
    SubAmount=0   
    totalBal = 0
    Do Until Rs.EOF
    Cds1=Rs.Fields("CdsGroup")
		totalBal = totalBal + Rs.Fields("Amount").Value 		
		SubAmount=SubAmount + Rs.Fields("Amount").Value 
		
	if (Cds1<>Cds2) and (first<>0) then %>
		<tr>
        <td colspan="7">
        &nbsp;</td>
        </tr> 
		<tr>
      	<td colspan="2"><b>Sub Total </b></td>
      	<td colspan="5" align="right"><b><%= FormatNum(SubAmount) %></b>&nbsp;</td>     
      	</td>
    	</tr>
    	<tr>
        <td colspan="7">
        &nbsp;</td>
        </tr> 

		<%		
		SubAmount=0 		
		end if
		first=1		 	
			If Rs.Fields("Amount").Value <> 0 Then%>

			<tr>
			  <td align="left"><%= rs("Entity") %>&nbsp;</td>				
			  <td><%= rs("Party") %>&nbsp;</td>			  
			  <td align="left"><%= rs("AccountCode") %>&nbsp;</td>
			  <td align="left"><%= rs("Receipt_No") %>&nbsp;</td>
			  <td align="left"><%= rs("Reference") %>&nbsp;</td>
			  <td><%= rs("Payment_Date") %>&nbsp;</td>
			  <td align="right"><%= FormatNum(rs("Amount")) %>&nbsp;</td>	  

			</tr>
    
    <%		End If
    	Cds2=Cds1
		Rs.MoveNext
    Loop%>
	<!--
    <tr>
      <td colspan="7">
        &nbsp;</td>
    </tr> 
       
	<tr>
      	<td colspan="2"><b>Sub Total </b></td>
      	<td colspan="5" align="right"><b><%'= FormatNum(SubAmount) %></b>&nbsp;</td>     
      	</td>
    	</tr>
    	<tr><td colspan="7"></td>
    </tr>
	-->
	<tr>
      <td colspan="7">
        &nbsp;</td>
    </tr> 

    <tr>
      <td colspan="2"><b>Total Amount: </b></td>
      <td colspan="5" align="right"><b><%= FormatNum(totalBal) %></b>&nbsp;</td>     
      </td>
    </tr>

  </table>

</body>

</html>