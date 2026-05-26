<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Contract NSE Compensation Funds</title>
  
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
selectedTradeDate = Request.Form("txtDate")

If genReport <> "1" Or Not IsDate(selectedTradeDate) Then%>
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
	<form method="POST" action="ContractNSEByDate.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">	
		<table>
			<tr>
				<td>Select any day of month</td>
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



	headerDescription = Year(selectedTradeDate) & " " & MonthName(Month(selectedTradeDate))


	DrawPageFunctions True, True, True %>
	

  
<%
	Dim conn 
   Dim sqlStr
   Dim rs
	
		 sqlStr = "SELECT * FROM [ContractCompensationByDate] WHERE DATEPART(month, TransDate) = " & Month(selectedTradeDate) & " AND DATEPART(yy, TransDate) = " & Year(selectedTradeDate)  
		 
		 Set conn = GetActiveConnection("KBroker")
        
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then
                %>
                <Script Language="JavaScript">
					alert("No contract stamps available");
					window.parent.history.go(-1);			
                </Script>
                <% Set Rs = Nothing
                Set Conn = Nothing
                Response.End
        End If
        
        rs.MoveFirst
        
%>


<table border="0" cellspacing="0" cellpadding="4" style="font-family: Arial Narrow; LEFT-MARGIN:100PX">
		
     <tr>
		     <td colspan=2>
		        <b><font face="Arial Narrow" size="4">Contract NSE Compensation Funds</font></b></td>
		     <td nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>   
	</tr>	
     <tr>
		     <td colspan=3>
		        <font face="Arial" size="2">for Month of:  <%= headerDescription %></font></td>
		</tr>   
   
<tr>
	<td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial" size="2">Traded</font></b></td>		
	<td bgcolor="#000000" align=right><b><font color="#FFFFFF" face="Arial" size="2">Gross</font></b></td>
	<td bgcolor="#000000" align=right><b><font color="#FFFFFF" face="Arial" size="2">MSE</font></b></td>	
  </tr>
<%		
	totalLevyAmount = 0
	totalgross=0
	Do Until rs.EOF
			totalLevyAmount = totalLevyAmount + FormatNum(rs.Fields("Commission")/2)
			totalgross = totalgross + rs.Fields("Gross")
 %>
                <tr>
                        <td><%=FormatDate(rs.Fields("TransDate"))%></td>
                        <td align=right><%=FormatNum(rs.Fields("Gross"))%></td>
                        <td align=right><%= FormatNum(rs.Fields("Commission")/2) %></td>                        
                </tr>
                <%                
                rs.MoveNext
        Loop
        conn.Close
        Set conn = Nothing%>
        
        <tr>
				<td colspan=3>&nbsp;</td>
         </tr>        
         <tr>
						<td colspan=1 align=right><font face="Arial" size="2"><b>Monthly Totals:</b></font></td>
				<td align=right style="border-style: solid; border-color: #000000; border-width: 1" height="30px"><%= FormatNum(totalgross) %></td>
                        <td align=right style="border-style: solid; border-color: #000000; border-width: 1" height="30px"><%= FormatNum(totalLevyAmount) %></td>
                        <td>&nbsp;</td>
         </tr>        
</table>

</body>

</html>
