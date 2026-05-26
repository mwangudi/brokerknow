<html>

<head>
<title>Journal Entries Listing</title>

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
genReport = Request.Form("genReport")
selectedFromDate = Request("transFromDate")
selectedToDate = Request("transToDate") 

If genReport = "" Then%>
	<Script Language="JavaScript">
		report_SetBodyClass();
		function validateForm(frm){			
			frm.target = '_self';			
			frm.submit();
		}
		
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "transFromDate","cmdDate","<%= FormatDate(DateAdd("d", -30, Date)) %>",1);
		var cal2=new ctlSpiffyCalendarBox("cal2", "frmMain", "transToDate","cmdDate2","<%= FormatDate(Date) %>",1);
	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="JournalEntriesListing.asp" Name="frmMain" id="frmMain">		
		<table border=0>
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
			<input type="hidden" name="genReport" value="1">
			<tr>
				<td colspan=2><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.getElementById('frmMain'))" Value=" Generate... ">&nbsp;&nbsp; <input type="Button" class="Buttons" Value=" Close " OnClick="JavaScript: window.parent.self.close();"></td>
			</tr>
		</table>
	</form>
	
	<%Set rs = Nothing
	Set Conn = Nothing
	Response.End
End If

%>


<% DrawPageFunctions True, True, True


SqlStr = "SELECT TransDate, Entry, Narrative, Debit, Credit, Entity_DPA_ AS Code, AccountName, ControlAccount  FROM JournalEntriesList WHERE TransDate BETWEEN '" & FormatDate(selectedFromDate) & "' AND '" & FormatDate(DateAdd("d", 1, selectedToDate)) & "' " & _
" ORDER BY TransDate"

'Response.Write sqlstr
'Response.End 

Set Conn = GetActiveConnection("KBroker")

Set Rs = Conn.Execute(sqlStr)

 If Rs.EOF Or Rs.BOF Then%>
		<Script Language="JavaScript">	
			ShowMessage('No data was found using the criteria entered.');
			window.parent.history.go(-1);
		</Script>
		<%Set Conn = Nothing
		Set Rs = Nothing
		Response.End
  End If


 %>

<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
     <tr>
		<td nowrap><b><font face="Arial Narrow" size="3"><%= Ucase("Journals") %></font></b></td>
		<td nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>
	</tr>
	<tr>
      <td colspan=2>Between <%= FormatDate(selectedFromDate) %> And <%= FormatDate(selectedToDate)%></td>
    </tr>	
       <tr>
		  <td COLSPAN=2><font face="Arial" size="2">&nbsp;</font></td>
	</tr>
</table>			



    <table border="0" width="100%" cellPadding="2" cellSpacing=0>
    <tr bgColor="#000000">
			
	<%For i = 0 To Rs.Fields.Count - 1%>		
	 	 <td nowrap><b><font color="#FFFFFF"><%= Rs.Fields(i).Name %></font></b></td>
   <%Next %>
	</tr>
	
	<% Do Until rs.EOF%>
        		<tr>
      				
					<%
					For i = 0 To Rs.Fields.Count - 1
						'If IsDate(Rs.Fields(i).Value) Then
						If i = 0 Then
							If Rs.Fields(i).Value > 0 Then
								%><td nowrap><%= FormatDate(Rs.Fields(i).Value) %> </td><%
							Else
								%><td nowrap><%= Rs.Fields(i).Value %></td><%
							End If
						Else
							%><td nowrap><%= Rs.Fields(i).Value %></td><%
						End If
					Next
					%>			
      
	        </tr>
	 <%  Rs.MoveNext
	 Loop%>	  	
  </table>
  
	 <%	 
	 Set Rs = Nothing
	 Set Conn = Nothing
     %>	
</body>

</html>
