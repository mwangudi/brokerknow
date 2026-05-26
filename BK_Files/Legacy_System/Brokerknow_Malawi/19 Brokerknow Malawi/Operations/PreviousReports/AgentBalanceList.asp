<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Agent Balance List</title>

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
selectedBank = Request.Form("cboAccount")
selectedFromDate = Request.Form("transFromDate")
selectedToDate = Request.Form("txtToDate")
SelectedType=Request.Form("cboEntity")

FirstDay=DateSerial(Year(Date), Month(Date)-1 + iOffset, 1)
LastDay=DateSerial(Year(Date), Month(Date) + iOffset, 1)-1

thistype=Request.Form("Selectedtype")

If genReport <> "1" Or Not IsDate(selectedFromDate) Then%>
	<Script Language="JavaScript">
		report_SetBodyClass();
		
		function validateForm(frm){						
			frm.target = '_self';			
			frm.submit();
		}
		
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "transFromDate","cmdDate","<%= FormatDate(FirstDay) %>",1);
		var cal1=new ctlSpiffyCalendarBox("cal1", "frmMain", "txtToDate","cmdDate","<%= FormatDate(LastDay) %>",1);

	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="AgentBalanceList.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">
		<table border="0">
			<tr>
				<td>Select date from:</td>
				<td  colspan="2">
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>				
			</tr>
			<tr>
				<td>To date:</td>
				<td colspan="2">
					<SCRIPT language="JavaScript">cal1.writeControl();</SCRIPT>	
				</td>
				
			</tr>			

			<tr>
				<td></td>
				<td><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... ">&nbsp;&nbsp;</td>		
				
			</tr>
		</table>
		
	</form>
	
	<%Set rs = Nothing
	Set Conn = Nothing
	Response.End
End If

%>


<% DrawPageFunctions True, True, True %>


<table border="0" cellspacing="2" cellpadding="2" style="font-family: Arial Narrow" width="100%">
    <tr>
	  <td bgcolor="#000000" width="80%" nowrap align="left"><font color="#FFFFFF" face="Impact" size="2">AGENT BALANCE LIST</font></td>
      <td bgcolor="#000000" width="20%" nowrap align=right><font color="#FFFFFF" face="Impact" size="2"><%= Session("CompanyName") %></font></td>
      
    </tr>

  </table>
<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" width="100%" > 
	<tr>      
      			<td ><b>Date:&nbsp;&nbsp;<%= FormatDate(Date) %></b></td>      			
    </tr>
	<tr>
		<td align="right" height="8">
            &nbsp;			
		</td>		
	</tr>
  <tr>      
      			<td ><b>From :&nbsp;&nbsp;<%= FormatDate(selectedFromDate) %>&nbsp;&nbsp;&nbsp;&nbsp;To :&nbsp;&nbsp;<%= FormatDate(selectedToDate) %></b></td>      			
    </tr>
</table>
<%

'sqlStr = "SELECT " & selColumns & " FROM AgentBalanceList " & SelectedSearchArgs & " " &  orderByCols 
sqlStr = "SELECT dbo.Agent.AgentName, dbo.Client.Agent_DPA_ ,SUM(dbo.LevyContract.LevyAmount) AS Commission" & _
		 " FROM         dbo.LevyContract INNER JOIN " & _
		 "                       dbo.Lots ON dbo.LevyContract.Contract_DPA_ = dbo.Lots.Contract_DPA_ INNER JOIN " & _
		 "                       dbo.OrdDetail ON dbo.Lots.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN " & _
		 "                       dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN " & _
		 "                       dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN " & _
		 "                       dbo.Agent ON dbo.Client.Agent_DPA_ = dbo.Agent.Agent_DPA_ " & _
		 " WHERE     (dbo.LevyContract.SystemMaintained = 12) AND (dbo.tbOrder.OrderSecType_DPA_ = 2) AND (CAST(FLOOR(CAST(dbo.Lots.LotTDate AS Float))  " & _
		 "                       AS DateTime) between '" & CDate(selectedFromDate) & "' and '" & CDate(selectedToDate) & "') " & _
		 " GROUP BY dbo.Client.Agent_DPA_, dbo.Agent.AgentName"
			

'Response.write(sqlStr)
'Response.end

Set Conn = GetActiveConnection("KBroker")

Set Rs = Conn.Execute(sqlStr)

 If Rs.EOF Or Rs.BOF Then%>
		<Script Language="JavaScript">	
			ShowMessage('No information was found using the criteria entered.');
			window.parent.history.go(-1);
		</Script>
		<%Set Conn = Nothing
		Set Rs = Nothing
		Response.End
  End If


 %>
    <table border="0" width="400" cellPadding="2" cellSpacing=0>
    <tr bgColor="#000000">
			
	<%For i = 0 To Rs.Fields.Count - 1
		if(i=1) then
		else
			if(i=2) then
			%>		
			 <td nowrap align="right"><b><font color="#FFFFFF">
			 <%= Rs.Fields(i).Name %></font></b></td>
		   <%
		   else
		   %>		
			 <td nowrap><b><font color="#FFFFFF">
			 <%= Rs.Fields(i).Name %></font></b></td>
		   <%
		   end if
	   end if
	   Next %>
	</tr>
	
	<%
	Total=0
	Do Until rs.EOF%>
        		<tr>
      				
      				<%For i = 0 To Rs.Fields.Count - 1
					if(i=1) then
					else
						if(i=2) then
						%>		
							 <td nowrap align="right"><%= FormatNum(Rs.Fields(i).Value) %></td>
					   <%
					   Total=total+Rs.Fields(i).Value
					   else
					   %>		
							 <td nowrap><%= Rs.Fields(i).Value %></td>
					   <%
					   end if
				   end if
				   Next %>			
      
	        </tr>
	 <%  Rs.MoveNext
	 Loop%>
	 <tr><td><b>Total</b></td><td align="right" style="border: 1px solid #000000" valign="top"><b><%=FormatNum(Total)%></b></td></tr>
  </table>
  
	 <%	 
	 Set Rs = Nothing
	 Set Conn = Nothing
     %>	
</body>



</html>
