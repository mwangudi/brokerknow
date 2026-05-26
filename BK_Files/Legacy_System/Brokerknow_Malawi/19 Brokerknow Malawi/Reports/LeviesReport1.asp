<!--#include file="../libroutines.asp"-->

<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Portfolio BS</title>  
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
<SCRIPT language=VBScript src="../scripts/reports.vbs"></SCRIPT>
<SCRIPT language=Javascript src="../scripts/reports.js"></SCRIPT>



<style media="print">
		@page {
				size: landscape;
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



<%
FirstDay=DateSerial(Year(Date), Month(Date)-1 + iOffset, 1)
LastDay=DateSerial(Year(Date),Month(Date),1)-1
'LastDay=Date-(Day(Date)+1)
genReport = Request.Form("genReport")
selectedClient = Request.Form("cboClient")
selectedFromDate = Request.Form("txtFromDate")
selectedToDate = Request.Form("txtToDate")
If genReport <> "1" Then%>
		<Script Language="JavaScript">
		report_SetBodyClass();
		function validateForm(frm){			
			//if (frm.txtDate.value==''){
			//	alert("Select a date");
			//	frm.txtDate.focus();
			//	return;
			//}
			
			frm.target = '_self';			
			frm.submit();
		}
		
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "txtFromDate","cmdDate","<%= FormatDate(FirstDay) %>",1);
		var cal1=new ctlSpiffyCalendarBox("cal1", "frmMain", "txtToDate","cmdDate","<%= FormatDate(LastDay) %>",1);
	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="LeviesReport1.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">	
		<table>
			<tr>
				<td>Levies: </td>
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
					        
					        sqlStr = "SELECT DISTINCT LevyName,SystemMaintained FROM LevyContract ORDER BY LevyName"

					        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
					        If Not (rs.EOF Or rs.BOF) Then
					                rs.MoveFirst
					                Do Until rs.EOF
					                ClientName=rs.Fields("LevyName")
					                NameClient=Mid(ClientName,1,30)
					                %>					                        
					                        <option SearchCode = "<%=rs.Fields("SystemMaintained")%>" SearchText = "<%=NameClient%>" value = '<%=rs.Fields("SystemMaintained")%>'><%=NameClient%></option>

					                        <%rs.MoveNext
					                Loop
					        End If
					%>

					    </select>
				</td>
			</tr>
			
			<tr>
				<td>From Date:</td>
				<td>
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>
			</tr>
			<tr>
				<td>To date:</td>
				<td>
					<SCRIPT language="JavaScript">cal1.writeControl();</SCRIPT>	
				</td>
			</tr>
			
			<tr>
				<td colspan=2><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... ">&nbsp;&nbsp;</td>
			</tr>
		</table>
		
	</form>
	
	<%
	Response.End
End If

%>


<% DrawPageFunctions True, True, True%>
<%
		Set conn = GetActiveConnection("KBroker")

selectedTradeDate = FormatDate(selectedTradeDate)
headerDescription = " Between " & FormatDate(selectedFromDate) & " and " & FormatDate(selectedToDate) 

		If Len(selectedClient & "") = 0 then
			%>
			<script>
				alert("Please select a levy.");
				window.history.back();
			</script>
			<%
			Response.End
		End If

 		sqlStr = "Select * From TrialLeviesReport Where SystemMaintained="& selectedClient & " and " & _
			   " TransDate between '" & CDate(selectedFromDate) & "' and '" & CDate(selectedToDate) & "' order by TransDate,Contract_DPA_"

		Set Rs = CreateObject("ADODB.Recordset")		
		Rs.CursorLocation = adUseClient				

	Set Rs=Conn.Execute(sqlStr)
	
	if rs.EOF or rs.BOf then
		  %>
                <script language = 'javascript'>
                		alert ("No Levies Found");
                		window.parent.history.go(-1);          		
                </script>
                
                <% Response.End   
		 end if
%>

<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
     <tr>
		<td nowrap><b><font face="Arial Narrow" size="4">Traded Levies</font></b></td>
		<td nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>
	</tr>	
    <tr>
	   <td COLSPAN=2><font face="Arial" size="2">for the Dates :  <%= headerDescription %></font></td>
	</tr>
    <tr>
		  <td COLSPAN=2><font face="Arial" size="2">&nbsp;</font></td>
	</tr>
</table>				

    <table border="0" width="100%" cellPadding="2" cellSpacing=0>
    <tr>
	<td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial" size="2">Contract</font></b></td>
	<td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial" size="2">Traded</font></b></td>
	<td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial" size="2">Type</font></b></td>
	<td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial" size="2">Security</font></b></td>
	<td bgcolor="#000000" align=right><b><font color="#FFFFFF" face="Arial" size="2">Gross</font></b></td>
	<td bgcolor="#000000" align=right><b><font color="#FFFFFF" face="Arial" size="2"><%=rs("LevyShortName")%></font></b></td>
	<td bgcolor="#000000" align=right><b><font color="#FFFFFF" face="Arial" size="2">MSE Slip</font></b></td>
  </tr>
<%		
	totalLevyAmount = 0
	Do Until rs.EOF
			totalLevyAmount = totalLevyAmount + FormatNum(rs.Fields("LevyAmount")) %>
                <tr>
                        <td><%=rs.Fields("ContractNumber")%></td>
                        <td><%=FormatDate(rs.Fields("TransDate"))%></td>
                        <td><%=rs.Fields("OrdDetailType")%></td>
                        <td><%=rs.Fields("SecurityCode")%></td>
                        <td align=right><%=rs.Fields("LotPrice")%></td>
                        <td align=right><%= FormatNum(rs.Fields("LevyAmount")) %></td>
                        <td align=right><%=rs.Fields("LotSlipNo")%></td>
                </tr>
                <%                
                rs.MoveNext
        Loop
        conn.Close
        Set conn = Nothing%>
        
        <tr>
				<td colspan=7>&nbsp;</td>
         </tr>        
         <tr>
						<td colspan=5 align=right><font face="Arial" size="2"><b>Totals:</b></font></td>
                        <td align=right style="border-style: solid; border-color: #000000; border-width: 1" height="30px"><%= FormatNum(totalLevyAmount) %></td>
                        <td>&nbsp;</td>
         </tr>        

  </table>
 

</body>

</html>