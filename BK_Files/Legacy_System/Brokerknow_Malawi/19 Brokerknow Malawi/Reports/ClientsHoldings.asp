<html>
<%
	' DB Connection
Dim Conn 
    Set Conn = CreateObject("ADODB.Connection")
    theDBName = "KBroker" 
    Conn.ConnectionString =  "FILE NAME=" & GetUDLPath(theDBName) 
    Conn.Open
    
   Function GetUDLPath(theDBName) 
    Dim tmpStr
    
    tmpStr = StrReverse(Request.ServerVariables("APPL_PHYSICAL_PATH"))
    
    tmpStr = Mid(tmpStr, InStr(1, tmpStr, "\") + 1)
    
    tmpStr = StrReverse(tmpStr)
    
    GetUDLPath = tmpStr & "\UDL\" & Trim(theDBName) & ".UDL"

End Function
%>
<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>CDS HOLDING BALANCES</title>
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
selectedClient = Request.Form("cboClient")
selectedFromDate = Request.Form("transFromDate")

If genReport <> "1" Or selectedClient = "" Then%>
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
		
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "transFromDate","cmdDate","<%= FormatDate(DateAdd("d", -90, Date)) %>",1);
	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="ClientsHoldings.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">
		<table>
			<tr>
								 <td>&nbsp;Client</td>
    <td>
       <input type = 'text' name ='txtClientCode' id = 'txtClientCode' size="10" onBlur="txtval = this.value; selectItem(cboClient);" value="<%=client%>">
    <select name = 'cboClient' id = 'cboClient' size="1" 
			onKeypress="return (dodefaultaction()==''); " 
			onKeydown="return (dodefaultaction()==''); " 
			onKeyup="return (UpdateCode(change(cboClient,0),cboClient,txtClientCode));" 
			onChange="UpdateCode(true,cboClient,txtClientCode);"
			onfocus="txtval = '';inputIsItemCode = 1;" 
			onblur="txtval = '';inputIsItemCode = 1;" readonly >
    	<!--option selected SearchCode = "" SearchText = "" value = ''></option-->
    	<option selected SearchCode = "" SearchText = "" value = '' AccManager=""></option>
		<%
		set rs =server.createObject("Adodb.recordset")
		dim ClientName
		dim NameClient
		set rs = server.CreateObject("Adodb.recordset")
       '' sqlStr = "SELECT * FROM [BondClientList]"
		

	sqlStr =" SELECT * from fullclientlist"

        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF					                
                
                ClientName=rs.Fields("ClientName")
			    ''NameClient=rs.Fields("Client_DPA_") & " " & Mid(ClientName,1,20)
					    %>                    
                        <option  SearchCode = "<%=rs.Fields("Client_DPA_")%>" SearchText = "<%=rs.Fields("ClientName")%>" value = '<%=rs.Fields("Client_DPA_")%>' 
                        <%if (rs.Fields("Client_DPA_")=clng(client))then
							Response.Write "selected"
                         end if%> ><%=mid(ClientName,1,30)%></option>
                        <%rs.MoveNext
                Loop
        End If
              
		%>
    </select></td>
			</tr>
			<!--<tr>
				<td>Select date from:</td>
				<td>
					<SCRIPT language="JavaScript">//cal.writeControl();</SCRIPT>	
				</td>
			</tr>-->
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
	sqlStr ="SELECT * FROM ClientsHoldings WHERE (BalanceFree ='Y') and (Client_DPA_ = " & selectedClient &")ORDER BY Security_DPA_"
	
	
	'Response.Write sqlstr
	'Response.End
	'sqlStr = "ClientStatement"
	Rs.CursorLocation = adUseClient	
	Rs.Open SQLServerFormat(sqlStr), conn.ConnectionString, adOpenKeyset, adLockOptimistic
	'Rs.Filter = "Client_DPA_ = '" & selectedClient & "' AND TransDate >= '" & FormatDate(selectedFromDate) & "'"
	
	
	If rs.EOF Or rs.BOF Then%>
		<Script Language="JavaScript">
			alert("The specified client does not have any Holdings")
			window.history.go(-1);
		</Script>
		<%Set Rs = Nothing
		Set Conn = Nothing
		Response.End
	End If
	
	Set rsClient = Conn.Execute ("SELECT * FROM Client WHERE Client_DPA_ = " & selectedClient)
	
	If Not (rsClient.EOF Or rsClient.BOF) Then
		accountDesc = rsClient.Fields("ClientName").Value & " " & rsClient.Fields("Client_DPA_").Value
		accountAddress = rsClient.Fields("ClientAddr").Value
	End If
	
	Set rsClient = Nothing
	
	'Set Rs = Conn.Execute ("SELECT * FROM ClientsHoldings WHERE ClientDPA = " & selectedClient)	
%>	

<table border="0" cellspacing="2" cellpadding="2" style="font-family: Arial Narrow" width="100%">
    <tr>
		<td width="10%" nowrap><font face="Impact" size="4">CDS HOLDING BALANCES</font></td>
      <td width="60%" nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>
      
    </tr>

  </table>
<br>
<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
    <tr>
      <td width="1%"><b>Date:</b></td>
      <td width="48%"><%= FormatDate1(rs("TradeDate")) %></td>
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
<BR>


  <table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow; LEFT-MARGIN:100PX"  width="100%">
    <tr>
      <td style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"><b><font face="Arial Narrow" size="3">Security:</font></b></td>
      <td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"  align=right><b><font face="Arial Narrow" size="3">Quantity:</font></b></td>
      <td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"  align=right><b><font face="Arial Narrow" size="3">&nbsp;</font></b></td>
      <td style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"><b><font face="Arial Narrow" size="3">Account Status:</font></b></td>
      
    </tr>

   <%
     
    Do Until Rs.EOF
    %>
		<tr>	
		  <td><font size="1"><%= rs.Fields("SecurityCode").Value %></font></td>		  
		  <td  style="text-align: right"><font size="1" ><%= formatnumber(Rs.Fields("Quantity").Value,0) %></font></td>
		  <td><font size="1" >&nbsp;</font></td>
		  <td><font size="1" ><%= Rs.Fields("AccountStatus").Value %></font></td>
		</tr>
	
	<%	Rs.MoveNext
	Loop
	%>
	
    <tr>
      <td colspan="6" align="right" style="border-bottom-style: solid; border-bottom-width: 1">
        &nbsp;&nbsp;&nbsp; </td>

    </tr>


  </table>
   
<%Set Rs = Nothing
Set Conn = Nothing%>   
</body>

</html>