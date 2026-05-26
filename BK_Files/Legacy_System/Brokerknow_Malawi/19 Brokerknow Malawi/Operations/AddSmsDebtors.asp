<!--#include virtual="libroutines.asp"-->
<%
	const UDLName = "KBroker"
	const DataSource = "SmsDebtors"
	const DataEntity = "SmsDebtors"
	const DataEntityPlural = "SmsDebtors"
	const ActionFolder = "Operations"

	Dim UserId
	Dim action
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim guid
	Dim guidStr
	Dim ID
    

	ID = Request("ID")

	If Trim(ID) = "" Then%>
            <script language = 'vbscript'>
                	ShowMessage "No record specified for editing"
            </script>
            <% response.end
    End If
	
	UserId=Session("UserID")
	action = ucase(Request.Form("action"))
	
	if action = "EXECUTE" then
	        
		    Dim smsDebit
			client= Request.Form("ID")
			smsDebit=Request.Form("smsDebit")
			
			if smsDebit="" then smsDebit=0	
			'save data
					
		     Set conn = GetActiveConnection("KBroker")
		  
				 conn.BeginTrans	
				 conn.Execute (sqlStr)
			 	
				conn.CommitTrans
					
			 WritefraEnabledDialogCloseScript
		
			 conn.Close
			 Set conn = Nothing
					
			 Response.End	
   	End if
   	
         	
%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit <%=DataEntity%></title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/accountList.js"></SCRIPT>
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<script language="JavaScript" src="CALENDAR/calendar.js"></script>
<!--END CALENDAR -->

</head>
<body Class="Dialog">
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<SCRIPT language="JavaScript">
	var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "txtDate","cmdDate","<%=FormatDate(now())%>",1);
</SCRIPT>

<form name = 'frmMain' method = 'post' action = 'AddEquityPrice.asp' id = "frmMain" >
<table border="0" width="100%" cellspacing="1" cellpadding="1">
 <tr>
    <td width="15%">Code</td>
    <td width="54%">
    <input type="text" name="code" id="code" value="<%=rs.fields("Client_DPA_")%>" size="20" class="readonly" readonly></td>
  </tr>
 <tr>
    <td width="20%">Client</td>
    <td width="54%">
    <input type="text" name="Security" id="Security" value="<%=rs.fields("ClientName")%>" size="40" 
    style= "BORDER-RIGHT: silver 1px ridge;BORDER-TOP: silver 1px solid;FONT-WEIGHT: normal;FONT-SIZE: 8pt;BORDER-LEFT: silver 1px solid;WIDTH: 300px;COLOR: navy; BORDER-BOTTOM: silver 1px outset;FONT-FAMILY: verdana, arial, helvetica, sans-serif;BACKGROUND-COLOR: #c0c0c0" 
    readonly></td>
  
  </tr>
  
  <tr>
    <td width="15%">Sms Debit</td>
    <td width="54%">
    <input type="checkbox" name="smsDebit" id="smsDebit" value="1" ></td>
  </tr>
  <tr>
	  <td width="100%" colspan=2 align="center" valign=absBottom>
		<BR>
		<input type = 'submit' Class=Buttons name ='cmdSave' id = 'cmdSave' value=" Save ">
    	<input type = 'button' Class=Buttons name ='cmdCancel' id = "cmdCancel" value="Cancel" onclick = "JavaScript: window.self.close()">
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		<input type = 'text' name ='ID' id = 'ID' value="<%=ID%>">
	</td>
  </tr>
</table>

</form>
</body>

</html>