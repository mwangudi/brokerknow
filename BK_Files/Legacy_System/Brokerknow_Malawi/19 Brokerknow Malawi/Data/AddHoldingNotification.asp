<!--#include virtual="../libroutines.asp"-->
<%
	const UDLName = "KBroker"
	const DataSource = "AddHoldingNotification"
	const ActionPage = "HoldingsNotificationList"
	const DataEntity = "HoldingNotification"
	const DataEntityPlural = "Holding Notification"
	const ActionFolder = "Data"
	
	Dim action
	Dim conn 
	Dim sqlStr
	Dim rs
		
	action = ucase(Request.Form("action"))
	
	if action = "EXECUTE" then
		Dim buttonAction
		Dim reloadRequired
		
		reloadRequired = false
		buttonAction = Trim(Ucase(Request.Form("cmdAdd")))
		if buttonAction = "SAVE" then
				Dim Description
				Dim Entity
				Dim Title
			
				Description = Request.Form("txtDescription")
				Entity = Request.Form("cboTitle")
				
				'validate Description
				If Trim(Description) = "" Then%>
				         <script language = 'vbscript'>
				         		ShowMessage "Please specify a Description."
				         		
				         </script>
				         <% response.end
				End If
				
				if Entity = 1 then Title = "Email Address" 'A bit of Hard Coding
				'save data		
		        Set conn = GetActiveConnection("KBroker")
		        
				sqlStr = "INSERT INTO [SystemNotification] (Title,Description, Entity_DPA_)" & _
						" SELECT " & "'" & Title & "'" & " as Title " & _
						"       ," & "'" & Description & "'" & " as Description " & _
						"       ," & " " & Entity & " " & " as Entity_DPA_" & _
						"        FROM [SystemNotification]"
				
			    sqlStr = "INSERT INTO [SystemNotification] (Title,Description, Entity_DPA_) " & _
						" Values (" & "'" & Title & "', '" & Description & "', " & Entity & ")" 
						    
				conn.BeginTrans
						conn.Execute sqlStr
				conn.CommitTrans
				conn.Close
				Set conn = Nothing
				WritefraEnabledDialogCloseScript
				Response.End
			end if
			Dim clientCode
        
			clientCode = "var validNavigate = true;" & chr(13)
			%>
			<script>
				<%=clientCode%>
			</script>
			<%
			response.End
   	end If
%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add <%=DataEntity%></title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="../Operations/CALENDAR/calendar.css">
<script language="JavaScript" src="../Operations/CALENDAR/calendar.js"></script>
<!--END CALENDAR -->

<script language='javascript'>
		var validNavigate = false;
		function ReleaseRecord()
		{
			if(!validNavigate)
			{
 				event.returnValue = "Please use the cancel button to close the dialog"
 			}
		}
		
		function AllowedNavigation()
		{
			validNavigate = true;
		}
		
		

</script>
</head>

<body Class="Dialog">

<form name = 'frm<%=DataSource%>' method = 'post' action = '<%=DataSource%>.asp' id = "frmMain">
<table border="0" width="100%" cellspacing="1" cellpadding="1">
  <tr>
    <td width="40%"> Title</td>
    <td width="60%">
    <select name ='cboTitle' id = "cboTitle">
	<option value="1" selected>Email Address</option>
	</select></td>
  </tr>
  <tr>
    <td width="40%"> Description&nbsp; </td>
    <td width="60%">
	<input type = 'text' name ='txtDescription' id = 'txtDescription' value="" size = "40">
	</td>
  </tr>
  
  <tr>
	  <td width="100%" colspan=3 align="center" valign=absBottom>
		<BR><BR>
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value="Save" onclick = "AllowedNavigation()">
    	<input type = 'button' Class=Buttons name ='cmdCancel' id = "cmdCancel" value="Cancel" onclick = "JavaScript: window.self.close()">
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		
	</td>
  </tr>
</table>

</form>
</body>

</html>
