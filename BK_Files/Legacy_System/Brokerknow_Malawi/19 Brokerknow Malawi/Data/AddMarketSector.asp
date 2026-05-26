<!--#include file="../libroutines.asp"-->
<%
	const UDLName = "KBroker"
	const DataSource = "AddMarketSector"
	const DataEntity = "MarketSector"
	const DataEntityPlural = "MarketSectors"
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
		buttonAction = Trim(Ucase(Request.Form("buttonAction")))
		if buttonAction = "SAVE" then
				Dim strDesc
				Dim Ref
			
				strDesc = Request.Form("txtDescription")				
				
				'validate Description
				If Trim(strDesc) = "" Then%>
				         <script language = 'vbscript'>
				         		ShowMessage "Please specify a short Description."
				         		
				         </script>
				         <% 
						 ReloadPage(ID)
						 response.end
				End If
				
				'save data		
		
				sqlStr = "INSERT INTO [MarketSector] (ShortDescription,Sector_DPA_)" & _
						" SELECT " & "'" & strDesc & "'" & " as ShortDescription" & _
						"       ," & " " & "iif(isnull(max([sector_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'MarketSector'),max([Sector_DPA_]) + 1)" & " " & " as Sector_DPA_" & _
						"        FROM [MarketSector]"
				Set conn = GetActiveConnection("KBroker")
				
				sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))

				conn.BeginTrans
						conn.Execute sqlStr
				conn.CommitTrans
				conn.Close
				Set conn = Nothing
				WritefraEnabledDialogCloseScript2
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
			forceSubmit();
		}
		function forceSubmit()
		{
			setOpener();
			document.frm<%=DataSource%>.method='post';
			document.frm<%=DataSource%>.target='_self';
			document.frm<%=DataSource%>.submit();		
		}
		
		function setOpener()
		{
			window.self.opener = window.dialogArguments.opener;					
		}
</script>
</head>

<body Class="Dialog" onload="setOpener()">

<form name = 'frm<%=DataSource%>' method = 'post' action = '<%=DataSource%>.asp' id = "frmMain">
<table border="0" width="100%" cellspacing="1" cellpadding="1">
  <tr>
    <td width="40%"> Short Description</td>
    <td width="60%"><input type = 'text' name ='txtDescription' id = 'txtDescription' size="20"></td>
  </tr>    
  <tr>
	  <td width="100%" colspan=3 align="center" valign=absBottom>
		<BR><BR>
		<input type = 'button' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value="Save" onclick = "AllowedNavigation()">
    	<input type = 'button' Class=Buttons name ='cmdCancel' id = "cmdCancel" value="Cancel" onclick = "JavaScript: window.self.close()">
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		<input type = 'hidden' name ='buttonAction' id = 'buttonAction' value="Save">
		
	</td>
  </tr>
</table>

</form>
</body>

</html>
