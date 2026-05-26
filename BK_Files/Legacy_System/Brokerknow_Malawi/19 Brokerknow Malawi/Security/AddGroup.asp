<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add Group</title>
  <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 


<script language='javascript'>
	function forceSubmit()
		{
			setOpener();
			document.frmAddGroup.method='post';
			document.frmAddGroup.target='_self';
			document.frmAddGroup.submit();		
		}
		
		function setOpener()
		{
			window.self.opener = window.dialogArguments.opener;					
		}
</script>
</head>

<body Class="Dialog" onload="setOpener()">

<!--#include file="../libroutines.asp"-->

<%
	
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
	
	action = ucase(Request.Form("action"))
	
	if action = "EXECUTE" then
		Dim groupName
       Dim description
       
       groupName = Trim(Request.Form("txtGroup"))
       description = Request.Form("txtDescription")
      
       
        'validate Group Name
        If Trim(groupName) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Group Name"
                		
                </script>
                <% 
				ReloadPage(ID)
				response.end
        End If
        
        'validate Description
        If Trim(description) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Description"
                		
                </script>
                <% ReloadPage(ID)
				response.end
        End If
        'validate size of Group Name
        If Len(groupName) > 100 Then%>
                <script language = 'vbscript'>
                ShowMessage "Group Name can only be 100 characters in length"
                
                </script>
                <% ReloadPage(ID)
				response.end
        End If
         'validate size of Description
        If Len(description) > 100 Then%>
                <script language = 'vbscript'>
                ShowMessage "Description can only be 100 characters in length"
                
                </script>
                <% ReloadPage(ID)
				response.end
        End If
        
        Set conn = GetActiveConnection("KBroker")
        
		Set chkRs = Conn.Execute ("SELECT GroupName FROM [Groups] WHERE GroupName = '" & groupName & "'")
		If Not (chkRs.EOF Or chkRs.BOF) Then%>
			<Script Language="JavaScript">
				ShowMessage("The specified group already exists")
			</Script>
			<%
			Set chkRs = Nothing
			Set Conn = Nothing
			ReloadPage(ID)
				response.end
		End If
		
		Set chkRs = Nothing
		
        'save data
        sqlStr = "INSERT INTO [Groups] (GroupName,Description) SELECT " & "'" & groupName & "'" & " as GroupName," & "'" & description & "'" & " as Description"
        
        conn.BeginTrans
                conn.Execute sqlStr                
        conn.CommitTrans
        conn.Close
        Set conn = Nothing
        WriteFraEnabledDialogCloseScript2
        Response.End
   	end If
%>

<form name = 'frmAddGroup' method = 'post' action = 'AddGroup.asp' target="deleteFrame" OnSubmit="JavaScript: UpdateDialogHandle();">
<table border="0" width="100%" cellspacing="1" cellpadding="1">
  <tr>
    <td width="30%" > Group</td>
    <td width="70%" ><input type = 'text' name ='txtGroup' id = 'txtGroup' size="20"></td>
  </tr>
  <tr>
    <td width="30%" > Description</td>
    <td width="70%" ><input type = 'text' name ='txtDescription' id = 'txtDescription' size="20"></td>
  </tr>
  <tr>
   <td width="100%" colspan="2" align=right>
		<BR>
	
		<input type = 'button' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save " onclick="forceSubmit()">
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();">
		&nbsp;&nbsp;
		<input type = 'hidden' name ='action' id = 'action' value="Execute">&nbsp;
		<input type = 'hidden' name ='buttonAction' id = 'buttonAction' value="Save">

      </td>
  </tr>

</table>
</form>

</body>

