<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit Group</title>
   <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 

<script language='javascript'>
		function forceSubmit()
		{
			setOpener();
			document.frmEditGroup.method='post';
			document.frmEditGroup.target='_self';
			document.frmEditGroup.submit();		
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
   Dim ID
   Dim rsEdit
	
	action = ucase(Request.Form("action"))
	ID = Request("ID")
	guid = Request.Form("guid")

		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for editing"
                		
                </script>
                <% WriteDialogRefuseOpenScript
                response.end
        End If

	if action = "EXECUTE" then
		Dim groupName
       Dim description
       
       groupName = Request.Form("txtGroup")
       description = Request.Form("txtDescription")
      
       
        'validate Group Name
        If Trim(groupName) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Group Name"
                		
                </script>
                <% ReloadPage(ID)
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
        
        Set chkRs = Conn.Execute ("SELECT GroupName FROM [Groups] WHERE GroupName = '" & groupName & "' AND GroupID <> " & ID)
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
        sqlStr = "UPDATE [Groups] SET GroupName = '" & groupName & "', Description = '" & description & "' WHERE GroupID = " & ID
        
        
        conn.BeginTrans
                conn.Execute sqlStr
        conn.CommitTrans
        conn.Close
        Set conn = Nothing
        WriteFraEnabledDialogCloseScript2
        Response.End
   	end If
%>

<form name = 'frmEditGroup' method = 'post' action = 'EditGroup.asp'>
<table border="0" width="100%" cellspacing="1" cellpadding="1">
<%
        Set conn = GetActiveConnection("KBroker")
       
        sqlStr = "SELECT * FROM [Groups] WHERE GroupID  = " & ID        
        Set rs = conn.Execute(sqlStr)
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		ShowMessage "The selected Group cannot be retrieved for editing"
                		
                </script>
                <%WriteDialogRefuseOpenScript 
                response.end
        End If

		
%>
   <tr>
    <td width="30%" > Group</td>
    <td width="70%" ><input type = 'text' name ='txtGroup' id = 'txtGroup' size="20" Value="<%= Rs.Fields("GroupName").Value %>"></td>
  </tr>
  <tr>
    <td width="30%" > Description</td>
    <td width="70%" ><input type = 'text' name ='txtDescription' id = 'txtDescription' size="20" Value="<%= Rs.Fields("Description").Value %>"></td>
  </tr>
  <tr>
    <td width="100%" colspan="2" align=right>
		<BR>
		<BR>		
		<input type = 'button' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save " onclick="forceSubmit()">
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();">
		&nbsp;&nbsp;
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
    	<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
    	<input type = 'hidden' name ='guid' id = "guid" value="<%=guidStr%>">
		<input type = 'hidden' name ='buttonAction' id = 'buttonAction' value="Save">
      </td>
  </tr>
</table>
</form>

</body>

</html>
