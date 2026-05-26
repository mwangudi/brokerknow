<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit Password Settings</title>
 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>

<!--CALENDAR -->
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
 <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css"> 
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT>
<script language="JavaScript" src="CALENDAR/calendar.js"></script>
<SCRIPT language=Javascript src="../scripts/reports.js"></SCRIPT>

<!--END CALENDAR -->
</head>

<body Class="Dialog">

<!--#include file="../libroutines.asp"-->
<div id="spiffycalendar" class="text"></div>

<%
	Dim action
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim guidStr 
	Dim guid 
	Dim rsEdit
	
	UserID = Session("UserID")
	action = ucase(Request.Form("action"))
	
	if UCase(Request.QueryString("action")) = "CONFIRM" then
		%>
		<p><font size="4">The Password Settings have been successfully saved.</font></p>
		<p><font size="4">Back to <a href='#' onClick="javascript: window.parent.location.href='../Webdefault.asp'">home page</a></font></p>
		<%
		Response.End
	end if
	
	If Trim(UserID) = "" Then
		%>
		<script language = 'vbscript'>
			ShowMessage "An error occurred. Please log in again."
		</script>
		<%
		Response.End
    End If
	
	if action = "EXECUTE" then
	   Dim otherNames
       Dim userName
       Dim surName
       Dim Description
       Dim passWord
       Dim passWordConfirm
       Dim chkExtra
       Dim expiresDate
       Dim isExpires, isChangePass, isEnabled, isRemoteUser
       Dim client
       'stop
       otherNames = Request.Form("txtOtherNames")
       userName = Request.Form("txtUserName")
	   surName = Request.Form("txtSurName")
       passWord = Request.Form("txtPassword")
       passWordConfirm = Request.Form("txtPasswordConfirm")
       SecretQuestion = Request.Form("txtSecretQuestion")
       SecretAnswer = Request.Form("txtSecretAnswer")
       SecretAnswerConfirm = Request.Form("txtConfirmSecretAnswer")
       
        'validate userName
        If Trim(userName) = "" Then%>
			<script language="vbscript">
                ShowMessage "Please specify the User Name"                						
			</script>
			<% response.end
        End If
        
        'validate otherNames
        If Trim(otherNames) = "" Then%>
			<script language="vbscript">
            	ShowMessage "Please specify the user's first name"
			</script>
			<%response.end
        End If
       
        'validate size of userName
        If Len(userName) <= 3 Then%>
			<script language="vbscript">
				ShowMessage "User name should be more than 3 characters in length"
			</script>
			<%response.end
        End If
     	
		'validate passWord
		If Trim(passWord) = "" Then%>
				<script language="vbscript">
                	ShowMessage "Please specify the Password"
				</script>
				<%response.end
        End If
         
        'validate size of passWord
        If Len(passWord) <= 3 Then%>
			<script language="vbscript">
				ShowMessage "Password should be more than 3 characters in length"
			</script>
			<%response.end
        End If
        
		'validate Secret Question
		If Trim(SecretQuestion) = "" Then%>
				<script language="vbscript">
                	ShowMessage "Please specify the Secret Question"
				</script>
				<%response.end
        End If
         
		'validate Secret Answer
		If Trim(SecretAnswer) = "" Then%>
				<script language="vbscript">
                	ShowMessage "Please specify the Secret Answer"
				</script>
				<%response.end
        End If
         
        'validate Secret Question
        If SecretAnswer <> SecretAnswerConfirm Then%>
			<script language="vbscript">
            	ShowMessage "The secret answers do not match. Please retry."
			</script>
			<%response.end
        End If
        
        'validate passWord
        If passWord <> passWordConfirm Then%>
			<script language="vbscript">
            	ShowMessage "The passwords do not match. Please retry."
			</script>
			<%response.end
        End If
        
		Set conn = GetActiveConnection("KBroker")
		
		otherNames = Replace(otherNames,"'","''")
		surName = Replace(surName,"'","''")
		SecretQuestion = Replace(SecretQuestion,"'","''")
		SecretAnswer = Replace(SecretAnswer,"'","''")
       
        sqlStr = "UPDATE Users SET Password = N'" & EncryptWithALP(passWord) & "', Surname = N'" & surName & "', " & _
			" OtherNames = N'" & otherNames & "', SecretQuestion = N'" & SecretQuestion & "', SecretAnswer = N'" & EncryptWithALP(SecretAnswer) & "' WHERE " & _
			" (UserID = " & UserID & ")"
        
        conn.BeginTrans
			conn.Execute sqlStr	      
		conn.CommitTrans
        Set Conn = Nothing
   		%>
   		<Script Language="JavaScript">
   			window.location.href = "../web/PWordSettings.asp?action=confirm"
   		</Script>
   		<%
        Response.End 
   	end If
   	
   	Set conn = GetActiveConnection("KBroker")
   	
   	SqlStr = "SELECT UserName, Password, Surname, OtherNames, Removed, SecretQuestion, SecretAnswer, " & _
   		" RequiresSecretQuestion FROM Users WHERE (UserID = " & UserID & ")"
   	
   	Set rsEdit = Conn.Execute(SqlStr)
   	If (rsEdit.BOF Or rsEdit.EOF) Then
   		%>
   		<Script Language="JavaScript">
   			ShowMessage("An error occurred. Please log in again.")
   		</Script>
   		<%
   		Response.End
   	End If
%>
<div align="center" width="70%">
<form name = 'frmPWordSettings' method = 'POST' action = 'PWordSettings.asp' >
<table border="0" width="500" cellpadding='2' cellspacing='2'>
	<tr>
		<td width="40%">User Name</td>
		<td width="60%"><input type="text" name="txtUserName" id="txtUserName" size="25" Value="<%=rsEdit.Fields("userName").Value %>" readonly></td>
	</tr>
	<tr>
		<td width="40%">First Name</td>
		<td width="60%"><input type="text" name="txtOtherNames" id="txtOtherNames" size="25" Value="<%=rsEdit.Fields("OtherNames").Value %>"></td>
    </tr>
	<tr>
		<td width="40%">Surname</td>
		<td width="60%"><input type="text" name="txtSurName" id="txtSurName" size="25" Value="<%=rsEdit.Fields("SurName").Value %>"></td>
	</tr>
	<tr>
		<td colspan="2" width="100%"><HR></td>
	</tr>
	<tr>
		<td width="40%">Password</td>
		<td width="60%"><input type="Password" name="txtPassword" id="txtPassword" size="25" Value="<%=DecryptWithALP(rsEdit.Fields("Password").Value) %>"></td>
	</tr>
	<tr>
		<td width="40%">Confirm Password</td>
		<td width="60%"><input type="Password" name="txtPasswordConfirm" id="txtPasswordConfirm" size="25" Value="<%= DecryptWithALP(rsEdit.Fields("Password").Value) %>"></td>
	</tr>
	<tr>
		<td colspan="2" width="100%"><HR></td>
    </tr>    
	<tr>
		<td width="40%">Secret Question</td>
		<td width="60%"><input type="text" name="txtSecretQuestion" id="txtSecretQuestion" size="25" Value="<%=rsEdit.Fields("SecretQuestion").Value%>"></td>
	</tr>
	<tr>
		<td width="40%">Secret Answer</td>
		<td width="60%"><input type="Password" name="txtSecretAnswer" id="txtSecretAnswer" size="25" Value="<%=DecryptWithALP(rsEdit.Fields("SecretAnswer").Value)%>"></td>
	</tr>
	<tr>
		<td width="40%">Confirm Secret Answer</td>
		<td width="60%"><input type="Password" name="txtConfirmSecretAnswer" id="txtConfirmSecretAnswer" size="25" Value="<%=DecryptWithALP(rsEdit.Fields("SecretAnswer").Value)%>"></td>
	</tr>
	<tr>
		<td colspan="2" width="100%"><HR></td>
	</tr>
	<tr>
	   <td width="100%" colspan="2" align=right>
			<BR>
			<BR>
			<BR>
			<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save ">&nbsp;&nbsp;
			<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.location.replace('../webmenu.asp')">
			<input type = 'hidden' name ='action' id = 'action' value="Execute">
			<input type = 'hidden' name ='ID' id = 'ID' value="<%=UserID%>">
	    </td>
	</tr>
</table>
</form>
</div>
<Script Language="JavaScript">
	function swapExpires(chk)
	{
		if (chk.checked) calExpiresDiv.style.display="";
		else calExpiresDiv.style.display = "none";
	}
</Script>

</body>

</html>