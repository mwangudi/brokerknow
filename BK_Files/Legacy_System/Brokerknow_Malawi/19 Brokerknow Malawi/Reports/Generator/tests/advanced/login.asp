<%
Dim blnIsValidUser
Dim strMessage
If Request.Form.Count>0 Then	
	blnIsValidUser = IsValidUser(Request.Form("username"),Request.Form("password"))
	If blnIsValidUser Then
		SetUserSession
		
		Response.Cookies("ReportWizUser")=Session("UserName")
		Response.Cookies("ReportWizUser").Expires=DateAdd("d",365,Now)

		Response.Redirect "reports.asp"
	Else
		If usr<>"" then
			strMessage="<font color=red>Access Denied</font>"
		end if
	End if
End If

Function IsValidUser(usr,pwd)
	if usr=pwd and usr<>"" then
		IsValidUser=True
	Else
		IsValidUser=False
	end if
End Function

Sub SetUserSession
	Session("UserName")=Request.Form("username")	
	Select Case LCase(Session("UserName"))
		Case "admin" : Session("IsAdmin")=True
		Case "hr" : Session("Departments")="HR":Session("CanAdd")=True
		Case Else : Session("Departments")="Public"
	End Select	
	If Request.Form("images")="ON" Then
		Session("UseImages")=True
	End If
	Session.Timeout = 30
End Sub

Sub Test(s)
	Response.Write "Test: " & s
	Response.End 
End SUb
%>

<html>
	<head>
		<title>Reports Login</title>
		<link rel="stylesheet" type="text/css" href="default.css">
	</head>

<body bgcolor=#6699cc style="margin:0px;" scroll="auto">
<table width="100%" height="100%"  cellpadding="4" cellspacing="4">
	<tr>
		<td height="30">
			
		</td>
	</tr>
	<tr>
		<td height="20">&nbsp;</td>
	</tr>
	<tr>
		<td width="100%" height="70%" valign="top" >
			<p>&nbsp;</p>
			<blockquote>
			
			<!-- login form -->
			<form method=post action="login.asp" id=form1 name=form1>

<table align="center" cellpadding=2 cellspacing=1>
	<tr>
		<th colspan="2">Reports Login Form</th>
		
	</TR>
	<tr>
		<td align="right">User Name:</td>
		<td><input type="text" name="username" /></td>
	</tr>
	<tr>
		<td align="right">Password:</td>
		<td><input type="password" name="password"/></td>
	</tr>
	<%If strMessage<>"" Then%>
	<tr>
		<td align="center" bgcolor=white colspan=2><%=strMessage%></td>
	</tr>
	<%End If%>
	<tr>
		<td>&nbsp;</td>
		<td><input type="checkbox" value="ON" name="images" id=chk1>&nbsp;<label for=chk1>Use Images</label></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td><input type="submit" value="Login" id=submit1 name=submit1></td>
	</tr>
		
	<tr>
		<td colspan=2 align=center>
		
	<a href="../">Home</a>
	</td>		
	</tr>
	
</table>
</form>
<!-- end login form -->
			
			</blockquote>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
</table>
</body>
</html>
