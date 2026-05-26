<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit User</title>
 
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
	
	action = ucase(Request.Form("action"))
	ID = Request("ID")
	guid = Request.Form("guid")

		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for editing"
                		
                </script>
                <% response.end
        End If

	
	if action = "EXECUTE" then
	   
	   Dim buttonAction
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
       dim SecretQuestion, SecretAnswer, SecretAnswerConfirm, AskSecretQuestion,Accepted
       
	   buttonAction = Trim(Ucase(Request.Form("cmdAdd")))

	   if instr(1,buttonAction,"SAVE") > 0 then
			   otherNames = Request.Form("txtOtherNames")
			   userName = Request.Form("txtUserName")
			   surName = Request.Form("txtSurName")
			   Description = Request.Form("txtDescription")
			   passWord = Request.Form("txtPassword")
			   passWordConfirm = Request.Form("txtPasswordConfirm")
			   chkExtra =  Request.Form("chkExtra")
			   expiresDate =  Request.Form("txtEDate")
			   client =  "NULL"    
			   SecretQuestion = Request.Form("txtSecretQuestion")
			   SecretAnswer = Request.Form("txtSecretAnswer")
			   SecretAnswerConfirm = Request.Form("txtSecretAnswerConfirm")
			   
						
				'validate userName
				If Trim(userName) = "" Then%>
						<script language="vbscript">
							ShowMessage "Please specify the User name"                						
						</script>
						<% response.end
				End If
				
				'validate otherNames
				If Trim(otherNames) = "" Then%>
						<script language="vbscript">
												ShowMessage "Please specify the user's first name"
												
										</script>
						<% response.end
				End If
			   
				'validate size of userName
				If Len(userName) <= 3 Then%>
					<script language="vbscript">
									ShowMessage "User name should be more than 3 characters in length"
									
									</script>
					<% response.end
				End If
				 
				
			   
				isExpires = ""
				isChangePass = 0
				isEnabled = 1
				isRemoteUser = 0      
				AskSecretQuestion = 0
				Accepted=0

				If chkExtra <> "" Then
					Dim chkExtraArray, i
					chkExtraArray = Split(chkExtra, ",")			

					For i = 0 To UBound(chkExtraArray)
						Select Case LCase(Trim(chkExtraArray(i)))					
							Case "isexpires"
								isExpires = Request.Form("txtEDate")
								If DateDiff("d", Date, isExpires) < 0 Then%>
									<script language="vbscript">
												ShowMessage "The date of expiry is set in the past."                						
										</script>
								<%	Response.End
								End If
							Case "ischangepass"
								isChangePass = 1
							Case "isdisabled"
								isEnabled = 0	
							Case "isremoteuser"	
								isRemoteUser = 1
								client =  Request.Form("cboClient")
							case "asksecretquestion"	
								AskSecretQuestion = 1
							case "accepted"	
								Accepted = 1				
						End Select
					Next
				End If		
				
				Set conn = GetActiveConnection("KBroker")
				
				'check whether another similar username exists
				Set chkRs = Conn.Execute("SELECT UserName FROM [Users] WHERE UserName = '" & userName & "' AND UserID <> " & ID)
				If Not (chkRs.EOF Or chkRs.BOF) Then%>
					<script language="vbscript">
							ShowMessage "A similar username already exists."
					</script>
					<% Set chkRs = Nothing
					Response.End
				End If
				
				Set chkRs = Nothing		
				
					 
				sqlStr = "UPDATE [Users] SET UserName = '" & userName & "'," & _
						"        SurName = '" & surName & "'" & _
						"       , OtherNames = '" & otherNames & "'" & _
						"		, FirstTime =  " & isChangePass & " " & _
						"       , Description = '" & Description & "'" & _
						"       , RequiresSecretQuestion = " & AskSecretQuestion & "" & _
						"       , Accepted = " & Accepted & "" & _
						"       , Expires =  '" & isExpires & "'" & _                
						"       , Enabled = " & isEnabled & " WHERE UserID = " & ID
				
				
				conn.BeginTrans
					conn.Execute sqlStr	      
				conn.CommitTrans

				ClientEmail=request.form("email")

				Subject = "User login details"
				Body = "your login details to Dyer and blair intranet has changed" & vbCrLf
				Body = Body & "Username : " & userName & vbCrLf
				Body = Body & "Website : www.dyerandblair.com" & vbCrLf
				
				if(Isnull(ClientEmail) or ClientEmail="") then
				else
					If(Accepted=1) then
					SendMail ClientEmail, subject, Body
					end if
				end if
			
				Set Conn = Nothing
				WritefraEnabledDialogCloseScript
				Response.End 
				end If
		end if	
		
		if instr(1,buttonAction,"EMAIL") > 0 then
			
				ClientEmail=request.form("email")
				userName = Request.Form("txtUserName")

				passedPass=request.form("password")				

				passedPass = DecryptWithALP(passedPass)											

				Subject = "User login details"
				Body = "your login details to Dyer and blair intranet" & vbCrLf
				Body = Body & "Username : <b>" & userName & "</b>" & vbCrLf
				Body = Body & "Password : <b>" & passedPass & "</b>" & vbCrLf
				Body = Body & "Website : www.dyerandblair.com" & vbCrLf
				Body = Body & "Incase you cant log in with those details please contact us "  & vbCrLf
				
				if(Isnull(ClientEmail) or ClientEmail="") then
				else
				SendMail ClientEmail, subject, Body
				end if

				WritefraEnabledDialogCloseScript
				Response.End 
		end if

		sub SendMail(toRecipient, subject, bodyText)
		cc = ""
		bcc = ""
		
		Set Conn = GetActiveConnection("KBroker")

		SQL = "SELECT Email FROM         Users INNER JOIN " & _
			 "                       UserGroups ON Users.UserID = UserGroups.UserID " & _
			 " WHERE     (NOT (Users.email IS NULL)) AND (UserGroups.GroupID = 1) OR " & _
			 "                       (UserGroups.GroupID = 2) OR " & _
			 "                       (UserGroups.GroupID = 6)"

		Set Rs = Conn.Execute(SQL)
		
		if not (Rs.eof and Rs.bof) then
			Do while Rs.eof=false
			cc=cc & Rs("Email") & ","
			Rs.movenext
			loop
		end if

		SQL = "SELECT * FROM MAilConfigList"
		Set Rs = Conn.Execute(SQL)
		
		If (Rs.EOF Or Rs.BOF) Then%>
			<Script Language="JavaScript">
				alert("The mail configurations have not been set.");
			</Script>
			<%Response.End 
		End If
					
		Const cdoSchema = "http://schemas.microsoft.com/cdo/configuration/" 
		Set objMsg = CreateObject("CDO.Message") 
		
		With objMsg
		
			.Configuration.Fields.Item(cdoSchema & "sendusing") = Rs.Fields("SendUsingMethod").Value 
			.Configuration.Fields.Item(cdoSchema & "smtpserver") = Rs.Fields("SMTPServer").Value 
			.Configuration.Fields.Item(cdoSchema & "smtpserverport") = Rs.Fields("SMTPServerPort").Value 
			.Configuration.Fields.Update 		 
			.Subject = subject
			.Sender = Rs.Fields("SendDisplayName").Value 
			.To = toRecipient
			If cc <> "" Then
				.CC = cc
			End If
			
			If bcc <> "" Then
				.BCC = bcc
			End If
			
			If bodyText <> "" Then
				.TextBody = bodyText
			End If
			
			.HTMLBody = bodyText	' "file://" & pathTo
			
			On Error Resume Next
			.Send 
			
			If Err.number > 0 Then%>
				<Script Language="JavaScript">
					alert("The page could not be sent. An unexpected error occured: <%= Err.description %>")
				</Script>
				<%Response.End 
			End If
		End With
		
		Set Rs = Nothing
		Set Conn = Nothing
		Set objMsg = Nothing
		Set Fso = Nothing%>
			<Script Language="JavaScript">
				alert("The page was sent successfully.")		
			</Script>
		<%	
	end sub
   	Set conn = GetActiveConnection("KBroker")
   	
   	
   	Set rsEdit = Conn.Execute("SELECT * FROM [ExistingOnlineUserList] WHERE UserID = " & ID)
   	If rsEdit.EOF Or rsEdit.BOF Then%>
   		<Script Language="JavaScript">
   			ShowMessage("The specified user record cannot be retrieved")
   		</Script>
   		<%WriteDialogRefuseOpenScript
   		Response.End
   	End If
%>

<form name = 'frmEditUser' method = 'post' action = 'EditExistingOnlineUser.asp' >
<table border="0" width="100%" cellpadding=2 cellspacing=2>
	<tr>
                <td width="40%">User Name</td>
                <td width="60%"><input type="text" name="txtUserName" id="txtUserName" size="25" Value="<%= rsEdit.Fields("userName").Value %>"></td>
              </tr>
 <tr>
                <td width="40%">First Name</td>
                <td width="60%"><input type="text" name="txtOtherNames" id="txtOtherNames" size="25" Value="<%= rsEdit.Fields("OtherNames").Value %>"></td>
              </tr>
              <tr>
                <td width="40%">Surname</td>
                <td width="60%"><input type="text" name="txtSurName" id="txtSurName" size="25" Value="<%= rsEdit.Fields("SurName").Value %>"></td>
              </tr>
              <tr>
                <td width="40%">Description</td>
                <td width="60%"><input type="text" name="txtDescription" id="txtDescription" size="25" Value="<%= rsEdit.Fields("Description").Value %>"></td>
              </tr> 
              <tr>
              <td width="40%">Client Name</td>
              <td width="60%">		
				<input readonly = 'true' class=readonly STYLE="WIDTH: 175px; text-align: left" type="text" name="txtClient" id="txtClient" size="25" Value="<%= rsEdit.Fields("ClientName").Value %>"></td>
              </tr>                 
     <tr>
	 <td width="100%" colspan=2>
	 <%expDate = rsEdit.Fields("Expires").Value
	 If IsDate(expDate) Then %>	 
		<input type=checkbox name="chkExtra" Value="isExpires" checked Class="BorderLess" OnClick="JavaScript: swapExpires(this)" ID="chkExpires"><label for="chkExpires" style="cursor: hand">&nbsp;Account expires</label> 
		<SPAN ID="calExpiresDiv" Style="display: ">
  			<SCRIPT language="JavaScript">			
				var cal=new ctlSpiffyCalendarBox("cal", "frmEditUser", "txtEDate","cmdDate","<%= FormatDate(expDate) %>",1);
				cal.writeControl();
			</SCRIPT>
		</SPAN>
	<%Else%>
		<input type=checkbox name="chkExtra" Value="isExpires" Class="BorderLess" OnClick="JavaScript: swapExpires(this)" ID="chkExpires"><label for="chkExpires" style="cursor: hand">&nbsp;Account expires</label> 
		<SPAN ID="calExpiresDiv" Style="display: none">
  			<SCRIPT language="JavaScript">			
				var cal=new ctlSpiffyCalendarBox("cal", "frmEditUser", "txtEDate","cmdDate","<%= FormatDate(Date) %>",1);
				cal.writeControl();
			</SCRIPT>
		</SPAN>	
	<%End If

	'Response.write " we are here having fun"
	'Response.end

	If rsEdit.Fields("FirstTime").Value = "1" Then changePass = "checked"
	If rsEdit.Fields("Enabled").Value = "0" Then isDisabled = "checked"
	If rsEdit.Fields("RequiresSecretQuestion").Value = "1" Then AskSecretQuestion = "checked"
	If rsEdit.Fields("Accepted").Value Then Accepted = "checked"
	If rsEdit.Fields("RemoteUser").Value = "1" Then 
		isRemoteUser = "checked"
		showList = ""
	else
		showList = "none"
	end if
	
	
	%>	
		<P>		
		<input type=checkbox name="chkExtra" <%= changePass %> Value="isChangePass" Class="BorderLess" Checked ID="chkChangePass"><label for="chkChangePass" style="cursor: hand">&nbsp;User must change password at next logon</label> 
		<P>
		<input type=checkbox name="chkExtra" <%= isDisabled %> Value="isDisabled" Class="BorderLess" ID="chkDisabled"><label for="chkDisabled" style="cursor: hand">&nbsp;Account is disabled</label> 
		<P>
		<input type=checkbox name="chkExtra" <%= AskSecretQuestion %> Value="AskSecretQuestion" Class="BorderLess" ID="chkAskSecretQuestion"><label for="chkAskSecretQuestion" style="cursor: hand">&nbsp;Ask Secret Question</label>
		<P>
		<input type=checkbox name="chkExtra" <%= Accepted %> Value="Accepted" Class="BorderLess" ID="chkAccepted"><label for="chkAccepted" style="cursor: hand">&nbsp;Accept</label>
	 </td>
  </tr>  
  <tr>
     <td width="100%" colspan="2" align=right>
	
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save ">
		&nbsp;&nbsp;
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Email Password ">&nbsp;&nbsp;
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();">
		
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		<input type = 'hidden' name ='ID' id = 'ID' value="<%= ID %>">
		<input type = 'hidden' name ='email' id = 'email' value="<%= rsEdit.Fields("ClientEmail").Value %>">
		<input type = 'hidden' name ='password' id = 'password' value="<%= rsEdit.Fields("Password").Value %>">		
      </td>
  </tr>
</table>
</form>
<Script Language="JavaScript">
	function swapExpires(chk){
		if (chk.checked) calExpiresDiv.style.display="";
		else calExpiresDiv.style.display = "none";
	}
</Script>

</body>

</html>
