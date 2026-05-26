<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add Online User</title>
 
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

	Set conn = GetActiveConnection("KBroker")

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
       dim SecretQuestion, SecretAnswer, SecretAnswerConfirm, AskSecretQuestion
       
       otherNames = Request.Form("txtOtherNames")
       userName = Request.Form("txtUserName")
	   surName = Request.Form("txtSurName")
       Description = Request.Form("txtDescription")
       'passWord = Request.Form("txtPassword")
       passWordConfirm = Request.Form("txtPasswordConfirm")
       chkExtra =  Request.Form("chkExtra")
       expiresDate =  Request.Form("txtEDate")
       client = Request.Form("cboClient")    
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
        
        'validate Client Selected
        If client = "" Then%>
			<script language="vbscript">
							ShowMessage "You must Select a Client"
							
							</script>
			<% response.end
        End If  
               
     	
       'obtain header key value
		sqlStr = "SELECT * FROM [Client] WHERE [Client_DPA_] = " & client 
		
		Set rs = conn.Execute(sqlStr)
		
		If (rs.EOF Or rs.BOF) Then%>
				<script language = 'vbscript'>
						ShowMessage "A serious error has been encountered while saving the data. Try saving again"
						window.history.back
				</script>
				<% response.end
		End If	

        isExpires = ""
		isChangePass = 0
		isEnabled = 1
		isRemoteUser = 0      
		AskSecretQuestion = 0	
		

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
				End Select
			Next
        End If
		
		'Set conn = GetActiveConnection("KBroker")
		
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
		
		tmpUserName = userName
		password = EncryptWithALP(tmpUserName)
		tmpPassword = password					
        		
				sqlStr = "INSERT INTO [Users] (UserID,UserName,[Password],SurName,OtherNames" & _
							"       ,FirstTime,Description,Expires,RemoteUser,Client_DPA_,Enabled" & _
							"		,RequiresSecretQuestion) SELECT " & " " & "iif(isnull(max([UserID])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Users'),max([UserID]) + 1)" & " " & " as UserID," & "'" & userName & "'" & " as UserName," & "'" & tmpPassword & "'" & " as [Password]" & _
							"       ," & "'" & otherNames & "'" & " as SurName" & _
							"       ," & "'" & otherNames & "'" & " as OtherNames" & _
							"		," & " " & isChangePass & " " & " as FirstTime" & _
							"       ," & "'Online user'" & " as Description" & _
							"       ," & "'" & isExpires & "'" & " as Expires" & _
							"		," & " 1 " & " as RemoteUser" & _
							"		," & " " & rs.Fields("Client_DPA_") & " " & " as Client_DPA_" & _
							"       ," & " 1 " & " as Enabled" & _							
							"       ," & " 0 " & " as RequiresSecretQuestion From Users"
			
        conn.BeginTrans
			conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))	
		conn.CommitTrans       	

		ClientEmail=rs("ClientEmail")

		Subject = "User login details"
		Body = "The following are your login details to Dyer and blair intranet" & vbCrLf
		Body = Body & "Username : " & userName & vbCrLf
		Body = Body & "Password : " & userName & vbCrLf
		Body = Body & "Website : www.dyerandblair.com" & vbCrLf
		
		'if(Isnull(ClientEmail) or ClientEmail="") then
		'else
		'SendMail ClientEmail, subject, bodyText
		'end if
	
		Set Conn = Nothing
		WritefraEnabledDialogCloseScript
		Response.End 

   	end If	

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
%>

<form name = 'frmEditUser' method = 'post' action = 'AddOnlineUser.asp' >
<table border="0" width="100%" cellpadding=2 cellspacing=2>
	<tr>
                <td width="40%">User Name</td>
                <td width="60%"><input type="text" name="txtUserName" id="txtUserName" size="25" Value=""></td>
              </tr>
 <tr>
                <td width="40%">First Name</td>
                <td width="60%"><input type="text" name="txtOtherNames" id="txtOtherNames" size="25" Value=""></td>
              </tr>
              <tr>
                <td width="40%">Surname</td>
                <td width="60%"><input type="text" name="txtSurName" id="txtSurName" size="25" Value=""></td>
              </tr>
              <tr>
                <td width="40%">Description</td>
                <td width="60%"><input type="text" name="txtDescription" id="txtDescription" size="25" Value=""></td>
              </tr> 
              <tr>
	 <td width="100%" colspan=2>	 
		<input type=checkbox name="chkExtra" Value="isExpires" Class="BorderLess" OnClick="JavaScript: swapExpires(this)" ID="chkExpires"><label for="chkExpires" style="cursor: hand">&nbsp;Account expires</label> 
		<SPAN ID="calExpiresDiv" Style="display: none">
  			<SCRIPT language="JavaScript">			
				var cal=new ctlSpiffyCalendarBox("cal", "frmEditUser", "txtEDate","cmdDate","<%= FormatDate(Date) %>",1);
				cal.writeControl();
			</SCRIPT>
		</SPAN>
		<P>		
		<input type=checkbox name="chkExtra" Value="isChangePass" Class="BorderLess" Checked ID="chkChangePass"><label for="chkChangePass" style="cursor: hand">&nbsp;User must change password at next logon</label> 
		<P>
		<input type=checkbox name="chkExtra" Value="isDisabled" Class="BorderLess" ID="chkDisabled"><label for="chkDisabled" style="cursor: hand">&nbsp;Account is disabled</label>		 
		<P>
		
				Client:
				<input type = 'text' name ='txtClientCode' id = 'txtClientCode' size="10" onBlur="txtval = this.value; selectItem(cboClient);">
				<select name = 'cboClient' id = "cboClient" size="1" 
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
					        
					        sqlStr = "SELECT * FROM FullClientList order by ClientName"
					        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
					        If Not (rs.EOF Or rs.BOF) Then
					                rs.MoveFirst
					                Do Until rs.EOF
					                ClientName=rs.Fields("ClientName")
					                NameClient=Mid(ClientName,1,30)
					                %>					                        
					                        <option SearchCode = "<%=rs.Fields("Client_DPA_")%>" SearchText = "<%=NameClient%>" value = '<%=rs.Fields("Client_DPA_")%>'><%=NameClient%></option>

					                        <%rs.MoveNext
					                Loop
					        End If
					%>

					    </select>
				
	 </td>
  </tr>  
  <tr>
     <td width="100%" colspan="2" align=right>
		<BR>
		<BR>
		<BR>
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save ">
		&nbsp;&nbsp;
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();">
		
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		<input type = 'hidden' name ='ID' id = 'ID' value="<%= ID %>">
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
