<!--#include virtual="libroutines.asp"-->
<%

Session("UserID") = -1 'user does not exist yet, but should be 
						'allowed some system access to facilitate registration
	
	Dim action
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim guidStr 
	Dim guid
	
	action = ucase(Request.Form("action"))	
	
	if action = "EXECUTE" then
		Dim buttonAction
		Dim reloadRequired
		
		reloadRequired = false
		buttonAction = Trim(Ucase(Request.Form("cmdAdd")))
		if buttonAction = "SAVE" then
				Dim agent
				Dim classType
				Dim owner
				Dim branch
				Dim commission
				Dim residency
				Dim gender
				Dim addr
				Dim bDate
				Dim cell
				Dim email
				Dim contact
				Dim office
				Dim home
				Dim fax
				Dim idPass
				Dim photo
				Dim vip
				Dim gen1
				Dim gen2
				Dim pAddr
				dim comment
				Dim bankAccSpecified
				Dim OpeningBal
				Dim username
				Dim SecretQuestion
				Dim SecretAnswer, ConfirmSecretAnswer
				
				comment = Request.Form("txtComment")
				agent = Request.Form("cboAgent") 
				classType = Request.Form("cboClass")
				owner = Request.Form("cboOwner")
				commission = Request.Form("cboCommission")
				gender = Request.Form("cboGender")
				residency = Request.Form("cboResidency")
				name = Request.Form("txtName")
				addr = Request.Form("txtAddr")
				bDate = Request.Form("txtBDate")
				cell = Request.Form("txtCellTel")
				contact = Request.Form("txtContact")
				email = Request.Form("txtEmail")
				office = Request.Form("txtOfficeTel")
				home = Request.Form("txtHomeTel")
				fax = Request.Form("txtFax")
				idPass = Request.Form("txtIDPass")
				photo = Request.Form("txtPhoto")
				vip = Request.Form("cboVIP")
				gen1 = Request.Form("txtGeneric1")
				gen2 = Request.Form("txtGeneric2")
				pAddr = Request.Form("txtPAddr")
				OpeningBal = Request.Form("txtOpeningBal")
				username = Request.Form("txtUserName")
				SecretQuestion = Request.Form("txtSecretQuestion")
				SecretAnswer = Request.Form("txtSecretAnswer")
				ConfirmSecretAnswer = Request.Form("txtConfirmSecretAnswer")				
				
				
				If Trim(gender) = "" Then
					gender = "NULL"
				End If
							
				'validate Agent
				If Trim(Agent) = "" Then
					Agent = "NULL"
				End If
				
				'validate Owner
				If Trim(Owner) = "" Then
					Owner = "NULL"
				End If
				
				'validate Name
				If Trim(Name) = "" Then%>
					<script language = 'vbscript'>
                		ShowMessage "Please specify the Name"
                		window.history.back
					</script>
					<%					
					response.end
					
				End If
				
				'validate UserName
				If Trim(UserName) = "" Then%>
					<script language = 'vbscript'>
                		ShowMessage "Please specify the UserName"
                		window.history.back
					</script>
					<% response.end
				End If
				
				'validate size of Address
				If Len(Addr) > 500 Then%>
					<script language = 'vbscript'>
						ShowMessage "Address can only be 500 characters in length"
						window.history.back
					</script>
					<%response.end
				End If
				
				'validate size of Cell Phone
				If Len(Cell) > 100 Then%>
					<script language = 'vbscript'>
						ShowMessage "Cell Phone can only be 100 characters in length"
						window.history.back
					</script>
					<%response.end
				End If
				
				'validate size of Contact Name
				If Len(Contact) > 100 Then%>
					<script language = 'vbscript'>
						ShowMessage "Contact Name can only be 100 characters in length"
						window.history.back
					</script>
					<%response.end
				End If
				
				'validate size of Email
				If Len(Email) > 100 Then%>
					<script language = 'vbscript'>
						ShowMessage "Email can only be 100 characters in length"
						window.history.back
					</script>
					<% response.end
				End If
				
				'validate entry of secret question
				If Len(SecretQuestion) = 0 Then%>
					<script language = 'vbscript'>
						ShowMessage "You must enter a secret question"
						window.history.back
					</script>
					<%response.end
				End If
				
				'validate entry of secret answer
				If Len(SecretAnswer) = 0 Then%>
					<script language = 'vbscript'>
						ShowMessage "You must enter a secret answer"
						window.history.back
					</script>
					<%response.end
				End If
				
				'validate sameness of secret answers
				If SecretAnswer <> ConfirmSecretAnswer Then%>
					<script language = 'vbscript'>
						ShowMessage "The two secret answers must be the same"
						window.history.back
					</script>
					<%response.end
				End If
				
				OpeningBal = 0
				If Not IsNumeric(OpeningBal) Then%>
					<script language = 'vbscript'>
						ShowMessage "Opening Balance can only be numeric"						
						window.history.back
					</script>
					<% response.end
				End If
				
				'validate size of Fax
				If Len(Fax) > 100 Then%>
					<script language = 'vbscript'>
						ShowMessage "Fax can only be 100 characters in length"
						window.history.back
					</script>
					<% response.end
				End If
				
				'validate size of Generic1
				If Len(Gen1) > 100 Then%>
					<script language = 'vbscript'>
						ShowMessage "Generic1 can only be 100 characters in length"
						window.history.back
					</script>
					<% response.end
				End If
				
				'validate size of Generic2
				If Len(Gen2) > 100 Then%>
					<script language = 'vbscript'>
						ShowMessage "Generic2 can only be 100 characters in length"
						window.history.back
					</script>
					<% response.end
				End If
				
				'validate size of Physical Address
				If Len(pAddr) > 500 Then%>
					<script language = 'vbscript'>
					ShowMessage "Physical Address can only be 500 characters in length"
					window.history.back
					</script>
					<% response.end
				End If
				
				'validate size of Home Phone
				If Len(home) > 100 Then%>
					<script language = 'vbscript'>
					ShowMessage "Home Phone can only be 100 characters in length"
					window.history.back
					</script>
					<% response.end
				End If
				
				'validate size of ID or Passport
				If Len(IDPass) > 100 Then%>
					<script language = 'vbscript'>
					ShowMessage "ID or Passport can only be 100 characters in length"
					window.history.back
					</script>
					<% response.end
				End If
				
				'validate size of Name
				If Len(Name) > 100 Then%>
					<script language = 'vbscript'>
					ShowMessage "Name can only be 100 characters in length"
					window.history.back
					</script>
					<% response.end
				End If
				
				'validate size of Office Phone
				If Len(Office) > 100 Then%>
					<script language = 'vbscript'>
						ShowMessage "Office Phone can only be 100 characters in length"
						window.history.back
					</script>
					<% response.end
				End If
				
			
				Set conn = GetActiveConnection("KBroker")

				sqlStr="Select * From CompanyInfo"
				set rs=conn.execute(sqlStr)

				CompanyContacts=" CompanyName: " & rs("CompanyName") & Chr(10) & _
								", Address: " & rs("Address") & " (" & rs("PostalCode") & ")" & Chr(10) & _
								", Phone Number: " & rs("PhoneNumber") & Chr(10) & _
								", Fax Number: " & rs("FaxNumber") & Chr(10) & _
								", Email: Admin@dyerandblair.com" & rs("FaxNumber") & Chr(10) & _
								", City: " & rs("City") & "; Country: " & rs("Country")
	
				'save data
				If Photo <> "" Then
					photoNew = photo
				Else
					photoNew = "/Data/Photos/_blank.jpg"		
				End If
				
				set guid = server.createobject("NDUtils.CGUID")
				guidStr = guid.GenerateGUID	
				
				vip = 0
				
				sqlStr = "INSERT INTO [OnlineClient] (ClientAddr,ClientBDate,ClientCellTel,ClientContact,ClientEmail" & _
						",ClientFax,ClientComment,ClientPAddr" & _
						",ClientHomeTel,ClientIDPass,ClientName,ClientOfficeTel,ClientPhoto" & _
						",Client_DPA_,Gender_DPA_,Residency_DPA_,Commission_DPA_,ClientVIP,Class_DPA_, Branch_DPA_" & _
						",Client_EIT_,OnlineRegistration) SELECT " & "'" & addr & "'" & " as ClientAddr," & "#" & bDate & "#" & " as ClientBDate" & _
						"," & "'" & cell & "'" & " as ClientCellTel" & _
						"," & "'" & contact & "'" & " as ClientContact" & _
						"," & "'" & email & "'" & " as ClientEmail," & "'" & fax & "'" & " as ClientFax" & _
						"," & "'" & comment & "'" & " as ClientComment" & _
						"," & "'" & pAddr & "'" & " as ClientPAddr" & _
						"," & "'" & home & "'" & " as ClientHomeTel" & _
						"," & "'" & idPass & "'" & " as ClientIDPass," & "'" & name & "'" & " as ClientName" & _
						"," & "'" & Office & "'" & " as ClientOfficeTel" & _
						"," & "'" & photoNew & "'" & " as ClientPhoto" & _ 
						"," & " " & "iif(isnull(max([Client_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'OnlineClient'),max([Client_DPA_]) + 1)" & " " & " as Client_DPA_" & _
						"," & " " & gender & " " & " as Gender_DPA_" & _
						"," & " 0 " & " as Residency_DPA_" & _
						"," & " 0 " & " as Commission_DPA_" & _
						"," & " 0 " & " as ClientVIP" & _
						"," & " 0 " & " as Class_DPA_" & _
						"," & " 0 " & " as Branch_DPA_" & _
						"," & "'" & guidStr & "'" & " as Client_EIT_" & _
						"," & " 1 " & " as OnlineRegistration" & _
						" FROM [OnlineClient]"				
				
				conn.BeginTrans
						conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
						
						'obtain header key value
						sqlStr = "SELECT [Client_DPA_] FROM [OnlineClient] WHERE [Client_EIT_] = " & "'" & guidStr & "'"
	        
						Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
						If (rs.EOF Or rs.BOF) Then%>
								<script language = 'vbscript'>
						    			ShowMessage "A serious error has been encountered while saving the data. Try saving again"
						    			window.history.back
								</script>
								<% response.end
						End If
												
						'create a login account
						Dim userRS
						Dim password
						Dim isChangePass
						Dim isExpires
						Dim tmpUserName
						Dim tmpPassword
						
						sqlStr = "SELECT [UserID] FROM [Users] WHERE [UserName] = " & "'" & UserName & "'"
						Set userRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
						If Not(userRS.EOF Or userRS.BOF) Then%>
								<script language = 'vbscript'>
						    			ShowMessage "The username you specified is already in use"
								    	window.history.back		
								</script>
								<% response.end
						End If
						tmpUserName = userName
						password = EncryptWithALP(tmpUserName)
						tmpPassword = password
						isChangePass = 1
						isExpires = ""
						
						sqlStr = "INSERT INTO [Users] (UserID,UserName,[Password],SurName,OtherNames" & _
							"       ,FirstTime,Description,Expires,RemoteUser,Client_DPA_,Enabled" & _
							"		, SecretQuestion, SecretAnswer, RequiresSecretQuestion) SELECT " & " " & "iif(isnull(max([UserID])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Users'),max([UserID]) + 1)" & " " & " as UserID," & "'" & userName & "'" & " as UserName," & "'" & EncryptWithALP(tmpPassword) & "'" & " as [Password]" & _
							"       ," & "'" & Name & "'" & " as SurName" & _
							"       ," & "''" & " as OtherNames" & _
							"		," & " " & isChangePass & " " & " as FirstTime" & _
							"       ," & "'Online user'" & " as Description" & _
							"       ," & "'" & isExpires & "'" & " as Expires" & _
							"		," & " 1 " & " as RemoteUser" & _
							"		," & " " & rs.Fields("Client_DPA_") & " " & " as Client_DPA_" & _
							"       ," & " 1 " & " as Enabled" & _
							"       ," & "'" & SecretQuestion & "'" & " as SecretQuestion" & _
							"       ," & "'" & EncryptWithALP(SecretAnswer) & "'" & " as SecretAnswer" & _
							"       ," & " 1 " & " as RequiresSecretQuestion From Users"
						
						'Response.write(sqlStr)
						'Response.end

						conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
				conn.CommitTrans
				'conn.Close
				'Set conn = Nothing
				
				SQL = "SELECT * FROM MAilConfigList"
	 Set Rs = Conn.Execute(SQL)
	
	If Rs.EOF Or Rs.BOF Then%>
		<Script Language="JavaScript">
			alert("The mail configurations have not been set.");
			//window.parent.self.close();    
		</Script>
		<%Response.End 
	End If
				
	ClientEmail=email

	Const cdoSchema = "http://schemas.microsoft.com/cdo/configuration/" 
	Set objMsg = CreateObject("CDO.Message") 
	
	With objMsg	
		.Configuration.Fields.Item(cdoSchema & "sendusing") = Rs.Fields("SendUsingMethod").Value 
		.Configuration.Fields.Item(cdoSchema & "smtpserver") = Rs.Fields("SMTPServer").Value 
		.Configuration.Fields.Item(cdoSchema & "smtpserverport") = Rs.Fields("SMTPServerPort").Value
		.Configuration.Fields.Item(cdoSchema & "smtpauthenticate") = Rs.Fields("SMTPAuthenticate").Value 
		.Configuration.Fields.Item(cdoSchema & "sendusername") = Rs.Fields("SendUserName").Value
		.Configuration.Fields.Item(cdoSchema & "sendpassword") = Rs.Fields("SendPassword").Value
		.Configuration.Fields.Item(cdoSchema & "smtpconnectiontimeout") = Rs.Fields("SMTPConnectionTimeout").Value
		.Configuration.Fields.Update 	
		 		 
		.Subject = "Login Details"
		.Sender = Rs.Fields("SendDisplayName").Value 
		.To = ClientEmail 
		If cc <> "" Then
			.CC = cc
		End If
		
		If bcc <> "" Then
			.BCC = bcc
		End If
		
		bodyText="The information provided has been received and you shall be contacted by one of our customer service representatives. For further information please contact US: " & CompanyContacts

		If bodyText <> "" Then
			.TextBody = bodyText
		End If
		
		'.CreateMHTMLBody "file://" & pathTo
		.HTMLBody = bodyText
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
	conn.Close
	Set Conn = Nothing
	Set objMsg = Nothing
	'Set Fso = Nothing
	'Response.End
	
				%>
				<script language="javascript">
            	alert('The information provided has been received and you shall be contacted by one of our customer service representatives. For further information please contact US:<%=CompanyContacts%>');
            	window.location.replace('http://www.dyerandblair.com');
				</script>
				<%
				
				Response.End
				
			
				'add mail send code here (really necessary?!!)
				%>
				<head>
				<title>BrokerKnow© Online</title>
				</head>
				<body>
				<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="50%">
					<tr>
						<td width="100%" colspan="2"><b>BrokerKnow© Online</font></b></td>
					</tr>
					<tr>
						<td width="100%" colspan="2">&nbsp;</td>
					</tr>
					<tr>
						<td width="100%" colspan="2">You have been successfully registered.</td>
					</tr>
					<tr>
						<td width="23%">Username</td>
						<td width="77%"><b><%=username%></b></td>
					</tr>
					<tr>
						<td width="23%">Password</td>
						<td width="77%"><b><%=password%></b></td>
					</tr>
					<tr>
						<td width="100%" colspan="2">Click <a href="../WebDefault.htm">here</a> to go to the login page</td>
					</tr>
					<tr>
						<td width="100%" colspan="2">&nbsp;</td>
					</tr>
				</table>
				</body>
				<%
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
   	
   	Set conn = GetActiveConnection("KBroker")
%>

<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Client Registration Form</title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="../Data/CALENDAR/calendar.css">
<script language="JavaScript" src="../Data/CALENDAR/calendar.js"></script>
<!--END CALENDAR -->

<script language = 'javascript'>
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
		
		ExemptFromDefaultTabIndex = true;
		
		function DisplayDefaultCommission(theList)
		{
			var commissionList = document.frmMain.cboCommission
			commissionList.value = theList.options[theList.selectedIndex].DefaultCommission;
		}
</script>
</head>

<body >
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>

<SCRIPT language="JavaScript">
	var cal=new ctlSpiffyCalendarBox("cal", "frmAddClient", "txtBDate","cmdDate","<%= FormatDate(Date) %>",1);
</SCRIPT>
<form name = 'frmAddClient' method = 'post' action = 'WebAddClient.asp' id = "frmMain">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td width="100%" colspan="4"><b>BrokerKnow© Online</b></td>
	</tr>
	<tr>
		<td width="100%" colspan="4"><b>Client Registration Form</b></td>
	</tr>
	<tr>
		<td width="100%" colspan="4">&nbsp;</td>
	</tr>
	<tr>
		<td width="20%">Username</td>
		<td width="30%" colspan="2"><input type = 'text' name ='txtUserName' STYLE="WIDTH: 200px" tabIndex='8' id = 'txtUserName'></td>
		<td width="50%" colspan="2" rowspan="7">
		<input type="hidden" name="txtPhoto">
		<IFRAME FRAMEBORDER=0 SCROLLING=NO SRC="Webupload.asp?filetext=txtPhoto" width="170px" height="170px" tabIndex='16'></IFRAME>
		</td>
	</tr>
	
	<tr>
		<td width="20%">Name</td>
		<td width="30%"><input type = 'text' name ='txtName' STYLE="WIDTH: 200px" tabIndex='8' id = 'txtName'></td>
	</tr>
	<tr>
		<td width="20%">Contact Name</td>
		<td width="30%"><input type = 'text' name ='txtContact' STYLE="WIDTH: 200px" tabIndex='9' id = 'txtContact' size="20"></td>
	</tr>
	<tr>
		<td width="20%">Cell Phone</td>
		<td width="30%"><input type = 'text' name ='txtCellTel' STYLE="WIDTH: 200px" tabIndex='10' id = 'txtCellTel' size="20"></td>
	</tr>
	<tr>
		<td width="20%">Home Phone</td>
		<td width="30%"><input type = 'text' name ='txtHomeTel' STYLE="WIDTH: 200px" tabIndex='11' id = 'txtHomeTel' size="20"></td>
	</tr>
	<tr>
		<td width="20%">Office Phone</td>
		<td width="30%"><input type = 'text' STYLE="WIDTH: 200px" name ='txtOfficeTel' tabIndex='12' id = 'txtOfficeTel' size="20"></td>
	</tr>
	<tr>
		<td width="20%">Email</td>
		<td width="30%"><input type = 'text' STYLE="WIDTH: 200px" name ='txtEmail' STYLE="WIDTH: 220px" tabIndex='13' id = 'txtEmail' size="20"></td>
	</tr>  
	<tr>
		<td width="20%">Fax</td>
		<td width="30%"><input type = 'text' STYLE="WIDTH: 200px" name ='txtFax' tabIndex='14' id = 'txtFax' size="20"></td>
		<td width="20%" align=right nowrap>ID/Passport&nbsp;</td>
		<td width="30%"><input type = 'text' STYLE="WIDTH: 170px" name ='txtIDPass' tabIndex='17' id = 'txtIDPass' size="20"></td>
	</tr>
	<tr>
		<td width="20%">Gender</td>
		<td width="30%">
		<select name = 'cboGender' tabIndex='4' id = 'cboGender' size="1" STYLE="WIDTH: 200px">
    		<option selected value = ''></option>
			<%
			sqlStr = "SELECT * FROM [GenderList]"
			Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
			If Not (rs.EOF Or rs.BOF) Then
				rs.MoveFirst
				Do Until rs.EOF
				    if cbool(rs.Fields("DefaultSelection")) then%>
							<option selected value = '<%=rs.Fields("Gender_DPA_")%>'><%=rs.Fields("GenderGender")%></option>
				    <%else%>
							<option value = '<%=rs.Fields("Gender_DPA_")%>'><%=rs.Fields("GenderGender")%></option>
				    <%end if
				    rs.MoveNext
				Loop
			End If
			%>
		</select>
		</td>
		<td width="20%" align=right>&nbsp;&nbsp;Date of Birth&nbsp;</td>
		<td width="30%"><SCRIPT language="JavaScript">cal.writeControl();  document.all.item('txtBDate').tabIndex='18'</SCRIPT>&nbsp;</td>
	</tr>
	<tr>
		<td width="20%"  valign="top">Address</td>
		<td width="30%" ><textarea rows=3 name ='txtAddr' tabIndex='16' id = 'txtAddr' size="35" cols="33" STYLE="WIDTH: 200px"></textarea></td>
		<td width="20%" align=right valign="top">Physical Address &nbsp;&nbsp;&nbsp;</td>
		<td width="30%"><textarea rows=3 name ='txtPAddr' tabIndex='22' id = 'txtPAddr' size="35" cols="25" STYLE="WIDTH: 170px"></textarea></td>
		<input type = 'hidden' name ='txtGeneric2' id = 'txtGeneric2' size="20">
	</tr>   
	<tr>
		<td width="20%">Secret Question</td>
		<td width="30%"><input type = 'text' STYLE="WIDTH: 200px" name ='txtSecretQuestion' tabIndex='15' id = 'txtSecretQuestion' size="20"></td>
		<td width="20%" align='right' nowrap>Secret Answer&nbsp;</td>
		<td width="30%"><input type = 'password' STYLE="WIDTH: 170px" name ='txtSecretAnswer' tabIndex='18' id = 'txtSecretAnswer' size="20"></td>
	</tr>
	<tr>
		<td width="50%" colspan="2">&nbsp;</td>
		<td width="20%" align='right' nowrap>Confirm Secret Answer&nbsp;</td>
		<td width="30%"><input type = 'password' STYLE="WIDTH: 170px" name ='txtConfirmSecretAnswer' tabIndex='18' id = 'txtConfirmSecretAnswer' size="20"></td>
	</tr>
</table>
<table align=right>
	<tr>
		<td width="100%" align="right" valign=absBottom >
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save " tabIndex='25' onclick = "AllowedNavigation()">
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.location.replace('../webdefault.htm')">
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		</td>
	</tr>	
</table>

</form>
</body>

</html>
