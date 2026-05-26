<!--#include file="libroutines.asp"-->


<%

Dim retUser, retPass, retVal

retUser = Request.Form("username")
retPass = Request.Form("password")
screenHeight = Request.Form("screenHeight")

retVal = AuthentiX(retUser, retPass)

If retVal <> "1" Then%>
	<SCRIPT LANGUAGE="JAVASCRIPT">
		window.alert("Invalid user credentials. Try again.")
		window.location.replace("Webdefault.htm")
	</SCRIPT>
	<%Response.End
Else
	%>
			<script language = 'javascript'>
					window.location.replace("WebDefault.asp");
										    			
			</script>
	<%
	Response.End
End If

Function AuthentiX(username, password)    
	Dim passedPass
       
    AuthentiX = 0  
    
    If username <> "" And password <> "" Then
        
        'no authentication here!   
        Set conn = GetActiveConnectionEx("KBroker")
        
        Set Rs = CreateObject("ADODB.Recordset")        
        SQL = "SELECT * FROM Users"
        Rs.CursorLocation = adUseClient
        Rs.Open SQL, conn.ConnectionString, 1, 3
        
        Rs.Filter = "Username = '" & username & "'"
        
        If Not (Rs.BOF Or Rs.EOF) Then
			passedPass = DecryptWithALP(Rs.Fields("Password").Value)
			
			If passedPass = password Then
			
				isEnabled = Rs.Fields("Enabled").Value
				isExpires = Rs.Fields("Expires").Value
				toResetPassword = Rs.Fields("FirstTime").Value
			
				'check for enabled state
				If Abs(isEnabled) <> "1" Then%>
					<Script Language="JavaScript">					
						alert("This account is disabled. Please contact your administrator.");
						window.location.replace("default.htm");
					</Script>
					<%Set Rs = Nothing
					Response.End	
				End If
			
				'check for expiry state
				If IsDate(isExpires) Then
					If DateDiff("d", FormatDate(Date), FormatDate(isExpires)) < 0 Then %>
						<Script Language="JavaScript">					
							alert("This account has expired. Please contact your administrator.");
							window.location.replace("default.htm");
						</Script>
						<%Set Rs = Nothing
					Response.End	
					End If	
				End If
			
				AuthentiX = 1         
				Session("Client_DPA_") = Rs.Fields("Client_DPA_").Value
				Session("UserID") = Rs.Fields("UserID").Value
				Session("CurrentUser") = Rs.Fields("OtherNames").Value & " " & Rs.Fields("Surname").Value
				Session("screenHeight") = screenHeight		
				
				'set company info into session variables
				 Set Rs = Conn.Execute("SELECT * FROM CompanyInfoView")
		 
				If Not (Rs.EOF Or Rs.BOF) Then
					companyLogo = Rs.Fields("Photo")
					companyName = Rs.Fields("CompanyName")
					
					'set company name session variable
					Session("CompanyName") = companyName
					Session("Branch") = Rs.Fields("Branch")					
					Session("companyLogo") = companyLogo
					Session("Broker_DPA_") = Rs.Fields("Broker_DPA_").Value 					
				End If
		 
				Set Rs = Nothing			
				
				'check for reset password
				If Abs(toResetPassword) = "1" Then%>
						<Script Language="JavaScript">											
							alert("You are required to change your password.");
							var targetPage = "Dialog.asp?titleDoc=Change Password&dialogDoc=Security/ChangePass.asp";
							var dWin = window.showModalDialog(targetPage, null, "dialogWidth:20em;dialogHeight:12em");
							if (dWin=="1") window.location.replace("Webdefault.asp");
							else if (dWin=="0") alert("The change password process was timed out.");			
						</Script>
						<%Response.End
				End If
			Else
				AuthentiX = 0
				Set Rs = Nothing
				Exit Function
			End If
			
            Exit Function
        Else
            AuthentiX = 0
            Set Rs = Nothing
            Exit Function
        End If        
        
    Else
        AuthentiX = 0
    End If    

End Function


%>
