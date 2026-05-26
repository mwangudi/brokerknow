<!--#include file="libroutines.asp"-->

<Script Language="JavaScript">
	function launchBrokerKnow(){
		var winFeatures;
		var origWin = window.parent;	
		winFeatures = "fullscreen=no, location=no, menubar=no, status=no, toolbar=no, scrollbars=yes, resizable=yes, top=0, left=0, width=" + (screen.availWidth - 10) + ", height=" + (screen.availHeight - 30);
		var newWin = window.open("default.asp", "KNWNG", winFeatures);	
		newWin.opener = null;
		origWin.opener = null;
		origWin.self.close();	
		newWin.parent.focus();

	}	
</Script>
<%
Dim retUser, retPass, retVal

retUser = Request.Form("username")
retPass = Request.Form("password")
screenHeight = Request.Form("screenHeight")

retVal = AuthentiX(retUser, retPass)

If retVal <> "1" Then%>
	<SCRIPT LANGUAGE="JAVASCRIPT">
		window.alert("Invalid user credentials. Try again.")
		window.location.replace("default.htm")
	</SCRIPT>
	<%Response.End
Else
	%>
	<SCRIPT LANGUAGE="JAVASCRIPT">
		window.location.href = 'default.asp';
		//launchBrokerKnow();
	</SCRIPT>
	<%
	Response.End
	
	
	IF ENABLED THEN
		If LCase(trim(retUser)) = "jthiore" OR LCase(trim(retUser)) = "Cgathigi" OR LCase(trim(retUser)) = "webadmin" OR LCase(trim(retUser)) = "user" Then
			%>
			<SCRIPT LANGUAGE="JAVASCRIPT">
				window.location.href = 'default.asp';
				//launchBrokerKnow();
			</SCRIPT>
			<%
			Response.End
		Else
			%>
			<SCRIPT LANGUAGE="JAVASCRIPT">
				//window.location.href = 'default.asp';
				launchBrokerKnow();
			</SCRIPT>
			<%
			Response.End
		End If
	END IF
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
			
				if Rs.Fields("RemoteUser").Value = 1 then%>
					<Script Language="JavaScript">					
						alert("This account can only be used for remote access. Please contact your administrator.");
						window.location.replace("default.htm");
					</Script>
					<%Set Rs = Nothing
					Response.End	
				End If
				
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
					Session("Branch_DPA_") = Rs.Fields("Branch_DPA_")										
					Session("companyLogo") = companyLogo
					Session("Broker_DPA_") = Rs.Fields("Broker_DPA_").Value 					
					Session("BrokerName") = Rs.Fields("BrokerName").Value
					Session("BrokerCode") = Rs.Fields("BrokerCode").Value
					'Session("PinNo") = Rs.Fields("PinNo").Value
				End If
		 
				Set Rs = Nothing			
				
				'check for reset password
				If Abs(toResetPassword) = "1" Then%>
						<Script Language="JavaScript">											
							alert("You are required to change your password.");
							var targetPage = "Dialog.asp?titleDoc=Change Password&dialogDoc=Security/ChangePass.asp";
							var dWin = window.showModalDialog(targetPage, null, "dialogWidth:20em;dialogHeight:12em");
							if (dWin=="1") launchBrokerKnow();
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