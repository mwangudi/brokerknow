<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add User</title>
 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>

<!--CALENDAR -->
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
 <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css"> 
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT>
<script language="JavaScript" src="CALENDAR/calendar.js"></script>
<SCRIPT language=Javascript src="../scripts/reports.js"></SCRIPT>

<!--END CALENDAR -->

<script language='javascript'>
	function forceSubmit()
	{
		//alert(window.parent.opener.name);
		document.frmAddUser.target="_self";
		document.frmAddUser.submit();	
	}
	function setOpener()
	{
		//alert ()
		window.self.opener = window.dialogArguments.opener;
	}
</script>
</head>

<body Class="Dialog" onload="setOpener()">

<!--#include file="../libroutines.asp"-->
<div id="spiffycalendar" class="text"></div>

<%
	
   Dim action
   Dim conn 
   Dim sqlStr
   Dim rs
   Dim guidStr 
   Dim guid 
	
	action = ucase(Request.Form("action"))
	
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
       
       otherNames = Request.Form("txtOtherNames")
       userName = Request.Form("txtUserName")
	   surName = Request.Form("txtSurName")
       Description = Request.Form("txtDescription")
       passWord = Request.Form("txtPassword")
       passWordConfirm = Request.Form("txtPasswordConfirm")
       chkExtra =  Request.Form("chkExtra")
       expiresDate =  Request.Form("txtEDate")    
       client =  "NULL"
       
        'validate userName
        If Trim(userName) = "" Then%>
				<script language="vbscript">
					ShowMessage "Please specify the User name"                						
			</script>
			<script language="javascript">					
					window.location="AddUser.asp?id=<%=ID%>";					
			</script>
				<% 
				 reloadpage(ID)
				response.end
        End If
        
        'validate otherNames
        If Trim(otherNames) = "" Then%>
				<script language="vbscript">
					ShowMessage "Please specify the user's first name"
					
			</script>
				<% 
				 reloadpage(ID)
				response.end
        End If
       
        'validate size of userName
        If Len(userName) <= 3 Then%>
			<script language="vbscript">
							ShowMessage "User name should be more than 3 characters in length"
							
							</script>
			<% 
			 reloadpage(ID)
			response.end
        End If
     	'validate passWord
        If Trim(passWord) = "" Then%>
				<script language="vbscript">
                						ShowMessage "Please specify the Password"
                						
								</script>
				<% 
				 reloadpage(ID)
				response.end
        End If
         'validate size of passWord
        If Len(passWord) <= 3 Then%>
			<script language="vbscript">
							ShowMessage "Password should be more than 3 characters in length"
							
							</script>
			<% 
			 reloadpage(ID)
			response.end
        End If       
        
        'validate passWord
        If passWord <> passWordConfirm Then%>
				<script language="vbscript">
                						ShowMessage "The passwords do not match. Please retry."
                						
								</script>
				<% 
				 reloadpage(ID)
				response.end
        End If
        
        isExpires = ""
		isChangePass = 0
		isEnabled = 1  
		isRemoteUser = 0    
		  
        If chkExtra <> "" Then
			Dim chkExtraArray, i
			chkExtraArray = Split(chkExtra, ",")
			For i = 0 To UBound(chkExtraArray)
				Select Case LCase(Trim(chkExtraArray(i)))					
					Case "isexpires"
						isExpires = Request.Form("txtEDate")
						If DateDiff("d", FormatDate(Date), FormatDate(isExpires)) < 0 Then%>
							<script language="vbscript">
                						ShowMessage "The date of expiry is set in the past."                						
								</script>
						<%	
						 reloadpage(ID)
						Response.End
						End If
					Case "ischangepass"
						isChangePass = 1
					Case "isdisabled"
						isEnabled = 0	
					Case "isremoteuser"	
						isRemoteUser = 1
						client =  Request.Form("cboClient")			
				End Select
			Next
        End If
		
		Set conn = GetActiveConnection("KBroker")
		
		'check whether another similar username exists
		Set chkRs = Conn.Execute("SELECT UserName FROM [Users] WHERE UserName = '" & userName & "'")
		If Not (chkRs.EOF Or chkRs.BOF) Then%>
			<script language="vbscript">
            		ShowMessage "A similar username already exists."
			</script>
			<% Set chkRs = Nothing
			
			 reloadpage(ID)
			Response.End
		End If
		
		Set chkRs = Nothing
		 
        sqlStr = "INSERT INTO [Users] (UserName,[Password],SurName,OtherNames" & _
                "       ,FirstTime,Description,Expires,RemoteUser,Client_DPA_,Enabled) SELECT " & "'" & userName & "'" & " as UserName," & "'" & EncryptWithALP(Password) & "'" & " as [Password]" & _
                "       ," & "'" & surName & "'" & " as SurName" & _
                "       ," & "'" & otherNames & "'" & " as OtherNames" & _
                "		," & " " & isChangePass & " " & " as FirstTime" & _
                "       ," & "'" & Description & "'" & " as Description" & _
                "       ," & "'" & isExpires & "'" & " as Expires" & _
                "		," & " " & isRemoteUser & " " & " as RemoteUser" & _
                "		," & " " & client & " " & " as Client_DPA_" & _
                "       ," & " " & isEnabled & " " & " as Enabled"
              
        conn.BeginTrans
			conn.Execute sqlStr	      
		conn.CommitTrans
        Set Conn = Nothing		
        WritefraEnabledDialogCloseScript2
        
		 reloadpage(ID)
		Response.End 
   	end If
%>

<form name = 'frmAddUser' method = 'post' action = 'AddUser.asp' >
<table border="0" width="100%" cellpadding=2 cellspacing=2>
	<tr>
                <td width="40%">User Name</td>
                <td width="60%"><input type="text" name="txtUserName" id="txtUserName" size="25"></td>
              </tr>
 <tr>
                <td width="40%">First Name</td>
                <td width="60%"><input type="text" name="txtOtherNames" id="txtOtherNames" size="25"></td>
              </tr>
              <tr>
                <td width="40%">Surname</td>
                <td width="60%"><input type="text" name="txtSurName" id="txtSurName" size="25"></td>
              </tr>
              <tr>
                <td width="40%">Description</td>
                <td width="60%"><input type="text" name="txtDescription" id="txtDescription" size="25"></td>
              </tr>
              <tr>
				<td colspan="2" width="100%"><HR></td>
              </tr>
              <tr>
                <td width="40%">Password</td>
                <td width="60%"><input type="Password" name="txtPassword" id="txtPassword" size="25"></td>
              </tr>
              <tr>
                <td width="40%">Confirm Password</td>
                <td width="60%"><input type="Password" name="txtPasswordConfirm" id="txtPasswordConfirm" size="25"></td>
              </tr>
              <tr>
			<tr>
				<td colspan="2" width="100%"><HR></td>
              </tr>    
     <tr>
	 <td width="100%" colspan=2>	 
		<input type=checkbox name="chkExtra" Value="isExpires" Class="BorderLess" OnClick="JavaScript: swapExpires(this)" ID="chkExpires"><label for="chkExpires" style="cursor: hand">&nbsp;Account expires</label> 
		<SPAN ID="calExpiresDiv" Style="display: none">
  			<SCRIPT language="JavaScript">			
				var cal=new ctlSpiffyCalendarBox("cal", "frmAddUser", "txtEDate","cmdDate","<%= FormatDate(Date) %>",1);
				cal.writeControl();
			</SCRIPT>
		</SPAN>
		<P>		
		<input type=checkbox name="chkExtra" Value="isChangePass" Class="BorderLess" Checked ID="chkChangePass"><label for="chkChangePass" style="cursor: hand">&nbsp;User must change password at next logon</label> 
		<P>
		<input type=checkbox name="chkExtra" Value="isDisabled" Class="BorderLess" ID="chkDisabled"><label for="chkDisabled" style="cursor: hand">&nbsp;Account is disabled</label>
		<P>
		<input type=checkbox name="chkExtra" Value="isRemoteUser" Class="BorderLess" ID="chkRemoteUser" OnClick="JavaScript: switchDisplay (document.all.item('cboClient')); "><label for="chkRemoteUser" style="cursor: hand">&nbsp;Remote User</label>  
		<select name = 'cboClient' id = 'cboClient' size="1" style="display:none">
   <%
        Set conn = GetActiveConnection("KBroker")
        sqlStr = "SELECT * FROM [ClientList]"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF%>
                        <option value = '<%=rs.Fields("Client_DPA_")%>'><%=rs.Fields("ClientName")%></option>
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
		<input type = 'button' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save " onClick="forceSubmit()">
		&nbsp;&nbsp;
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();">
		
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
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
