<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Change Password</title>
 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>

<!--CALENDAR -->
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
 <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css"> 
  <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT>

</head>

<body Class="Dialog">

<!--#include file="../libroutines.asp"-->

<%
	
   Dim action
   Dim conn 
   Dim sqlStr
   Dim rs
   Dim guidStr 
   Dim guid 
	
	action = ucase(Request.Form("action"))
	
	if action = "EXECUTE" then
	  
       passWord = Request.Form("txtPassword")
       passWordConfirm = Request.Form("txtPasswordConfirm")
       
         'validate size of passWord
        If Len(passWord) <= 3 Then%>
			<script language="vbscript">
							alert "Password should be more than 3 characters in length"
							
							</script>
			<% response.end
        End If       
        
        'validate passWord
        If passWord <> passWordConfirm Then%>
				<script language="vbscript">
                						alert "The passwords do not match. Please retry."
                						
								</script>
				<% response.end
        End If
        
       
		Set conn = GetActiveConnection("KBroker")
		sqlStr = "UPDATE [Users] SET [Password] = '" & EncryptWithALP(Password) & "', FirstTime = 0 WHERE UserID = " & Session("UserID")
		
        conn.BeginTrans
			conn.Execute sqlStr	      
		conn.CommitTrans
        Set Conn = Nothing%>
        <Script Language="JavaScript">
			//flag success
			window.parent.returnValue = "1";
			//close dialog window
			window.parent.close();
        </Script>
        <%
        Response.End 
   	end If
%>
<Div style="display: none">
	<IFRAME marginwidth="0" marginheight="0" FRAMEBORDER=0 SRC="" ID="hiddenFrame" NAME="deleteFrame" TAG=""></IFRAME>	
</Div>
<form name = 'frmChangePass' method = 'post' action = 'ChangePass.asp' >
<table border="0" width="100%" cellpadding=2 cellspacing=2>
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
     <td width="100%" colspan="2" align=right>
		<BR>
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save ">
		&nbsp;&nbsp;
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();">
		
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
      </td>
  </tr>
</table>
</form>

<Script Language="JavaScript">
	var timeOutID;
	
	function cancelChangePassword(){		
		window.clearTimeout(timeOutID);
		window.parent.returnValue = "0";
		window.parent.close();			
	}
	
	//give the user a minute max to change the password
	timeOutID = window.setTimeout("cancelChangePassword()", 60000);
		
</Script>

</body>

</html>
