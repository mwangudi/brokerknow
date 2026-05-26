<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit Mail Configuration</title>
   <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 

 <script language='javascript'>
		function forceSubmit()
		{
			setOpener();
			//var targetPage = window.dialogArguments.opener.document.all.item("frmMain").elements("EditPage").value;
					
			document.frmEditMailConfiguration.method='post';
			document.frmEditMailConfiguration.target='_self';
			document.frmEditMailConfiguration.submit();		
		}
		
		function setOpener()
		{
			window.self.opener = window.dialogArguments.opener;					
		}
 </script>
</head>

<body Class="Dialog"  onload="setOpener()">

<!--#include file="../libroutines.asp"-->


<%
	
	Dim action
	Dim conn 
	Dim rs

	Set conn = GetActiveConnection("KBroker")
	Set Rs = CreateObject("ADODB.Recordset")
	Rs.Open "SELECT * FROM MailConfiguration", Conn, 2, 3
	
	
	If Request.Form("submitme") <> "" Then
		
			For i = 0 To Rs.Fields.Count -1 
				thisFldValue = Trim(Request.Form("cntrl" & Rs.Fields(i).Name))
				If thisFldValue <> "" Then
					Rs.Fields(i).Value =  thisFldValue
				End If
			Next
		
			Rs.Update 
		
			Set Rs = Nothing
			Set Conn = Nothing
			
			WriteFraEnabledDialogCloseScript
			Response.End
	
	End If 
		
	If Not (Rs.EOF or Rs.BOF) Then
		SMTPIPAddress = Rs.Fields("SMTPServer").Value
		SMTPIPPort = Rs.Fields("SMTPServerPort").Value 
		SMTPTimeout = Rs.Fields("SMTPConnectionTimeout").Value 
		SMTPAuthMethod = Rs.Fields("SMTPAuthenticate").Value  
		SMTPUserName = Rs.Fields("SendUserName").Value 
		SMTPPassword = Rs.Fields("SendPassword").Value 
		SMTPDisplayName = Rs.Fields("SendDisplayName").Value 
		SMTPSendMethod = Rs.Fields("SendUsingMethod").Value 
		 
	End If

	With Response
			.write "<form name='frmEditMailConfiguration'  method='POST' action='EditMailConfiguration.asp' ACTION='_self'>" & vbCrLf
			.write "	<table height='196'><tr><td height='25'>SMTP Server (IP Address)</td><td height='25'><input type='text' value=""" &  SMTPIPAddress & """ name ='cntrlSMTPServer' size='37' ></td></tr><tr><td height='25'>SMTP" & vbCrLf
			.write "			        Server Port (Mail)</td><td height='25'>" & vbCrLf
			.write "			<input type='text' value=""" &  SMTPIPPort & """ name ='cntrlSMTPServerPort' size='11'  ></td></tr><tr><td height='25'>Connection        Timeout (secs)</td><td height='25'>" & vbCrLf
			.write "			        <input type='text'  value=""" &  SMTPTimeout & """ name ='cntrlSMTPConnectionTimeout'  size='11' ></td></tr><tr><td height='25'>SMTP" & vbCrLf
			.write "			        Server Authentication Protocol</td><td height='25'><select size='1' name='cntrlSMTPAuthenticate' >" & vbCrLf
			.write "			          <option id=0 value='0'>Anonymous</option>" & vbCrLf
			.write "			          <option id=1 value='1'>Basic</option>" & vbCrLf
			.write "			          <option id=2 value='2'>NTLM</option>" & vbCrLf
			.write "			        </select></td></tr><tr><td height='16'>Username</td><td height='16'><input type= 'text' value=""" &  SMTPUserName & """ name ='cntrlSendUserName' size='37' ></tr>" & vbCrLf
			.write "		<tr>" & vbCrLf
			.write "			      <td id='lbSupCode' height='25'>Password</td><td height='25'><input type= 'password' value=""" &  SMTPPassword & """ name='cntrlSendPassword' size='37' ></td>" & vbCrLf
			.write "			    </tr>" & vbCrLf
			.write "			    <tr><td id='lbSupCode' height='25'>Mail Message Display Name (with mail" & vbCrLf
			.write "			        address)</td><td height='25'><input type= 'text' name ='cntrlSendDisplayName' value=""" &  SMTPDisplayName & """ size='37' ></td></tr><tr><td id='lbSupCode' height='25'>Mail" & vbCrLf
			.write "			        Sending Method</td><td height='25'><select size='1' name='cntrlSendUsingMethod' >" & vbCrLf
			.write "					  <option value='1'>Send Using Pick Up Directory</option>" & vbCrLf
			.write "			          <option value='2'>Send Using Port</option>          " & vbCrLf
			.write "			        </select></td></tr><tr><td colspan='2' align='right' height='27'><input type='button' class='buttons' value=' Save ' name='submitme'   OnClick='Javascript: forceSubmit();'>&nbsp&nbsp;" & vbCrLf
			.Write "				<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=' Cancel ' OnClick='JavaScript: window.self.close();'></td></tr></table>" & vbCrLf
			.write "			<input type = 'hidden' name ='buttonAction' id = 'buttonAction' value='Save'>	</form>"
	End With
	
	Set Rs = Nothing  
	Set Conn = Nothing
	
	%>
	
<Script Language="JavaScript">	
	document.all.item("cntrlSMTPAuthenticate").selectedIndex = "<%=SMTPAuthMethod%>";
	document.all.item("cntrlSendUsingMethod").selectedIndex = "<%=SMTPSendMethod - 1%>";	
</SCRIPT>

</body>

</html>
