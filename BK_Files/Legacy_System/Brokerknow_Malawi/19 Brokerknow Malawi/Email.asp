<html>

<head>
	<meta http-equiv="Content-Language" content="en-uk">
	<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
	<title>Send Page as Email</title>
	<LINK href="STYLE/default.css" type=TEXT/CSS rel=STYLESHEET> 
	<LINK href="STYLE/webparts.css" type=TEXT/CSS rel=STYLESHEET>
	<SCRIPT language=Javascript src="scripts/common.js"></SCRIPT>
</head>

<body Class="Dialog">

<!--#include file="libroutines.asp"-->

<%

enterDetailsPointer = Request("enterDetailsPointer")

'get HTML output
htmlOut = Request.Form("emailDoc")

'include body text
bodyText = trim(Request("bodyText"))
htmlOut = bodyText & "<BR><BR><BR>" & htmlOut

'get path of referrer, as this is where the temp HTML page ought to be saved
fromPath = Request.Form("emailDocSourcePath")

'get the name of the report
reportName = Request.Form ("emailDocName")

If enterDetailsPointer = "" Or enterDetailsPointer = "0" Then
	toRecipient = Request.Form("To")
	cc = Request.Form("cc")
	bcc = Request.Form("bcc")
	subject = Request.Form("subject")
	
	bodyText = trim(Request.Form("bodytext"))
	attachment = Request.Form("attachment")
	
	If htmlOut = "" Or fromPath = "" Or toRecipient = ""  Or subject = ""  Then%>
		<Script Language="JavaScript">
			alert("This page <%= subject %> does not have enough information to proceed.");
			//window.parent.self.close();
		</Script>
		<%	   
		Response.End
	End If

	if instr(1,attachment,".pdf",1)=0 then

	%>
		<Script Language="JavaScript">
			alert("Please  a valid file pdf");
			//window.parent.self.close();
		</Script>

	<%

	response.end
	end if
	
   

	'write to output file
	Dim FSO, fTextStream, pathTo, tmpMark, pageFrom
	
	Randomize
	
	tmpMark = "-" & Rnd(2) & "-PNW"

	pathTo = fromPath & "\TMPReport" &  tmpMark & ".html"
	
	Set FSO = CreateObject("Scripting.FileSystemObject")

	Set fTextStream = Fso.CreateTextFile(pathTo, True, False)
	fTextStream.Write htmlOut
	fTextStream.Close
	Set fTextStream = Nothing
	
	
	Set Conn = GetActiveConnection("KBroker")
	SQL = "SELECT * FROM MAilConfigList"
	Set Rs = Conn.Execute(SQL)
	
	If Rs.EOF Or Rs.BOF Then%>
		<Script Language="JavaScript">
			alert("The mail configurations have not been set.");
			//window.parent.self.close();    
		</Script>
		<%Response.End 
	End If
				
	
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
		.Subject = subject
		'.Sender = Rs.Fields("SendDisplayName").Value 
		.Sender = "securities@africanalliance.co.ke"
		.AddAttachment attachment
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
		
		'.CreateMHTMLBody "file://" & pathTo
		'.CreateMHTMLBody "file://C:\Knownig\Brokerknow\Email\blank.htm" 
		'.CreateMHTMLBody "D:\08 - Brokerknow KENYA\E-Mailer\blank.htm"

		'.CreateMHTMLBody ""
		On Error Resume Next
		.Send 
		
		If Err.number > 0 Then%>
			<Script Language="JavaScript">
				alert("The page could not be sent. An unexpected error occured: <%= Err.description %>")
			</Script>
			<%Response.End 

		else 

		Insert toRecipient,cc, bcc, Subject,attachment,bodyText
		End If
	End With
	
	FSO.DeleteFile pathTo, True	
	
	Set Rs = Nothing
	Set Conn = Nothing
	Set objMsg = Nothing
	Set Fso = Nothing%>
		<Script Language="JavaScript">
			alert("The page was sent successfully.")		
		</Script>
	<%	
	Response.End 	
Else
	
%>

<Script Language="JavaScript">
	function DoSendMail()
	{
		var to = document.all.item('to');
		var cc = document.all.item('cc');
		var bcc = document.all.item('bcc');
		var subject = document.all.item('subject');
		var bodytext = document.all.item('bodytext');
		var attachment = document.all.item('attachment');
		
		if (to.value=='') 
		{
			alert('Ensure that the recipient address is entered in the "To" field.')
			to.focus();
			return;
		}
		
		if (subject.value=='') 
		{
			alert('Ensure that the mail subject is entered.');
			subject.focus();			
			return;
		}
		
		
		window.parent.dialogArguments.document.all.item('to').value = to.value;
		window.parent.dialogArguments.document.all.item('cc').value = cc.value;
		window.parent.dialogArguments.document.all.item('bcc').value = bcc.value;
		window.parent.dialogArguments.document.all.item('subject').value = subject.value;
		window.parent.dialogArguments.document.all.item('bodytext').value = bodytext.value;	
		window.parent.dialogArguments.document.all.item('attachment').value = attachment.value;	
		
		window.parent.dialogArguments.emailDoc();
		window.parent.self.close();   
		
		
		
	
	}
</Script>
<form method="POST" action="sendEmail">
	<input type=hidden name="sendPointer" value=1>
	<input type=hidden name="filePath" value="<%= pathTo %>">
	
 <table border="0" width="100%">
    <tr>
      <td width="21%"><font face="Tahoma" size="2" color="#000080"><b>To:</b></font></td>
      <td width="79%"><input type="text" name="To" size="40"></td>
    </tr>
    <tr>
      <td width="21%"><font face="Tahoma" size="2" color="#000080"><b>cc:</b></font></td>
      <td width="79%"><input type="text" name="cc" size="40"></td>
    </tr>
    <tr>
      <td width="21%"><font face="Tahoma" size="2" color="#000080"><b>bcc:</b></font></td>
      <td width="79%"><input type="text" name="bcc"  size="40"></td>
    </tr>
    <tr>
      <td width="21%"><font face="Tahoma" size="2" color="#000080"><b>Subject</b></font></td>
      <td width="79%"><input type="text" name="subject" size	="40"></td>
    </tr>
	<tr>
      <td width="21%"><font face="Tahoma" size="2" color="#000080"><b>Attachment</b></font></td>
      <td width="79%"><input type="file" name="attachment"  id ="attachment" size="40"></td>
    </tr>
    <tr>
      <td width="21%"><font face="Tahoma" size="2" color="#000080"><b>Body Text
        (Optional)</b></font></td>
      <td width="79%"><textarea rows="7" name="bodyText" cols="40"></textarea></td>
    </tr>
  </table>
  <p align="right"><input type="button" class="Buttons" value=" Send " name="B1" OnClick="JavaScript: DoSendMail()">&nbsp;&nbsp;&nbsp;<input type="button" class="Buttons" value=" Close " name="B2" OnClick="JavaScript: window.self.close();"></p>
</form>


<%End If


'function to record the send emails

function Insert(EmailTo,EmailCC, EmailBCC, Subject,attachment,bodyText)
	Set Conn = GetActiveConnection("KBroker")
	set rsAdd = server.createobject("adodb.recordset")
	rsAdd.Open "SELECT * FROM EmailedDocs", conn,3,2
	rsAdd.Addnew
		rsAdd("EmailTo")=EmailTo
		rsAdd("EmailCC")=EmailCC
		rsAdd("EmailBCC")=EmailBCC
		rsAdd("Subject")=Subject
		rsAdd("Message")=bodyText
		rsAdd("DocumentName")=attachment
		rsAdd("EmailedTime")=now()
		rsAdd("EmailedBy")=session("UserId")
		rsAdd("Emailed")=1
	rsAdd.Update

rsAdd.close
conn.close
set conn= nothing
set rsAdd= nothing

end function 

%>

</body>

</html>
