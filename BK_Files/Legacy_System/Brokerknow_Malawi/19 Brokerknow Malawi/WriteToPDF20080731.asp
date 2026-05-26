<%
'get HTML output
htmlOut = Request.Form("PDFDoc")
theOrientation = Trim(Request.Form("PDFOrient"))

'Response.Write theOrientation
'Response.End 

If theOrientation <> "L" Then
	''PORTRAIT
	If Instr(1,htmlOut,"CLIENTS STATEMENT") > 0 Then
		str = ""
		str = str & "<html>"
		str = str & "<head>"
		str = str & "<title>Client Statement</title>  "
		str = str & "<SCRIPT language=Javascript src=""../../scripts/common.js""></SCRIPT> "
		str = str & "<SCRIPT language=Javascript src=""../../scripts/fhsupport.js""></SCRIPT>"
		str = str & "<link rel=""stylesheet"" type=""text/css"" href=""CALENDAR/calendar.css"">"
		str = str & "<SCRIPT language=Javascript src=""Calendar/calendar.js""></SCRIPT>"
		str = str & "<LINK REL=""STYLESHEET"" TYPE=""TEXT/CSS"" HREF=""../../STYLE/default.css""> "
		str = str & "<LINK REL=""STYLESHEET"" TYPE=""TEXT/CSS"" HREF=""../../STYLE/webparts.css"">"
		str = str & "<SCRIPT language=VBScript src=""../../scripts/reports.vbs""></SCRIPT>"
		str = str & "<SCRIPT language=Javascript src=""../../scripts/reports.js""></SCRIPT>"
		str = str & "<style media=""print"">"
		str = str & "@page {"
		str = str & "size: portrait;"
		str = str & "margin-left: 0cm;"
		str = str & "margin-right: 0cm;"
		str = str & "margin-top: 0cm;"
		str = str & "margin-bottom: 0cm;"
		str = str & "writing-mode: tb-rl;"
		'str = str & "height: 80%;"
		'str = str & "margin: 10% 0%;"
		str = str & "br.newpage{"
		str = str & "page-break-before:always;"
		str = str & "}"		
		str = str & "}"
		str = str & "</style>"
		str = str & "</head>"
		str = str & "<body Class=""Reports"" leftmargin=""25"">"
		
		'htmlOut = Replace(htmlOut,"../data/photos/aaprintlogofooter.jpg","Include\aaprintlogo.jpg")
	Else
		str = ""
		str = str & "<html>"
		str = str & "<head>"
		str = str & "<title>REPORT</title>  "
		str = str & "<SCRIPT language=Javascript src=""../../scripts/common.js""></SCRIPT> "
		str = str & "<SCRIPT language=Javascript src=""../../scripts/fhsupport.js""></SCRIPT>"
		str = str & "<link rel=""stylesheet"" type=""text/css"" href=""CALENDAR/calendar.css"">"
		str = str & "<SCRIPT language=Javascript src=""Calendar/calendar.js""></SCRIPT>"
		str = str & "<LINK REL=""STYLESHEET"" TYPE=""TEXT/CSS"" HREF=""../../STYLE/default.css""> "
		str = str & "<LINK REL=""STYLESHEET"" TYPE=""TEXT/CSS"" HREF=""../../STYLE/webparts.css"">"
		str = str & "<SCRIPT language=VBScript src=""../../scripts/reports.vbs""></SCRIPT>"
		str = str & "<SCRIPT language=Javascript src=""../../scripts/reports.js""></SCRIPT>"
		str = str & "<style media=""print"">"
		str = str & "@page {"
		str = str & "size: portrait;"
		str = str & "margin-left: 0cm;"
		str = str & "margin-right: 0cm;"
		str = str & "margin-top: 0cm;"
		str = str & "margin-bottom: 0cm;"
		str = str & "writing-mode: tb-rl;"
		'str = str & "height: 80%;"
		'str = str & "margin: 10% 0%;"
		str = str & "br.newpage{"
		str = str & "page-break-before:always;"
		str = str & "}"		
		str = str & "}"
		str = str & "</style>"
		str = str & "</head>"
		str = str & "<body Class=""Reports"" leftmargin=""25"">"
		'str = str & "<table border=0 cellspacing=2 cellpadding=2 align=center width=""90%"">"
		'str = str & "<tr>"
		'str = str & "<td align=center><img src=""Include\aaprintlogo.jpg"" width=""480"" height=""179""></td>"
		'str = str & "</tr>"
		'str = str & "</table>"
	End If
else
	''LANDSCAPE
	If Instr(1,htmlOut,"Portfolio Valuation Statement") > 0 Then
		str = ""
		str = str & "<html>"
		str = str & "<head>"
		str = str & "<title>Portifolio Valuation Statement</title>  "
		str = str & "<SCRIPT language=Javascript src=""../../scripts/common.js""></SCRIPT> "
		str = str & "<SCRIPT language=Javascript src=""../../scripts/fhsupport.js""></SCRIPT>"
		str = str & "<link rel=""stylesheet"" type=""text/css"" href=""CALENDAR/calendar.css"">"
		str = str & "<SCRIPT language=Javascript src=""Calendar/calendar.js""></SCRIPT>"
		str = str & "<LINK REL=""STYLESHEET"" TYPE=""TEXT/CSS"" HREF=""../../STYLE/default.css""> "
		str = str & "<LINK REL=""STYLESHEET"" TYPE=""TEXT/CSS"" HREF=""../../STYLE/webparts.css"">"
		str = str & "<SCRIPT language=VBScript src=""../../scripts/reports.vbs""></SCRIPT>"
		str = str & "<SCRIPT language=Javascript src=""../../scripts/reports.js""></SCRIPT>"
		str = str & "<style media=""print"">"
		str = str & "@page {"
		str = str & "size: landscape;"
		str = str & "margin-left: 0cm;"
		str = str & "margin-right: 0cm;"
		str = str & "margin-top: 0cm;"
		str = str & "margin-bottom: 0cm;"
		str = str & "writing-mode: tb-rl;"
		'str = str & "height: 80%;"
		'str = str & "margin: 10% 0%;"
		str = str & "br.newpage{"
		str = str & "page-break-before:always;"
		str = str & "}"		
		str = str & "}"
		str = str & "</style>"
		str = str & "</head>"
		str = str & "<body Class=""Reports"" leftmargin=""25"">"
		
		'htmlOut = Replace(htmlOut,"../data/photos/aaprintlogofooter.jpg","Include\aaprintlogo.jpg")
	Else
		str = ""
		str = str & "<html>"
		str = str & "<head>"
		str = str & "<title>REPORT</title>  "
		str = str & "<SCRIPT language=Javascript src=""../../scripts/common.js""></SCRIPT> "
		str = str & "<SCRIPT language=Javascript src=""../../scripts/fhsupport.js""></SCRIPT>"
		str = str & "<link rel=""stylesheet"" type=""text/css"" href=""CALENDAR/calendar.css"">"
		str = str & "<SCRIPT language=Javascript src=""Calendar/calendar.js""></SCRIPT>"
		str = str & "<LINK REL=""STYLESHEET"" TYPE=""TEXT/CSS"" HREF=""../../STYLE/default.css""> "
		str = str & "<LINK REL=""STYLESHEET"" TYPE=""TEXT/CSS"" HREF=""../../STYLE/webparts.css"">"
		str = str & "<SCRIPT language=VBScript src=""../../scripts/reports.vbs""></SCRIPT>"
		str = str & "<SCRIPT language=Javascript src=""../../scripts/reports.js""></SCRIPT>"
		str = str & "<style media=""print"">"
		str = str & "@page {"
		str = str & "size: portrait;"
		str = str & "margin-left: 0cm;"
		str = str & "margin-right: 0cm;"
		str = str & "margin-top: 0cm;"
		str = str & "margin-bottom: 0cm;"
		str = str & "writing-mode: tb-rl;"
		'str = str & "height: 80%;"
		'str = str & "margin: 10% 0%;"
		str = str & "br.newpage{"
		str = str & "page-break-before:always;"
		str = str & "}"		
		str = str & "}"
		str = str & "</style>"
		str = str & "</head>"
		str = str & "<body Class=""Reports"" leftmargin=""25"">"
		str = str & "<table border=0 cellspacing=2 cellpadding=2 align=center width=""90%"">"
		str = str & "<tr>"
		str = str & "<td align=center><img src=""Include\aaprintlogo.jpg"" width=""480"" height=""179""></td>"
		str = str & "</tr>"
		str = str & "</table>"
	End If
End If

If Instr(1,htmlOut,"../data/photos/") > 0 Then
	htmlOut = Replace(htmlOut,"../data/photos/","Include\")
End If

If Instr(1,htmlOut,"../images/") > 0 Then
	htmlOut = Replace(htmlOut,"../images/","Include\")
End If

htmlOut = str & htmlOut

'Response.Write htmlOut
'Response.End 

If htmlOut = "" Then
	%>
	<Script Language="JavaScript">
		alert("Error in conversion to PDF.");
		window.history.back(1);
	</Script>
	<%
	Response.End
End If

'write to output file
Dim FSO, fTextStream, pathTo, tmpMark, pageFrom, xlFileName
	
Randomize
	
tmpMark = "-" & Rnd(2) & "-PNW"

fromPath = server.MapPath("./") & "\Mailer\bin"

'pathTo = fromPath & "\PDF\" &  tmpMark & ".html"
'xlFileName = "PDF" &  tmpMark & ".pdf"
'newPathTo = fromPath & "\PDF\" & xlFileName

'Response.Write newPathTo
'Response.End 

strDestFile = tmpMark & ".html"
pathTo = fromPath & "\" &  strDestFile
xlFileName = "PDF" &  tmpMark & ".pdf"
newPathTo = fromPath & "\" & xlFileName

'Response.Write newPathTo
'Response.End 

Set FSO = CreateObject("Scripting.FileSystemObject")

'delete earlier HTML output file if it exists
If FSO.FileExists(pathTo) Then
	FSO.DeleteFile pathTo, True
End If

Set fTextStream = Fso.CreateTextFile(pathTo, True, False)
fTextStream.Write htmlOut
fTextStream.Close
Set fTextStream = Nothing
	
Set FSO = Nothing

'Response.Write pathTo & "<BR>"
'Response.Write newPathTo
'Response.End 

''CONVERSION TO PDF
''------------------------------------------------------------------------------------------
Dim sName, sPort, sPath, thisPath
	
sName = Request.ServerVariables("SERVER_NAME")
sPort = Request.ServerVariables("SERVER_PORT")
	
If sPort <> "" Then
	sPath = "http://" & sName & ":" & sPort & "/"
Else
	sPath = "http://" & sName & "/"
End If
	
thisPath = Server.MapPath(".")
	
sPath = sPath & "Mailer/bin/" & strDestFile

'Response.Write sPath
'Response.End 
	
Dim theDoc

Set theDoc = Server.CreateObject("ABCpdf5.Doc")

theDoc.Rect.SetRect 40, 72, 520, 640

If theOrientation = "L" Then
	' apply a rotation transform
	w = 520
	h = 640
	l = 40
	b = 72
	theDoc.Transform.Rotate 90, l, b
	theDoc.Transform.Translate w, 0

	' rotate our rectangle
	theDoc.Rect.Width = h
	theDoc.Rect.Height = w
End If

theURL = sPath

If Left(theURL, 4) <> "http" Then theURL = "http://www.google.com/"

theID = theDoc.AddImageUrl(theURL, True, 0, False)
For i = 1 To 50 ' add up to 50 pages
  If theDoc.GetInfo(theID, "Truncated") <> "1" Then Exit For
  theDoc.Page = theDoc.AddPage()
  theID = theDoc.AddImageToChain(theID)
Next
theDoc.PageNumber = 1

If theOrientation = "L" Then
	' adjust the default rotation and save
	theID = theDoc.GetInfo(theDoc.Root, "Pages")
	theDoc.SetInfo theID, "/Rotate", "90"
End If

theDoc.Save newPathTo

'Response.end
''------------------------------------------------------------------------------------------
''END CONVERSION TO PDF

IF ENABLED THEN
	''TO ALLOW THE USER TO SAVE THE DOCUMENT
	Response.ContentType = "application/x-unknown"

	fn = xlFileName

	FPath = newPathTo
	    
	Response.AddHeader "Content-Disposition", "attachment; filename=""" & fn & """"

	Set adoStream = CreateObject("ADODB.Stream") 
	adoStream.Open() 
	adoStream.Type = 1 
	adoStream.LoadFromFile(FPath) 
	Response.BinaryWrite adoStream.Read() 
	adoStream.Close 
	Set adoStream = Nothing 
	%>
	<Script Language=JavaScript>
		alert('PDF generated.');
		window.history.back(-1)		
	</Script>
	<%
END IF

PDF_FILE_NAME = newPathTo

If PDF_FILE_NAME = "" Then%>
	<Script Language="JavaScript">
		alert("This page <%= subject %> does not have enough information to proceed.");
		//window.parent.self.close();
	</Script>
	<%	   
	Response.End
End If

toRecipient = Request.Form("To2")
cc = Request.Form("cc2")
bcc = Request.Form("bcc2")
subject = Request.Form("subject2")
bodyText = Request.Form("bodytext2")
attachment = PDF_FILE_NAME
	
'Response.write toRecipient & "<br>"
'Response.write cc & "<br>"
'Response.write bcc & "<br>"
'Response.write subject & "<br>"
'Response.write bodyText & "<br>"
'Response.write attachment & "<br>"
'Response.End 
		
if instr(1,attachment,".pdf",1)=0 then
	%>
	<Script Language="JavaScript">
		alert("Please a valid file pdf");
		//window.parent.self.close();
	</Script>
	<%
	response.end
end if

'response.write htmlOut
'response.end
%>

<!--#include file="libroutinesTEST.asp"-->

<%
Set Conn = GetActiveConnection("KBroker")
Set Rs = CreateObject("ADODB.Recordset")	

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
	
'toRecipient = "jcmgakingo@yahoo.com"
	
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
	.Sender = "securities@africanalliance.co.ke"
	.From = "securities@africanalliance.co.ke"
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
		
	'.CreateMHTMLBody "file://E:\Knowing\BrokerknowUG\Email\blank.htm" 

	'On Error Resume Next
	
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
Set Fso = Nothing
		
%>
<Script Language="JavaScript">
	alert("The pdf was created and sent successfully.")	
	//window.parent.close();	
</Script>

<br><br>
<p align="center" style="font-family:arial;font-weight:bold;font-size:12pt;">Message sent.

<%	
Response.End 
%>