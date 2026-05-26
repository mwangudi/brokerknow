<%@ Language=VBScript %>
<HTML>
<HEAD></HEAD>
<BODY>
<!--#include file="../libroutinesVB.asp"-->
<%
Response.Buffer = True

theData = Request.Form("hidData")
theClient = Request.Form("hidClient")
theCategory = Request.Form("hidCategory")

Dim strDestFilePath
Dim outFileName
Dim str

Select Case theCategory
	Case "Statement"
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
		str = str & "<table border=0 cellspacing=0 cellpadding=2 align=center width=""90%"">"
		str = str & "<tr>"
		str = str & "<td align=center><img src=""Include\aaprintlogo.jpg"" width=""480"" height=""179""></td>"
		str = str & "</tr>"
		str = str & "</table>"
		
		theData = str & theData
		
	Case "Contract"
		str = ""
		str = str & "<html>"
		str = str & "<head>"
		str = str & "<title>Client Contract</title>  "
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
		
		theData = str & theData
	
	Case "ContractCompounded"
		str = ""
		str = str & "<html>"
		str = str & "<head>"
		str = str & "<title>Client Contract Compounded</title>  "
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
		
		theData = str & theData
		
	Case "HoldingsValuation"
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
		
		theData = str & theData
		
		theData = Replace(theData,"xx-xx-xx","Include\aaprintlogo.jpg")
End Select

'Response.Write thedata
'Response.End 

Randomize

strDestFile = "-MAIL" & Rnd(2) & "-GEN.htm"
strDestFilePath = Server.MapPath(".") & "\bin\" & strDestFile
                                
Dim FSO
Set FSO = Server.CreateObject("Scripting.FileSystemObject")

Dim fTextStream
Set fTextStream = Fso.CreateTextFile(strDestFilePath, True, False)
fTextStream.Write theData
fTextStream.Close
Set fTextStream = Nothing
                              
outFileName = Replace(strDestFilePath, "htm", "pdf")
 
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
	
sPath = sPath& "Mailer/bin/" & strDestFile

'Response.Write sPath
'Response.End 
	
Dim theDoc

Set theDoc = Server.CreateObject("ABCpdf5.Doc")

theDoc.Rect.SetRect 40, 72, 520, 640

If theCategory = "HoldingsValuation" Then
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

If theCategory = "HoldingsValuation" Then
	' adjust the default rotation and save
	theID = theDoc.GetInfo(theDoc.Root, "Pages")
	theDoc.SetInfo theID, "/Rotate", "90"
End If

theDoc.Save outFileName

'Response.end
''------------------------------------------------------------------------------------------
''END CONVERSION TO PDF

If Not Len(theClient) > 0 Then
	theClient = 0
End If

Set Conn = GetActiveConnection("KBroker")
Set RST = CreateObject("ADODB.Recordset")	

sqlStr = "SELECT Client_DPA_, ClientEmail FROM Client WHERE (Client.Client_DPA_ = "& theClient &")"
'Set RST = Conn.Execute(SQL)

RST.CursorLocation = adUseClient	
RST.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
	        
If Not (RST.EOF Or RST.BOF) Then
	toRecipient = RST("ClientEmail")
Else
	toRecipient = ""
	Response.Write "<br><br><p align=center style=""font-family:tahoma;font-size:10pt;"">Invalid Email Address."
	Response.End 
End If

'toRecipient = "jcmgakingo@yahoo.com"

Select Case theCategory
	Case "Statement"
		subject = "AAKS - Client Statement"	
	
	Case "Contract"
		subject = "AAKS - Client Contract"
		
	Case "ContractCompounded"
		subject = "AAKS - Client Contract"
	
	Case "HoldingsValuation"
		subject = "AAKS - Holdings Valuation"
		
	Case Else
		subject = " "
End Select

Dim theCDO
Set theCDO = Server.CreateObject("CDO.Message") 

SQL = "SELECT * FROM MailConfigList"
Set Rs = Conn.Execute(SQL)

If Not (Rs.eof or Rs.bof) Then  
	theCDO.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = Rs.Fields("SendUsingMethod").Value
	theCDO.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = Rs.Fields("SMTPServer").Value
	theCDO.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = Rs.Fields("SMTPServerPort").Value
	theCDO.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = Rs.Fields("SMTPAuthenticate").Value
	theCDO.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = Rs.Fields("SendUserName").Value
	theCDO.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Rs.Fields("SendPassword").Value
	theCDO.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = Rs.Fields("SMTPConnectionTimeout").Value
	theCDO.Configuration.Fields.Update
	    
	theCDO.AddAttachment outFileName
	theCDO.To = toRecipient
	theCDO.Subject = subject

	theCDO.From = Rs.Fields("SendDisplayName").Value
	theCDO.Sender = Rs.Fields("SendDisplayName").Value
	theCDO.HTMLBody = theData
	theCDO.CreateMHTMLBody "file://E:\Knowing\Brokerknow\Mailer\blank.htm", cdoSuppressNone 
	
	On Error Resume Next  
	'theCDO.Send
End If

Response.Write "<br><br><p align=center style=""font-family:tahoma;font-size:10pt;"">Message Sent."

%>
<script language="vbscript">
	'window.close()
</script>
<%

Response.End 
%> 
 
</BODY>
</HTML>
