<%@ Language=VBScript %>
<HTML>
<HEAD></HEAD>
<BODY>

<!--
METADATA TYPE="TypeLib"
NAME="BCL easyPDF Printer SDK 1.2 Type Library"
UUID="{514D3817-1AA8-4CF8-96D9-C56139C79373}"
VERSION="1.2"
-->

<%
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
		str = str & "<table border=0 cellspacing=2 cellpadding=2 align=center width=""90%"">"
		str = str & "<tr>"
		str = str & "<td align=center><img src=""Include\aaprintlogo.jpg"" width=""482"" height=""178""></td>"
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
    
strDestFilePath = Server.MapPath(".") & "\bin\" & "-MAIL" & Rnd(2) & "-CS.htm"
                                
Dim FSO
Set FSO = Server.CreateObject("Scripting.FileSystemObject")

Dim fTextStream
Set fTextStream = Fso.CreateTextFile(strDestFilePath, True, False)
fTextStream.Write theData
fTextStream.Close
Set fTextStream = Nothing
                              
outFileName = Replace(strDestFilePath, "htm", "pdf")
 
''CONVERSION TO PDF
Dim oLoader
Dim oPrinter
Dim oPrintJob
Dim oSetting

Set oLoader = Server.CreateObject("easyPdfSdk.Loader")

Set oPrinter = oLoader.Printer
Set oPrintJob = oPrinter.IEPrintJob
Set oSetting = oPrinter.PrinterSetting

Select Case theCategory
	Case "HoldingsValuation"
		'oSetting.CreatePaper "MyCustomPaper", 11.69, 8.27
		'oSetting.LayoutPaperSize = oSetting.GetPaperSize("MyCustomPaper")
		'oSetting.Save
		oSetting.LayoutPaperSize = oSetting.GetPaperSize("A4")
		oSetting.LayoutPaperOrientation = PRN_PAPER_ORIENT_LANDSCAPE
		oSetting.Save
	Case Else
		oSetting.LayoutPaperSize = oSetting.GetPaperSize("A4")
		oSetting.LayoutPaperOrientation = PRN_PAPER_ORIENT_PORTRAIT
		oSetting.Save
End Select

oPrintJob.Footer = ""
oPrintJob.Header = ""
oPrintJob.MarginBottom = 1
oPrintJob.MarginLeft = 1
oPrintJob.MarginRight = 1
oPrintJob.MarginTop = 1
oPrintJob.SaveSetting

On Error Resume Next
Call oPrintJob.PrintOut(strDestFilePath, outFileName)

If Not Len(theClient) > 0 Then
	theClient = 0
End If

SQL = "SELECT Client.Client_DPA_FROM Client WHERE (Client.Client_DPA_ = "& theClient &")"
Set RST = oConn.Execute(SQL)
        
If Not (RST.EOF Or RST.BOF) Then
	toRecipient = RST("ClientEmail")
Else
	toRecipient = ""
	Response.Write "<br><br><p align=center style=""font-family:tahoma;font-size:10pt;"">Invalid Email Address."
	Response.End 
End If

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

Set Conn = GetActiveConnection("KBroker")
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
	theCDO.TextBody = "Please find attached your file from AAKS."
	 
	On Error Resume Next  
	'theCDO.Send
End If

Response.Write "<br><br><p align=center style=""font-family:tahoma;font-size:10pt;"">Message Sent."

Response.End 
%> 
 
</BODY>
</HTML>
