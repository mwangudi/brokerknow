<!--#include file="../libroutines.asp"-->
<HTML>
<HEAD>
<TITLE>Run Batch File</TITLE>

<%
fp = Request.QueryString("fp")
fpath = server.MapPath(".") & "\batch\" & fp

Set RST = Server.CreateObject("ADODB.RecordSet")

'Set Conn = GetActiveConnection("KBroker")
Set Conn = Server.CreateObject("ADODB.Connection")
'Conn.Open("BrokerKnow")

Conn.Open("BMA")

UserID = 10

SQL = "INSERT INTO SendSMS (SMSFileName, UserID)" & _
" VALUES ('"& fpath &"', "& UserID &")"

'SQL = "INSERT INTO [SendSMS] (SMSFileName, UserID) " & _
'		"		SELECT " & "'" & fp & "'" & " as SMSFileName " & _
'		"		," & " " & UserID & " " & " as UserID "
											
'SQL = SQLServerFormatWithCustomMax(HandleQuote(SQL))

Conn.Execute SQL

'RST.Open SQL, Conn, 0, 1

'conn.Execute sqlStr
									
Set RST = Nothing
Set Conn = Nothing
%>

<SCRIPT LANGUAGE="VBScript">
'Function RunFile(theFile)
'	Set WshShell = CreateObject("WScript.Shell")

'	WshShell.Exec(theFile)
'End Function
</SCRIPT>

</HEAD>

<!--<BODY onload="VBScript: RunFile('<%=fpath%>')">-->
<BODY>

</BODY>
</HTML>

<SCRIPT LANGUAGE="VBScript">
	ShowMessage "File has been sent."
	window.location.href = "downloadSMS.asp"
</SCRIPT>