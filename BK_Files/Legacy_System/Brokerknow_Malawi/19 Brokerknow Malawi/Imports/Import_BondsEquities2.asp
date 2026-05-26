<HTML>
<HEAD>
	<TITLE>Bonds & Equities Import</TITLE>
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
	<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT>
	<style type="text/css">
		td {border: 1 solid gray; text-align:left;}
	</style>
</head>

<body>

<!--#include virtual="libroutines.asp"-->
<!--#include file="flUploader.asp"-->	

<%
action = Request.QueryString("action")
If action = "EXECUTE_SAVE" Then
	
	Bond = Request.Form("hidBond")
	Equity = Request.Form("hidEquity")
	
	sBond = Split(Bond,",")
	sEquity = Split(Equity,",")
	
	Set Conn = Server.CreateObject("ADODB.Connection")
	Set Conn = GetActiveConnection("KBroker")
	
	SQL = "INSERT INTO _ImportBondsEquitiesTurnover (Bonds, Equities)" & _
	"VALUES ('"& sBond(1) &"', '"& sEquity(0) &"')"
	Conn.Execute SQLServerFormat(HandleQuote(SQL))
	%>
	<SCRIPT LANGUAGE="JavaScript">
		window.alert("Bonds and Equity Turnovers succesfully commited.");
		window.location='import.asp';
	</SCRIPT>
	<%
	Response.End 
End If
%>
		
<%
Dim load
Set load = new Loader
				
' calling initialize method
load.initialize
				
' File name
Dim fileName, import_UDLPath
fileName = LCase(load.getFileName("sourcefile"))			
	
' Path where file will be uploaded
Dim pathToFile, CurrentDirectory
CurrentDirectory = "."	
		
pathToFile = Server.MapPath(CurrentDirectory) & "\" & filename
		
Dim fso
Dim msgExists
Dim SqlStr2
		
Set fso = server.CreateObject("Scripting.FileSystemObject") 
		
'If (fso.FileExists(pathToFile) = True) Or (fso.FolderExists(pathToFile) = True) Then
	'msgExists = "Could not upload file. A file or folder with that name already exists"
	'fso.DeleteFile pathToFile
	'Set fso = Nothing
	%>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
		//alert("<%=msgExists%>");
		//window.history.back(0);
	//-->
	</SCRIPT>
	<%
	'response.end
'Else		
	Dim fileUploaded
	fileUploaded = load.saveToFile ("sourcefile", pathToFile)
'End If
		
' destroy load object
Set load = Nothing	
	
' if the file is uploaded then
If (fileUploaded = True) Then
	'File = "C:\Documents and Settings\Administrator.BROKERKNOWSERV\Desktop\13-Jan-2006.xls"
	File = pathToFile
Else
	%>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
		alert("<%=msgExists%>");
		window.history.back(0);
	//-->
	</SCRIPT>
	<%
	Response.End 
End If
%>
<form method="post" name="frmPrices" action="Import_BondsEquities2.asp?action=EXECUTE_SAVE">

<TABLE border=0 cellspacing=0 cellpadding=2 width=80% align=center style="border: 1 solid gray;">
<%
Response.Write("<TR><TD><b>NSE Daily Price List ["& FormatDate(Date) &"]</b></TD></TR>")

Set objConn = Server.CreateObject("ADODB.Connection")

objConn.Open "DBQ=" & File & ";" & "DRIVER={Microsoft Excel Driver (*.xls)};" 

''BONDS
Set objRS = Server.CreateObject("ADODB.Recordset")

objRS.ActiveConnection = objConn
objRS.CursorType = 3'Static cursor
objRS.Source = "Select * from [Price List$H176:I178]"
objRS.Open

Response.Write("<TR><TD>Turnover in Bonds Today</TD></TR>")

do until objRS.EOF
		Response.Write "<TR>"
		Response.write "<TD>" & FormatNum(objRS.Fields.Item(0).Value) & "<TD>"
		Response.write "<input type=hidden name=hidBond id=hidBond value="& objRS.Fields.Item(0).Value &">"
		Response.Write "</TR>" & vbCrLf 
	objRS.MoveNext
loop
     
objRS.close
set objRS=nothing

''EQUITY
Set objRS = Server.CreateObject("ADODB.Recordset")

objRS.ActiveConnection = objConn
objRS.CursorType = 3'Static cursor
objRS.Source = "Select * from [Price List$H182:I184]"
objRS.Open

Response.Write("<TR><TD>Equity Turnover Today</TD></TR>")

do until objRS.EOF
		Response.Write "<TR>"
		Response.write "<TD>" & FormatNum(objRS.Fields.Item(0).Value) & "<TD>"
		Response.write "<input type=hidden name=hidEquity id=hidEquity value="& objRS.Fields.Item(0).Value &">"
		Response.Write "</TR>" & vbCrLf 
	objRS.MoveNext
loop
     
objRS.close
set objRS=nothing

objConn.close
set objConn=nothing
%>
<tr>
	<td style="border: 0;">
		<input type="hidden" value="1" name="commit">
		<input type="submit" name="submit" value="Commit Turnover">
	</td>
</tr>
</table>
</form>
<%
Response.Write("</TABLE>")
%>
</body>
</html>


