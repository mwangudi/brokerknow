<!--#include file="../libroutines.asp"-->
<HTML>
<HEAD>
<TITLE> SEND SMS </TITLE>
<BODY>
<% 
    fn = Request.QueryString("f")
    fpath = server.MapPath(".") & "\bin\" & fn
    
    EntryLine = ""
    
    Set FSO = Server.CreateObject("Scripting.FileSystemObject")
    
    Dim UniqueLine
    Dim FileStruct
    Dim FilePath
    
    UniqueLine = Day(Date)&"-"&MonthName(Month(Date),true)&"-"&Year(Date)&"_"&Hour(Time)&"-"&Minute(Time)&"-"&Second(Time)
    
    Select Case Mid(fn,1,2)
		
		Case "ca"
			
			FileStruct = "ca_"& UniqueLine &".bat"
			FilePath = server.MapPath(".") & "\batch\" & FileStruct
			
		Case "db"
			
			FileStruct = "db_"& UniqueLine &".bat"
			FilePath = server.MapPath(".") & "\batch\" & FileStruct
			
		Case "sw"
			
			FileStruct = "sw_"& UniqueLine &".bat"
			FilePath = server.MapPath(".") & "\batch\" & FileStruct
			
		Case Else
			''Do Nothing
			FileStruct = UniqueLine &".bat"
			FilePath = server.MapPath(".") & "\batch\" & FileStruct
			
    End Select
	
	EntryLine = EntryLine & "@ECHO OFF" & vbCrLf 
	EntryLine = EntryLine & "> script.ftp ECHO USER 1362866" & vbCrLf 
	EntryLine = EntryLine & ">>script.ftp ECHO aaks12" & vbCrLf 
	EntryLine = EntryLine & ">>script.ftp ECHO binary" & vbCrLf 
	EntryLine = EntryLine & ">>script.ftp ECHO prompt n" & vbCrLf 
	EntryLine = EntryLine & ">>script.ftp ECHO put " & fpath & vbCrLf 
	EntryLine = EntryLine & ">>script.ftp ECHO quit" & vbCrLf 
	EntryLine = EntryLine & "FTP -v -n -s:script.ftp ftpupload.clickatell.com" & vbCrLf 
	EntryLine = EntryLine & "TYPE NUL >script.ftp" & vbCrLf 
	EntryLine = EntryLine & "DEL script.ftp" & vbCrLf 
	EntryLine = EntryLine & "GOTO End" & vbCrLf 
	EntryLine = EntryLine & vbCrLf 
	EntryLine = EntryLine & ":End" & vbCrLf 
	EntryLine = EntryLine & vbCrLf 
	EntryLine = EntryLine & "exit" & vbCrLf 
			    
    sBuff = EntryLine
	
	
	Set FSOFile = FSO.OpenTextFile(FilePath, 8, True)
							
	FSOFile.WriteLine sBuff
	FSOFile.Close()
		
	Set FSOFile = Nothing 
	Set FSO = Nothing
    
    %>
    <SCRIPT LANGUAGE="VBScript">
		window.location.href = "batch.asp?fp=<%=FileStruct%>"
	</SCRIPT>
    <%
    Response.End 
%>
</BODY>
</HTML>