<%@ Language=VBScript %>
<HTML>
<HEAD></HEAD>
<BODY>
<%
Dim FSO
Set FSO = Server.CreateObject("Scripting.FileSystemObject")

Dim FL
Set FL = FSO.GetFolder(Server.MapPath(".") & "\bin")

For each fle in FL.Files
	If Instr(1,fle,"htm") > 0 Or Instr(1,fle,"pdf") > 0 Or Instr(1,fle,"html") > 0 Then
		FSO.DeleteFile fle, True
	End If
Next
%>
</BODY>
</HTML>
