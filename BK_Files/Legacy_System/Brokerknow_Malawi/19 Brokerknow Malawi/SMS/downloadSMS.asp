<!--#include file="../libroutines.asp"-->
<HTML>
<HEAD>
<TITLE> Download SMS </TITLE>
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css">
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">

<SCRIPT LANGUAGE="JavaScript">
	function SaveFile(theFile)
	{
		window.location.href = 'downloadSMS.asp?Action=SaveFile&theFile=' + theFile;
	}
</SCRIPT>

<script language="VBScript">
	Function selectFolder()
		theDrive = 	window.document.all("sltDrives").value
		theFile = window.document.all("hidFile").value
		
		window.location.href = "downloadSMS.asp?Action=SaveFile&theFile="& theFile &"&drive="& theDrive
	End Function
		
	Function selectSubFolder()
		theFolder = window.document.all("sltFolders").value 
		theDrive = 	window.document.all("sltDrives").value
		theFile = window.document.all("hidFile").value
			
		window.document.all("txtFolder").value = theFolder
			
		window.location.href = "downloadSMS.asp?Action=SaveFile&theFile="& theFile &"&drive="& theDrive & "&folder=" & theFolder
	End Function
		
	Function putSubFolder()
		theFolder = window.document.all("sltFolders").value 
		
		window.document.all("txtFolder").value = theFolder
	End Function
	
	Function SaveFileToPath()
		theFolder = window.document.all("txtFolder").value
		theFile = window.document.all("hidFile").value
		
		window.location.href = "downloadSMS.asp?Action=SaveFileToPath&theFile="& theFile &"&folder=" & theFolder
	End Function
	
	Function UploadFileToPath()
		theFolder = window.document.all("txtFolder").value
		theFile = window.document.all("hidFile").value
		
		window.location.href = "downloadSMS.asp?Action=UploadFileToPath&theFile="& theFile &"&folder=" & theFolder
	End Function
	
	Function GoBack()
		'window.history.back(-1)
		window.location.href="SMS.asp"
	End Function
</script>
	
<STYLE TYPE="TEXT/CSS">
	a {text-decoration:none}
	a:hover {text-decoration:underline}
	a:visited {text-decoration:none}
	TD {BORDER-BOTTOM:1 SOLID GRAY;BORDER-LEFT:1 SOLID GRAY;}
</STYLE>
</HEAD>

<BODY>

<%
''variables to use
Action = Request.QueryString("Action")

If Action="" Then
	Action = "ListFiles"
End If
%>

<%
If (Action = "ListFiles") Then
	Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	
	Dim FSOFolder
	Dim FSOFiles
	
	FolderPath = Server.MapPath(".") & "\bin\"
	
    Set FSOFolder = FSO.GetFolder(FolderPath)
    Set FSOFiles = FSOfolder.Files
    
    ''DIM ARRAY
    Dim TheFiles()
    ReDim TheFiles(500)
    
    ''n AS RECORDCOUNT FOR ARRAY
    n = -1
    
    ''GET FILES AND PUT THEM IN AN ARRAY
    For Each FSOFile In FSOFiles
		If (Instr(1,FSOFile.Name,"txt")>0) Then
			fname = FSOFile.Name
			fdate1 = FSOFile.DateLastModified
			
			''MASSAGE DATE FOR SORTING
			if not (len(fdate1)=0) then
				theday = day(fdate1)
				themonth = month(fdate1)
				theyear = year(fdate1)
							
				if (len(theday)=1) then
					theday = "0" & theday
				end if
									
				if (len(themonth)=1) then
					themonth = "0" & themonth
				end if
									
				if (len(theyear)=1) then
					theyear = "0" & theyear
				end if
			end if
			
			fdate1 = theyear & "" & themonth & "" & theday
			fdate2 = FormatDate(FSOFile.DateLastModified)
			fdate3 = FormatDateTime(FSOFile.DateLastModified,vbLongTime)
			fsize = FSOFile.Size
			
			n = n + 1
			
			If (n > UBound(TheFiles)) Then
			    ReDim Preserve TheFiles(n + 99)
			End If
			
			''BUILD ARRAY
			TheFiles(n) = Array(fname,fdate1,fdate2,fdate3,fsize)
		End If
	Next
	
	''SET CURRENT FILECOUNT
	FileCount = n
	ReDim Preserve TheFiles(n)
	
	''SORT BY
	'0 = Filename
	'1 = Datelastmodified
	'3 = Filesize
	
	''SORT THE ARRAY
	For i = FileCount TO 0 Step -1
		MinMax = TheFiles(0)(1)
		MinMaxSlot = 0
			
		For j = 1 To i
			'SORT AS STRING
			'Mark = (strComp(TheFiles(j)(1), MinMax, vbTextCompare) > 0)
			
			'SORT AS NON-STRING, DATE
			Mark = (TheFiles(j)(1) < MinMax)
			
			If Mark Then 
			    MinMax = TheFiles(j)(1)
			    MinMaxSlot = j
			End If
		Next
			
		If MinMaxSlot <> i Then 
		    Temp = TheFiles(MinMaxSlot)
		    TheFiles(MinMaxSlot) = TheFiles(i)
		    TheFiles(i) = Temp
		End If
	Next	
    %>
	<DIV style="WIDTH: 100%">
	<BR>
	<p><input type ="button" name ="back" value="<< Back" onclick="VBScript: GoBack()"></p>
	<BR>
	<TABLE cellPadding="5" cellSpacing="2" width="75%" style="border:1 solid gray;font-family:tahoma;font-size:8pt;">
		<TR bgcolor="gainsboro">
			<TD COLSPAN="6" STYLE="BORDER:0">
				<b>SMS Files Generated</b>
			</TD>
		</TR>
		
		<TR>
			<TD WIDTH="20%"><b>Name</b></TD>
			<TD WIDTH="20%"><b>Date Created</b></TD>
			<TD ALIGN="right" WIDTH="20%"><b>Size (bytes)</b></TD>
			<TD ALIGN="center" WIDTH="20%"><b>Download</b></TD>
			<!--<TD ALIGN="center" WIDTH="20%"><b>Save As</b></TD>
			<TD ALIGN="center" WIDTH="20%"><b>Send</b></TD>-->
		</TR> 
		<%
		''DISPLAY FILES FROM ARRAY
		P = 0
		For i = 0 To FileCount
			P = P + 1
		    %>
			<TR>
				<TD NOWRAP WIDTH="25%">
					<%=P%>.&nbsp;&nbsp;<a href="downloadSMS.asp?Action=ViewFile&FilePath=<%=TheFiles(i)(0)%>" title="PREVIEW"><%=TheFiles(i)(0)%></a>
				</TD>
				<TD NOWRAP WIDTH="20%"><%=TheFiles(i)(2)%> (<%=TheFiles(i)(3)%>)</TD>
				<TD NOWRAP ALIGN="right" WIDTH="20%"><%=TheFiles(i)(4)%></TD>
				<TD NOWRAP ALIGN="center" WIDTH="20%"><!--<a title="DOWNLOAD" href="download.asp?f=<%=TheFiles(i)(0)%>">DOWNLOAD</a>-->
				
				<input type=button style="font-family:tahoma;font-weight:bold;font-size:8pt;background-color:silver;color:black;border:0;" name="download" Value="Download" onclick="window.location.href='download.asp?f=<%=TheFiles(i)(0)%>';"></TD>
				
				<!--<TD NOWRAP ALIGN="center" WIDTH="20%"><a title="SAVE AS" href="bin\<%=TheFiles(i)(0)%>">SAVE AS</a>&nbsp;</TD>
				<TD NOWRAP ALIGN="center" WIDTH="20%"><a title="SEND" href="send.asp?f=<%=TheFiles(i)(0)%>">SEND</a>&nbsp;</TD>-->
			</TR>
			<%
		Next
		%>
	</TABLE>
    </DIV>
	<%
End If
%>

<% 
If (Action = "ViewFile") Then
	%>
	<BR><BR>
	
	<DIV align="left" style="WIDTH: 100%">
	<%
	FilePath = Request.QueryString("FilePath")
	%>
	<DIV align="left" style="WIDTH: 100%">
		<b>SMS File: <%=FilePath%></b>
	</DIV>
	<%
	FilePath = Server.MapPath(".") & "\bin\" & FilePath
	
	Set FSO = Server.CreateObject("Scripting.FileSystemObject")

	'If the file doesn't exist give an error message.
	If Not (FSO.FileExists(FilePath)) Then
		Response.Write("File " & FilePath & " does not exist! Read aborted.")
		'Otherwise, open it and display the contents.
	Else
		Set FSOFile = FSO.OpenTextFile(FilePath, 1, True)
		
		'Response.Write("Reading <b>" & FilePath & "<br></b>")
		Response.Write("<pre>")
		
		Do While Not FSOFile.AtEndOfStream
			Response.Write(FSOFile.ReadLine & vbCrLf)
		Loop
		
		Response.Write("</pre>")
		FSOFile.Close()
		
		Set FSOFile = Nothing 
		Set FSO = Nothing
	End If
	%>
	</DIV>
	<BR>
	<input type ="button" name ="back" value="<< Back" onclick ="javascript:history.back(1);">
	<%
End If
%>

</FORM>
	
</BODY>
</HTML>
