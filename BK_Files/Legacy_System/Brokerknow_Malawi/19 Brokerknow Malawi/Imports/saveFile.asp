<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE>REF File Upload</TITLE>
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT>
</head>
<!--#include virtual="libroutines.asp"-->
<body class="Reports" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!--#INCLUDE FILE="clsUpload.asp"-->

<%

'the file is given a temporary unique name before saving to the server.
Dim guidStr 

set guid = server.createobject("NDUtils.CGUID")
guidStr = guid.GenerateGUID
strFilename = guidStr & ".txt"


Dim Upload
Dim FileName
Dim Folder

Set Upload = New clsUpload

' Grab the file name
FileName = strFilename 'Upload.Fields("File").FileName

' Get path to save file to
Folder = Server.MapPath(".") & "\"

' Save the binary data to the file system
Upload("File").SaveAs Folder & FileName

' Release upload object from memory
Set Upload = Nothing


				Dim Fso, fl, ts
				
				strFilePath = Folder & FileName
				
				Set Fso = Server.CreateObject("Scripting.FileSystemObject")

				
				Set fl = Fso.GetFile(strFilePath)      
				
				Set ts = fl.OpenAsTextStream
				
				Do Until ts.AtEndOfStream
					Response.Write Replace(ts.ReadLine, Chr(32), "&nbsp;")				
					Response.Write "<br>"
				Loop
				
				'Response.Write ts.ReadAll


				Set ts = Nothing
				Set fl = Nothing
				Set Fso = Nothing	
	
%>
  
   </body>
  </html>