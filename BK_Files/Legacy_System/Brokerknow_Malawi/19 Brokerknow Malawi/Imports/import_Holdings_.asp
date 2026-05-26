<%OPTION EXPLICIT%>
<html>
	

	<head>
		<title>CDS Import</title>
	</head>
	
	<body>
	<!--#include virtual="libroutines.asp"-->
	<!--#include file="flUploader.asp"-->	
		<%
		
		Set Rs = Server.CreateObject("ADODB.Recordset")			
		Set Conn = Server.CreateObject("ADODB.Connection")
		Set Conn = GetActiveConnection("KBroker")

		' load object
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
		
		If (fso.FileExists(pathToFile) = True) Or (fso.FolderExists(pathToFile) = True) Then
			msgExists = "Could not upload file. A file or folder with that name already exists"
			fso.DeleteFile pathToFile
			Set fso = Nothing
		''response.redirect("importHoldings.asp")
		Else		
			'validate file
			'	check for proper broker
			'Dim filNameOnly
			
			'filNameOnly = 
			' Uploading file data
			Dim fileUploaded
			fileUploaded = load.saveToFile ("sourcefile", pathToFile)
			
		End If
		
		' destroy load object
		Set load = Nothing	
		
		If (fileUploaded = True) Then
			'save CDS info in db
			Dim objConn, schemaRs, objRS, Conn, Rs, sheetName
			Dim txtStream, CDSData, dataStrip, newLineChecker, i
			Dim CDSDLL
			
			Conn.Execute("Delete From _CDS_Imported_Holdings_")
			
			newLineChecker = vbCrLf
			
			'strip header and footer info using FSO...			
			Set txtStream = fso.GetFile(pathToFile).OpenAsTextStream(1) 'open for reading..  
			CDSData = txtStream.ReadAll
			Set txtStream = Nothing
			
				'strip header here...
				dataStrip = InStr(1, CDSData, newLineChecker)
				If dataStrip > 0 Then
					CDSData = Mid(CDSData, dataStrip + 2)
				End If 
				
				'strip footer here
				dataStrip = InStrRev(CDSData, newLineChecker)
				If dataStrip > 0 Then
					CDSData = Mid(CDSData, 1, dataStrip - 2) 
				End If
				
			'copy new data into file
			'the file is given a temporary unique name before saving to the server.
			Dim guidStr, strFilename, guid
			set guid = server.createobject("NDUtils.CGUID")
			guidStr = guid.GenerateGUID
			guidStr = Replace(Replace(guidStr, "{", ""), "}", "")
			strFilename = Server.MapPath(CurrentDirectory) & "\" & guidStr & ".txt"			
			Set txtStream = fso.CreateTextFile(strFilename, True)
			txtStream.Write CDSData
			txtStream.Close
			Set txtStream = Nothing
			
			
			
			'swap new file
			pathToFile = strFilename
			
			
			Set Rs = Server.CreateObject("ADODB.Recordset")			
			Set Conn = Server.CreateObject("ADODB.Connection")
			Set Conn = GetActiveConnection("KBroker")
			
			
			import_UDLPath = Server.MapPath("/") & "\UDL\KBroker.UDL"
			
			Set CDSDLL = Server.CreateObject("CDSHoldings.ImportHoldings")
				CDSDLL.importCds ""& pathToFile ,""& import_UDLPath	,"_CDS_Imported_Holdings_"	
			Set CDSDLL = Nothing   

'delete previous working file
			fso.DeleteFile pathToFile, True			
							
			
			'destroy objects
			Set Rs = Nothing
			Set Conn = Nothing
			Set objRs = Nothing
			Set objConn = Nothing
			
			'delete the uploaded xl file	
			On Error Resume Next	
			
			fso.DeleteFile pathToFile, True
			Set fso = Nothing
			%>
				<script language="JavaScript">
					window.parent.self.alert ('The information has been imported successfully.');
					window.location.replace("importHoldings.asp");
				</Script>
			<%
			

		Else
			Dim str_message
			If msgExists = "" Then
				str_message = fileName & " could not be uploaded."
			Else
				str_message = msgExists
			End If%>
				<script language="JavaScript">
					window.parent.self.alert ('An error occured:\n<%= str_message %>')
				</Script>
			
			<%			
		End If
				
		Set fso = Nothing
		%>
	
	</body>
</html>