<%OPTION EXPLICIT
		
			Class CFieldDef
                    Public FieldName 
                    Public StartPos 
                    Public FieldSize 
            End Class
%>
<html>
	
	<head>
		<title>CDS Import</title>
	</head>
	
	<body>
	<!--#include file="..\libroutines.asp"-->
	<!--#include file="flUploader.asp"-->	
		<%
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
		
		Set fso = server.CreateObject("Scripting.FileSystemObject") 
		
		If (fso.FileExists(pathToFile) = True) Or (fso.FolderExists(pathToFile) = True) Then
			msgExists = "Could not upload file. A file or folder with that name already exists"
		Else		
			'validate file
			Dim brokerCode

			brokerCode = mid(filename,4,2)
			if brokerCode = Session("BrokerCode") then
					'check whether the file has been imported before
					Dim rs
					Dim sqlStr
					Dim conn
					
					Set conn = GetActiveConnection("KBroker")
					
					sqlStr = "SELECT * FROM _CDS_Imported_Files_ WHERE CDSFileName = '" & filename & "' AND CDSFileImported = 1"
					Set Rs = conn.Execute (SQLServerFormat(HandleQuote(sqlStr)))
					
					if not(rs.BOF or rs.EOF) then
							msgExists = "The selected file has already been imported"
					else
							'Clear previously failed attempts
							sqlStr = "DELETE FROM _CDS_Imported_Files_ WHERE CDSFileName = '" & Replace(filename,"'","''") & "'"
							conn.Execute SQLServerFormat(sqlStr)
							
							'record the file
							
							sqlStr = "INSERT INTO _CDS_Imported_Files_ (CDSFileName,CDSFileImported) SELECT " & _
							        "       " & "'" & Replace(filename,"'","''") & "'" & " as CDSFileName" & _
							        "       ," & " " & 0 & " " & " as CDSFileImported"	
							
							conn.Execute SQLServerFormat(sqlStr) 
							
							' Uploading file data
							Dim fileUploaded
							fileUploaded = load.saveToFile ("sourcefile", pathToFile)
					
					end if
			else
					msgExists = "The selected file contains trades for a broker other than " & Session("BrokerName")
			end if

		End If
		
		' destroy load object
		Set load = Nothing	

		If (fileUploaded = True) Then
			'save CDS info in db
			Dim objConn, schemaRs, objRS,   sheetName
			Dim txtStream, CDSData, dataStrip, newLineChecker, i
			Dim CDSDLL
			
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
			
			
			'delete previous working file
			'fso.DeleteFile pathToFile, True			
			Set fso = Nothing
			
			'swap new file
			
			pathToFile = strFilename
			
			
			'import_UDLPath = Server.MapPath("/") & "\UDL\KBroker.UDL"
			
			'Set CDSDLL = Server.CreateObject("CDS.CImportEngine")
			'CDSDLL.TradeFilePath = pathToFile
			'CDSDLL.UDLPath = import_UDLPath 
			'CDSDLL.ImportTrades 			
			'Set CDSDLL = Nothing    
'***********************************************************************************************
'					IMPORT FILE INTO DB
'***********************************************************************************************
			
            'Clear previous imports
            sqlStr = "Delete From _CDS_IMPORTED_TRADES_"
            conn.Execute SQLServerFormat(sqlStr)
            'response.end
            Const FIELD_COUNT = 29
            
            Const TRADE_DATE = 1
            Const TRADE_TIME = 2
            Const CDS_REF = 3
            Const BUY_SELL = 4
            Const PARTICIPANT_ID = 5
            Const PARTICIPANT_TYPE = 6
            Const CLIENT_PREFIX = 7
            Const CLIENT_SUFFIX = 8
            Const CLIENT_ACCOUNT_NO = 9
            Const CATEGORY = 10
            Const ISSUER_CODE = 11
            Const DEBT_TYPE = 12
            Const SECURITY_DESCRIPTION = 13
            Const MAIN_TYPE = 14
            Const SUB_TYPE = 15
            Const QUANTITY = 16
            Const PRICE = 17
            Const BROKERAGE_FEE = 18
            Const CDS_FEE = 19
            Const INTEREST = 20
            Const STATUS = 21
            Const CONTRA_BROKER_ID = 22
            Const CONTRA_BROKER_TYPE = 23
            Const CROSSING = 24
            Const SETTLEMENT_AMOUNT = 25
            Const SETTLEMENT_DATE = 26
            Const TRADE_TYPE = 27
            Const RESERVED = 28
            Const RECORD_TYPE = 29
            
            Dim FIELD_DEFS(29) 
			
			For i = 1 To FIELD_COUNT
			        Select Case i
			                Case 1
			                         SetupFieldDef FIELD_DEFS, i, "TradeDate", 1, 8
			                Case 2
			                         SetupFieldDef FIELD_DEFS, i, "TradeTime", 9, 19
			                Case 3
			                         SetupFieldDef FIELD_DEFS, i, "CDSRef", 28, 10
			                Case 4
			                         SetupFieldDef FIELD_DEFS, i, "BuySell", 38, 1
			                Case 5
			                         SetupFieldDef FIELD_DEFS, i, "ParticipantID", 39, 4
			                Case 6
			                         SetupFieldDef FIELD_DEFS, i, "ParticipantType", 43, 1
			                Case 7
			                         SetupFieldDef FIELD_DEFS, i, "ClientPrefix", 44, 13
			                Case 8
			                         SetupFieldDef FIELD_DEFS, i, "ClientSuffix", 57, 2
			                Case 9
			                         SetupFieldDef FIELD_DEFS, i, "ClientAccountNo", 59, 2
			                Case 10
			                         SetupFieldDef FIELD_DEFS, i, "Category", 61, 1
			                Case 11
			                         SetupFieldDef FIELD_DEFS, i, "IssuerCode", 62, 4
			                Case 12
			                         SetupFieldDef FIELD_DEFS, i, "DebtType", 66, 2
			                Case 13
			                         SetupFieldDef FIELD_DEFS, i, "SecurityDescription", 68, 30
			                Case 14
			                         SetupFieldDef FIELD_DEFS, i, "MainType", 98, 1
			                Case 15
			                         SetupFieldDef FIELD_DEFS, i, "SubType", 99, 4
			                Case 16
			                         SetupFieldDef FIELD_DEFS, i, "Quantity", 103, 19
			                Case 17
			                         SetupFieldDef FIELD_DEFS, i, "Price", 122, 10
			                Case 18
			                         SetupFieldDef FIELD_DEFS, i, "BrokerageFee", 132, 17
			                Case 19
			                         SetupFieldDef FIELD_DEFS, i, "CDSFee", 149, 17
			                Case 20
			                         SetupFieldDef FIELD_DEFS, i, "Interest", 166, 13
			                Case 21
			                         SetupFieldDef FIELD_DEFS, i, "Status", 179, 1
			                Case 22
			                         SetupFieldDef FIELD_DEFS, i, "ContraBrokerID", 180, 4
			                Case 23
			                         SetupFieldDef FIELD_DEFS, i, "ContraBrokerType", 184, 1
			                Case 24
			                         SetupFieldDef FIELD_DEFS, i, "Crossing", 185, 1
			                Case 25
			                         SetupFieldDef FIELD_DEFS, i, "SettlementAmount", 186, 32
			                Case 26
			                         SetupFieldDef FIELD_DEFS, i, "SettlementDate", 218, 8
			                Case 27
			                         SetupFieldDef FIELD_DEFS, i, "TradeType", 226, 1
			                Case 28
			                         SetupFieldDef FIELD_DEFS, i, "Reserved", 227, 573
			                Case 29
			                         SetupFieldDef FIELD_DEFS, i, "RecordType", 800, 1

			                Case Else%>
										<script language="JavaScript">
											window.parent.self.alert ('Invalid field index')
										</Script>
			
									<%
			        End Select
			Next
			        
			If i < FIELD_COUNT Then%>
						<script language="JavaScript">
							window.parent.self.alert ('The field initialization has not covered all declared fields')
						</Script>
			
					<%
			End If
			
			'do import
			Dim txtLine 
        
        
			Set fso = Server.CreateObject("Scripting.FileSystemObject")
			Set txtStream = fso.GetFile(pathToFile).OpenAsTextStream(1) 'open for reading..
       
			conn.BeginTrans
			        Do Until txtStream.AtEndOfStream
			                txtLine = txtStream.ReadLine
			                if len(trim(txtLine)) = 800 then
                                    sqlStr = "INSERT INTO [_CDS_Imported_Trades_]" & _
                                           "(TradeDate,TradeTime,CDSRef,BuySell,ParticipantID,ParticipantType," & _
                                           "ClientPrefix,ClientSuffix,ClientAccountNo,Category,IssuerCode," & _
                                           "DebtType,SecurityDescription,MainType,SubType,Quantity,Price," & _
                                           "BrokerageFee,CDSFee,Interest,Status,ContraBrokerID,ContraBrokerType," & _
                                           "Crossing,SettlementAmount,SettlementDate,TradeType,Reserved,RecordType) "
                                             ' " " & "'" & Replace(GetFieldValue(FIELD_DEFS, TRADE_DATE, txtLine), "'", "''") & "'" & "as TradeDate" & _
                                    sqlStr = sqlStr & "SELECT" & _
                                               " " & "'" & Replace(GetFieldValue(FIELD_DEFS, TRADE_DATE, txtLine), "'", "''") & "'" & "as TradeDate" & _
                                               "," & "'" & Replace(GetFieldValue(FIELD_DEFS, TRADE_TIME, txtLine), "'", "''") & "'" & "as TradeTime" & _
                                               "," & "'" & Replace(GetFieldValue(FIELD_DEFS, CDS_REF, txtLine), "'", "''") & "'" & "as CDSRef" & _
                                               "," & "'" & Replace(GetFieldValue(FIELD_DEFS, BUY_SELL, txtLine), "'", "''") & "'" & "as BuySell" & _
                                               "," & "'" & Replace(GetFieldValue(FIELD_DEFS, PARTICIPANT_ID, txtLine), "'", "''") & "'" & "as ParticipantID" & _
                                               "," & "'" & Replace(GetFieldValue(FIELD_DEFS, PARTICIPANT_TYPE, txtLine), "'", "''") & "'" & "as ParticipantType" & _
                                               "," & "'" & Replace(GetFieldValue(FIELD_DEFS, CLIENT_PREFIX, txtLine), "'", "''") & "'" & "as ClientPrefix" & _
                                               "," & "'" & Replace(GetFieldValue(FIELD_DEFS, CLIENT_SUFFIX, txtLine), "'", "''") & "'" & "as ClientSuffix" & _
                                               "," & "'" & Replace(GetFieldValue(FIELD_DEFS, CLIENT_ACCOUNT_NO, txtLine), "'", "''") & "'" & "as ClientAccountNo" & _
                                               "," & "'" & Replace(GetFieldValue(FIELD_DEFS, Category, txtLine), "'", "''") & "'" & "as Category" & _
                                               "," & "'" & Replace(GetFieldValue(FIELD_DEFS, ISSUER_CODE, txtLine), "'", "''") & "'" & "as IssuerCode" & _
                                               "," & "'" & Replace(GetFieldValue(FIELD_DEFS, DEBT_TYPE, txtLine), "'", "''") & "'" & "as DebtType" & _
                                               "," & "'" & Replace(GetFieldValue(FIELD_DEFS, SECURITY_DESCRIPTION, txtLine), "'", "''") & "'" & "as SecurityDescription" & _
                                               "," & "'" & Replace(GetFieldValue(FIELD_DEFS, MAIN_TYPE, txtLine), "'", "''") & "'" & "as MainType"
                                    sqlStr = sqlStr & "" & _
                                               "," & "'" & Replace(GetFieldValue(FIELD_DEFS, SUB_TYPE, txtLine), "'", "''") & "'" & "as SubType" & _
                                               "," & "'" & Replace(GetFieldValue(FIELD_DEFS, QUANTITY, txtLine), "'", "''") & "'" & "as Quantity" & _
                                               "," & "'" & Replace(GetFieldValue(FIELD_DEFS, PRICE, txtLine), "'", "''") & "'" & "as Price" & _
                                               "," & "'" & Replace(GetFieldValue(FIELD_DEFS, BROKERAGE_FEE, txtLine), "'", "''") & "'" & "as BrokerageFee" & _
                                               "," & "'" & Replace(GetFieldValue(FIELD_DEFS, CDS_FEE, txtLine), "'", "''") & "'" & "as CDSFee" & _
                                               "," & "'" & Replace(GetFieldValue(FIELD_DEFS, INTEREST, txtLine), "'", "''") & "'" & "as Interest" & _
                                               "," & "'" & Replace(GetFieldValue(FIELD_DEFS, Status, txtLine), "'", "''") & "'" & "as Status" & _
                                               "," & "'" & Replace(GetFieldValue(FIELD_DEFS, CONTRA_BROKER_ID, txtLine), "'", "''") & "'" & "as ContraBrokerID" & _
                                               "," & "'" & Replace(GetFieldValue(FIELD_DEFS, CONTRA_BROKER_TYPE, txtLine), "'", "''") & "'" & "as ContraBrokerType" & _
                                               "," & "'" & Replace(GetFieldValue(FIELD_DEFS, CROSSING, txtLine), "'", "''") & "'" & "as Crossing" & _
                                               "," & "'" & Replace(GetFieldValue(FIELD_DEFS, SETTLEMENT_AMOUNT, txtLine), "'", "''") & "'" & "as SettlementAmount" & _
                                               "," & "'" & Replace(GetFieldValue(FIELD_DEFS, SETTLEMENT_DATE, txtLine), "'", "''") & "'" & "as SettlementDate" & _
                                               "," & "'" & Replace(GetFieldValue(FIELD_DEFS, TRADE_TYPE, txtLine), "'", "''") & "'" & "as TradeType" & _
                                               "," & "'" & Replace(GetFieldValue(FIELD_DEFS, RESERVED, txtLine), "'", "''") & "'" & "as Reserved" & _
                                               "," & "'" & Replace(GetFieldValue(FIELD_DEFS, RECORD_TYPE, txtLine), "'", "''") & "'" & "as RecordType" '& _
                                               '"    FROM [_CDS_Imported_Trades_]"
                                     
                                    conn.Execute SQLServerFormat(sqlStr)
                              end if


			        Loop
			conn.CommitTrans
        
			txtStream.Close
			Set txtStream = Nothing
			Set fso = Nothing
			
			Sub SetupFieldDef(FldArray() , FldIndex , FldName , FldStartPos , FldSize )
			        set FldArray(FldIndex) = new CFieldDef
			        FldArray(FldIndex).FieldName = FldName
			        FldArray(FldIndex).StartPos = FldStartPos
			        FldArray(FldIndex).FieldSize = FldSize
			End Sub

			Function GetFieldValue(FldArray() , FldIndex , TextLine )
			        Dim theValue 
			        
			        theValue = Mid(TextLine, FldArray(FldIndex).StartPos, FldArray(FldIndex).FieldSize)
			        GetFieldValue = theValue
			End Function

'***********************************************************************************************
'					END IMPORT FILE INTO DB
'***********************************************************************************************		
		
			'mark file as imported
			sqlStr = "UPDATE _CDS_Imported_Files_ SET " & _
					"       CDSFileImported = 1 WHERE CDSFileName = " & "'" & filename & "'" 	
			conn.Execute SQLServerFormat(HandleQuote(sqlStr))
			
			'destroy objects
			Set Rs = Nothing
			Set Conn = Nothing
			Set objRs = Nothing
			Set objConn = Nothing
			
			'delete the uploaded xl file	
			On Error Resume Next		
			fso.DeleteFile pathToFile%>
				<script language="JavaScript">
					window.parent.self.alert ('The information has been imported successfully.');
					window.location.replace("CDSMatchedTradesList.asp");
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