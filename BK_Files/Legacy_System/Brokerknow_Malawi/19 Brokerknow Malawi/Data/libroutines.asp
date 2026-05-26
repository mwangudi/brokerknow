<!-- METADATA TYPE="typelib" UUID="00000200-0000-0010-8000-00AA006D2EA4" NAME="ADO Type Library"-->
<%
Dim intPageSize 
	
If IsNumeric(Session("screenHeight")) Then
	If CInt(Session("screenHeight")) <= 600 Then
		intPageSize = 20
	ElseIf	CInt(Session("screenHeight")) <= 800 Then
		intPageSize = 30
	ElseIf	CInt(Session("screenHeight")) <= 1024 Then
		intPageSize = 44	
	Else
		intPageSize = 54		
	End If
Else
	'default
	intPageSize = 30
End If
	
WriteClientMsgBoxScript

WriteClientFormatNumScript

Function RoundPoint05(theNum)
	If Not IsNumeric(theNum) Then
		RoundPoint05 = theNum
		Exit Function
	End If	
	
	raw_number = theNum
	raw_number1 = raw_number * 100 
	raw_number2 = Int(raw_number * 10) * 10
	raw_number3 = raw_number1 - raw_number2 

	If raw_number3 < 2.5 Then
	    raw_number4 = raw_number2 + 0
	ElseIf raw_number3 < 7.5 Then
	    raw_number4 = raw_number2 + 5
	Else
	    raw_number4 = raw_number2 + 10
	End If

	raw_number5 = CDbl(raw_number4 / 100)


	RoundPoint05 = raw_number5

End Function


Function WriteClientFormatNumScript

	With Response
		.write "<Script Language=""VBScript"">" & vbCrLf
		.write "Function FormatNum(theNum) " & vbCrLf		
		.write "	Dim tempNum " & vbCrLf
		.write "	On Error Resume Next " & vbCrLf
		.write "	tempNum = FormatNumber(theNum, 2) " & vbCrLf		
		.write "	If Err.Number > 0 Then " & vbCrLf
		.write "		FormatNum = theNum " & vbCrLf
		.write "		Err.Clear " & vbCrLf
		.write "	Else " & vbCrLf
		.write "		If tempNum < 0 Then " & vbCrLf
		.write "			FormatNum = ""("" & FormatNumber(Abs(tempNum), 2) & "")"" " & vbCrLf 
		.write "		Else	 " & vbCrLf
		.write "			FormatNum = tempNum " & vbCrLf
		.write "		End If	 " & vbCrLf
		.write "	End If	 " & vbCrLf
		.Write "End Function " & vbCrLf
				
		.write "Function FormatNumCommasOnly(theNum) " & vbCrLf		
		.write "	Dim tempNum " & vbCrLf
		.write "	On Error Resume Next " & vbCrLf
		.write "	tempNum = FormatNumber(theNum, 0) " & vbCrLf		
		.write "	If Err.Number > 0 Then " & vbCrLf
		.write "		FormatNumCommasOnly = theNum " & vbCrLf
		.write "		Err.Clear " & vbCrLf
		.write "	Else " & vbCrLf
		.write "		If tempNum < 0 Then " & vbCrLf
		.write "			FormatNumCommasOnly = ""("" & FormatNumber(Abs(tempNum), 0) & "")"" " & vbCrLf 
		.write "		Else	 " & vbCrLf
		.write "			FormatNumCommasOnly = tempNum " & vbCrLf
		.write "		End If	 " & vbCrLf
		.write "	End If	 " & vbCrLf
		.Write "End Function " & vbCrLf
		.write "</Script>	 " & vbCrLf
		
		
		
	End With
		
End Function


Function WriteClientMsgBoxScript
	With Response
		.write "<Script Language=""JavaScript"">"  & vbCrLf
		.write "	function TryFramedMsg(msg){" & vbCrLf
		.write "		try{" & vbCrLf
		.write "			window.parent.editDocOpener.alert(msg);" & vbCrLf
		.write "			return true;" & vbCrLf
		.write "		}" & vbCrLf
		.write "		catch(e){return false;}" & vbCrLf
		.write "		}" & vbCrLf
		.write "	function ShowMessage(msg){" & vbCrLf
		.Write "		if (TryFramedMsg(msg)==false) window.parent.alert(msg);" & vbCrLf
		.write "	}" & vbCrLf
        .write "</Script>" & vbCrLf
	End With	

End Function

Function WriteDialogCancelScript
	With Response
		.write "<Script Language=""JavaScript"">" & vbCrLf
		.write "		try{" & vbCrLf
		.write "		window.opener.self.close();" & vbCrLf
		.write "		window.self.close();" & vbCrLf
		.write "	}" & vbCrLf
		.write "	catch(e){window.self.close();}	" & vbCrLf
        .write "</Script>" & vbCrLf
	End With	
End Function


Function WriteDeleteCloseScript
	With Response
		.write "<Script Language=""JavaScript"">" & vbCrLf
		.write "		try{" & vbCrLf
		.write "		window.opener.location.reload();  " & vbCrLf
		.write "		window.self.close();" & vbCrLf
		.write "	}" & vbCrLf
		.write "	catch(e){window.self.close();}	" & vbCrLf
        .write "</Script>" & vbCrLf
	End With	
End Function

Function WriteRefreshDialogScript
   		With Response
   			.write "<script language=""javascript"">" & vbCrLf
   			.write "	window.parent.doRefreshOpener();" & vbCrLf
   			.write "</script>" & vbCrLf
   		End With
   	
End Function

Function WriteDialogRelocateScript(pageTo)
	With Response
		.write "<Script Language=""JavaScript"">" & vbCrLf
		.write "		try{" & vbCrLf
		.write "		window.parent.relocateDocOpener('" & pageTo & "');" & vbCrLf
		.write "	}" & vbCrLf
		.write "	catch(e){window.parent.closeDocOpener();}	" & vbCrLf
        .write "</Script>" & vbCrLf
	End With
End Function

Function WritefraEnabledDialogCloseScript
	With Response
		.write "<Script Language=""JavaScript"">" & vbCrLf
		.write "		try{" & vbCrLf
		.write "		window.parent.parent.frames[""maininfo""].location.reload(); " & vbCrLf
		.write "		window.parent.closeDocOpener();" & vbCrLf
		.write "	}" & vbCrLf
		.write "	catch(e){window.parent.closeDocOpener();}	" & vbCrLf
        .write "</Script>" & vbCrLf
	End With
End Function


Function WriteDialogRefuseOpenScript
	With Response
		.write "<Script Language=""JavaScript"">" & vbCrLf
		.write "		try{" & vbCrLf
		.write "		window.self.close();" & vbCrLf
		.write "	}" & vbCrLf
		.write "	catch(e){window.self.close()}	" & vbCrLf
        .write "</Script>" & vbCrLf
	End With	
End Function

Function WriteDialogCloseScript
	With Response
		.write "<Script Language=""JavaScript"">" & vbCrLf
		.write "		try{" & vbCrLf
		.write "		window.parent.dialogArguments.opener.location.reload(); " & vbCrLf
		.write "		window.self.close();" & vbCrLf
		.write "	}" & vbCrLf
		.write "	catch(e){window.self.close()}	" & vbCrLf
        .write "</Script>" & vbCrLf
	End With	
End Function

Function GetActiveConnectionEx(theDBName) 
	'no access!! Only for the
	'grant access pages list (separate using commas)
	Const GRANT_ACCESS_PAGES_LIST = "validateLog.asp"
	
	If InStr(1, Request.ServerVariables("PATH_INFO"), GRANT_ACCESS_PAGES_LIST, vbTextCompare) < 1 Then%>
		<Script Language="JavaScript">
			window.parent.document.write ('<H2>Access Denied</H2>.<p><H4>An exclusive procedure (GetActiveConnectionEx) was accessed illegally.</H4>.')				
		</Script>
		<%Response.End
	End If
	
    Dim tmpConn 
    Set tmpConn = CreateObject("ADODB.Connection")

    tmpConn.ConnectionString = "FILE NAME=" & GetUDLPath(theDBName)
    tmpConn.CursorLocation = adUseClient
    tmpConn.Open
    
    Set GetActiveConnectionEx = tmpConn
End Function

Function GetActiveConnection(theDBName) 
    
    'no page should have access to this function unless user is logged in!
    
    If Session("UserID") = "" Then%>
		<Script Language="JavaScript">
			window.parent.document.write ('<HTML><HEAD><TITLE>Session Expired</TITLE></HEAD><BODY><H2>Access Denied: Session Expired</H2>.<p><h4>Your session has expired. Please close this browser and begin the login process again.</h4></p></BODY></HTML>');
		</Script>
		<%Response.End
    End If
    
    Dim tmpConn 
    Set tmpConn = CreateObject("ADODB.Connection")

    tmpConn.ConnectionString = "FILE NAME=" & GetUDLPath(theDBName)
    tmpConn.CursorLocation = adUseClient
    tmpConn.Open
    
    Set GetActiveConnection = tmpConn
End Function

Function FormatDate(theDate)
	On Error Resume Next
	FormatDate = Day(theDate) & "-" & MonthName(Month(theDate), True) & "-" & Year(theDate)
	If Err.Number > 0 Then
		FormatDate = theDate
	End If
End Function


Function WriteBackgroundImg
	For Each Thing In Request.ServerVariables
		Response.Write Thing & " = " & Request.ServerVariables(Thing) & "<BR>"
	Next
	
	imgPath = Request.ServerVariables("APPL_PHYSICAL_PATH") & "\images/bodyBack.jpg"
	
	imgPath = Server.MapPath(imgPath)
	

	With Response
		.Write "<Script Language=JavaScript>	" & vbCrLf
		.Write "	function resizeWin(){" & vbCrLf
		.Write "	var theDocImg = document.all.item(""BackImg"")" & vbCrLf
		.Write "	theDocImg.style.top = document.body.offsetTop  ;" & vbCrLf
		.Write	"	theDocImg.style.left = document.body.offsetLeft    ; 		" & vbCrLf
		.Write	"	theDocImg.height = document.body.offsetHeight ;" & vbCrLf
		.Write "	theDocImg.width = document.body.offsetWidth; 		" & vbCrLf
		.Write "	}" & vbCrLf
	
		.Write "window.onresize = resizeWin;	" & vbCrLf
		.Write " window.onscroll = resizeWin; " & vbCrLf
		.Write "  window.onload = resizeWin; " & vbCrLf
		.Write "</Script>" & vbCrLf
		.Write "<IMG id=BackImg style=""Z-INDEX: -1; POSITION: absolute"" src =""images/bodyBack.jpg"" align=absLeft>" & vbCrLf
	End With

End Function


Function GetUDLPath(theDBName) 
    Dim tmpStr
    
    tmpStr = StrReverse(Request.ServerVariables("APPL_PHYSICAL_PATH"))
    
    tmpStr = Mid(tmpStr, InStr(1, tmpStr, "\") + 1)
    
    tmpStr = StrReverse(tmpStr)
    
    GetUDLPath = tmpStr & "\UDL\" & Trim(theDBName) & ".UDL"

End Function


Function InitializeDefaultDB
	strConnectionString = GetActiveConnection("KBroker")
	
	Set objCn = Server.CreateObject("ADODB.Connection")

	With objCn
		.CursorLocation = 2
		.ConnectionTimeout = 15
		.CommandTimeout = 30
		.ConnectionString = strConnectionString
		.Open
	End With

End Function

Function Paging(ByVal intPage, ByVal intPageCount, ByVal intRecordCount)
	Dim strQueryString
	Dim strScript
	Dim intStart
	Dim intEnd
	Dim strRet
	Dim i

	If intPage > intPageCount Then
		intPage = intPageCount
	ElseIf intPage < 1 Then 
		intPage = 1
	End If
	
	If intRecordCount = 0 Then
		strRet = "No Records Found"
	ElseIf intPageCount = 1 Then
		strRet = "End Of Hits"
	Else
		strScript = ""
	
		If intPage <= 10 Then
			intStart = 1
		Else
			If (intPage Mod 10) = 0 Then
				intStart = intPage - 9
			Else
				intStart = intPage - (intPage Mod 10) + 1
			End If
		End If

		intEnd = intStart + 9
		If intEnd > intPageCount Then intEnd = intPageCount
	
		strRet = "Page " & intPage & " of " & intPageCount & ": "
	
		If intPage <> 1 Then 
			strRet = strRet & "<font face=Arial title=""Previous"" OnClick=""JavaScript: Paging('" & strScript
			strRet = strRet &  intPage - 1 & "')"
			strRet = strRet & """>&lt;&lt;Prev&nbsp;&nbsp;</font> "
		End If
	
		For i = intStart To intEnd
			If i = intPage Then
				strRet = strRet & "<font color=silver><b>" & i & "&nbsp;&nbsp;</b></font> "
			Else
				strRet = strRet & "<font face=Arial title=""Page " & i & " of " & intEnd & """ OnClick=""JavaScript: Paging('" & strScript
				strRet = strRet  & i & "')"
				strRet = strRet & """><b>" & i & "&nbsp;&nbsp;</b></font>"
				If i <> intEnd Then strRet = strRet & " "
			End If
		Next
	
		If intPage <> intPageCount Then
			strRet = strRet & " <font face=Arial title=""Next"" OnClick=""JavaScript: Paging('" & strScript
			strRet = strRet & intPage + 1 & "')"
			strRet = strRet & """>Next&gt;&gt;</font> "
		End If
	End If
	
	strRet = strRet & "<input type = 'hidden' name ='Page' id=Page value=1>"
	strRet = strRet & "<Script Language=JavaScript>" & Chr(13)
	strRet = strRet & "		window.onload = resizePagingDisplay; window.onresize = resizePagingDisplay; window.document.onscroll = resizePagingDisplay; " & Chr(13)
	strRet = strRet & "</script>"
	
	Paging = strRet
End Function


Function IIf(ByVal Expr, ByVal TruePart, ByVal FalsePart)
	If (Expr) Then
		IIf = TruePart
	Else
		IIf = FalsePart
	End If
End Function


Function IntToNull(ByVal varValue)
	If IsNull(varValue) Then
		IntToNull = Null
	ElseIf IsEmpty(varValue) Then
		IntToNull = Null
	ElseIf Not IsNumeric(varValue) Then
		IntToNull = Null
	Else
		IntToNull = CLng(varValue)
	End If
End Function

Function CharToNull(ByVal varValue)
	If IsNull(varValue) Then
		CharToNull = Null
	ElseIf IsEmpty(varValue) Then
		CharToNull = Null
	ElseIf varValue = "" Then
		CharToNull = Null
	Else
		CharToNull = CStr(varValue)
	End if
End Function

Function BitToNull(ByVal varValue)
	If IsNull(varValue) Then
		BitToNull = Null
	ElseIf IsEmpty(varValue) Then
		BitToNull = Null
	ElseIf Not IsNumeric(varValue) Then
		BitToNull = Null
	ElseIf CInt(varValue) = 0 Then
		BitToNull = 0
	Else
		BitToNull = 1
	End If
End Function

Function BoolToNull(ByVal varValue)
	If IsNull(varValue) Then
		BoolToNull = Null
	ElseIf IsEmpty(varValue) Then
		BoolToNull = Null
	Else 
		Select Case LCase(CStr(varValue))
			Case "0", "false", "no", "off"
				BoolToNull = False
			Case Else
				BoolToNull = True
		End Select
	End If
End Function

Function DateToNull(ByVal varValue)
	If IsNull(varValue) Then
		DateToNull = Null
	ElseIf IsEmpty(varValue) Then
		DateToNull = Null
	ElseIf Not IsDate(varValue) Then
		DateToNull = Null
	Else
		DateToNull = CDate(varValue)
	End If
End Function

Function DblToNull(ByVal varValue)
	If IsNull(varValue) Then
		DblToNull = Null
	ElseIf IsEmpty(varValue) Then
		DblToNull = Null
	ElseIf Not IsNumeric(varValue) Then
		DblToNull = Null
	Else
		DblToNull = CDbl(varValue)
	End If
End Function



Function EncryptWithALP(strData)
    Dim strALPKey
    Dim strALPKeyMask
    Dim lngIterator
    Dim blnOscillator
    Dim strOutput
    Dim lngHex

    Const lngALPKeyLength = 8

    If Len(strData) = 0 Then
        Exit Function
    End If
    Randomize
    For lngIterator = 1 To lngALPKeyLength
        strALPKey = strALPKey + Trim(Hex(Int(16 * Rnd)))
        strALPKeyMask = strALPKeyMask + Trim(Int(2 * Rnd))
    Next 
    
    lngIterator = 0
    Do Until Len(strData) = 0
        blnOscillator = Not blnOscillator
        lngIterator = lngIterator + 1
        If lngIterator > lngALPKeyLength Then
            lngIterator = 1
        End If
         
        If blnOscillator Then
			lngHex = CLng(Asc(Left(strData, 1)) + Asc(Mid(strALPKey, lngIterator, 1)))
		Else
			lngHex = CLng(Asc(Left(strData, 1)) - Asc(Mid(strALPKey, lngIterator, 1)))
		End	If
                
        If lngHex > 255 Then
            lngHex = lngHex - 255
        ElseIf lngHex < 0 Then
            lngHex = lngHex + 255
        End If
        strOutput = strOutput + Right(String(2, "0") + Hex(lngHex), 2)
        strData = Right(strData, Len(strData) - 1)
    Loop
    For lngIterator = 1 To lngALPKeyLength
        If Mid(strALPKeyMask, lngIterator, 1) = "1" Then
            strOutput = Mid(strALPKey, lngIterator, 1) + strOutput
        Else
            strOutput = strOutput + Mid(strALPKey, lngIterator, 1)
        End If
    Next
    EncryptWithALP = Right(String(2, "0") + Hex(BinaryToDouble(strALPKeyMask)), 2) + strOutput
End Function

Function BinaryToDouble(ByVal strData)
    Dim dblOutput
    Dim lngIterator
    Dim tempDblOutPut
    Do Until Len(strData) = 0		
		dblOutput = dblOutput + IIf(Right(strData, 1) = "1", (2 ^ lngIterator), 0)
        
        strData = Left(strData, Len(strData) - 1)
        lngIterator = lngIterator + 1
    Loop
    
    BinaryToDouble = dblOutput
End Function

Function DoubleToBinary(ByVal dblData)
    Dim strOutput
    Dim lngIterator
    Dim tempStrOutPut
    Do Until (2 ^ lngIterator) > dblData
		strOutput = IIf(((2 ^ lngIterator) And dblData) > 0, "1", "0") + strOutput
        lngIterator = lngIterator + 1
    Loop
    DoubleToBinary = strOutput
End Function

Function DecryptWithALP(strData)
    Dim strALPKey
    Dim strALPKeyMask
    Dim lngIterator
    Dim blnOscillator
    Dim strOutput
    Dim lngHex
    Const lngALPKeyLength = 8	
    
    If Len(strData) = 0 Then
        Exit Function
    End If
    
    strALPKeyMask = Right(String(lngALPKeyLength, "0") + DoubleToBinary(CLng("&H" + Left(strData, 2))), lngALPKeyLength)
    strData = Right(strData, Len(strData) - 2)
    For lngIterator = lngALPKeyLength To 1 Step -1
        If Mid(strALPKeyMask, lngIterator, 1) = "1" Then
            strALPKey = Left(strData, 1) + strALPKey
            strData = Right(strData, Len(strData) - 1)
        Else
            strALPKey = Right(strData, 1) + strALPKey
            strData = Left(strData, Len(strData) - 1)
        End If
    Next
    
    lngIterator = 0
    Do Until Len(strData) = 0
        blnOscillator = Not blnOscillator
        lngIterator = lngIterator + 1
        If lngIterator > lngALPKeyLength Then
            lngIterator = 1
        End If
		If blnOscillator Then
			lngHex = CLng("&H" + Left(strData, 2) - Asc(Mid(strALPKey, lngIterator, 1)))
		Else
			lngHex = CLng("&H" + Left(strData, 2) + Asc(Mid(strALPKey, lngIterator, 1)))
		End If	
        
        If lngHex > 255 Then
            lngHex = lngHex - 255
        ElseIf lngHex < 0 Then
            lngHex = lngHex + 255
        End If
        
        strOutput = strOutput + Chr(lngHex)
        strData = Right(strData, Len(strData) - 2)
    Loop
    
    DecryptWithALP = strOutput
End Function

Function HandleQuote(sqlStr)
        Dim RE 
        Dim matches 
        Dim match 
        Dim i 
        Dim subStr 
        Dim strArray 
        Set RE = server.createobject("VBScript.RegExp")
        
        RE.Pattern = "([^,\s,',)]'[^,\s,',)])"
        Do
            Set matches = RE.Execute(sqlStr)
            For Each match In matches
                    subStr = match.Value
                    strArray = Split(subStr, "'")
                    subStr = Join(strArray, "''")
                    sqlStr = Replace(sqlStr, match.Value, subStr, 1, 1)
            Next
        Loop Until matches.Count = 0
        
        HandleQuote = sqlStr
End Function


Function UpdateAuditTrail(sqlStr)
	Dim sqlParse, auditSQLStr, newIDRs, newID
	Dim i, thisField, thisFieldValue
	Set sqlParse = CreateObject("SQLParser.CSQLParser")	
	
  
    sqlParse.ParseSQL (sqlstr)
    
    If sqlParse.SQLType = "" Then
		Set sqlParse = Nothing
		Exit Function
    End If
    
    auditSQLStr =  "INSERT INTO AuditTrail (UserID, AuditTrailAction, AuditTrailTable, AuditTrailSQL, AuditTrailCriteria) " & _
				" SELECT '" & Session("UserID") & "', '" & sqlParse.SQLType & "', '" & sqlParse.TableName  & "', '" & Replace (sqlStr, "'", "''") & "', '" & Replace (sqlParse.Criteria, "'", "''") & "'"
			
				
    conn.Execute auditSQLStr
    
    Set newIDRs = conn.Execute ("SELECT SCOPE_IDENTITY() AS [SCOPE_IDENTITY]")
    If Not (newIDRs.EOF or newIDRs.BOF) Then
		newID = newIDRs.Fields("SCOPE_IDENTITY").Value
		For i = 1 To sqlParse.SQLFields.Count
			thisField = sqlParse.SQLFields(i).FieldName 
			thisFieldValue = sqlParse.SQLFields(i).FieldValue 
			
			auditSQLStr = "INSERT INTO AuditTrailItem  " & _
						" (AuditTrail_DPA_, AuditTrailItemField, AuditTrailItemValue) " & _
						"	SELECT     " & newID & ", '" & thisField & "', '" & Replace(thisFieldValue, "'", "''") & "'"
			On Error REsume Next			
			Conn.Execute auditSQLStr 
			If Err Then
				Response.Write "sql str caused an unexpected error: <br>" & Err.Description & "<br>" & auditSQLStr
				Response.End
			End If
							
		Next  
    
    End If
    
    Set sqlParse = Nothing

End Function

Function SQLServerFormatWithCustomMax(sqlStr)
        Const startStr = "iif(isnull(max("
        Const midStr1 = ")),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = '"
        Const midStr2 = "'),max("
        Const endStr = ") + 1)"
        Const SQLstartStr = "ISNULL(MAX("
        Const SQLmidStr = ") + 1,(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = '"
        Const SQLendStr = "'))"
        Dim startStrPos 
        Dim midStrPos1
        Dim midStrPos2 
        Dim endStrPos 
        Dim fldName 
        Dim fldNameStartPos 
        Dim tblName
        Dim tblNameStartPos
        Dim oldStr 
        Dim newStr
                
        sqlStr = Replace(sqlStr, "[", "")
        sqlStr = Replace(sqlStr, "]", "")
        sqlStr = Replace(sqlStr, "#", "'")
        sqlStr = Replace(sqlStr, vbCrLf, "' + CHAR(13) + CHAR(10) + '")
        
        'sqlStr = lcase(sqlStr)
        do while 1 < 2
				startStrPos = InStr(1, sqlStr, startStr, vbTextCompare)
				If startStrPos < 1 Then
				        UpdateAuditTrail sqlStr
				        SQLServerFormatWithCustomMax = sqlStr
				        Exit Function
				End If
        
				midStrPos1 = InStr(startStrPos, sqlStr, midStr1, vbTextCompare)
				fldNameStartPos = startStrPos + Len(startStr)
				fldName = Mid(sqlStr, fldNameStartPos, midStrPos1 - fldNameStartPos)
				midStrPos2 = InStr(midStrPos1, sqlStr, midStr2, vbTextCompare)
				tblNameStartPos = midStrPos1 + Len(midStr1)
				tblName = Mid(sqlStr, tblNameStartPos, midStrPos2 - tblNameStartPos)
        
				newStr = SQLstartStr & fldName & SQLmidStr & tblName & SQLendStr
        
				endStrPos = InStr(midStrPos2, sqlStr, endStr, vbTextCompare)
				oldStr = Mid(sqlStr, startStrPos, (endStrPos + Len(endStr) - startStrPos))
				sqlStr = Replace(sqlStr, oldStr, newStr)  
		loop      

End Function

Function SQLServerFormat(sqlStr)
        Const startStr = "iif(isnull(max("
        Const midStr = ")),1,max("
        Const endStr = ") + 1)"
        Const SQLstartStr = "ISNULL(MAX("
        Const SQLendStr = ")+1, 1)"
        Dim startStrPos 
        Dim midStrPos 
        Dim endStrPos 
        Dim fldName 
        Dim fldNameStartPos 
        Dim oldStr 
                
        sqlStr = Replace(sqlStr, "[", "")
        sqlStr = Replace(sqlStr, "]", "")
        sqlStr = Replace(sqlStr, "#", "'")
        sqlStr = Replace(sqlStr, vbCrLf, "' + CHAR(13) + CHAR(10) + '")
        
        'sqlStr = lcase(sqlStr)
        startStrPos = InStr(1, sqlStr, startStr, vbTextCompare)
        If startStrPos < 1 Then
                SQLServerFormat = sqlStr
                UpdateAuditTrail sqlStr
                Exit Function
        End If
        midStrPos = InStr(startStrPos, sqlStr, midStr, vbTextCompare)
        fldNameStartPos = startStrPos + Len(startStr)
        fldName = Mid(sqlStr, fldNameStartPos, midStrPos - fldNameStartPos)
        fldName = SQLstartStr & fldName & SQLendStr
        
        endStrPos = InStr(1, sqlStr, endStr, vbTextCompare)
        oldStr = Mid(sqlStr, startStrPos, (endStrPos + Len(endStr) - startStrPos))
        sqlStr = Replace(sqlStr, oldStr, fldName)        

        UpdateAuditTrail sqlStr
        
        SQLServerFormat = sqlStr
        

End Function

Function FormatNumEx(theNum, decimalPlaces)
	Dim tempNum
		
	If Not IsNumeric(theNum) Then
		If Trim(theNum) <> "" Then
			numBuffer = Trim(theNum)
			inStrNum = ""
			numLength = Len(numBuffer)
			If numLength > 1 Then
				If Left(numBuffer, 1) = "-" Then
					numBuffer = Mid(numBuffer, 2)
					numLength = Len(numBuffer)
				End If
			End If
			
			For z = 1 To numLength
				tempNumTest = Mid(numBuffer, z, 1) 
				If IsNumeric(tempNumTest) Or tempNumTest = "." Or tempNumTest = "," Then
					inStrNum = inStrNum & tempNumTest					
				Else
					Exit For	
				End If	
			Next
			
			If IsNumeric(inStrNum) Then				
				appendStr = Mid(numBuffer, z) 
				If inStrNum < 0 Then
					FormatNumEx = "(" & FormatNumber(Abs(inStrNum), decimalPlaces) & ")" & appendStr
				Else
					FormatNumEx = FormatNumber(inStrNum, decimalPlaces) & appendStr
				End If	
				Exit Function
			End If
			
		End If
		FormatNumEx = theNum
		Exit Function
	End If
	
	On Error Resume Next
		tempNum = FormatNumber(theNum, decimalPlaces)		
	If Err.Number > 0 Then
		FormatNumEx = theNum
		Err.Clear
	Else
		If tempNum < 0 Then
			FormatNumEx = "(" & FormatNumber(Abs(tempNum), decimalPlaces) & ")"
		Else
			FormatNumEx = tempNum
		End If	
	End If	
End Function


Function FormatNum(theNum)
	Dim tempNum
	
	If Not IsNumeric(theNum) Then
		If Trim(theNum) <> "" Then
			numBuffer = Trim(theNum)
			inStrNum = ""
			numLength = Len(numBuffer)
			If numLength > 1 Then
				If Left(numBuffer, 1) = "-" Then
					numBuffer = Mid(numBuffer, 2)
					numLength = Len(numBuffer)
				End If
			End If
			
			For z = 1 To numLength
				tempNumTest = Mid(numBuffer, z, 1) 
				If IsNumeric(tempNumTest) Or tempNumTest = "." Or tempNumTest = "," Then
					inStrNum = inStrNum & tempNumTest					
				Else
					Exit For	
				End If	
			Next
			
			If IsNumeric(inStrNum) Then				
				appendStr = Mid(numBuffer, z) 
				If inStrNum < 0 Then
					FormatNum = "(" & FormatNumber(Abs(inStrNum), 2) & ")" & appendStr
				Else
					FormatNum = FormatNumber(inStrNum, 2) & appendStr
				End If	
				
				Exit Function
			End If
			
		End If
		
		FormatNum = theNum
		Exit Function
	End If
	
	On Error Resume Next
		tempNum = FormatNumber(theNum, 2)		
	If Err.Number > 0 Then
		FormatNum = theNum
		Err.Clear
	Else
		If tempNum < 0 Then
			FormatNum = "(" & FormatNumber(Abs(tempNum), 2) & ")"
		Else
			FormatNum = tempNum
		End If	
	End If	
End Function

Function FormatNumCommasOnly(theNum)
	Dim tempNum
	On Error Resume Next
		tempNum = FormatNumber(theNum, 0)		
	If Err.Number > 0 Then
		FormatNumCommasOnly = theNum
		Err.Clear
	Else
		If tempNum < 0 Then
			FormatNumCommasOnly = "(" & FormatNumber(Abs(tempNum), 0) & ")"
		Else
			FormatNumCommasOnly = tempNum
		End If	
	End If	
End Function

Function DrawPageFunctions(includePrint, includeEmailThisPage, includeBackBtn)
	Dim clientSideScript
	Dim sName, sPort, sPath, thisPath
	
	sName = Request.ServerVariables("SERVER_NAME")
	sPort = Request.ServerVariables("SERVER_PORT")
	
	If sPort <> "" Then
		sPath = "http://" & sName & ":" & sPort & "/"
	Else
		sPath = "http://" & sName & "/"
	End If
	
	
	thisPath = Server.MapPath(".")
	
	sPath = "/"
	
	
	barTdClientScript = " nowrap class=nav onMouseover=""JavaScript: if (this.className=='nav') this.className='nav_over';"" onMouseout=""JavaScript: if (this.className=='nav_over') this.className='nav';""" 
	
	With Response
		.Write "<!--BEGINPAGEFUNCS-->" & vbCrLf
		clientSideScript =  "<Script Language=""JavaScript"">" & vbCrLf
		
		If includePrint = True Then
			.Write "<table border=0 cellpadding=2 cellspacing=4 id=""printTable"" class=""printTable"">"  & vbCrLf
			.Write "<tr name=functionRow id=functionRow>" & vbCrLf			
			.Write "    <td " & barTdClientScript & " OnClick='javascript:printReportDoc()'><img src=""../images/printLink.gif"" border=0 title=""Print"">Print" & vbCrLf
			WritePrinterDialogDiv sPath, thisPath, reportName
			'.Write "<form method=post action='" & sPath & "exportXL.asp' id=""frmExport"" style=""display: none""><textarea name=htmlDoc></textarea>" & vbCrLf
			
			'.Write "</form>" & vbCrLf
			.Write "    </td>" & vbCrLf			
				
			clientSideScript = clientSideScript & "	function printReportDoc(){"  & vbCrLf
			clientSideScript = clientSideScript & "		document.all.item('printTable').style.display = 'none';"  & vbCrLf
			clientSideScript = clientSideScript & "		window.self.print();"  & vbCrLf
			clientSideScript = clientSideScript & "		document.all.item('printTable').style.display = '';	"  & vbCrLf
			clientSideScript = clientSideScript & "	}" & vbCrLf
			
			clientSideScript = clientSideScript & "	function exportDoc(){"  & vbCrLf
			clientSideScript = clientSideScript & "		var thisHTMLDoc = getBodyHTML();"  & vbCrLf
			clientSideScript = clientSideScript & "		document.all.item('htmlDoc').value = thisHTMLDoc;"  & vbCrLf
			clientSideScript = clientSideScript & "		document.all.item('reportName').value = window.document.title;"  & vbCrLf
			clientSideScript = clientSideScript & "		document.all.item('frmPrinterDialog').target = 'deleteFrame';"  & vbCrLf
			clientSideScript = clientSideScript & "		document.all.item('frmPrinterDialog').submit();"  & vbCrLf
			clientSideScript = clientSideScript & "	}" & vbCrLf	  
		End If
	
	
	
		If includeEmailThisPage = True Then
			.Write "    <td " & barTdClientScript & " OnClick=""javascript: launchMailer()""><font color=blue face=""Wingdings"" size=3>+</font>Email this Page" & vbCrLf
			.Write "	<form method='POST' style='position:absolute;visibility: hidden' name='frmSendMail'>" & vbCrLf
			.Write "		<input type=hidden name=emailDoc>"  & vbCrLf	
			.Write "		<input type=hidden name=emailDocSourcePath value=""" & thisPath & """>" & vbCrLf
			.Write "		<input type=hidden name=emailDocName>" & vbCrLf
			.Write "		<input type=hidden name=to>" & vbCrLf
			.Write "		<input type=hidden name=cc>" & vbCrLf
			.Write "		<input type=hidden name=bcc>" & vbCrLf
			.Write "		<input type=hidden name=subject>" & vbCrLf
			.Write "		<input type=hidden name=bodyText>" & vbCrLf
			.Write "	</form>" & vbCrLf
			.Write "	<Div style='display: none'>" & vbCrLf
			.Write "	<IFRAME marginwidth=0 marginheight=0 FRAMEBORDER=0 ID='hiddenEmailFrame' NAME='EmailFrame'></IFRAME>" & vbCrLf
			.Write "	</Div>" & vbCrLf			
			.Write "    </td>" & vbCrLf	
			
			clientSideScript = clientSideScript & "	function emailDoc(){"  & vbCrLf
			clientSideScript = clientSideScript & "		var thisHTMLDoc = getBodyHTML();"  & vbCrLf
			clientSideScript = clientSideScript & "		document.all.item('emailDoc').value = thisHTMLDoc;"  & vbCrLf
			clientSideScript = clientSideScript & "		document.all.item('emailDocName').value = window.document.title;"  & vbCrLf
			clientSideScript = clientSideScript & "		document.all.item('frmSendMail').action = '" & sPath & "Email.asp';"  & vbCrLf
			clientSideScript = clientSideScript & "		document.all.item('frmSendMail').target = 'EmailFrame';"  & vbCrLf
			clientSideScript = clientSideScript & "		document.all.item('frmSendMail').submit();"  & vbCrLf
			clientSideScript = clientSideScript & "	}" & vbCrLf			
			clientSideScript = clientSideScript & "	function launchMailer(){"  & vbCrLf
			clientSideScript = clientSideScript & "		window.showModalDialog('" & sPath & "Email.asp?enterDetailsPointer=1', window.self, 'dialogWidth:30em;dialogHeight:22em;status:0;dialogHide:false;help:no;scroll:yes;resizable:no;edge:sunken;unadorned:yes');"  & vbCrLf
			clientSideScript = clientSideScript & "	}" & vbCrLf			
			
			
		End If
		
		If includeBackBtn Then
			.Write "    <td " & barTdClientScript & " OnClick=""javascript: window.parent.history.go(-1)""><font color=blue face=""Webdings"" size=1>3</font>Back" & vbCrLf
			.Write "    </td>" & vbCrLf		
		End If
		
		.Write "    <td " & barTdClientScript & " OnClick=""javascript: window.parent.self.close()""><font color=blue face=""Wingdings"" size=3>x</font>Close" & vbCrLf
		.Write "    </td>" & vbCrLf	
			
		.Write "</tr></table>" & vbCrLf
	
		clientSideScript = clientSideScript & vbCrLf & "</Script>" & vbCrLf & "<!--ENDPAGEFUNCS-->"				
		
		.Write clientSideScript
		
			
	End With 
End Function


Function FormatDateFull(theDate)
	FormatDateFull = WeekDayName(WeekDay(theDate, 1)) & ", " & Day(theDate) & "-" & MonthName(Month(theDate)) & "-" & Year(theDate)
End Function


Function JumpUpFromWeekendToWeek(theDate)
		Dim theSelDay
		
		theSelDay = WeekDay(theDate)
		
		If (theSelDay = vbSunday) Then
			JumpUpFromWeekendToWeek = FormatDate(DateAdd("d", 1, theDate))
		ElseIf (theSelDay = vbSaturday) Then
			JumpUpFromWeekendToWeek = FormatDate(DateAdd("d", 2, theDate))
		Else
			JumpUpFromWeekendToWeek = theDate	
		End If			
		
End Function	

Function JumpBackFromWeekendToWeek(theDate)
		Dim theSelDay
		
		theSelDay = WeekDay(theDate)
		
		If (theSelDay = vbSunday) Then
			JumpBackFromWeekendToWeek = FormatDate(DateAdd("d", -2, theDate))
		ElseIf (theSelDay = vbSaturday) Then
			JumpBackFromWeekendToWeek = FormatDate(DateAdd("d", -1, theDate))
		Else
			JumpBackFromWeekendToWeek = theDate	
		End If			
		
End Function		


Function WritePrinterDialogDiv(thePath, thisPath, reportName)
	With Response
	
		.Write "	<form method='POST' action='" & thePath &  "exportXL.asp' style='position:absolute;visibility: hidden' name='frmPrinterDialog'>" & vbCrLf
		.Write "<input type=hidden name=sourcePath value=""" & thisPath & """>"
		.Write "<input type=hidden name=reportName value=""" & reportName & """>"
		.Write "<input type=hidden name=htmlDoc>"
		.Write "<input type=hidden name=LandScape value=0>"
		.Write "<input type=hidden name=ZoomRatio value=100>"					
		.Write "  <table bgColor='#CCCCCC' border='0' cellspacing='0' cellpadding='0' style='border: 2 outset #C0C0C0' width='300px'>" & vbCrLf
		.Write "<tr OnClick=""JavaScript: document.all.item('frmPrinterDialog').style.visibility='hidden';"" style='cursor: hand'>" & vbCrLf
		.Write "<td nowrap style='background-color: #000080; color: #FFFFFF'>" & vbCrLf
		.Write "<p align='left'><font size='2' face='Verdana'><b>&nbsp;Printer Dialog</b></font></td>" & vbCrLf
		.Write "<td nowrap  style='background-color: #000080; color: #FFFFFF'>" & vbCrLf
		.Write "<p align='right'><font size='2' face='Verdana'><input type='button' value=' X ' name='B2' OnMouseUp=""JavaScript: document.all.item('frmPrinterDialog').style.visibility='hidden';""></font></td>" & vbCrLf
		.Write "</tr>" & vbCrLf
		.Write "<tr>" & vbCrLf
		.Write "<td></td>" & vbCrLf
		.Write "<td></td>" & vbCrLf
		.Write "</tr>" & vbCrLf
		.Write "<tr>" & vbCrLf
		.Write "<td nowrap ><b><font size='2' face='Tahoma' color='#000080'>Printer</font></b></td>" & vbCrLf
		.Write "<td nowrap ><select size='1' name='cboPrinter'>" & vbCrLf
		
				Set pConn = GetActiveConnection("KBroker")
				Set printersRs = pConn.Execute ("SELECT * FROM Printers")
				If Not (printersRs.EOF OR printersRS.BOF) Then
					Do Until printersRs.EOF
						.Write "<option value='" & printersRs("PrinterActualName").Value & "'>" & printersRs("PrinterName").Value  & "</Option>" & vbCrLf
						printersRs.MoveNext
					Loop	
				
				End If
				
				Set printersRs = Nothing
				Set pConn = Nothing
				
				
				.Write "</select></td>" & vbCrLf
				.Write "</tr>" & vbCrLf
				.Write "<tr>" & vbCrLf
				.Write "<td nowrap ><b><font size='2' face='Tahoma' color='#000080'>No. of Copies</font></b></td>" & vbCrLf
				.Write "<td nowrap ><input type='text' name='txtNumbers' size='5' value='1'></td>" & vbCrLf
				.Write "</tr>" & vbCrLf
				.Write "<tr>" & vbCrLf
				.Write "<td nowrap align=right colspan='2'><b><font size='2' face='Tahoma' color='#000080'>&nbsp;</font></b><input type='button' value=' Send To Printer ...' name='B1' OnClick=""JavaScript: exportDoc()""></td>" & vbCrLf
				.Write "</tr>" & vbCrLf
				.Write "</table>" & vbCrLf
				.Write "<p>&nbsp;</p>" & vbCrLf
				.Write "</form>" & vbCrLf

	End With

End Function

Function CreditDebitValueRev(theVal)
	CreditDebitValueRev = Replace(Replace(theVal, " Dr", ""), " Cr", "")
End Function

Function CreditDebitValue(theVal)
	If IsNumeric(theVal) Then
		If theVal >= 0 Then
			CreditDebitValue =	theVal & " Cr"
		Else
			CreditDebitValue = theVal & " Dr"
		End If
	Else
		CreditDebitValue = theVal
	End If
End Function

function ApplyDisplayColor(displayText)
		displayText = "<font color = " & displayColor & ">" & displayText & "</font>"
		ApplyDisplayColor = displayText
end function
%>