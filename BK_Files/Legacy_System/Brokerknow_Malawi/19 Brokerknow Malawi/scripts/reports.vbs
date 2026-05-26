Dim mySearchArray
Dim myFieldTypes()
Dim checkToDefaultOn()
Dim myOptionsGuide
Dim zeroBasedMaxSearch
	
zeroBasedMaxSearch = 3
	
Dim mySortArray
	
Const fixedMaxSortFields = 3
	

Function DoSort
	Dim mySortArgs, optSelVal		
	Dim i	
	
	For i = 1 To fixedMaxSortFields
		thisField = document.all.item("SortFld" & i).value
		If thisField <> "" Then
			If document.all.item("ASC" & i).checked Then
				sortCat = "ASC"
			Else
				sortCat = "DESC"
			End If
			If mySortArgs <> "" Then
				mySortArgs = mySortArgs & ", " & thisField & " " & sortCat
			Else
				mySortArgs =  thisField & " " & sortCat
			End If
		End If	
	Next
		
	If mySortArgs <> "" Then
		document.getElementById("SelectedSortArgs").value = mySortArgs											
	End If
		
End Function

Function report_initSort(mySortArrayData)
	
	dim myOptions
	
	mySortArray = Split(mySortArrayData, ";")
		
	myOptions = "<OPTION VALUE="""">--</OPTION>"
	For i = 0 To UBound(mySortArray)
		thisField = Trim(mySortArray(i))
		If InStr(1, thisField, ":") > 0 Then
			thisFieldName = Mid(thisField, 1, InStr(1, thisField, ":") - 1)
			thisFieldDesc = Mid(thisField, InStr(1, thisField, ":") + 1)
		Else
			thisFieldName = thisField
			thisFieldDesc = thisField
		End	If
		myOptions = myOptions & "<OPTION VALUE=""" & thisFieldName & """>" & thisFieldDesc & "</OPTION>"
	Next
		
	For i = 1 To fixedMaxSortFields	
		With document
			If i > 1 Then
				.write "<TR>"
				.write "<td colspan=2><table border=0 cellspacing=0 cellpadding=1><tr><td width=25%><font face=Verdana size=1>Then by</font></td>"
				.write "<td align=left width=""75%"" nowrap>&nbsp;"
				.write "</td></tr></table></td></tr>"
			End If
			.write "<TR>"
			.write "<TD>"
			.write "<Select NAME='SortFld" & i & "'>"
			.write myOptions
			.write "</Select>"
			.write "</TD>"
			.write "<TD>"
			.write "<INPUT TYPE=RADIO NAME=SortValue" & i & " Class=""BorderLess"" ID=ASC" & i & "  VALUE=ASC Checked><LABEL FOR=ASC" & i & " STYLE=""CURSOR: HAND""><font face=Verdana size=1>Ascending</font></LABEL>" & Chr(13)
			.write "<INPUT TYPE=RADIO NAME=SortValue" & i & " Class=""BorderLess"" ID=DESC" & i & " VALUE=DESC><LABEL FOR=DESC" & i & " STYLE=""CURSOR: HAND""><font face=Verdana size=1>Descending</font></LABEL>"
			.write "</TD>"
			.write "</TR>"	 		
				
		End	With
	Next	
End Function




Function report_init_Filter(mySearchArrayData)
	Dim myOptions
	Dim myWorkingOptions
			
	mySearchArray = Split(mySearchArrayData, ";")		
			
	For i = 0 To UBound(mySearchArray)
		thisField = Trim(mySearchArray(i))
		If InStr(1, thisField, ":") > 0 Then
			thisFieldName = Mid(thisField, 1, InStr(1, thisField, ":") - 1)
			thisFieldDesc = Mid(thisField, InStr(1, thisField, ":") + 1)
		Else
			thisFieldName = thisField
			thisFieldDesc = thisField
		End	If
		ReDim Preserve myFieldTypes(i)	
				
		If  InStr(1, thisFieldDesc, "*") > 0 Then 				
			myFieldTypes(i) = Mid(thisFieldDesc, InStr(1, thisFieldDesc, "*") + 1)
			thisFieldDesc = Mid(thisFieldDesc, 1, InStr(1, thisFieldDesc, "*") - 1)
			'field types:
			'0 = text, default
			'1 = date
			'2 = number
			'3 = date range
			'4 = number range
			'5 = boolean
					
		Else
			myFieldTypes(i)	= "0"
		End If
				
		If myOptions <> "" Then			
			myOptions = myOptions & "<OPTION VALUE=""" & thisFieldName & """ TAG=""" & i & """>" & thisFieldDesc & "</OPTION>"
		Else
			myOptions = "<OPTION SELECTED VALUE="""" TAG="""">--</OPTION>"
			myOptions = myOptions & "<OPTION VALUE=""" & thisFieldName & """ TAG=""" & i & """>" & thisFieldDesc & "</OPTION>"
		End If		
	Next
			
	myWorkingOptions = "<OPTION SELECTED VALUE=""0"">Contains</OPTION>"
	myWorkingOptions = myWorkingOptions & "<OPTION VALUE=""1"">Equals</OPTION>"
	myWorkingOptions = myWorkingOptions & "<OPTION VALUE=""2"">Starts with</OPTION>"
	myWorkingOptions = myWorkingOptions & "<OPTION VALUE=""3"">Ends with</OPTION>"
	myWorkingOptions = myWorkingOptions & "<OPTION VALUE=""4"">Not Contains</OPTION>"
			
	For i = 0 To zeroBasedMaxSearch	
		ReDim Preserve checkToDefaultOn(i)
		checkToDefaultOn(i) = False
		With document
			If i > 0 Then
				.write "<TR>"
				.write "<td  COLSPAN=3> &nbsp;</td>"
				.write "</tr>"
				.write "<TR><td COLSPAN=2><font face=Verdana size=1><b>"
				.write "<INPUT TYPE=RADIO Class=""BorderLess"" NAME=concactVal" & i & " ID=concactAnd" & i & "  VALUE=AND Checked><LABEL FOR=concactAnd" & i & " STYLE=""CURSOR: HAND""><font face=Verdana size=1>And</font></LABEL>&nbsp;"
				.write "<INPUT TYPE=RADIO Class=""BorderLess"" NAME=concactVal" & i & " ID=concactOr" & i & "  VALUE=OR><LABEL FOR=concactOr" & i & " STYLE=""CURSOR: HAND""><font face=Verdana size=1>Or</font></LABEL>&nbsp;"
				.write "</b></font>"
				.write "</td></tr>"
			End If
					
			.write "<TR>"
			.write "<TD>"
			.write "<Select NAME='SearchFld" & i & "' OnChange=""JavaScript: report_drawFilterInterface(this.options[this.selectedIndex].TAG, '" & i & "')"">"
			.write myOptions
			.write "</Select>"				
			.write "</TD>"
			.write "<TD>"
			.write "<Select NAME='SearchTerms" & i & "' OnChange=""JavaScript: checkToDefault(this, '" & i & "');"">"
			.write myWorkingOptions
			.write "</Select>"
			.write "</TD>"				
			.write "<TD ID=""inputTD" & i & """>"
			.write "&nbsp;"
			.write "</TD>"
			.write "</TR>"
									
		End	With
	Next
			
	With document			
				.write "<TR align=right>"
				.write "<td COLSPAN=3><br>"
				.write "</td></tr>"
	End With			
			
	On Error Resume Next
	Dim theFirstSel
	Set theFirstSel = document.all.item("SearchFld0")
	theFirstSel.options(1).selected = True
	report_drawFilterInterface theFirstSel.options(theFirstSel.selectedIndex).TAG, "0"


End Function
	
	
Function IsANumber(theVal)
	IsANumber = IsNumeric(Chr(theVal))
End Function
	
Function GetMidCompareSearchStr(searchVal, docSearchTerms)
	If docSearchTerms.value = "0" Then
		termStr = " LIKE '%" & searchVal & "%'"
	ElseIf docSearchTerms.value = "1" Then
		termStr = " LIKE '" & searchVal & "'"
	ElseIf docSearchTerms.value = "2" Then
		termStr = " LIKE '" & searchVal & "%'"
	ElseIf docSearchTerms.value = "3" Then
		termStr = " LIKE '%" & searchVal & "'"
	ElseIf docSearchTerms.value = "4" Then
		termStr = " NOT LIKE '%" & searchVal & "%'"
	Else
		termStr = " LIKE '%" & searchVal & "%'"
	End If
		
	GetMidCompareSearchStr = termStr
End Function
	
Function DoSearch
	Dim searchStr, concactStr, termStr
	Dim myDocSearchTerms, myfieldTypei
	Dim i
		
	On Error Resume Next
		
		
		
	For i = 0 To zeroBasedMaxSearch
		thisField = document.all.item("SearchFld" & i).value
		If thisField <> "" Then
			If i > 0 Then
				If document.all.item("concactAnd" & i).checked Then
					concactStr = "AND"
				Else
					concactStr = "OR"
				End If
			End If
			
			Set myDocSearchTerms =  document.all.item("SearchTerms" & i)
			myfieldTypei = document.all.item("SearchFld" & i).options(document.all.item("SearchFld" & i).selectedIndex).TAG
				
			If (myFieldTypes(myfieldTypei) = "5" Or myFieldTypes(myfieldTypei) = "6")  Then 'boolean
				searchStr = thisField & " = " & Abs(CLng(CBool(document.all.item("SearchValueBool" & i & "1").checked)))
			ElseIf myFieldTypes(myfieldTypei) = "3"  Then 'date range
				searchStr = thisField & " <= #" & document.all.item("SearchValueTo" & i).value & "# AND " & thisField & " >= #" & document.all.item("SearchValueFrom" & i).value & "#"
			ElseIf myFieldTypes(myfieldTypei) = "4"  Then 'number range
				searchStr = thisField & " <= " & document.all.item("SearchValueTo" & i).value & " AND " & thisField & " >= " & document.all.item("SearchValueFrom" & i).value
			Else
				termStr = GetMidCompareSearchStr(document.all.item("SearchValue" & i).value, myDocSearchTerms)
				termStr = thisField & termStr
			End If
							
				
			If searchStr <> "" Then
				searchStr = searchStr & " " & concactStr & " " & termStr
			Else
				searchStr =  termStr
			End If
	
		End If
	Next
		
		
	If searchStr <> "" Then
		document.getElementById("SelectedSearchArgs").value = searchStr
	End If			
		
End Function
