<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html;charset=UTF-8">
<META HTTP-EQUIV="Expires" CONTENT="0">
<TITLE>Search</TITLE>
<LINK href="../STYLE/default.css" type=TEXT/CSS rel=STYLESHEET>
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<script language="JavaScript" src="CALENDAR/calendar.js"></script>
</HEAD>

<BODY CLASS="FadeBlue">

<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>

<!--#include file="../libroutines.asp"-->

<SCRIPT language="JavaScript">
	var cal, cal2;
	
	function drawFilterInterface(fieldTypei, i){
		var interfaceTD, r;
		interfaceTD = document.all.item("inputTD" + i);
		
		r = parseInt(fieldTypei);
		if (r=="NaN" || fieldTypei==""){
			interfaceTD.innerHTML = "&nbsp;";			
			return;
		}
		
		checkToDefaultOn(i) = false;
		
		switch (myFieldTypes(fieldTypei)){
			case "0": // 0 = text, default
				interfaceTD.innerHTML = "<INPUT TYPE=TEXT NAME=SearchValue" + i + ">";
				document.all.item('SearchValue' + i).focus();
				break;
			case "1": 	//1 = date
				cal = new ctlSpiffyCalendarBox("cal", "frmSearch", "SearchValue" + i, "cmdDate", "<%= FormatDate(Date) %>", 1)
				cal.returnOutStringOnWrite();
				interfaceTD.innerHTML = cal.writeControl();
				break;
			case "2": 	//2 = number	
				interfaceTD.innerHTML = '<INPUT TYPE=TEXT NAME=SearchValue' + i + ' OnKeyDown="persistNumbers();">';
				document.all.item('SearchValue' + i).focus();
				break;
			case "3": 	//3 = date range
				cal = new ctlSpiffyCalendarBox("cal", "frmSearch", "SearchValueFrom" + i, "cmdDate", "<%= FormatDate(Date) %>", 1)
				cal2 = new ctlSpiffyCalendarBox("cal2", "frmSearch", "SearchValueTo" + i, "cmdDate2", "<%= FormatDate(Date) %>", 1)
				cal.returnOutStringOnWrite();
				cal2.returnOutStringOnWrite();
				interfaceTD.innerHTML = "From:&nbsp;" + cal.writeControl() + "<p>To:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" +   cal2.writeControl() ;
				checkToDefaultOn(i) = true;
				break;
			case "4": 	//4 = number range
				interfaceTD.innerHTML = 'From:&nbsp;<INPUT TYPE=TEXT NAME=SearchValueFrom' + i + ' OnKeyDown="persistNumbers();">';
				interfaceTD.innerHTML += '<p>To:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME=SearchValueTo' + i + ' OnKeyDown="persistNumbers();">';
				checkToDefaultOn(i) = true;
				break;	
			case "5":	//5 = boolean field	
				interfaceTD.innerHTML = '<INPUT TYPE=radio class=borderless NAME=SearchValueBool' + i + '1 value=True id=chkOpt' + i + '><label for=chkOpt' + i + '>True</label>';
				interfaceTD.innerHTML += '&nbsp;<INPUT TYPE=radio class=borderless NAME=SearchValueBool' + i + ' value=False id=chkOpt2' + i + '><label for=chkOpt2' + i + '>False</label>';
				checkToDefaultOn(i) = true;
				break;	
			case "6":	//6 = boolean field, yes no	
				interfaceTD.innerHTML = '<INPUT TYPE=radio class=borderless NAME=SearchValueBool' + i + '1 value=True id=chkOpt1><label for=chkOpt1>Yes</label>';
				interfaceTD.innerHTML += '&nbsp;<INPUT TYPE=radio class=borderless NAME=SearchValueBool' + i + ' value=False id=chkOpt2><label for=chkOpt2>No</label>';
				checkToDefaultOn(i) = true;
				break;	
				
			default:
				interfaceTD.innerHTML = "<INPUT TYPE=TEXT NAME=SearchValue" + i + ">";
				document.all.item('SearchValue' + i).focus();
				break;	
		}
		
		DoResizeWin();
	}
	
	function persistNumbers(){
			 event.returnValue = IsANumber(event.keyCode);	 
	}
	
	
	function checkToDefault(theSel, iTurn){
		if (checkToDefaultOn(iTurn)) theSel.selectedIndex = 0;
	}
	
</SCRIPT>

<Script Language="VBScript">
	Dim mySearchArray
	Dim myFieldTypes()
	Dim checkToDefaultOn()
	Dim myOptionsGuide
	Dim zeroBasedMaxSearch
	
	zeroBasedMaxSearch = 3
	
	
	Function IsANumber(theVal)
		IsANumber = IsNumeric(Chr(theVal))
	End Function
	
	Function DoInit		
		If window.dialogArguments <> "" Then
			window.name = "searchWindow"			
			window.defaultStatus = "Search Window"
			mySearchArray = Split(window.dialogArguments.searchArgs, ";")			
    
		Else
			window.self.close
		End If
    
	End Function
	
	Function DoResizeWin
		window.dialogHeight = (document.all.item("mainTable").clientHeight / 10) - 5  & "em"
		window.dialogWidth = (document.all.item("mainTable").clientWidth / 10) - 5 & "em"
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
		
		
		If searchStr = "" Then
			window.alert("The search information could not be built")
		Else		
			window.dialogArguments.opener.document.all.item("SelectedSearchArgs").value = searchStr
			window.dialogArguments.opener.document.all.item("frmMain").action = window.dialogArguments.searchTarget
			window.dialogArguments.opener.document.all.item("frmMain").submit
			window.dialogArguments.searchTD.className = "footerHighlightNavOn"
			window.dialogArguments.cancelButton.style.display = ""			
			window.self.close
		End If	
		
		
	End Function
	
	DoInit
</Script> 

<Form ID="frmSearch">
<table ID="mainTable">
	<TR>
		<TD><font face=Verdana size=1><b>Search by</b></font>			
		</TD>
		<TD COLSPAN=2><font face=Verdana size=1><b>Qualification</b></font>			
		</TD>
	</TR>
	<Script Language="VBScript">
		Dim myOptions
		Dim myWorkingOptions
		
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
					.write "<td  COLSPAN=3><hr width=100%></td>"
					.write "</tr>"
					.write "<TR><td COLSPAN=2><font face=Verdana size=1><b>"
					.write "<INPUT TYPE=RADIO Class=""BorderLess"" NAME=concactVal" & i & " ID=concactAnd" & i & "  VALUE=AND Checked><LABEL FOR=concactAnd" & i & " STYLE=""CURSOR: HAND""><font face=Verdana size=1>And</font></LABEL>&nbsp;"
					.write "<INPUT TYPE=RADIO Class=""BorderLess"" NAME=concactVal" & i & " ID=concactOr" & i & "  VALUE=OR><LABEL FOR=concactOr" & i & " STYLE=""CURSOR: HAND""><font face=Verdana size=1>Or</font></LABEL>&nbsp;"
					.write "</b></font>"
					.write "</td></tr>"
				End If
				
				.write "<TR>"
				.write "<TD>"
				.write "<Select NAME='SearchFld" & i & "' OnChange=""JavaScript: drawFilterInterface(this.options[this.selectedIndex].TAG, '" & i & "')"">"
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
					.write "<td COLSPAN=3><INPUT TYPE=BUTTON NAME=SEARCH VALUE=Search OnClick=""VBScript: DoSearch"">&nbsp;&nbsp;<INPUT TYPE=BUTTON NAME=close VALUE=Close OnClick=""window.self.close()"">"
					.write "</td></tr>"
		End With			
		
		On Error Resume Next
		Dim theFirstSel
		Set theFirstSel = document.all.item("SearchFld0")
		theFirstSel.options(1).selected = True
		drawFilterInterface theFirstSel.options(theFirstSel.selectedIndex).TAG, "0"
		
	</Script>
	
</table>
</form>
</BODY>
</HTML>