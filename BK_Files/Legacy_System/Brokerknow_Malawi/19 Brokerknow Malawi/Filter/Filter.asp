<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html;charset=UTF-8">
<META HTTP-EQUIV="Expires" CONTENT="0">
<TITLE>Filter</TITLE>
<LINK href="../STYLE/default.css" type=TEXT/CSS rel=STYLESHEET>
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<script language="JavaScript" src="CALENDAR/calendar.js"></script>
</HEAD>

<BODY CLASS="FadeBlue">
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<!--#include file="../libroutines.asp"-->

<SCRIPT language="JavaScript">
	var cal, cal2;
	var checkToDefaultOn = new Boolean();
	checkToDefaultOn = false;
	function drawFilterInterface(i){
		var interfaceTD, r;
		interfaceTD = document.all.item("inputTD");
		currSearchFieldType = i;
		r = parseInt(i);
		if (r=="NaN" || i==""){
			interfaceTD.innerHTML = "&nbsp;";			
			return;
		}
		
		checkToDefaultOn = false;
		
		switch (myFieldTypes(i)){
			case "0": // 0 = text, default
				interfaceTD.innerHTML = "<INPUT TYPE=TEXT NAME=SearchValue" + i + ">";
				document.all.item('SearchValue' + i).focus();
				break;
			case "1": 	//1 = date
				cal = new ctlSpiffyCalendarBox("cal", "frmFilter", "SearchValue" + i, "cmdDate", "<%= FormatDate(Date) %>", 1)
				cal.returnOutStringOnWrite();
				interfaceTD.innerHTML = cal.writeControl();
				break;
			case "2": 	//2 = number	
				interfaceTD.innerHTML = '<INPUT TYPE=TEXT NAME=SearchValue' + i + ' OnKeyDown="persistNumbers();">';
				document.all.item('SearchValue' + i).focus();
				break;
			case "3": 	//3 = date range
				cal = new ctlSpiffyCalendarBox("cal", "frmFilter", "SearchValueFrom" + i, "cmdDate", "<%= FormatDate(Date) %>", 1)
				cal2 = new ctlSpiffyCalendarBox("cal2", "frmFilter", "SearchValueTo" + i, "cmdDate2", "<%= FormatDate(Date) %>", 1)
				cal.returnOutStringOnWrite();
				cal2.returnOutStringOnWrite();
				interfaceTD.innerHTML = "From:&nbsp;" + cal.writeControl() + "<p>To:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" +   cal2.writeControl() ;
				checkToDefaultOn = true;
				break;
			case "4": 	//4 = number range
				interfaceTD.innerHTML = 'From:&nbsp;<INPUT TYPE=TEXT NAME=SearchValueFrom' + i + ' OnKeyDown="persistNumbers();">';
				interfaceTD.innerHTML += '<p>To:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME=SearchValueTo' + i + ' OnKeyDown="persistNumbers();">';
				checkToDefaultOn = true;
				break;	
			case "5":	//5 = boolean field	
				interfaceTD.innerHTML = '<INPUT TYPE=radio class=borderless NAME=SearchValueFrom' + i + ' value=True id=chkOpt1><label for=chkOpt1>True</label>';
				interfaceTD.innerHTML += '&nbsp;<INPUT TYPE=radio class=borderless NAME=SearchValueFrom' + i + ' value=False id=chkOpt2><label for=chkOpt2>False</label>';
				checkToDefaultOn = true;
				break;	
			case "6":	//6 = boolean field, yes no	
				interfaceTD.innerHTML = '<INPUT TYPE=radio class=borderless NAME=SearchValueFrom' + i + ' value=True id=chkOpt1><label for=chkOpt1>Yes</label>';
				interfaceTD.innerHTML += '&nbsp;<INPUT TYPE=radio class=borderless NAME=SearchValueFrom' + i + ' value=False id=chkOpt2><label for=chkOpt2>No</label>';
				checkToDefaultOn = true;
				break;		
			default:
				interfaceTD.innerHTML = "<INPUT TYPE=TEXT NAME=SearchValue" + i + ">";
				document.all.item('SearchValue' + i).focus();
				break;	
		}
	}
	
	function persistNumbers(){
			 event.returnValue = IsANumber(event.keyCode);	 
	}
	
	function checkToDefault(theSel){
		if (checkToDefaultOn) theSel.selectedIndex = 0;
	}
</SCRIPT>


<Script Language="VBScript">
	Dim myFilterArray
	Dim myFieldTypes()
	Dim currSearchFieldType
	Dim myOptionsGuide
	
	Function IsANumber(theVal)
		IsANumber = IsNumeric(Chr(theVal))
	End Function
		
	Function DoInit		
		If window.dialogArguments <> "" Then
			window.name = "filterWindow"			
			window.defaultStatus = "Filter Window"
			myFilterArray = Split(window.dialogArguments.filterArgs, ";")			
    
		Else
			window.self.close
		End If
    
	End Function
	
	Function DoResizeWin
		window.dialogHeight = document.all.item("mainTable").clientHeight / 10 & "em"
	End Function

	Function GetMidCompareSearchStr(searchVal)
		If document.all.item("SearchTerms").value = "0" Then
			termStr = " LIKE '%" & searchVal & "%'"
		ElseIf document.all.item("SearchTerms").value = "1" Then
			termStr = " LIKE '" & searchVal & "'"
		ElseIf document.all.item("SearchTerms").value = "2" Then
			termStr = " LIKE '" & searchVal & "%'"
		ElseIf document.all.item("SearchTerms").value = "3" Then
			termStr = " LIKE '%" & searchVal & "'"
		ElseIf document.all.item("SearchTerms").value = "4" Then
			termStr = " NOT LIKE '%" & searchVal & "%'"
		Else
			termStr = " LIKE '%" & searchVal & "%'"
		End If
		
		GetMidCompareSearchStr = termStr
	End Function
	
	Function DoFilter
		Dim searchStr, thisField, optSelVal, optCompound
		Dim optClearSearch, optClearSort
		
		optCompound = True
		optClearSearch = False
		optClearSort = False
	
	
		Set thisField = document.all.item("SearchFld")
		If thisField.value = "" Then
			alert "Select a field for the filter process"
			thisField.focus
			Exit Function
		End If
	
		Select Case myFieldTypes(currSearchFieldType)
			Case "0" 'text, default
				searchStr = thisField.value & GetMidCompareSearchStr(document.all.item("SearchValue" & currSearchFieldType).value)
			Case "1" 'date
				searchStr = thisField.value & " = '" & document.all.item("SearchValue" & currSearchFieldType).value & "'"
			Case "2" 'number
				searchStr = thisField.value & GetMidCompareSearchStr(document.all.item("SearchValue" & currSearchFieldType).value)
			Case "3", "4" 'date range
				searchStr = thisField.value & " <= #" & document.all.item("SearchValueTo" & currSearchFieldType).value & "# AND " & thisField.value & " >= #" & document.all.item("SearchValueFrom" & currSearchFieldType).value & "#"
			Case "4" 'number ranges
				searchStr = thisField.value & " <= " & document.all.item("SearchValueTo" & currSearchFieldType).value & " AND " & thisField.value & " >= " & document.all.item("SearchValueFrom" & currSearchFieldType).value 
			Case "5", "6" 'boolean data
				If document.all.item("chkOpt1").checked Then
					searchStr = thisField.value & " = " & Abs(CLng(CBool(document.all.item("chkOpt1").value)))
				Else
					searchStr = thisField.value & " = " & Abs(CLng(CBool(document.all.item("chkOpt2").value)))
				End If
				
		End Select
		
		If searchStr = "" Then
			window.alert("The filter information could not be built")
		Else
			On Error Resume Next
			optSelVal = document.all.item("SortOptions").value
			
			Select Case myOptionsGuide
				Case "0"
					optCompound = False
				Case "1"
					If optSelVal = "0" Then optClearSearch = True
					optCompound = False
						
				Case "2"					
					If optSelVal = "0" Then 
						optClearSort = True
						optCompound = False
					End If	
				Case "3"
					optCompound = False
					If optSelVal = "0" Then optClearSort = True
				Case "4"
					
					If optSelVal = "0" Then 
						optClearSearch = True
						optCompound = False
					ElseIf optSelVal = "2" Then 
						optCompound = False
					ElseIf optSelVal = "3" Then 
						optClearSearch = True
					End If	
				Case "5"
					If optSelVal = "2" Then 
						optClearSort = True
					ElseIf optSelVal = "3" Then 
						optClearSort = True
						optCompound = False
					ElseIf optSelVal = "4" Then 
						optCompound = False						
					End If	
				Case "6"
					optCompound = False
					If optSelVal = "0" Then 
						optClearSort = True
					ElseIf optSelVal = "2" Then 
						optClearSearch = True
					ElseIf optSelVal = "3" Then 
						optClearSearch = True	
						optClearSort = True
					End If	
				Case "7"
					If optSelVal = "0" Then 
						optClearSort = True
					ElseIf optSelVal = "2" Then 
						optClearSearch = True
						optCompound = False
					ElseIf optSelVal = "3" Then 
						optClearSearch = True
					ElseIf optSelVal = "4" Then
						optCompound = False
					ElseIf optSelVal = "5" Then 
						optClearSort = True
					ElseIf optSelVal = "6" Then 
						optClearSort = True
						optClearSearch = True
						optCompound = False
					ElseIf optSelVal = "7" Then 
						optClearSort = True
						optClearSearch = True	
					End If
				Case Else
					optCompound = False	
			End Select
			
	
			If optCompound Then
				window.dialogArguments.opener.document.all.item("SelectedFilterArgs").value = window.dialogArguments.opener.document.all.item("SelectedFilterArgs").value & " AND (" & searchStr & ")"
			Else
				window.dialogArguments.opener.document.all.item("SelectedFilterArgs").value = searchStr
			End If	
			
			If optClearSort Then
				window.dialogArguments.opener.document.all.item("SelectedSortArgs").value = ""
			End If
			
			If optClearSearch Then
				window.dialogArguments.opener.document.all.item("SelectedSearchArgs").value = ""
			End If			
			
			window.dialogArguments.opener.document.all.item("frmMain").action = window.dialogArguments.filterTarget
			window.dialogArguments.opener.document.all.item("frmMain").submit
			window.dialogArguments.filterTD.className = "footerHighlightNavOn"
			window.dialogArguments.cancelButton.style.display = ""
			window.self.close
		End If	
		
		
	End Function
	
	Function GetFilterOptions
		Dim currSortArgs, currSearchArgs, currFilterArgs, myFirstOption, mySecondOption
		Dim myThirdOption, myFourthOption
		
		currSortArgs = window.dialogArguments.opener.document.all.item("SelectedSortArgs").value
		currSearchArgs = window.dialogArguments.opener.document.all.item("SelectedSearchArgs").value
		currFilterArgs = window.dialogArguments.opener.document.all.item("SelectedFilterArgs").value
		
		If currSortArgs = "" And currSearchArgs = "" And currFilterArgs = "" Then
			myOptionsGuide = "0"
			GetFilterOptions = ""
			Exit Function
		ElseIf currSortArgs = "" And currFilterArgs = "" And currSearchArgs <> "" Then
			myFirstOption = "<OPTION VALUE=0>Clear current search results and filter</OPTION>"
			mySecondOption = "<OPTION VALUE=1 SELECTED>Filter current search results</OPTION>"
			myOptionsGuide = "1"
		ElseIf currSortArgs = "" And currFilterArgs <> "" And currSearchArgs = "" Then
			myFirstOption = "<OPTION VALUE=0>Clear current filter and filter anew</OPTION>"
			mySecondOption = "<OPTION VALUE=1 SELECTED>Compound current filter</OPTION>"	
			myOptionsGuide = "2"
		ElseIf currSortArgs <> "" And currFilterArgs = "" And currSearchArgs = "" Then
			myFirstOption = "<OPTION VALUE=0>Clear current sort results and filter</OPTION>"
			mySecondOption = "<OPTION VALUE=1 SELECTED>Filter and maintain current sort</OPTION>"
			myOptionsGuide = "3"		
		ElseIf currSortArgs = "" And currFilterArgs <> "" And currSearchArgs <> "" Then
			myFirstOption = "<OPTION VALUE=0>Clear current search and filter anew</OPTION>"
			mySecondOption = "<OPTION VALUE=1 SELECTED>Compound current filter</OPTION>"		
			myThirdOption = "<OPTION VALUE=2>Clear current filter and filter anew</OPTION>"
			myFourthOption = "<OPTION VALUE=3>Clear current search and compound filter</OPTION>"				
			myOptionsGuide = "4"
		ElseIf currSortArgs <> "" And currFilterArgs <> "" And currSearchArgs = "" Then
			myFirstOption = "<OPTION VALUE=1 SELECTED>Compound current filter</OPTION>"
			mySecondOption = "<OPTION VALUE=2>Clear sort and compound current filter</OPTION>"
			myThirdOption = "<OPTION VALUE=3>Clear sort and filter anew</OPTION>"	
			myFourthOption = "<OPTION VALUE=4>Maintain current sort, and filter anew</OPTION>"				
			myOptionsGuide = "5"
		ElseIf currSortArgs <> "" And currFilterArgs = "" And currSearchArgs <> "" Then
			myFirstOption = "<OPTION VALUE=0>Clear current sort and filter anew</OPTION>"
			mySecondOption = "<OPTION VALUE=1 SELECTED>Filter current results</OPTION>"		
			myThirdOption = "<OPTION VALUE=2>Clear current search and filter anew</OPTION>"
			myFourthOption = "<OPTION VALUE=3>Clear both sort and search, and filter anew</OPTION>"	
			myOptionsGuide = "6"
		ElseIf currSortArgs <> "" And currFilterArgs <> "" And currSearchArgs <> "" Then
			myFirstOption = "<OPTION VALUE=0>Clear current sort and filter anew</OPTION>"
			mySecondOption = "<OPTION VALUE=1 SELECTED>Compound current filter</OPTION>"		
			myThirdOption = "<OPTION VALUE=2>Clear current search and filter anew</OPTION>"
			myThirdOption = myThirdOption & "<OPTION VALUE=3>Clear current search and compound filter</OPTION>"	
			myFourthOption = "<OPTION VALUE=4>Clear current filter and filter anew</OPTION>"		
			myFourthOption = myFourthOption & "<OPTION VALUE=5>Clear current sort and compound filter</OPTION>"		
			myFourthOption = myFourthOption & "<OPTION VALUE=6>Clear both search and sort, and filter anew</OPTION>"
			myFourthOption = myFourthOption & "<OPTION VALUE=7>Clear both search and sort, and compound filter</OPTION>"				
			myOptionsGuide = "7"
		End If
		
	
		GetFilterOptions = myFirstOption & mySecondOption & myThirdOption & myFourthOption
		
	End Function
	
	
	DoInit
</Script> 

<Form ID="frmFilter">
<table width="100%" ID="mainTable" cellpadding=2 align=left border=0>
	<TR STYLE="visibility: hidden;" ID="optionsDiv">
		<TD colspan=2 ALIGN=left>
			<table align=absBottom border=0 cellspacing=0 cellpadding=0>
					<TR align=left>
						<TD nowrap>
								<font face=verdana size=1><b>Filter Options:</b></font>
								<Script Language="JavaScript">
									var meOptions = GetFilterOptions();				
									if (meOptions!=""){
										document.write ('<SELECT NAME="SortOptions" STYLE="font-family: Verdana; font-size: 7pt">');
										document.write(meOptions);
										document.write ("</SELECT>")
										document.all.item("optionsDiv").style.visibility = "";
										
									}	
												
									
										
								</Script>		
						</TD>
					</TR>
					<TR>
						
					</tr>	
			</table>
			
			</TD>
		</TR>
	<P>	
	<Script Language="VBScript">
		Dim myOptions
		Dim myWorkingOptions
		Dim i
		
		thePage = window.dialogArguments.opener.document.all.item("ActionPage").value
		
		For i = 0 To UBound(myFilterArray)
			thisField = Trim(myFilterArray(i))
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
				
				If thePage="ClientList.asp" And Instr(1,"Name",myOptions)>0 Then
					myOptions = myOptions & "<OPTION SELECTED VALUE=""" & thisFieldName & """ TAG=""" & i & """>" & thisFieldDesc & thepage & "</OPTION>"
				Else
					myOptions = myOptions & "<OPTION VALUE=""" & thisFieldName & """ TAG=""" & i & """>" & thisFieldDesc & "</OPTION>"
				End If
				
			Else
				'myOptions = "<OPTION SELECTED VALUE="""" TAG="""">--</OPTION>"
				myOptions = myOptions & "<OPTION VALUE=""" & thisFieldName & """ TAG=""" & i & """>" & thisFieldDesc & "</OPTION>"
			End If	
		Next
		
		myWorkingOptions = "<OPTION SELECTED VALUE=""0"">Contains</OPTION>"
		myWorkingOptions = myWorkingOptions & "<OPTION VALUE=""1"">Equals</OPTION>"
		myWorkingOptions = myWorkingOptions & "<OPTION VALUE=""2"">Starts with</OPTION>"
		myWorkingOptions = myWorkingOptions & "<OPTION VALUE=""3"">Ends with</OPTION>"
		myWorkingOptions = myWorkingOptions & "<OPTION VALUE=""4"">Not Contains</OPTION>"
		
		
		With document
			.write "<TR>"
			.write "<TD><font face=Verdana size=1><b>Filter by: </b></font></TD>"
			.write "<TD><Select NAME='SearchFld' OnChange=""JavaScript: drawFilterInterface(this.options[this.selectedIndex].TAG)"">"
			.write myOptions
			.write "</Select>"
			.write "</TD></TR>"
			.write "<TR><TD><font face=Verdana size=1><b>Qualification: </b></font></TD>"
			.write "<TD><Select NAME='SearchTerms' OnChange=""JavaScript: checkToDefault(this);"">"
			.write myWorkingOptions
			.write "</Select>"
			.write "</TD>"		
			.write "</TR>"
			.write "<TR>"		
			.write "<TD><font face=Verdana size=1><b>Criterion: </b></font></TD>"
			.write "<TD ID=""inputTD"" align=left>"
			.Write "&nbsp;"
			.write "</TD>"
			.write "</TR>"
			.write "<BR>"
		End	With
		
		
		On Error Resume Next
		Dim theFirstSel
		Set theFirstSel = document.all.item("SearchFld")
		theFirstSel.options(1).selected = True
		drawFilterInterface theFirstSel.options(theFirstSel.selectedIndex).TAG, "0"
	</Script>

	
		<TR ALIGN=RIGHT>	
				<td colspan=2 style="position: absolute; top: 180px; left: 280px; z-index: -1" ALIGN=RIGHT>
					
					<INPUT TYPE=BUTTON NAME=FILTER VALUE=Filter OnClick="VBScript: DoFilter">
					&nbsp;&nbsp;
					<INPUT TYPE=BUTTON NAME=close VALUE=Close OnClick="window.self.close()">
							
				</td>			
	</TR>

</table>

</form>

</BODY>
</HTML>