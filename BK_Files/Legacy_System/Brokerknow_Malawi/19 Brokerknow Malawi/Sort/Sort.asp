<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html;charset=UTF-8">
<META HTTP-EQUIV="Expires" CONTENT="0">
<TITLE>Sort</TITLE>
<LINK href="../STYLE/default.css" type=TEXT/CSS rel=STYLESHEET>
</HEAD>
<BODY CLASS="FadeBlue">

<!--#include file="../libroutines.asp"-->

<Script Language="VBScript">
	Dim mySortArray
	Dim myOptionsGuide
	
	Const fixedMaxSortFields = 3
	
	Function DoInit		
		If window.dialogArguments <> "" Then
			window.name = "sortWindow"			
			window.defaultStatus = "Sort Window"
			mySortArray = Split(window.dialogArguments.sortArgs, ";")			
    
		Else
			window.self.close
		End If
   
	End Function
	

	Function DoSort
		Dim mySortArgs, optSelVal, optCompound
		Dim optClearSearch, optClearFilter
		
		optCompound = True
		optClearSearch = False
		optClearFilter = False
		
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
			optSelVal = document.all.item("SortOptions").value
			
			Select Case myOptionsGuide
				Case "0"
					optCompound = False
				Case "1"
					If optSelVal = "0" Then optClearSearch = True
					optCompound = False
						
				Case "2"
					optCompound = False
					If optSelVal = "0" Then optClearFilter = True
				Case "3"
					If optSelVal = "0" Then optCompound = False
				Case "4"
					optCompound = False
					If optSelVal = "2" Then 
						optClearSearch = True
					ElseIf optSelVal = "3" Then 
						optClearFilter = True
					ElseIf optSelVal = "4" Then 
						optClearSearch = True
						optClearFilter = True
					End If	
				Case "5"
					If optSelVal = "0" Then 
						optCompound = False
					ElseIf optSelVal = "2" Then 
						optClearFilter = True
						optCompound = False
					ElseIf optSelVal = "3" Then 
						optClearFilter = True
					End If	
				Case "6"
					If optSelVal = "0" Then 
						optCompound = False
					ElseIf optSelVal = "2" Then 
						optClearSearch = True
						optCompound = False
					ElseIf optSelVal = "3" Then 
						optClearSearch = True	
					End If	
				Case "7"
					If optSelVal = "0" Then 
						optCompound = False
					ElseIf optSelVal = "2" Then 
						optClearSearch = True
						optCompound = False
					ElseIf optSelVal = "3" Then 
						optClearSearch = True
					ElseIf optSelVal = "4" Then 
						optClearFilter = True
						optCompound = False
					ElseIf optSelVal = "5" Then 
						optClearFilter = True
					ElseIf optSelVal = "6" Then 
						optClearFilter = True
						optClearSearch = True
						optCompound = False
					ElseIf optSelVal = "7" Then 
						optClearFilter = True
						optClearSearch = True	
					End If
			End Select
				
			If optCompound Then
				window.dialogArguments.opener.document.all.item("SelectedSortArgs").value = window.dialogArguments.opener.document.all.item("SelectedSortArgs").value & ", " & mySortArgs
			Else
				window.dialogArguments.opener.document.all.item("SelectedSortArgs").value = mySortArgs
			End If	
			
			If optClearFilter Then
				window.dialogArguments.opener.document.all.item("SelectedFilterArgs").value = ""
			End If
			
			If optClearSearch Then
				window.dialogArguments.opener.document.all.item("SelectedSearchArgs").value = ""
			End If
			
			
			window.dialogArguments.opener.document.all.item("frmMain").action = window.dialogArguments.sortTarget
			window.dialogArguments.opener.document.all.item("frmMain").submit
			window.dialogArguments.sortTD.className = "footerHighlightNavOn"
			window.dialogArguments.cancelButton.style.display = ""
			window.self.close				
		Else
			window.alert("The sort facility does not have enough information to proceed." & Chr(13) & "Ensure that at least one sort option is selected.")
		End If
		
	End Function
	
	Function GetSortOptions
		Dim currSortArgs, currSearchArgs, currFilterArgs, myFirstOption, mySecondOption
		Dim myThirdOption, myFourthOption
		
		currSortArgs = window.dialogArguments.opener.document.all.item("SelectedSortArgs").value
		currSearchArgs = window.dialogArguments.opener.document.all.item("SelectedSearchArgs").value
		currFilterArgs = window.dialogArguments.opener.document.all.item("SelectedFilterArgs").value
		
		If currSortArgs = "" And currSearchArgs = "" And currFilterArgs = "" Then
			myOptionsGuide = "0"
			GetSortOptions = ""
			Exit Function
		ElseIf currSortArgs = "" And currFilterArgs = "" And currSearchArgs <> "" Then
			myFirstOption = "<OPTION VALUE=0>Clear current search results and sort</OPTION>"
			mySecondOption = "<OPTION VALUE=1 SELECTED>Sort current search results</OPTION>"
			myOptionsGuide = "1"
		ElseIf currSortArgs = "" And currFilterArgs <> "" And currSearchArgs = "" Then
			myFirstOption = "<OPTION VALUE=0>Clear current filter results and sort</OPTION>"
			mySecondOption = "<OPTION VALUE=1 SELECTED>Sort current filter results</OPTION>"	
			myOptionsGuide = "2"
		ElseIf currSortArgs <> "" And currFilterArgs = "" And currSearchArgs = "" Then
			myFirstOption = "<OPTION VALUE=0>Clear current sort results and sort anew</OPTION>"
			mySecondOption = "<OPTION VALUE=1 SELECTED>Compound current sort</OPTION>"
			myOptionsGuide = "3"		
		ElseIf currSortArgs = "" And currFilterArgs <> "" And currSearchArgs <> "" Then
			myFirstOption = "<OPTION VALUE=1 SELECTED>Sort search and filter results</OPTION>"
			mySecondOption = "<OPTION VALUE=2>Clear current search results and sort</OPTION>"
			myThirdOption = "<OPTION VALUE=3>Clear current filter results and sort</OPTION>"	
			myFourthOption = "<OPTION VALUE=4>Clear both filter and search results, and sort</OPTION>"
			myOptionsGuide = "4"
		ElseIf currSortArgs <> "" And currFilterArgs <> "" And currSearchArgs = "" Then
			myFirstOption = "<OPTION VALUE=0>Clear current sort and sort anew</OPTION>"
			mySecondOption = "<OPTION VALUE=1 SELECTED>Compound current sort</OPTION>"		
			myThirdOption = "<OPTION VALUE=2>Clear current filter and sort anew</OPTION>"
			myFourthOption = "<OPTION VALUE=3>Clear current filter and compound sort</OPTION>"	
			myOptionsGuide = "5"
		ElseIf currSortArgs <> "" And currFilterArgs = "" And currSearchArgs <> "" Then
			myFirstOption = "<OPTION VALUE=0>Clear current sort and sort anew</OPTION>"
			mySecondOption = "<OPTION VALUE=1 SELECTED>Compound current sort</OPTION>"		
			myThirdOption = "<OPTION VALUE=2>Clear current search and sort anew</OPTION>"
			myFourthOption = "<OPTION VALUE=3>Clear current search and compound sort</OPTION>"	
			myOptionsGuide = "6"
		ElseIf currSortArgs <> "" And currFilterArgs <> "" And currSearchArgs <> "" Then
			myFirstOption = "<OPTION VALUE=0>Clear current sort and sort anew</OPTION>"
			mySecondOption = "<OPTION VALUE=1 SELECTED>Compound current sort</OPTION>"		
			myThirdOption = "<OPTION VALUE=2>Clear current search and sort anew</OPTION>"
			myThirdOption = myThirdOption & "<OPTION VALUE=3>Clear current search and compound sort</OPTION>"	
			myFourthOption = "<OPTION VALUE=4>Clear current filter and sort anew</OPTION>"		
			myFourthOption = myFourthOption & "<OPTION VALUE=5>Clear current filter and compound sort</OPTION>"		
			myFourthOption = myFourthOption & "<OPTION VALUE=6>Clear both search and filter, and sort anew</OPTION>"
			myFourthOption = myFourthOption & "<OPTION VALUE=7>Clear both search and filter, and compound sort</OPTION>"				
			myOptionsGuide = "7"
		End If
		
	
		GetSortOptions = myFirstOption & mySecondOption & myThirdOption & myFourthOption
		
	End Function
	
	DoInit
</Script> 
<table border=0 cellspacing=1 cellpadding=2 width=100%>
	<TR>
		<TD COLSPAN=2>
			<table border=0 cellspacing=0 cellpadding=1>
				<tr>
					<td width=25%><font face=Verdana size=1>Sort by</font></td>
					<td width=100% align=left nowrap><hr width="100%"> </td>
				</tr>
			</table>	
		</TD>
		
	</TR>
	<Script Language="VBScript">
		dim myOptions
		
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
					.write "<td align=left width=""75%"" nowrap><hr width=100%>"
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
	</Script>
	
</table>


	<table align=Top width="100%" border=0 cellspacing=0 cellpadding=2>
			<TR align=left STYLE="visibility: hidden" ID="optionsDiv">
				<TD nowrap>
						<font face=verdana size=1>Sort Options:</font>
						<SELECT NAME="SortOptions" STYLE="font-family: Verdana; font-size: 7pt">
							<Script Language="JavaScript">
								var meOptions;
								try{
									meOptions = GetSortOptions();
									if (meOptions!=""){
										document.write(meOptions);
										document.all.item("optionsDiv").style.visibility = "";
									}									
								}
								catch(e){}
							</Script>
						</SELECT>					
				</TD>
			</TR>
			<TR align=right>
				<td valign=absTop align=right>
					<INPUT TYPE=BUTTON NAME=SORT VALUE=Sort OnClick="VBScript: DoSort">
					&nbsp;&nbsp;
					<INPUT TYPE=BUTTON NAME=close VALUE=Close OnClick="window.self.close()">
					&nbsp;&nbsp;
					
				</td>
					
			</tr>	
	</table>


</BODY>
</HTML>