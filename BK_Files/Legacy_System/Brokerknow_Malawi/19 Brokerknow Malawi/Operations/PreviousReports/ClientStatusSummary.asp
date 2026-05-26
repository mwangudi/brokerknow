<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Client Status Summary</title>
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
	 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
	 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>

<style media="print">
	
		@page {
			margin-left: 2cm;
			margin-right: 5cm;
			margin-top: 1cm;    
			margin-bottom: 2cm;
			writing-mode: tb-rl;
			height: 80%;
			margin: 10% 0%;						
			br.newpage{
				page-break-before:always;
			}		
		}		 
		
	</style>

</head>

<body Class="Reports">

<!--#include file="../libroutines.asp"-->

<%


genReport = Request.Form("genReport")
report_description = "Client Status Summary"
selColumns = Request.Form("customCols")
useCustomization = Request.Form("useOwnFields")
SelectedSearchArgs = Request.Form("SelectedSearchArgs")
orderByCols = Request.Form("SelectedSortArgs")

If genReport = ""  Then%>
	<Script Language="JavaScript">
		try{
			if (window.parent.name !== 'KNWNG')
					document.body.className = 'dialog';
		}			
		catch(e){}
		
		function validateForm(frm){			
			var doc = document.all.item('customCols');
			if (doc.length >= 0) SelectAll(doc);			
			DoSearch();
			DoSort(); 
			frm.target = '_self';			
			frm.submit();
		}
		
		function switchDisplay(obj){
			if (obj.style.display=='none') obj.style.display = '';
			else obj.style.display = 'none';
		}
		
		 function SelectAll(Object){
		 //select all upwards
		  for (loop=Object.length-1; loop>-1;	   loop--)	
		    {	
		     Object.options[loop].selected = true
		    }	
		 }

		function moveDown(){
			var doc = document.all.item('customCols');
			var currIndex = doc.selectedIndex;
			if (currIndex >= 0){
				try{
					var nextOptionIndex = currIndex - 1 + 2;
					var tempOptionText = doc.options[nextOptionIndex].text;
					var tempOptionValue = doc.options[nextOptionIndex].value;
					doc.options[nextOptionIndex].text = doc.options[currIndex].text;
					doc.options[nextOptionIndex].value = doc.options[currIndex].value;
					doc.options[currIndex].text = tempOptionText;
					doc.options[currIndex].value = tempOptionValue;
					doc.selectedIndex = nextOptionIndex;
				}	
				catch(e){}
			}

		}

		function moveUp(){
			var doc = document.all.item('customCols');
			var currIndex = doc.selectedIndex;
			if (currIndex >= 0){
				try{
					var nextOptionIndex = currIndex - 1;
					var tempOptionText = doc.options[nextOptionIndex].text;
					var tempOptionValue = doc.options[nextOptionIndex].value;
					doc.options[nextOptionIndex].text = doc.options[currIndex].text;
					doc.options[nextOptionIndex].value = doc.options[currIndex].value;	
					doc.options[currIndex].text = tempOptionText;
					doc.options[currIndex].value = tempOptionValue;
					doc.selectedIndex = nextOptionIndex;
				}
				catch(e){}
			}
		}
		
		function AddOption(Input,Output){    
   		    	 NewOption = new Option();   			    
			    NewOption.text = Input.id;
			    NewOption.value = Input.value;
			    NewOption.selected = false;	
				Output.add(NewOption, 0)   		
    	    
 		 }
 		 
 		 function RemoveOption(Input, Field){
			Selection = new Boolean();
			for (loop=Field.length - 1; loop >= 0; loop--) {
	   			 var GoneOption = Field.options[loop]
				if (GoneOption.value==Input.value) {
	      			Selection = true;
	      			Field.remove(GoneOption.index);
	      		}
	      	}	
		}
		
		function evalCheck(chk){
			if (chk.checked){
				AddOption(chk, document.getElementById('customCols'));
			}
			else {
				RemoveOption(chk, document.getElementById('customCols'));
			}
		}


	</Script>
	<form method="POST" action="ClientStatusSummary.asp" Name="frmMain" id="frmMain">
		
		<input type="hidden" name="genReport" id="genReport" value="1">
		<input type="hidden" name="SelectedSearchArgs" id="SelectedSearchArgs">
		<input type="hidden" name="SelectedSortArgs" id="SelectedSortArgs">		
		<table>
			<tr>
				<td colspan="2"> <input type="checkbox" OnClick="JavaScript: switchDisplay (document.all.item('OwnFieldsSelectRow')); " class="BorderLess" name="useOwnFields" id="useOwnFields" value="1"> &nbsp; &nbsp; <label for="useOwnFields" style="cursor: hand">Customize this report (optional)</label></td>
			</tr>
			
			

			<tr style="display: none;" id="OwnFieldsSelectRow" align="right">
				<td colspan=2 valign="top" style="PADDING-LEFT: 25px">
					<TABLE border="0" cellpadding="2" cellspacing="5">
						<TR>
							<TD nowrap valign="top" STYLE="border: 1px ridge #000080">
 								Select columns to be included in the report
						<%
						Set Conn = GetActiveConnection("KBroker")
						Set Rs = Conn.OpenSchema (adSchemaColumns, Array(Empty, Empty, "ClientStatusSummary"))

						If Not (Rs.EOF Or Rs.BOF) Then
							Do Until Rs.EOF%>
								<br>
								<input type="checkbox" OnClick="JavaScript: evalCheck(this)" class="BorderLess" name="selColumns" value="[<%= Rs.Fields("COLUMN_NAME").Value %>]" id="<%= Rs.Fields("COLUMN_NAME").Value %>">
								<label for="<%= Rs.Fields("COLUMN_NAME").Value %>"><font face="Arial" color="navy" style="cursor: hand" TITLE="<%= Rs.Fields("COLUMN_NAME").Value %>"><%= Rs.Fields("COLUMN_NAME").Value %></font></label>
								
							<%
								dataSearchValue = "[" & Rs.Fields("COLUMN_NAME").Value & "]:" & Rs.Fields("COLUMN_NAME").Value & "*0"
							   	dataSortValue = "[" & Rs.Fields("COLUMN_NAME").Value & "]:" & Rs.Fields("COLUMN_NAME").Value   
								If 	dataFldStr = "" Then
									dataFldStr =  dataSearchValue
									SortFieldsStr = dataSortValue   
								Else
									dataFldStr = dataFldStr & ";" & dataSearchValue 
									SortFieldsStr = SortFieldsStr & ";" & dataSortValue   
								End If
																
 
								Rs.MoveNext
							Loop
						End If%>								
						</TD>
							<TD STYLE="border: 1px ridge #000080" valign="top">
								Change order of appearance of columns included in the report
								<TABLE border="0" cellpadding="2" cellspacing="5">
									<TR>
										<TD nowrap valign="top">
 
	
										<select name="customCols" multiple size="8">
															
											</select>
										</td>
										<td wrap valign="middle">	
											
										<img  STYLE="cursor: hand" class="buttons" TITLE="Move Up" src="../images/moveUp.gif" OnClick="JavaScript: moveUp();">
										<br>
										<img class="buttons" TITLE="Move Down" src="../images/moveDown.gif" OnClick="JavaScript: moveDown();" STYLE="cursor: hand">
																			
										</td>
									</tr>
								</table>		
					</TD>
					</tr>
					<tr>
						<td>
							<TABLE border="0" cellpadding="2" cellspacing="5" width="100%">
							<TR>
								<TD nowrap valign="top" STYLE="border: 1px ridge #000080">
 									Add custom filters
 									<br>
 									
 									
 									<table ID="mainTable">
	<TR>
		<TD><font face=Verdana size=1><b>Search by</b></font>			
		</TD>
		<TD COLSPAN=2><font face=Verdana size=1><b>Qualification</b></font>			
		</TD>
	</TR>
		
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

			mySearchArray = Split("<%= dataFldStr  %>", ";")			
       
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
		
		
		If searchStr <> "" Then
			document.getElementById("SelectedSearchArgs").value = searchStr
		End If			
		
	End Function
	
	DoInit
	
</Script> 

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
					.write "<td COLSPAN=3><br>"
					.write "</td></tr>"
		End With			
		
		On Error Resume Next
		Dim theFirstSel
		Set theFirstSel = document.all.item("SearchFld0")
		theFirstSel.options(1).selected = True
		drawFilterInterface theFirstSel.options(theFirstSel.selectedIndex).TAG, "0"
		
	</Script>
	
</table>
 									
 									
 									
 									
 									
 									
 									
 									
 									
 								</td>
 							</tr>
 							</table>	
 						</td>
 						<td valign="top">
 							<TABLE border="0" cellpadding="2" cellspacing="5" width="100%">
							<TR>
								<TD nowrap valign="top" STYLE="border: 1px ridge #000080">
 									Add custom sort
 									<br>
 									
 									
 									
 									
 									
<Script Language="VBScript">
	Dim mySortArray
	Dim myOptionsGuide
	
	Const fixedMaxSortFields = 3
	
	Function DoInitSort		
			mySortArray = Split("<%= SortFieldsStr %>", ";")	    
	End Function
	

	Function DoSort
		Dim mySortArgs, optSelVal		

	
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
	
		
	DoInitSort
</Script> 
<table border=0 cellspacing=1 cellpadding=2 width=100%>
	<TR>
		<TD COLSPAN=2>
			<table border=0 cellspacing=0 cellpadding=1>
				<tr>
					<td width=25%><font face=Verdana size=1>Sort by</font></td>
					<td width=100% align=left nowrap>&nbsp;</td>
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
	</Script>
	
</table>

 									
 									
 									
 									
 									
 									
 									
 									

								</td>
							</tr>
							</table>	
						</td>
					</tr>
					
					</table>
				</td>
			</tr>
			
			
			<tr>
				<td colspan=2><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.getElementById('frmMain'))" Value=" Generate... ">&nbsp;&nbsp; <input type="Button" class="Buttons" Value=" Close " OnClick="JavaScript: window.parent.self.close();"></td>
			</tr>
		</table>
		
	</form>
	
	
	
	


	
	
	<%Set rs = Nothing
	Set Conn = Nothing
	Response.End
End If

DrawPageFunctions True, True, True



If useCustomization <> "1" Then
	selColumns = "*"
	orderByCols = ""
	SelectedSearchArgs = ""
Else
	If selColumns = "" Then
		selColumns = "*"		
	End If
	
	If orderByCols <> "" Then
		orderByCols = " ORDER BY " & orderByCols 	
	End If	
	
	If SelectedSearchArgs  <> "" Then
		SelectedSearchArgs = "WHERE " & SelectedSearchArgs 
	End If
	
End If	



sqlStr = "SELECT " & selColumns & " FROM ClientStatusSummary " & SelectedSearchArgs & " " &  orderByCols 
	

Set Conn = GetActiveConnection("KBroker")

Set Rs = Conn.Execute(sqlStr)

 If Rs.EOF Or Rs.BOF Then%>
		<Script Language="JavaScript">	
			ShowMessage('No information was found using the criteria entered.');
			window.parent.history.go(-1);
		</Script>
		<%Set Conn = Nothing
		Set Rs = Nothing
		Response.End
  End If


 %>

<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
     <tr>
		<td nowrap><b><font face="Arial Narrow" size="4"><%= report_description %></font></b></td>
		<td nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>
	</tr>	
       <tr>
		  <td COLSPAN=2><font face="Arial" size="2">&nbsp;</font></td>
	</tr>
</table>			



    <table border="0" width="100%" cellPadding="3" cellSpacing=0>
    <tr bgColor="#000000">
			
	<%For i = 0 To Rs.Fields.Count - 1%>		
	 	 <td nowrap><b><font color="#FFFFFF"><%= Rs.Fields(i).Name %></font></b></td>
   <%Next %>
	</tr>
	
	<% Do Until rs.EOF%>
        		<tr>
      				
      				<%For i = 0 To Rs.Fields.Count - 1%>		
					 	 <td nowrap><%= Rs.Fields(i).Value %></td>
				   <%Next %>			
      
	        </tr>
	 <%  Rs.MoveNext
	 Loop%>	  	
  </table>
  
	 <%	 
	 Set Rs = Nothing
	 Set Conn = Nothing
     %>	
</body>



</html>
