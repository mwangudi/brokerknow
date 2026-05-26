<!--#include file="libroutines.asp"-->

<%

'get HTML output
htmlOut = Request.Form("htmlDoc")

'get path of referrer, as this is where the temp HTML page ought to be saved
fromPath = Request.Form("sourcePath")

noOfCopies = Request.Form("txtNumbers")
selPrinter = Request.Form("cboPrinter")
selLandscape = Request.Form("LandScape")
ZoomRatio = Request.Form("ZoomRatio")


'get the name of the report
reportName = Request.Form("reportName")

If htmlOut = "" Or fromPath = "" OR Not IsNumeric(noOfCopies) Or selPrinter = "" Then%>
	<Script Language="JavaScript">
		alert("The printing process cannot proceed. Ensure that you have entered valid details.");
	</Script>
	<%Response.End
End If


If ZoomRatio = "" Or ZoomRatio = "0" Or Not IsNumeric(ZoomRatio) Then
	ZoomRatio = 0 'default, 100
End If

If selLandscape = "" Or selLandScape = "0" Then
	selPaperOrientation = 1 'default, xlpotrait
Else
	selPaperOrientation = 2
End If

	
	'write to output file

	Dim FSO, fTextStream, pathTo, tmpMark, pageFrom, xlFileName
	
	Randomize
	
	tmpMark = "-" & Rnd(2) & "-PNW"

	pathTo = fromPath & "\TMPReport" &  tmpMark & ".html"
	xlFileName = "TMPReport" &  tmpMark & ".xls"
	newPathTo = fromPath & "\" & xlFileName

	Set FSO = CreateObject("Scripting.FileSystemObject")

	Set fTextStream = Fso.CreateTextFile(pathTo, True, False)
	fTextStream.Write htmlOut
	fTextStream.Close
	Set fTextStream = Nothing
	
	
	'end write
	
	'export to xl
	Dim appX, wkBk, wkSheet
    
    Set appX = CreateObject("Excel.Application")
    appX.Visible = False   
    
    'Set wkBk = appX.Workbooks.Add
    Set wkBk = appX.WorkBooks.Open(pathTo)
    Set wkSheet = wkBk.Worksheets(1)
    'wkSheet.Name = reportName    
    
    'save file first
    wkBk.SaveAs newPathTo,  -4143, "", "", False, False
    
    'On Error Resume Next
     'With wkSheet.QueryTables.Add("URL;" & pathTo, wkSheet.Range("A1"))
     '  .Name = reportName
     '   .FieldNames = True
     '   .RowNumbers = False
     '   .FillAdjacentFormulas = False
     '   .PreserveFormatting = False
     '   .RefreshOnFileOpen = False
     '   .BackgroundQuery = True
     '   .RefreshStyle = 1 '=xlInsertDeleteCells
     '   .SavePassword = False
     '   .SaveData = True
     '   .AdjustColumnWidth = True
     '   .RefreshPeriod = 0
     '   .WebSelectionType = 1 '=xlEntirePage
     '   .WebFormatting = 1  '=xlWebFormattingAll
     '   .WebPreFormattedTextToColumns = True
     '   .WebConsecutiveDelimitersAsOne = True
     '   .WebSingleBlockTextImport = False
     '   .WebDisableDateRecognition = True
     '   .Refresh False        
   ' End With
    
    'wkSheet.Columns("A:A").ColumnWidth = 15  
    
    With wkBk
        .UpdateRemoteReferences = False
        .PrecisionAsDisplayed = False
        .Windows(1).DisplayGridlines = False
        .Windows(1).GridlineColorIndex = 2
        .Windows(1).DisplayHeadings = False
        .Windows(1).DisplayWorkbookTabs = False
    End With
    
     'format print data
    'On error Resume Next
		With wkSheet.PageSetup
			.PrintTitleRows = ""
			.PrintTitleColumns = ""
			.PrintArea = ""    
			.LeftHeader = ""
			.CenterHeader = ""
			.RightHeader = "&""Arial,Bold""African Alliance" & Chr(10) & "&""Arial,Regular""Page &P of &N"
			.LeftMargin = wkSheet.Application.InchesToPoints(0.75)
			.RightMargin = wkSheet.Application.InchesToPoints(0.75)
			.TopMargin = wkSheet.Application.InchesToPoints(1)
			.BottomMargin = wkSheet.Application.InchesToPoints(1)
			.HeaderMargin = wkSheet.Application.InchesToPoints(0.5)
			.FooterMargin = wkSheet.Application.InchesToPoints(0.5)
			.PrintHeadings = False
			.PrintGridlines = False
			.PrintComments = -4142 'xlPrintNoComments			
			.CenterHorizontally = False
			.CenterVertically = False
			.Orientation = selPaperOrientation 'xlPortrait =1, xlLandScape = 2
			.Draft = False
			.PaperSize = 9 'xlPaperA4 = 9, xlPaperA3 = 8, xlPaperLetter = 1
			.FirstPageNumber = -4105 '=xlAutomatic
			.Order = 1 '=xlDownThenOver, xlOverThenDown = 2
			.BlackAndWhite = False
			.Zoom = CDbl(ZoomRatio) '75			
			'.PrintErrors = 3 '=xlPrintErrorsNA, xlPrintErrorsDash = 2, xlPrintErrorsBlank = 1
			
		End With
    Err.Clear
    
    'end format print data
       
    'print info
    'wkSheet.PrintOut ,, noOfCopies, False, selPrinter, False
    
    ''protect data sheet
    'Randomize
    'wkSheet.Protect "Rep" & rnd(8) & "CK", True, True, True
    'end protect
    
    
    wkBk.Close True
    
    Set wkSheet = Nothing
    Set wkBk = Nothing        
    appX.Quit
    Set appX = Nothing
    
    'end export

	'transfer user o file
	
	'pageFrom = Request.ServerVariables("HTTP_REFERER")
	'pageFrom = StrReverse(pageFrom)
	'pageFrom = Mid(pageFrom, InStr(1, pageFrom, "/", vbTextCompare))
	'pageFrom = StrReverse(pageFrom)
		
	
	With Response		
		.Write "<Script Language=JavaScript>" & vbCrLf
		'.Write "	window.open('XLReport.asp?xlFile=" & pageFrom & xlFileName & "', null, 'height=600,width=700,status=no,scrollbars=yes,toolbar=no,menubar=no,location=no,resizable=yes'); " & vbCrLf
		'.Write "	window.open('" & pageFrom & xlFileName & "', null, 'height=600,width=700,status=no,scrollbars=yes,toolbar=no,menubar=yes,location=no,resizable=yes'); " & vbCrLf				
		'.Write "	window.location.replace('XLReport.asp?xlFile=" & pageFrom & xlFileName & "');"  & vbCrLf		
		.Write "	alert('Kindly pick the print out from the selected printer');"
		.Write "</Script>"		
	End With
	
	
	'delete earlier HTML output file
	FSO.DeleteFile pathTo, True
	

	
	'FSO.DeleteFile newPathTo, True
	
	Set FSO = Nothing

%>



