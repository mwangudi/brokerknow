<%@ Language=VBScript %>
<HTML>
<BODY>
<Script Language=VBScript>
Sub CreateDataDoc(oApp)
  ' Declare variables.
  Dim sServer,oDoc,oRS,sTemp,sHead,oRange,oField
  
  ' Place your server's name here.
  sServer = "user"
  ' Create a new document.
  Set oDoc = oApp.Documents.Add
  ' Create a new recordset.
  Set oRS = CreateObject("ADODB.Recordset")
  ' Open the XML recordset from the server and pass the SQL statement
  ' to the Getdata.asp page.
  sSQL = "SELECT * FROM AUTHORS"
  oRS.Open "http://" & sServer & ":80/WordMailMerge/Getdata.asp?SQL=" & sSql
  ' Convert the recordset to a string.
  sTemp = oRS.GetString(2, -1, vbTab)  ' 2 = adClipString
         
  ' Append the field names to the front of the string.
  For Each oField In oRS.Fields
    sHead = sHead & oField.Name & vbTab
  Next
        
  ' Strip off the last tab.
  sTemp = Mid(sHead, 1, Len(sHead) - 1) & vbCrLf & sTemp
         
  ' Get a range object and insert the text into the document.
  Set oRange = oDoc.Range
  oRange.Text = sTemp
  
  ' Convert the text to a table.
  oRange.ConvertToTable vbTab
  ' Save the document to a temp file.
  oDoc.SaveAs "d:\data.doc"
  ' Close the document (no save).
  oDoc.Close False
End Sub


Sub ButtonClick()
  Dim oApp
  Dim oDoc
  Dim oMergedDoc
  
	'MsgBox "Creating word instance..."
  ' Create an instance of Word.     
  Set oApp = CreateObject("Word.Application")
 	
	'MsgBox "Calling data file..."
  ' Create our data file.
  CreateDataDoc oApp
  
  ' Add a new document.
	MsgBox "Setting up new doc..."
  Set oDoc = oApp.Documents.Add
  With oDoc.MailMerge
    ' Add our fields.
    .Fields.Add oApp.Selection.Range, "au_fname"
    oApp.Selection.TypeText " "
    .Fields.Add oApp.Selection.Range, "au_lname"
    oApp.Selection.TypeParagraph
    .Fields.Add oApp.Selection.Range, "city"
    oApp.Selection.TypeText ", "
    .Fields.Add oApp.Selection.Range, "state"
    oApp.Selection.TypeParagraph
    .Fields.Add oApp.Selection.Range, "zip"
    oApp.Selection.TypeParagraph
           
    ' Create an autotext entry.
    Dim oAutoText
    Set oAutoText = oApp.NormalTemplate.AutoTextEntries.Add _
    ("MyLabelLayout", oDoc.Content)
    oDoc.Content.Delete
    .MainDocumentType = 1  ' 1 = wdMailingLabels
         
    ' Open the saved data source.
    .OpenDataSource "d:\data.doc"

    ' Create a new document.
    oApp.MailingLabel.CreateNewDocument "5160", "", _
         "MyLabelLayout", , 4  ' 4 = wdPrinterManualFeed

    .Destination = 0  ' 0 = wdSendToNewDocument
    ' Execute the mail merge.
    .Execute

    oAutoText.Delete
  End With

'	MsgBox "closing..."
       
  ' Close the mail merge edit document.
  oDoc.Close False
  ' Get the current document.
  Set oMergedDoc = oApp.ActiveDocument
  ' Show Word to the user.
  oApp.Visible = True
  
  ' Uncomment these lines to save the merged document locally.
  'oMergedDoc.SaveAs "d:\test.doc"
  'oMergedDoc.Close False
  'oApp.Quit False
End Sub

</Script>

<INPUT type=button value="Create Word Document" onclick="VBScript:ButtonClick">
</BODY>
</HTML>
