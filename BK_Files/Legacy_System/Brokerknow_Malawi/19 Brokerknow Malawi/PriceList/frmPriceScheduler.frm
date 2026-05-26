VERSION 5.00
Begin VB.Form frmPriceScheduler 
   Caption         =   "Price List Scheduler"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   2415
   End
End
Attribute VB_Name = "frmPriceScheduler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Conn As ADODB.Connection
Dim Rs As ADODB.Recordset
Public import_UDLPath, sqlStr, fileName As String
Public strPath, strPathContracts, strPathStockwatch, strPathDebtors As String
Private Sub Command1_Click()
   
    Set Rs = New ADODB.Recordset
   
    sqlStr = "select * from PathConfigurations order by DocumentID"
    Rs.Open sqlStr, Conn, 0, 1
    While Not Rs.EOF
       If Rs("DocumentID") = 1 Then strPath = Rs("Path")
       If Rs("DocumentID") = 2 Then strPathStockwatch = Rs("Path")
       If Rs("DocumentID") = 3 Then strPathDebtors = Rs("Path")
       If Rs("DocumentID") = 4 Then strPathContracts = Rs("Path")
        Rs.MoveNext
    Wend
    
    Rs.Close
    If CheckIfFileExists Then
        importPrices
        commitPrices
        GeneratSms
    End If
    Conn.Close
    Set Conn = Nothing
    Unload Me
End Sub

Function CheckIfFileExists() As Boolean
    'Create he FSO objects
     Dim oFileSystem As New FileSystemObject
     Dim oFolder As Folder
     Dim oCurrentFile As File
     Dim oFileColl As Files
     Dim EmailTo, RecieverCC, RecieverBCC As String
     Dim importTime As Variant
     On Error GoTo errhandler
     ' get the file location
     Set oFolder = oFileSystem.GetFolder(strPath)
     Set oFileColl = oFolder.Files
     
     'go through specified folder and pick the todays file
 
     If oFileColl.Count > 0 Then
     For Each oCurrentFile In oFileColl
        If (FormatDateTime(oCurrentFile.DateCreated, vbShortDate) = FormatDateTime(Now(), vbShortDate)) And (FormatDateTime(oCurrentFile.DateLastModified, vbShortDate) = FormatDateTime(Now(), vbShortDate)) And (Len(oCurrentFile.Name) > 15) Then
            If (GetFileExt(oCurrentFile.Name) = ".xls") Then
               fileName = oCurrentFile.Name
            End If
        End If
     Next
     End If
     
    'if the file exists return true else false
     If Trim(fileName) <> "" Then
        CheckIfFileExists = True
     Else
        CheckIfFileExists = False
        'EmailTo = "synjoro@yahoo.com"
        ErrMsg = "Today's Price List was NOT found" & Chr(13)
        ErrMsg = ErrMsg & "Please download and save the price list in the price list folder"
        'get the mail confiturateions here
        Dim rsEmail As ADODB.Recordset
        Set rsEmail = New ADODB.Recordset
        rsEmail.Open "Select * from EmailConfigurations where EmailID=1", Conn, 0, 1
        If Not rsEmail.EOF Or Not rsEmail.BOF Then
           EmailTo = Trim(rsEmail("EmailTo"))
           RecieverCC = Trim(rsEmail("EmailCc"))
           RecieverBCC = Trim(rsEmail("EmailCc"))
           importTime = Trim(rsEmail("importTime"))
        End If
        rsEmail.Close
        Set rsEmail = Nothing
        importTime1 = Split(importTime, ":")
        saaHii = FormatDateTime(Now(), vbLongTime)
        saaHii1 = Split(saaHii, ":")
        If importTime1(0) = saaHii1(0) Then
            SendMail "securities@africanalliance.co.ke", EmailTo, "Price List", ErrMsg, AttachmentPath, RecieverCC, RecieverBCC
        End If
     End If
     'Kill the FSO objects
     Set oFileSystem = Nothing
     Set oFolder = Nothing
     Set oFileColl = Nothing
     Set oCurrentFile = Nothing
  
    Exit Function
errhandler:
    ErrMsg = Err.Description
 End Function
 
Private Function GetFileExt(ByVal fileName As String) As String
    On Error GoTo ErrorHandler
    Dim sFileExt As String
    sFileExt = Mid$(fileName, InStrRev(fileName, "."), Len(fileName) - InStrRev(fileName, ".") + 1)
    GetFileExt = sFileExt
    Exit Function
ErrorHandler:
    GetFileExt = ""
    Exit Function
End Function

Private Function importPrices()
    import_UDLPath = App.Path & "\KBroker.UDL"
    Dim Rs1 As ADODB.Recordset
    'Set Conn = New ADODB.Connection
    Set Rs1 = New ADODB.Recordset
    Set Rs = New ADODB.Recordset
    'Conn.ConnectionString = "FILE NAME=" & import_UDLPath
    'Conn.Open
    sqlStr = "select * from PathConfigurations where DocumentID=1"
    
    
    ''
       Dim xlApp As Excel.Application
       Dim wb As Workbook
       Dim ws As Worksheet
       Dim var As Variant
       
       Set xlApp = New Excel.Application
       
       Set wb = xlApp.Workbooks.Open(strPath & "\" & fileName)
    
       Set ws = wb.Worksheets("price list") 'Specify your worksheet name
       'var = ws.Range("E12").Value
       'or
      ' var = ws.Cells(1, 1).Value
      
      'first delete the previous imports
      
      Conn.Execute ("delete from _ImportPriceList")
      'select from the imports table in readynes for addition
      Rs.Open "Select * from _ImportPriceList where 1=1", Conn, 3, 2
      For i = 12 To 100
         Debug.Print ws.Range("C" & i)
         If UCase(Trim(Mid(ws.Range("C" & i), 1, 23))) = "FIXED INCOME SECURITIES" Then
            Exit For
         End If
         'to get the cells that contain something
         
       If Trim(ws.Range("B" & i)) <> "" Then
        Rs.AddNew
            Rs1.Open "select security_DPA_ from security where LTrim(Rtrim(NSEName)) like '%" & UCase(Trim(Mid(ws.Range("C" & i), 1, 15))) & "%'", Conn, 0, 1
            If Not Rs1.EOF Or Not Rs1.BOF Then
                secDPA = Rs1("security_DPA_")
            Else
                secDPA = Null
            End If
            Rs1.Close
            Rs("secName") = Trim(ws.Range("C" & i))
            Rs("price") = Trim(ws.Range("G" & i))
            Rs("SecKnow_DPA_") = secDPA
        Rs.Update
       End If
       'now insert the  records to the _importpricelist table
    
    
      Next
      
      
      
      ' MsgBox var
       wb.Close 0
       
       xlApp.Quit
       Set Rs1 = Nothing
       Rs.Close
      Set Rs = Nothing
       Set ws = Nothing
       Set wb = Nothing
       Set xlApp = Nothing

       ' Unload Me
        'Conn.Execute ("delete from _ImportPriceList")
End Function

Private Function commitPrices()
    Dim sqlstrDel, sql As String
    Dim maxNumber As Long
    import_UDLPath = App.Path & "\KBroker.UDL"
    Dim Rs1 As ADODB.Recordset
    Dim Rs2 As ADODB.Recordset
    Set Rs1 = New ADODB.Recordset
    Set Rs2 = New ADODB.Recordset
    Set Rs = New ADODB.Recordset
    
    'Conn.ConnectionString = "FILE NAME=" & import_UDLPath
   ' Conn.Open
        
        
      'insert the record to the database
      
   ' Conn.BeginTrans
        'first delete only the Price list imported for that day if any b4  adding the new records
        'sqlstrDel = "Delete from Datastream_Market where MktDate = '" & CDate(Date) & "'"
        sqlstrDel = "Delete from Datastream_Market where MktDate ='" & FormatDate(Now()) & "'"
        Conn.Execute (sqlstrDel)
        ' Rs.Open sqlstrDel, Conn, 3, 2
         'Rs.Close
        sql = "SELECT _ImportPriceList.SecKnow_DPA_, isnull(_ImportPriceList.price,0) as price, _ImportPriceList.importdate, Security.SecurityCode FROM _ImportPriceList INNER JOIN Security ON _ImportPriceList.SecKnow_DPA_ = Security.Security_DPA_ order by secknow_DPA_"
        Rs.Open sql, Conn, 0, 1
        
        Rs1.Open "select * from DataStream_Market where 1=1", Conn, 3, 2
        
        While Not Rs.EOF
        'Add record to Datastream Market
        Rs2.Open "select Max(MktUnique) as MaxUnique from Datastream_Market", Conn, 0, 1
            maxNumber = Rs2("MaxUnique")
        Rs2.Close
           
        If Rs("price") <> 0 And Trim(Rs("price")) <> "" Then
            Rs1.AddNew
                 Rs1("MktUnique") = maxNumber
                 Rs1("MktClose") = Rs("price")
                 Rs1("MktCode") = Rs("SecurityCode")
                 Rs1("MktDate") = Date
            Rs1.Update
        
        Else
            Rs1.AddNew
                 Rs1("MktUnique") = maxNumber
                 Rs1("MktClose") = Null
                 Rs1("MktCode") = Rs("SecurityCode")
                 Rs1("MktDate") = Date
            Rs1.Update

        End If
            
        Rs.MoveNext
        Wend
            
       ' Conn.CommitTrans
        Rs1.Close
        Set Rs1 = Nothing
        Set Rs2 = Nothing
        Rs.Close
        Set Rs = Nothing
        'Conn.Close
        'Set Conn = Nothing
End Function

Private Sub Form_Load()
    import_UDLPath = App.Path & "\KBroker.UDL"
    Set Conn = New ADODB.Connection
    Conn.ConnectionString = "FILE NAME=" & import_UDLPath
    Conn.Open
    
    Command1_Click
End Sub
Function FormatDate(theDate)
    On Error Resume Next
    FormatDate = Day(theDate) & "-" & MonthName(Month(theDate), True) & "-" & Year(theDate)
    If Err.Number > 0 Then
        FormatDate = theDate
    End If
End Function

Function GeneratSms()
    Dim HeaderText, Path As String
    Dim Rs1 As ADODB.Recordset
    Set Rs1 = New ADODB.Recordset
    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset
    sqlStr = "SELECT DISTINCT TOP 100 PERCENT Client_DPA_,ClientCellTel FROM StockWatchList order by client_DPA_"
        Rs.Open sqlStr, Conn, 0, 1
        Dim smsStr3, PricenSec, SmsStr, cellNo, client As String
        'Dim first
        
        'Show = 1
        If Rs.EOF Or Rs.BOF Then
            'send mail here
        Else
             Rs.MoveFirst
            
                While Not Rs.EOF
                'create a string for the sms here
                
                clientNo = Rs("Client_DPA_")
                cellNo = Replace(Rs("ClientCellTel"), " ", "")
                                
                sqlstr1 = "select * from StockWatchList where client_DPA_ =" & Rs("Client_DPA_")

                Rs1.Open sqlstr1, Conn, 0, 1
                
                    i = 0
                    While Not Rs1.EOF
                    'get the securities here
                    
                    If Trim(secStr) = "" Then
                        secStr = Rs1("SecurityCode") & " " & Rs1("Price")
                    Else
                        'check if the characters are greater than 150
                        If Len(secStr & " " & Rs1("SecurityCode") & " " & Rs1("Price")) > 150 Then
                            
                            secStr1 = secStr1 & " " & Rs1("SecurityCode") & " " & Rs1("Price")
                        Else
                        secStr = secStr & " " & Rs1("SecurityCode") & " " & Rs1("Price")
                        End If
                    End If
                    
                    Rs1.MoveNext
                Wend
                 Rs1.Close
                
                'make sure the cell fone number are ok

                If Trim(cellNo) = "" Or cellNo = Null Then
                    cellNo = "*"
                    SmsStr = ""
                Else
                If Mid(cellNo, 1, 1) = 0 Then cellNo = Mid(cellNo, 2, 10) Else cellNo = Mid(cellNo, 1, 9)
                If Not IsNumeric(cellNo) Or Trim(cellNo) = "" Then
                    cellNo = cellNo & " *"
                    SmsStr = ""
                Else
                    ' add the prefix to the cell no
                    cellNo = "CSV: 254" & cellNo
                    SmsStr = cellNo & "|" & secStr & Chr(13)
                    
                End If
                
                End If
                If secStr1 <> "" Then smsStr_1 = cellNo & "|" & secStr1 & Chr(13)
                    smsStr1 = smsStr1 & SmsStr & smsStr_1
                    secStr = ""
                    smsStr_1 = ""
                    secStr1 = ""
                    Rs.MoveNext
                Wend

        End If

        HeaderText = "api-id: 1127740" & Chr(13)
        HeaderText = HeaderText & "user: african alliance" & Chr(13)
        HeaderText = HeaderText & "password: alliance1" & Chr(13)
        HeaderText = HeaderText & "text: #field1#" & Chr(13)
        HeaderText = HeaderText & "Delimiter: |" & Chr(13)
        
        Path = strPathStockwatch & "\sw_" & Replace(Date, "/", "") & ".txt"
        
        smsStr1 = HeaderText & "" & SmsStr
        Debug.Print HeaderText & Chr(13) & smsStr1
        Dim fso As New FileSystemObject
        fso.CreateTextFile Path, True
        
        Open Path For Output As #1
        Print #1, HeaderText & Chr(13) & smsStr1
        Close #1
    
        Set fso = Nothing
       
        
        
End Function
Public Function SendMail(Sender, Receiver, Subject, Message, AttachmentFilePath, ReceiverCC, ReceiverBCC)
    On Error GoTo errhandler
    'On Error Resume Next
    'code to send mail
    
    
    'Set Conn = New ADODB.Connection
    'Dim fso As New FileSystemObject
    'Conn.ConnectionString = "FILE NAME=" & import_UDLPath
    'Conn.Open
         
    Dim RptMail As jmail.Message
    Set RptMail = New jmail.Message
    RptMail.AddRecipient Trim(Receiver)
    If ReceiverCC <> "" Then
        RptMail.AddRecipientCC Trim(ReceiverCC)
    End If
    If ReceiverBCC <> "" Then
        RptMail.AddRecipientBCC Trim(ReceiverBCC)
    End If
    RptMail.From = Sender
    RptMail.Body = Trim(Message)
    RptMail.Subject = Trim(Subject)
    If AttachmentFilePath <> "" Then
        RptMail.AddAttachment AttachmentFilePath
    End If
    'to get mailserver
    Dim rsMail As New ADODB.Recordset
    rsMail.CursorLocation = adUseClient
    rsMail.Open "Select SMTPServer from MailConfiguration", Conn, 0, 1
    If Not rsMail.EOF Or rsMail.BOF Then
       RptMail.Send rsMail("SMTPServer")
       'RptMail.Send "smtp.wananchi.com"
       
    End If
    rsMail.Close
   ' Conn.Close
    Set rsMail = Nothing
    'RptMail.Execute
    Exit Function
errhandler:
    ErrMsg = Err.Number & " " & Err.Description
    Unload Me
    Exit Function
End Function
