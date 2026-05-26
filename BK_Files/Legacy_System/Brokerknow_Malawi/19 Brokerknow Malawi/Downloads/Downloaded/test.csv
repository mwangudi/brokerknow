VERSION 5.00
Begin VB.Form frmHoldingImport 
   Caption         =   "HoldingsImport"
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
      Top             =   960
      Width           =   2415
   End
End
Attribute VB_Name = "frmHoldingImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Private MyScriptingContext As ScriptingContext
    Private Myscript As ScriptingContext
    Private MyRequest As Request
    Private MyResponse As Response
    
    Dim pathToFile, CurrentDirectory As String
    Dim fileName, import_UDLPath, FilePart As String
    Dim conn As ADODB.Connection
    Dim rsImport As ADODB.Recordset
    Dim rsCommit As ADODB.Recordset
    Dim rst As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim fomartedText As String
    Dim clientRegistration  As String
    Dim security As String
    Dim shareQty As String
    Dim accStatus As String
    Dim tradeDate As String
    Dim tradeTime As String
    Dim clientReg As String
    Dim i As Integer
    Dim TextLine As String
    Dim Errmsg As String
    Public strPath As String
    Dim ImportSuccess As Boolean
    
   
 Function CheckIfFileExists() As Boolean
    'Create he FSO objects
     Dim oFileSystem As New FileSystemObject
     Dim oFolder As Folder
     Dim oCurrentFile As File
     Dim oFileColl As Files
     On Error GoTo errhandler
     ' get the file location
     
     Set oFolder = oFileSystem.GetFolder(strPath)
     
     Set oFileColl = oFolder.Files
     
     'go through specified folder and pick the todays file
 
     If oFileColl.Count > 0 Then
     For Each oCurrentFile In oFileColl
        If (FormatDateTime(oCurrentFile.DateCreated, vbShortDate) = FormatDateTime(Now(), vbShortDate)) And (FormatDateTime(oCurrentFile.DateLastModified, vbShortDate) = FormatDateTime(Now(), vbShortDate)) And (Len(oCurrentFile.Name) > 15) Then
            If (GetFileExt(oCurrentFile.Name) = ".txt") Then
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
        Errmsg = "No holding file was found for import "
     End If
     
     'Kill the FSO objects
     Set oFileSystem = Nothing
     Set oFolder = Nothing
     Set oFileColl = Nothing
     Set oCurrentFile = Nothing
    Exit Function
errhandler:
    If Err.Number = 76 Then
       Errmsg = "Invalid Client Holdings Import File Path."
    Else
       Errmsg = Err.Description
    End If
 
 End Function
 
 Function CopyFile() As Boolean
  
    If CheckIfFileExists = True Then
        CurrentDirectory = App.Path & "\"
        Dim fso As New FileSystemObject
        ' check whether the file already exists and if not copy the file to the current directory
        If fso.FileExists(CurrentDirectory & fileName) Then
           
            Errmsg = "Brokerknow could not import " & fileName & ". File already exists"
            CopyFile = False
        Else
           'Ensure root folder is referenced correctly
           If Mid(strPath, Len(strPath), 1) <> "\" Then strPath = strPath & "\"
    
            fso.CopyFile strPath & fileName, CurrentDirectory & fileName, False
            
            Errmsg = fileName & " was imported successfuly."
            CopyFile = True
        End If
        
        Set fso = Nothing
    Else
           Unload Me
            'MsgBox Errmsg
     End If
 End Function
 
Private Sub Command1_Click()

   'Get the CDS_holding path
    If App.PrevInstance = True Then
        Unload Me
        Exit Sub
    End If
    
    Dim fso As New FileSystemObject
  
    If fso.FileExists(App.Path & "\path.txt") Then
        Open App.Path & "\path.txt" For Input As 1
        Do While Not EOF(1)
            Line Input #1, strPath
            strPath = Trim(strPath)
        Loop
     Else
        MsgBox "Please include the path.txt file in the imports folder"
        Errmsg = "Please include the path.txt file in the imports folder"
         'Notify Users
          SendMail Errmsg
        Unload Me
        
        Exit Sub
     End If
     Close #1
     Set fso = Nothing
     
    'Call CheckIfFileExists
    If CopyFile = True Then
    
        import_UDLPath = App.Path & "\UDL\KBroker.UDL"
        importCds App.Path & "\" & fileName, "" & import_UDLPath, "_CDS_Imported_Holdings_"
        
        If ImportSuccess = True Then
            commitHoldings (import_UDLPath)
        Else
            If Trim(Errmsg) = "" Then Errmsg = "An Unspecified Error occured.Client Holdings were not imported."
            SendMail Errmsg
           
            Unload Me
        Exit Sub
        End If
    Else
        If Trim(Errmsg) = "" Then Errmsg = "An Unspecified Error occured.Client Holdings were not imported."
        SendMail Errmsg
        Unload Me
        Exit Sub
    End If
End Sub


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

Public Function importCds(fileName As String, dbfile As String, tableName As String)
    On Error GoTo errhandler
    Set conn = New ADODB.Connection
    Dim fso As New FileSystemObject
    conn.ConnectionString = "FILE NAME=" & dbfile
     conn.Open
    Set rsImport = New ADODB.Recordset
    'delete the previous imports
    conn.Execute ("Delete from " & tableName)
      i = 0
    rsImport.Open "select * from " & tableName, conn, 3, 2
        Open fileName For Input As 1
        
        'conn.BeginTrans
        Do While Not EOF(1)
        Line Input #1, TextLine
        TextLine = Trim(TextLine)
        If TextLine <> "" And i = 0 Then
            Errmsg = fileName & " is not a valid holdings file"
            Close #1
            fso.DeleteFile (fileName)
            Exit Function
        End If
        If i = 1 Then
            tradeDate = Trim(Mid$(TextLine, 24, 11))
            tradeTime = Trim(Mid$(TextLine, 35, 13))
       
        'check if the file is valid
        If IsDate(tradeDate) = False Then
            
            Errmsg = fileName & " is not a valid holdings file."
            Close #1
            fso.DeleteFile (fileName)
            Exit Function
        End If
        End If
      
        If Trim(TextLine) <> "" And i > 8 Then
         rsImport.AddNew
            clientRegistration = Trim(Mid$(TextLine, 1, 15))
            security = Trim(Mid$(TextLine, 21, 5))
            shareQty = Trim(Mid$(TextLine, 33, 22))
            accStatus = Trim(Mid$(TextLine, 55, 100))
            rsImport("TradeDate") = FormatDate(tradeDate)
            rsImport("TradeTime") = tradeTime
            rsImport("CDSNO") = clientRegistration
            rsImport("SecurityImportCode") = security
            rsImport("Quantity") = shareQty
            rsImport("AccountStatus") = accStatus
            If Trim(accStatus) = "BALANCE FREE" Then
                rsImport("BalanceFree") = 1
            Else
                rsImport("BalanceFree") = 0
            End If
            rsImport.Update
            
        End If
        i = i + 1
        Loop
        Close #1
        fso.DeleteFile (fileName)
      
        rsImport.Close
        Set rsImport = Nothing
        conn.Close
    
errhandler:
    If Trim(Err.Description) <> "" Then
        Errmsg = "Brokerknow did not successfully import client Holdings. Error Description: " & Err.Description
        ImportSuccess = False
        fso.DeleteFile (fileName)
        Exit Function
    Else
        Errmsg = "Brokerknow successfully imported but failed to commit Client Holdings. "
        ImportSuccess = True
        Exit Function
    End If
    
End Function
Function commitHoldings(dbfile As String)
On Error GoTo errhandler
    Set conn = New ADODB.Connection
    conn.ConnectionString = "FILE NAME=" & dbfile
    conn.Open
    Set rsCommit = New ADODB.Recordset
    Set rst = New ADODB.Recordset
    'delete the previous holdings
    conn.Execute ("Delete From Holdings")
    rst.Open "select * from CDSMatchedHoldings where Imported <> 1", conn.ConnectionString, 3, 2
    If rst.EOF Or rst.BOF Then
        Errmsg = "There are currently no holdings imported"
    End If
    rsCommit.Open "select * from Holdings", conn, 3, 2
    ' Add/ commit the records to the holdings table
   
        rst.MoveFirst
        Dim clientDPA As Variant
        Dim securityDPA As Variant
        
        Set rs = New ADODB.Recordset
        Dim SQLstr As String
        While Not rst.EOF
        SQLstr = "select Client_DPA_ from client where ClientCDSNo= '" & rst("CDSNo") & "'"
            rs.Open SQLstr, conn, 0, 1
                If Not rst.EOF Or Not rst.BOF Then
                    rs.MoveFirst
                    clientDPA = rs("Client_DPA_").Value
                End If
            rs.Close
            rs.Open "Select Security_DPA_ from security where SecurityCode='" & rst("securityCode") & "'", conn, 0, 1
                If Not rst.EOF Or Not rst.BOF Then
                    rs.MoveFirst
                    securityDPA = rs("Security_DPA_").Value
                End If
            rs.Close
            rsCommit.AddNew
            rsCommit("TradeDate") = FormatDate(rst("TradeDate"))
            rsCommit("TradeTime") = rst("TradeTime")
            rsCommit("Client_DPA_") = clientDPA
            rsCommit("Security_DPA_") = securityDPA
            rsCommit("Quantity") = rst("Quantity")
            rsCommit("AccountStatus") = rst("AccountStatus")
            If rst("BalanceFree") = 1 Then
                rsCommit("BalanceFree") = "Y"
            Else
                rsCommit("BalanceFree") = "N"
            End If
            rst("Imported") = 1
            rst.Update
            rsCommit.Update
        rst.MoveNext
        Wend
        
       rsCommit.Close
        rst.Close
        Set rst = Nothing
        Set rsCommit = Nothing
        Set conn = Nothing
  '  conn.CommitTrans
  
errhandler:
 If Trim(Err.Description) <> "" Then
        Errmsg = "Brokerknow did not successfully commit Client Holdings. Error Description: " & Err.Description
        Exit Function
    Else
        Errmsg = "Brokerknow successfully imported and committed Client Holdings."
        SendMail Errmsg
        
        Unload Me
    End If
End Function

Function GetUDLPath(theDBName)
   Dim tmpStr As String
    
    tmpStr = StrReverse(App.Path)
    tmpStr = App.Path
    
    tmpStr = Mid(tmpStr, InStr(1, tmpStr, "\") + 1)
    
    GetUDLPath = tmpStr & "\UDL\" & Trim(theDBName) & ".UDL"
End Function

Private Sub Form_Load()
    Command1_Click
End Sub

Public Function SendMail(Message)
 On Error GoTo errhandler
'Send mail to those persons specified in the brokerknow system
    
    Dim Sender
    Dim Subject
    Dim AttachmentFilePath, ReceiverCC
    Dim SQLstr, SMTPServer
    Dim fso As New FileSystemObject
    Dim RptMail As jmail.Message
    Dim rs As New ADODB.Recordset
    Dim rsMail As New ADODB.Recordset
    Dim SenderName
    Dim Password
    Dim Username
    Dim Displayname, test, i
    
    Set RptMail = New jmail.Message
    Set conn = New ADODB.Connection
    
    ' Append date and time of operation to the message
    
    Message = Message & vbCrLf & "Task Performed at: " & FormatDate(Now()) & "  " & TimeValue(Now())
    
    import_UDLPath = App.Path & "\UDL\KBroker.UDL"
    conn.ConnectionString = "FILE NAME=" & import_UDLPath
    conn.Open
    
    rsMail.CursorLocation = adUseClient
    rsMail.Open " Select * from MailConfiguration ", conn, 0, 1
    
    If Not (rsMail.EOF Or rsMail.BOF) Then
     'Dimension variables
     SMTPServer = rsMail("SMTPServer")
     Sender = rsMail("SendDisplayName")
     SenderName = "BrokerKnow"
     Subject = "Client Holdings Import Report"
     AttachmentFilePath = ""
     Password = rsMail("SendPassword")
     Username = rsMail("SendUserName")
     Displayname = rsMail("SendDisplayName")
     
        ' Fetch email addresses set for holdings operations
        SQLstr = " Select * from Systemnotification where Entity_DPA_ = 1 And Not(Systemnotification.Description IS NULL) "
        rs.CursorLocation = adUseClient
        rs.Open SQLstr, conn, adOpenDynamic, adLockReadOnly
        
        If Not (rs.EOF Or rs.BOF) Then
         
           i = 1 ' Addreessee counter
          
          'Build Message
          RptMail.From = Sender
          RptMail.FromName = Trim(SenderName)
          RptMail.Body = Trim(Message)
          RptMail.Subject = Trim(Subject)
          RptMail.Priority = 1 ' Send message immediately
                   
          If AttachmentFilePath <> "" Then
           RptMail.AddAttachment AttachmentFilePath
          End If
          
          'RptMail.MailServerPassWord = Trim(Password)
          'RptMail.MailServerUserName = Trim(Username)
        
           ' Add Addressees
           rs.MoveFirst
           
           Do Until rs.EOF
            
              test = Trim(rs.Fields("Description").Value)
                   
               If i = 1 Then
                   RptMail.AddRecipient Trim(rs.Fields("Description").Value)
               Else
                   RptMail.AddRecipientBCC Trim(rs.Fields("Description").Value)
               End If
                    
                i = i + 1
                rs.MoveNext
            Loop
             
            RptMail.Send SMTPServer
            
            rs.Close
            Set rs = Nothing
        End If
       
        rsMail.Close
        Set rsMail = Nothing
        Set RptMail = Nothing
    
    End If
    
    Exit Function
errhandler:
    Errmsg = Err.Number & " " & Err.Description
    Unload Me
    Exit Function
End Function

Function FormatDate(theDate)
  If Not IsDate(theDate) Then
   FormatDate = theDate
   Exit Function
  End If
 
 FormatDate = Day(theDate) & "-" & MonthName(Month(theDate), True) & "-" & Year(theDate)
End Function
