VERSION 5.00
Begin VB.Form frmEmails 
   Caption         =   "Send Documents"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   5145
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBCC 
      Height          =   405
      Left            =   840
      TabIndex        =   6
      Top             =   1080
      Width           =   3975
   End
   Begin VB.TextBox txtCC 
      Height          =   405
      Left            =   840
      TabIndex        =   5
      Top             =   600
      Width           =   3975
   End
   Begin VB.TextBox txtTo 
      Height          =   405
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "BCC"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "CC"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "TO"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmEmails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim pathToFile, CurrentDirectory, strPath, fileName, ErrMsg As String
Dim import_UDLPath As String
Dim EmailTo, RecieverCC, RecieverBCC As String
Private Sub Command1_Click()

'validate for empty record
    If Trim(txtTo.Text) = "" Then
        MsgBox "Please enter ther Email To Person"
        txtTo.SetFocus
      Exit Sub
    End If
    If 1 <> 1 Then
        MsgBox "Please Enter a valid email address"
        txtTo.SetFocus
        Exit Sub
    End If
    
    Exit Sub
    'get the send to email address
    EmailTo = txtTo.Text
    RecieverCC = txtCC.Text
    RecieverBCC = txtBCC.Text
    CheckIfFileExists
 End Sub

Function CheckIfFileExists() As Boolean
    'Create he FSO objects
    strPath = App.Path
     Dim oFileSystem As New FileSystemObject
     Dim oFolder As Folder
     Dim oCurrentFile As File
     Dim oFileColl As Files
     Dim rsEmail As New ADODB.Recordset
     Dim sql, AttachmentPath As String
     On Error GoTo errhandler
     ' get the file location
     Set oFolder = oFileSystem.GetFolder(strPath)
     Set oFileColl = oFolder.Files
     
     'go through specified folder and pick the todays file
    Dim DateCreated As Date
    sql = "select * from EmailedDocs where 1=1"
    Set conn = New ADODB.Connection
    conn.ConnectionString = "FILE NAME=" & import_UDLPath
    conn.Open
    rsEmail.Open sql, conn, 3, 2
     If oFileColl.Count > 0 Then
     For Each oCurrentFile In oFileColl
        
        If (FormatDateTime(oCurrentFile.DateCreated, vbShortDate) = FormatDateTime(Now(), vbShortDate)) And (FormatDateTime(oCurrentFile.DateLastModified, vbShortDate) = FormatDateTime(Now(), vbShortDate)) And (Len(oCurrentFile.Name) > 15) Then
            If (GetFileExt(oCurrentFile.Name) = ".pdf") Then
                    fileName = oCurrentFile.Name
                    MsgBox fileName
                   DateCreated = oCurrentFile.DateCreated
                  'insert the files u want to email in the emaildocs table
                  ' and also email the documents
                    rsEmail.AddNew
                    rsEmail("DocumentName") = fileName
                    rsEmail("DateCreated") = DateCreated
                    rsEmail("EmailedTime") = Now()
                    rsEmail("Emailed") = 1
                    
             'get the attachement
                 AttachmentPath = App.Path & "\" & fileName
                 'send the mail here
                 
                SendMail "securities@africanalliance.co.ke", EmailTo, fileName, ErrMsg, AttachmentPath, RecieverCC, RecieverBCC
                AttachmentPath = ""
                rsEmail.Update
                     
        End If
      End If
     Next
     End If
     rsEmail.Close
     Set rsEmail = Nothing
    'if the file exists return true else false
     If Trim(fileName) <> "" Then
        CheckIfFileExists = True
       ' MsgBox fileName
     Else
        CheckIfFileExists = False
        ErrMsg = "No holding file was found for import "
        'MsgBox ErrMsg
     End If
     'Kill the FSO objects
     Set oFileSystem = Nothing
     Set oFolder = Nothing
     Set oFileColl = Nothing
     Set oCurrentFile = Nothing
    Exit Function
errhandler:
    ErrMsg = Err.Description
  '  MsgBox ErrMsg
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
Public Function SendMail(Sender, Receiver, Subject, Message, AttachmentFilePath, ReceiverCC, ReceiverBCC)
    On Error GoTo errhandler
    'On Error Resume Next
    'code to send mail
    
    
    Set conn = New ADODB.Connection
    Dim fso As New FileSystemObject
    conn.ConnectionString = "FILE NAME=" & import_UDLPath
    conn.Open
         
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
    rsMail.Open "Select SMTPServer from MailConfiguration", conn, 0, 1
    If Not rsMail.EOF Or rsMail.BOF Then
       RptMail.Send rsMail("SMTPServer")
       'RptMail.Send "smtp.wananchi.com"
       
    End If
    rsMail.Close
    conn.Close
    Set rsMail = Nothing
    'RptMail.Execute
    Exit Function
errhandler:
    ErrMsg = Err.Number & " " & Err.Description
    Unload Me
    Exit Function
End Function

Private Sub Form_Load()
    import_UDLPath = App.Path & "\UDL\KBroker.UDL"
End Sub
