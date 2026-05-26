VERSION 5.00
Begin VB.Form frmDebitSms 
   Caption         =   "DEBIT SMS"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "frmDebitSms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conn As ADODB.Connection
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
    GenerateDebit
    Unload Me
End Sub

Private Sub Form_Load()
'open the connection here
    import_UDLPath = App.Path & "\KBroker.UDL"
    Set Conn = New ADODB.Connection
    Conn.ConnectionString = "FILE NAME=" & import_UDLPath
    Conn.Open
    
    Command1_Click
End Sub

Function GenerateDebit()
    Dim sqlStr1, smsStr, HeaderText As String
    Dim Bal As Double
    Dim Rs1 As ADODB.Recordset
    Set Rs1 = New ADODB.Recordset
    Set Rs = New ADODB.Recordset
    Dim selectedDate As Date
            
     sqlStr = "select * from debita"
     Rs.Open sqlStr, Conn, 0, 1
         Rs.MoveFirst
            
        While Not Rs.EOF
        'create a string for the sms here
        ' format the celltell number
        cellNo = Replace(Rs("ClientCellTel"), " ", "")
        If Mid(cellNo, 1, 1) = 0 Then cellNo = Mid(cellNo, 2, 10) Else cellNo = Mid(cellNo, 1, 9)
        If Not IsNumeric(cellNo) Or Trim(cellNo) = "" Then
            cellNo = cellNo & " *"
        Else
            ' add the prefix to the cell no
            cellNo = "CSV: 254" & cellNo
        End If
        'get the the current balance to compare with
        
        Rs1.Open "select Balance from debitb where client_DPA_ =" & Rs("Client_DPA_"), Conn, 0, 1
        If Not Rs1.EOF Or Not Rs1.BOF Then
            If Rs1("Balance") < Rs("Balance") Then Bal = Rs("Balance") Else Bal = Rs1("Balance")
        Else
            Bal = Rs("Balance")
        End If
        Rs1.Close
        Bal = FormatNumber(0 - Bal, 2)
            smsStr = smsStr & cellNo & "|" & "Please note that you have an outstanding balance of KES " & FormatNumber(Bal, 2) & " overdue. Please settle." & Chr(13)
        Rs.MoveNext
        Wend
     
     
        HeaderText = "api-id: 1127740" & Chr(13)
        HeaderText = HeaderText & "user: african alliance" & Chr(13)
        HeaderText = HeaderText & "password: alliance1" & Chr(13)
        HeaderText = HeaderText & "text: #field1#" & Chr(13)
        HeaderText = HeaderText & "Delimiter: |" & Chr(13)
        
        Path = strPathDebtors & "\db_" & Replace(Date, "/", "") & ".txt"
        
        smsStr1 = HeaderText & "" & smsStr
        Debug.Print HeaderText & Chr(13) & smsStr1
        Dim fso As New FileSystemObject
        fso.CreateTextFile Path, True
        
        Open Path For Output As #1
        Print #1, smsStr1
        Close #1
    
        Set fso = Nothing
       
     
End Function

Function FormatDate(theDate)
    On Error Resume Next
    FormatDate = Day(theDate) & "-" & MonthName(Month(theDate), True) & "-" & Year(theDate)
    If Err.Number > 0 Then
        FormatDate = theDate
    End If
End Function
