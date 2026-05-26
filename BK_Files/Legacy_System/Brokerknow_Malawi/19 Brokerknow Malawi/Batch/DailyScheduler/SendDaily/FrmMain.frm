VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form FrmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SendDaily"
   ClientHeight    =   4350
   ClientLeft      =   10455
   ClientTop       =   540
   ClientWidth     =   4710
   FillColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   4710
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4200
      Top             =   0
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   4695
      ExtentX         =   8281
      ExtentY         =   6800
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label lblCounter 
      BackColor       =   &H00E0E0E0&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Conn As ADODB.Connection
Dim RS As ADODB.Recordset
Dim RS2 As ADODB.Recordset
Dim RS3 As ADODB.Recordset
Dim RS4 As ADODB.Recordset
Dim Client_DPA_, startDate, theDate2, urlPath, SqlStr, UpdateSqlStr, Rep

Private Sub Form_Load()
    On Error GoTo ErrApp
    
    GetSettingsDailyClientStatement
    
    UpdateSqlStr = ""
    If (Trim(UCase(reqdStatus)) = UCase("Pending")) And (Val(Format(Now, "yyyymmddhhmmss")) < Val(Format(reqdDate, "yyyymmddhhmmss"))) Then
        Unload Me
        Exit Sub
    Else
        Set Conn = ActiveConnection
        
        SqlStr = "SELECT *" & _
        " FROM (SELECT Client.Client_DPA_" & _
        " FROM Client INNER JOIN" & _
        " SendClientReports ON Client.Client_DPA_ = SendClientReports.Client_DPA_ INNER JOIN" & _
        " GenericSetting ON SendClientReports.GenericSetting_DPA_2 = GenericSetting.GenericSetting_DPA_" & _
        " WHERE (SendClientReports.Deleted = 0) AND (GenericSetting.GenericSetting_DPA_ = 2)) a LEFT OUTER JOIN" & _
        " SendDaily ON a.Client_DPA_ = SendDaily.Client_DPA_"
        Set RS = New ADODB.Recordset
        
        Set RS = Conn.Execute(SqlStr)
        
        If Not (RS.BOF Or RS.EOF) Then
            Do Until RS.EOF
                    If IsNull(RS("ClientEmail")) Then
                        SqlStr = ""
                        SqlStr = SqlStr & " if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SendDaily]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) "
                        SqlStr = SqlStr & " drop table [dbo].[SendDaily] "
                        Conn.Execute SqlStr
                        
                        SqlStr = ""
                        SqlStr = SqlStr & " SELECT     Client.Client_DPA_, GenericSetting.GenericSettingDescription, Client.ClientEmail, SendClientReports.ClientStatements,  "
                        SqlStr = SqlStr & "                       SendClientReports.ClientContracts, SendClientReports.ClientContractsCompounded, SendClientReports.HoldingsValuation, 0 AS Sent1, 0 AS Sent2, 0 AS Sent3, 0 AS Sent4 "
                        SqlStr = SqlStr & " INTO            SendDaily "
                        SqlStr = SqlStr & " FROM         Client INNER JOIN "
                        SqlStr = SqlStr & "                       SendClientReports ON Client.Client_DPA_ = SendClientReports.Client_DPA_ INNER JOIN "
                        SqlStr = SqlStr & "                       GenericSetting ON SendClientReports.GenericSetting_DPA_2 = GenericSetting.GenericSetting_DPA_ "
                        SqlStr = SqlStr & " WHERE     (SendClientReports.Deleted = 0) AND (GenericSetting.GenericSetting_DPA_ = 2) "
                        Conn.Execute SqlStr
                        
                        Exit Do
                    End If
                RS.MoveNext
            Loop
        End If
        
        ''CLIENT STATEMENTS
        SqlStr = "SELECT * FROM SendDaily WHERE Sent1 = 0 AND ClientStatements = 1"
        Set RS = New ADODB.Recordset
        
        RS.Open SqlStr, Conn, adOpenDynamic, adLockOptimistic
        
        If Not (RS.BOF Or RS.EOF) Then
            If (Trim(UCase(reqdStatus)) = UCase("Pending")) And (Val(Format(Now, "yyyymmddhhmmss")) > Val(Format(reqdDate, "yyyymmddhhmmss"))) Then
                Client_DPA_ = Trim(RS.Fields("Client_DPA_"))
                startDate = Format(DateAdd("d", -1, Now), "dd-mmm-yyyy")
                theDate2 = Format(DateAdd("d", 0, Now), "dd-mmm-yyyy")
                
                If RS("ClientStatements") = 1 And Len(RS("ClientEmail")) > 0 Then
                    urlPath = "http://10.64.10.27/Mailer/SendBatchClientStatement.asp?a=" & Client_DPA_ & "&b=" & startDate & "&c=" & theDate2
                    WebBrowser1.Navigate urlPath
                End If
                
                Rep = "ClientStatement"
                UpdateSqlStr = "UPDATE SendDaily SET Sent1 = 1 WHERE (Client_DPA_ = " & Client_DPA_ & ")"
                Timer1.Enabled = True
            End If
        Else
            ''CLIENT CONTRACTS
            SqlStr = "SELECT * FROM SendDaily WHERE Sent2 = 0 AND ClientContracts = 1"
            Set RS2 = New ADODB.Recordset
            
            RS2.Open SqlStr, Conn, adOpenDynamic, adLockOptimistic
            
            If Not (RS2.BOF Or RS2.EOF) Then
                If (Trim(UCase(reqdStatus)) = UCase("Pending")) And (Val(Format(Now, "yyyymmddhhmmss")) > Val(Format(reqdDate, "yyyymmddhhmmss"))) Then
                    Client_DPA_ = Trim(RS2.Fields("Client_DPA_"))
                    startDate = Format(DateAdd("d", -1, Now), "dd-mmm-yyyy")
                    theDate2 = Format(DateAdd("d", 0, Now), "dd-mmm-yyyy")
                    
                    If RS2("ClientContracts") = 1 And Len(RS2("ClientEmail")) > 0 Then
                        urlPath = "http://10.64.10.27/Mailer/SendBatchClientContract.asp?a=" & Client_DPA_ & "&b=" & startDate & "&c=" & theDate2
                        WebBrowser1.Navigate urlPath
                    End If
                    
                    Rep = "ClientContract"
                    UpdateSqlStr = "UPDATE SendDaily SET Sent2 = 1 WHERE (Client_DPA_ = " & Client_DPA_ & ")"
                    Timer1.Enabled = True
                End If
            Else
                ''CLIENT CONTRACTS COMPOUNDED
                SqlStr = "SELECT * FROM SendDaily WHERE Sent3 = 0 AND ClientContractsCompounded = 1"
                Set RS3 = New ADODB.Recordset
                
                RS3.Open SqlStr, Conn, adOpenDynamic, adLockOptimistic
                
                If Not (RS3.BOF Or RS3.EOF) Then
                    If (Trim(UCase(reqdStatus)) = UCase("Pending")) And (Val(Format(Now, "yyyymmddhhmmss")) > Val(Format(reqdDate, "yyyymmddhhmmss"))) Then
                        Client_DPA_ = Trim(RS3.Fields("Client_DPA_"))
                        startDate = Format(DateAdd("d", -1, Now), "dd-mmm-yyyy")
                        theDate2 = Format(DateAdd("d", 0, Now), "dd-mmm-yyyy")
                        
                        If RS3("ClientContractsCompounded") = 1 And Len(RS3("ClientEmail")) > 0 Then
                            urlPath = "http://10.64.10.27/Mailer/SendBatchClientContractCompounded.asp?a=" & Client_DPA_ & "&b=" & startDate & "&c=" & theDate2
                            WebBrowser1.Navigate urlPath
                        End If
                        
                        Rep = "ClientContractCompounded"
                        UpdateSqlStr = "UPDATE SendDaily SET Sent3 = 1 WHERE (Client_DPA_ = " & Client_DPA_ & ")"
                        Timer1.Enabled = True
                    End If
                Else
                    ''HOLDINGS VALUATION
                    SqlStr = "SELECT * FROM SendDaily WHERE Sent4 = 0 AND HoldingsValuation = 1"
                    Set RS4 = New ADODB.Recordset
                    
                    RS4.Open SqlStr, Conn, adOpenDynamic, adLockOptimistic
                    
                    If Not (RS4.BOF Or RS4.EOF) Then
                        If (Trim(UCase(reqdStatus)) = UCase("Pending")) And (Val(Format(Now, "yyyymmddhhmmss")) > Val(Format(reqdDate, "yyyymmddhhmmss"))) Then
                            Client_DPA_ = Trim(RS4.Fields("Client_DPA_"))
                            startDate = Format(DateAdd("d", -1, Now), "dd-mmm-yyyy")
                            theDate2 = Format(DateAdd("d", 0, Now), "dd-mmm-yyyy")
                            
                            If RS4("HoldingsValuation") = 1 And Len(RS4("ClientEmail")) > 0 Then
                                urlPath = "http://10.64.10.27/Mailer/SendBatchHoldingsValuation.asp?a=" & Client_DPA_ & "&b=" & startDate & "&c=" & theDate2
                                WebBrowser1.Navigate urlPath
                            End If
                            
                            Rep = "HoldingsValuation"
                            UpdateSqlStr = "UPDATE SendDaily SET Sent4 = 1 WHERE (Client_DPA_ = " & Client_DPA_ & ")"
                            Timer1.Enabled = True
                        End If
                    Else
                        ''REPORT SENDING COMPLETE
                        SqlStr = ""
                        SqlStr = SqlStr & " if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SendDaily]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) "
                        SqlStr = SqlStr & " drop table [dbo].[SendDaily] "
                        Conn.Execute SqlStr
                        
                        SqlStr = ""
                        SqlStr = SqlStr & " SELECT     Client.Client_DPA_, GenericSetting.GenericSettingDescription, Client.ClientEmail, SendClientReports.ClientStatements,  "
                        SqlStr = SqlStr & "                       SendClientReports.ClientContracts, SendClientReports.ClientContractsCompounded, SendClientReports.HoldingsValuation, 0 AS Sent1, 0 AS Sent2, 0 AS Sent3, 0 AS Sent4 "
                        SqlStr = SqlStr & " INTO            SendDaily "
                        SqlStr = SqlStr & " FROM         Client INNER JOIN "
                        SqlStr = SqlStr & "                       SendClientReports ON Client.Client_DPA_ = SendClientReports.Client_DPA_ INNER JOIN "
                        SqlStr = SqlStr & "                       GenericSetting ON SendClientReports.GenericSetting_DPA_2 = GenericSetting.GenericSetting_DPA_ "
                        SqlStr = SqlStr & " WHERE     (SendClientReports.Deleted = 0) AND (GenericSetting.GenericSetting_DPA_ = 2) "
                        Conn.Execute SqlStr
                        
                        Open App.Path & "\TimeLogDaily.txt" For Output As 1
                        Print #1, Format(DateAdd("d", 1, sendDay), "dd-mmm-yyyy") & " " & sendTime
                        Print #1, "Pending"
                        Close 1
                        
                        Unload Me
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If

    Exit Sub
    
ErrApp:
    
    LogEvent "Error: " & Err.Description
    Exit Sub
    Unload Me
End Sub

Private Sub Timer1_Timer()
    On Error GoTo ErrApp
    
    lblCounter.Caption = lblCounter.Caption + 1
    
    If lblCounter.Caption = 100 Then
        RS.Close
        
        SqlStr = UpdateSqlStr
        Conn.Execute SqlStr
        
        SqlStr = "INSERT INTO SendBatchItems (Frequency, Report, Client_DPA_, DateSent, TimeSent)" & _
        " VALUES('Daily','" & Rep & "','" & Client_DPA_ & "','" & Format(Date, "dd-mmm-yyyy") & "', '" & Format(Now, "dd-mmm-yyyy hh:mm:ss") & "')"
        Conn.Execute SqlStr
        
        Conn.Close
        
        Shell App.Path & "\SendDaily.exe", vbNormalFocus
        Unload Me
    End If
    
    Exit Sub
    
ErrApp:
    
    LogEvent "Error: " & Err.Description
    Exit Sub
    Unload Me
End Sub

Private Sub LogEvent(str As String)
    Open VB.App.Path & "\errLog.txt" For Output As 1
    Print #1, str
    Close #1
End Sub
