Attribute VB_Name = "ModMain"
Option Explicit

Public FSO, FSOTextStream, sendDay, sendTime, reqdStatus, reqdDate, theData

Public Function ActiveConnection() As ADODB.Connection
    On Error GoTo ErrApp
    
    Dim oconn As ADODB.Connection
    Set oconn = New ADODB.Connection
    
    oconn.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=BrokerKnowAALatest;Data Source=BROKERKNOW_SERV"
    oconn.CommandTimeout = 0
    oconn.ConnectionTimeout = 0
    oconn.Open
    
    Set ActiveConnection = oconn
    
    Exit Function
    
ErrApp:
    MsgBox "Error: " & Err.Description
    Exit Function
End Function

Public Sub GetSettingsMonthlyClientStatement()
    Open App.Path & "\TimeLogMonthly.txt" For Input As 2
    Line Input #2, reqdDate
    Line Input #2, reqdStatus
    
    sendDay = Format(reqdDate, "dd-mmm-yyyy")
    sendTime = Format(reqdDate, "hh:mm:ss")
    Close #2
End Sub




