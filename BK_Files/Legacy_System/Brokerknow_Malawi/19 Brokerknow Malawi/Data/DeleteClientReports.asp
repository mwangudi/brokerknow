<html>

<head>

<title>Delete Client Reports</title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css"> 
<script language="JavaScript" src="../scripts/common.js"></script>
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
 
</head>

<body Class="Dialog">
<!--#include file="../libroutines.asp"-->

<%
const LinkedIndependent = 1
const LinkedDependent = 2
UserId=Session("UserId")	
Dim conn 
Dim sqlStr
Dim rs
	
Set conn = GetActiveConnection("KBroker")
    
ID = Request("ID")
If Trim(ID) = "" Then
	%>
    <script language = 'vbscript'>
    		ShowMessage "No record specified for deletion."
    </script>
    <%
    response.end
End If
conn.BeginTrans
        sqlStr = "DELETE FROM [SendClientReports] WHERE ClientReportsID  = " & ID
            
    '    sqlStr = "UPDATE [SendClientReports]" & _
	'		" SET Deleted = 1" & _
	'		" WHERE ClientReportsID  = " & ID
				
        conn.Execute(sqlStr)
conn.CommitTrans
Set Conn = Nothing
response.redirect "../Reports/SendClientReports.asp"
Response.End
   	
response.End
%>

</body>
</html>
