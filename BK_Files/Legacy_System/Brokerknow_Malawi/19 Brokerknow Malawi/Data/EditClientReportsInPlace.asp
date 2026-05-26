<!--#include file="../libroutines.asp"-->
<%
Dim conn 
Dim sqlStr
Dim rs
Dim ID
	
ID = Request.Form("ID")

If Trim(ID) = "" Then
	%>
	<script language = 'vbscript'>
	ShowMessage "No record specified for editing"
	                		
	</script>
	<% response.end
End If
		
conn.BeginTrans
'sqlStr = "DELETE FROM [SendClientReports] WHERE ClientReportsID  = " & ID
            
sqlStr = "UPDATE [SendClientReports]" & _
" SET Deleted = 1" & _
" WHERE ClientReportsID  = " & ID
				
conn.Execute(sqlStr)
conn.CommitTrans
conn.Close
Set conn = Nothing

response.redirect "Reports/SendClientReports.asp"
%>
















