<!--#include file="../libroutines.asp"-->
<%
Dim conn 
Dim sqlStr
Dim rs	
	
Set conn = GetActiveConnection("KBroker")    

action = ucase(Request("delAction"))
Dim reloadRequired
reloadRequired = false
if action = "EXECUTE" then
	ID = Request("ID")

	If Trim(ID) = "" Then		%>
		<script language = 'vbscript'>
			ShowMessage "No record specified for deletion"
			window.self.close
		</script>
		<%		response.end
	End If
		
	conn.BeginTrans
		
		BatchNo = "NULL"
		BatchSeq = "NULL"
            
		sqlStr = "UPDATE Offerings SET Downloaded=0, LastDownloaded=0 WHERE Batch_No = " & ID
        Conn.execute(SQLServerFormat(HandleQuote(sqlStr)))		

		sqlStr = "UPDATE Offerings SET Batch_no=" & BatchNo & ", BatchSeq=" & BatchSeq & " WHERE Batch_No = " & ID		
      	Conn.execute(SQLServerFormat(HandleQuote(sqlStr)))		
		conn.CommitTrans
		
	WriteDeleteCloseScript
	Response.End
end If
Dim clientCode
        
clientCode = "var validNavigate = true;" & chr(13)
clientCode = clientCode & "window.self.close();" & chr(13)%>
<script>
<%=clientCode%>
</script>
<%
response.End
%>
