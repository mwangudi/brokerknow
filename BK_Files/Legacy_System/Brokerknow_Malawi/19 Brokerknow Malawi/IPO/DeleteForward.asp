<!--#include file="../libroutines.asp"-->
<%
	Dim conn 
	Dim sqlStr
	Dim rs
	
	Set conn = GetActiveConnection("KBroker")
    
	action = ucase(Request("delAction"))

	if action = "EXECUTE" then
		ID = Request("ID")
		
		If Trim(ID) = "" Then
			%>
			<script language = 'vbscript'>
				ShowMessage "No record specified for deletion"
				window.self.close
			</script>
			<%
			response.end
		End If
				
		sqlstr = "SELECT  Offering_DPA_, Downloaded " & _
		" FROM   dbo.Offerings " & _
		" WHERE  (NOT (Batch_No IS NULL)) AND (Offering_DPA_ = " & ID & ")"
		set rs = conn.execute(sqlstr)

   If (THISPIECEOFCODEISCOMMENTEDOUT) = TRUE THEN
		if not (rs.bof or rs.eof) then
			%>
			<script language="javascript">
				alert('The record cannot be deleted! It has already been batched.');
			</script>
			<%
			set conn = nothing
			Response.End
			end if	
    End IF
			
		sqlStr = "SELECT * FROM Offerings WHERE Offering_DPA_ = " & ID
		
		Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		If (rs.BOF Or rs.EOF) Then
			%>
		    <script language = 'vbscript'>
		    	ShowMessage "Serious error. The payment cannot be found for deletion"
		    	window.self.close
		    </script>
		    <%
		    response.end
		End If
				
		conn.BeginTrans
			sqlStr = "UPDATE Offerings SET Deleted = 1,status = 2 WHERE Offering_DPA_ = " & ID       
			sqlStr = SQLServerFormat(HandleQuote(sqlStr))	
		
			Conn.execute(sqlStr)	
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
