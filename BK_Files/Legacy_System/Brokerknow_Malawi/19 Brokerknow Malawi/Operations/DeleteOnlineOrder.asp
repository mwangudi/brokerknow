

<!--#include file="../libroutines.asp"-->
<%
	const LinkedIndependent = 1
   const LinkedDependent = 2
	
	Dim conn 
   Dim sqlStr
   Dim rs
	
	Set conn = GetActiveConnection("KBroker")
    
        
	action = ucase(Request.Form("delAction"))
	if action = "EXECUTE" then
		  
       ID = Request("ID")

		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for deletion"
                		
                </script>
                <%response.end
        End If
		
		
        'find out whether any child records exist
        sqlStr = "SELECT Child,DeletionMessage,ParentKey FROM [_Parent_Child_Links_] WHERE (Parent = 'WebtbOrder') AND (ChildType = " & LinkedIndependent & ")"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.BOF Or rs.EOF) Then
                Dim depRS
                Dim tableName
                
                rs.MoveFirst
                Do Until rs.EOF
                		tableName = rs.Fields("Child")
                        sqlStr = "SELECT TOP 1 * FROM [" & tableName & "] WHERE Order_DPA_ = " & ID
                        Set depRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
                        If Not (depRS.BOF Or depRS.EOF) Then%>
                				<script language = 'vbscript'>
                					ShowMessage "<%=rs.Fields("DeletionMessage")%>"
                					
                				</script>
                				<%response.end
                        End If
                        rs.MoveNext
                Loop
        End If
        
        
        
        conn.BeginTrans
                                
                'delete from database
				sqlStr = "DELETE FROM [WebtbOrder] WHERE Order_DPA_ = " & ID
                conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        conn.CommitTrans
        Set Conn = Nothing
        WriteDeleteCloseScript
        Response.End
        
   	end If

        
%>
