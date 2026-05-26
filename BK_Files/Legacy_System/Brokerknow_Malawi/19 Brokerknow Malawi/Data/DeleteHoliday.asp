<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Delete Holiday</title>


<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>

 
</head>

<body>
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
		  
       ID = Request.Form("ID")

		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for deletion"
                		
                </script>
                <%response.end
        End If
        
        'find out whether any child records exist
        sqlStr = "SELECT Child,DeletionMessage,ParentKey FROM [_Parent_Child_Links_] WHERE (Parent = 'Holidays') AND (ChildType = " & LinkedIndependent & ")"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.BOF Or rs.EOF) Then
                dim depRS
                Dim tableName
                
                rs.MoveFirst
                Do Until rs.EOF
                		tableName = rs.Fields("Child")
                        sqlStr = "SELECT TOP 1 * FROM [" & tableName & "] WHERE Holiday_DPA_ = " & ID
                        Set depRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
                        If Not (depRS.BOF Or depRS.EOF) Then%>
                				<script language = 'vbscript'>
                					ShowMessage <%=rs.Fields("DeletionMessage")%>
                					
                				</script>
                				<%response.end
                        End If
                        rs.MoveNext
                Loop
        End If
        
         'delete from database
        sqlStr = "DELETE FROM [Holidays] WHERE Holiday_DPA_ = " & ID
        conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        Set Conn = Nothing
        WriteDeleteCloseScript
        Response.End
  
   	End If
   	
 
%>

</body>

</html>
