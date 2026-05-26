<html>

<head>
<title>Delete Security</title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>

</head>

<body Class="Dialog">

<!--#include file="../libroutines.asp"-->
<%
	Dim conn 
	Dim sqlStr
	Dim rs
	
	Set conn = GetActiveConnection("KBroker")
    
	action = ucase(Request("delAction"))
	
	if action = "EXECUTE" then
			  
		ID = Request("ID")
		
		const LinkedIndependent = 1
		const LinkedDependent = 2
			
		If Trim(ID) = "" Then
			%>
		    <script language = 'vbscript'>
		    		ShowMessage "No record specified for deletion"
		    </script>
		    <%
		    response.end
		End If

        'find out whether any child records exist
        sqlStr = "SELECT Child,DeletionMessage,ParentKey FROM [_Parent_Child_Links_] WHERE (Parent = 'Security') AND (ChildType = " & LinkedIndependent & ")"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.BOF Or rs.EOF) Then
                dim depRS
                Dim tableName
                
                rs.MoveFirst
                Do Until rs.EOF
                		tableName = rs.Fields("Child")
                        sqlStr = "SELECT TOP 1 * FROM [" & tableName & "] WHERE Security_DPA_ = " & ID
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
        
        'check whether child records are linked to other data
        Dim childName
        Dim childRS
        
        sqlStr = "SELECT Child,DeletionMessage,ParentKey FROM [_Parent_Child_Links_] WHERE (Parent = 'Security') AND (ChildType = " & LinkedDependent & ")"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.BOF Or rs.EOF) Then
                rs.MoveFirst
                childName = rs.Fields("Child")
                sqlStr = "SELECT * FROM [" & childName & "] WHERE Security_DPA_ = " & ID
                Set childRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
                If Not (childRS.BOF Or childRS.EOF) Then
                        'find the key field
                        sqlStr = "SELECT Child,DeletionMessage,ParentKey FROM [_Parent_Child_Links_] WHERE (Parent = '" & childName & "') AND (ChildType = " & LinkedIndependent & ")"
                        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
                        If Not (rs.BOF Or rs.EOF) Then
                                Dim childKey
                                
                                childKey = rs.Fields("ParentKey")
                                Do Until childRS.EOF
                                        'find out whether any child records exist
                                        sqlStr = "SELECT Child,DeletionMessage,ParentKey FROM [_Parent_Child_Links_] WHERE (Parent = '" & childName & "') AND (ChildType = " & LinkedIndependent & ")"
                                        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
                                        If Not (rs.BOF Or rs.EOF) Then
                                                rs.MoveFirst
                                                
                                                Do Until rs.EOF
														tableName = rs.Fields("Child")
                                                        sqlStr = "SELECT TOP 1 * FROM [" & tableName & "] WHERE " & childKey & " = " & childRS.Fields(childKey)
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
                                        childRS.MoveNext
                                Loop
                                childRS.MoveFirst
                        End If
                End If
        End If
       
      
        'delete child records
        conn.BeginTrans
                
            If Not (childRS.BOF Or childRS.EOF) Then
                        
                    sqlStr = "DELETE FROM [" & childName & "] WHERE Security_DPA_ = " & ID
                    conn.Execute SQLServerFormat(HandleQuote(sqlStr))
            End If
                
            'delete from database
            sqlStr = "DELETE FROM [Security] WHERE Security_DPA_ = " & ID
            conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        
        conn.CommitTrans
        
        Set Conn = Nothing
        WriteDeleteCloseScript
        Response.End
   	end If
%>