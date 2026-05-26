<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Delete Bond Issue</title>
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
 <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css"> 
 <script language="JavaScript" src="../scripts/common.js"></script>
</head>

<body Class="Dialog">
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
        sqlStr = "SELECT Child,DeletionMessage,ParentKey FROM [_Parent_Child_Links_] WHERE (Parent = 'Security') AND (ChildType = " & LinkedIndependent & ")"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.BOF Or rs.EOF) Then
                Dim childRS
                Dim tableName
                
                rs.MoveFirst
                Do Until rs.EOF
                			tableName = rs.Fields("Child")
                			ParentKey = rs.Fields("ParentKey")
                        sqlStr = "SELECT TOP 1 * FROM [" & tableName & "] WHERE " & ParentKey & " = " & ID
                        
                        Set childRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
                        If Not (childRS.BOF Or childRS.EOF) Then%>
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
        
		sqlStr = "SELECT Child,DeletionMessage,ParentKey FROM [_Parent_Child_Links_] WHERE (Parent = 'Security') AND (ChildType = " & LinkedDependent & ")"
		Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		If Not (rs.BOF Or rs.EOF) Then
		        rs.MoveFirst
		        childName = rs.Fields("Child")
		        ParentKey = rs.Fields("ParentKey")
		        sqlStr = "SELECT * FROM [" & childName & "] WHERE " & ParentKey & " = " & ID
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
		
       
        'delete from database
        sqlStr = "DELETE FROM [Security] WHERE Security_DPA_ = " & ID
        conn.BeginTrans
         conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        conn.CommitTrans
        Set Conn = Nothing
        WriteDeleteCloseScript
        Response.End
   	end If

	
        
%>


</body>

</html>
