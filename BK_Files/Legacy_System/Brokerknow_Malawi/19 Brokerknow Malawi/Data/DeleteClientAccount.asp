<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Delete Client</title>
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
 <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css"> 
 <script language="JavaScript" src="../scripts/common.js"></script>
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
    
    UserId=Session("UserID")
        
	action = ucase(Request("delAction"))
	Dim reloadRequired
		
	reloadRequired = false
	if action = "EXECUTE" then
		
		
       ID = Request("ID")

		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for deletion"
                		window.self.close
                </script>
                <%response.end
        End If
		
		Dim childRS
		Dim tableName
		'find out whether any child records exist
		sqlStr = "SELECT Child,DeletionMessage,ParentKey FROM [_Parent_Child_Links_] WHERE (Parent = 'Client') AND (ChildType = " & LinkedIndependent & ")"					
		
		Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		If Not (rs.BOF Or rs.EOF) Then
		        rs.MoveFirst
		        Do Until rs.EOF
		        		tableName = rs.Fields("Child")
		        		if(trim(tableName)="tbOrder") then
		        		sqlStr = "SELECT TOP 1 * FROM [" & tableName & "] WHERE Client_DPA_ = " & ID & " and Deleted=0"
		        		else
		                sqlStr = "SELECT TOP 1 * FROM [" & tableName & "] WHERE Client_DPA_ = " & ID
				        end if        
		                Set childRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		                If Not (childRS.BOF Or childRS.EOF) Then%>
		        				<script language = 'vbscript'>
		        					ShowMessage "<%=rs.Fields("DeletionMessage")%>"
		        					window.self.close
		        				</script>
		        				<%response.end
		                End If
		                rs.MoveNext
		        Loop
		End If        
		
		'check whether child records are linked to other data
		Dim childName
        
		sqlStr = "SELECT Child,DeletionMessage,ParentKey FROM [_Parent_Child_Links_] WHERE (Parent = 'Client') AND (ChildType = " & LinkedDependent & ")"
		
		Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		If Not (rs.BOF Or rs.EOF) Then
		        rs.MoveFirst
		        childName = rs.Fields("Child")
				
		        sqlStr = "SELECT * FROM [" & childName & "] WHERE Client_DPA_ = " & ID
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
		                                                sqlStr = "SELECT TOP 1 * FROM [" & tableName & "] WHERE " & childKey & " = " & childRS.Fields(childKey) & " and Deleted=0"
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
		
		conn.BeginTrans								
				
		
                'delete child records
				
                If Not (childRS.BOF Or childRS.EOF) then
						if(childName="Agent") then
                        sqlStr = "Update [" & childName & "] Set deleted=1 WHERE Client_DPA_ = " & ID
						else
						sqlStr = "Delete  from [" & childName & "]  WHERE Client_DPA_ = " & ID						
						end if
                        conn.Execute SQLServerFormat(HandleQuote(sqlStr))
                End If
                
				'delete from database
				sqlStr = "Update [Client] Set Deleted=1,ChangedBy=" & UserId & ",Timechanged=GetDate() WHERE Client_DPA_ = " & ID
				conn.Execute SQLServerFormat(HandleQuote(sqlStr))
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

</body>

</html>
