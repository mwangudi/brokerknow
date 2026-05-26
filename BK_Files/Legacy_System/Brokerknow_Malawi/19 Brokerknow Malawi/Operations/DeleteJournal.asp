<!--#include file="../libroutines.asp"-->


<%
	const UDLName = "KBroker"
	const DataSource = "EditJournal"
	const DataEntity = "Journal"
	const ChildDataEntity = "JournalEntry"
	const DataEntityList = "JournalList"
	const DataEntityPlural = "Journals"
	const ActionFolder = "Operations"
	
	const LinkedIndependent = 1
	const LinkedDependent = 2
	
	Dim conn 
	Dim sqlStr
	Dim rs
	
		Set conn = GetActiveConnection("KBroker")
	 
	    UserId=Session("UserID")
	     
		action = ucase(Request.Form("delAction"))
		if action = "EXECUTE" then
			  
	    ID = Request("ID")

			If Trim(ID) = "" Then%>
	             <script language = 'vbscript'>
	             		ShowMessage "No record specified for deletion"
	             		
	             </script>
	             <%response.end
	     End If
			
			'convert item ID to journal ID
			Dim orderRS
			sqlStr = "SELECT " & DataEntity & "_DPA_ FROM " & DataEntityList & " WHERE  " & ChildDataEntity & "_DPA_ =" & ID '& " and Deleted=0"
			'sqlStr="Select Journal_DPA_ from JournalList"			

			Set orderRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
			
			If (orderRS.EOF Or orderRS.BOF) Then%>
					    <script language = 'vbscript'>
					         	ShowMessage "The " & <%=DataEntity%> & " cannot be retrieved for deletion"
					         	
					    </script>
					    <% response.end
			End If
			ID = orderRS.Fields("" & DataEntity & "_DPA_")
	     'find out whether any child records exist
	     sqlStr = "SELECT Child,DeletionMessage,ParentKey FROM [_Parent_Child_Links_] WHERE (Parent = '" & DataEntity & "') AND (ChildType = " & LinkedIndependent & ")"
	     Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
	     If Not (rs.BOF Or rs.EOF) Then
	             Dim depRS
	             Dim tableName
	             
	             rs.MoveFirst
	             Do Until rs.EOF
	             		tableName = rs.Fields("Child")
	                     sqlStr = "SELECT TOP 1 * FROM [" & tableName & "] WHERE " & DataEntity & "_DPA_ = " & ID & "and Deleted=0"
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
	     
	     'check whether child records are linked to other data
	     Dim childRS
	     Dim childName
	     
	     sqlStr = "SELECT Child,DeletionMessage,ParentKey FROM [_Parent_Child_Links_] WHERE (Parent = '" & DataEntity & "') AND (ChildType = " & LinkedDependent & ")"
	     Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
	     If Not (rs.BOF Or rs.EOF) Then
	             rs.MoveFirst
	             childName = rs.Fields("Child")
	             sqlStr = "SELECT * FROM [" & childName & "] WHERE " & DataEntity & "_DPA_ = " & ID & "and Deleted=0"
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
	     
	     'delete child records
	     conn.BeginTrans
	             If Not (childRS.BOF Or childRS.EOF) Then
	                     
	                     sqlStr = "Update [" & childName & "] set Deleted=1 WHERE " & DataEntity & "_DPA_ = " & ID
	                     conn.Execute SQLServerFormat(HandleQuote(sqlStr))
	             End If
	             
	             'delete from database
				sqlStr = "Update [" & DataEntity & "] Set Deleted=1,ChangedBy=" & UserId & " WHERE " & DataEntity & "_DPA_ = " & ID

	            conn.Execute SQLServerFormat(HandleQuote(sqlStr))
	     conn.CommitTrans
	     Set Conn = Nothing
	     WriteDeleteCloseScript
	     Response.End
	     
		end If

%>

