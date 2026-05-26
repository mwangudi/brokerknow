<!--#include file="../libroutines.asp"-->
<%
'======================= Begin_Alter_Across_Entities =================================
		
		const UDLName = "KBroker"
		const DataSource = "AccountList"
		const DataEntity = "Account"
		const EntityName = "Account"
		const DataEntityPlural = "Accounts"
		const ActionFolder = "Operations"
		const Entity_DPA_ = "Account_DPA_"
'======================= End_Alter_Across_Entities =================================		


	const LinkedIndependent = 1
   const LinkedDependent = 2
	
	Dim conn 
   Dim sqlStr
   Dim rs
	
	Set conn = GetActiveConnection("KBroker")
    
        
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


		'system maintained accounts cannot be deleted
		sqlStr = "SELECT * FROM Account WHERE SystemMaintained = 0 AND Account_DPA_=" & ID
        
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		window.self.ShowMessage "The selected <%=DataEntity%> cannot be deleted"
                		
                </script>
                <% response.end
        End If
		
        'find out whether any child records exist
				'obtain the entity type
				
		sqlStr = "SELECT EntityType_DPA_ FROM Account WHERE Account_DPA_ = " & ID
		Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		If (rs.EOF Or rs.BOF) Then%>
				<script language = 'vbscript'>
				        ShowMessage "A serious error has been encountered while deleting the data. Try deleting again"
												        
				</script>
				<% response.end
		End If
		
		Dim entType
		
		entType = rs.Fields("EntityType_DPA_")
        sqlStr = "SELECT TOP 1 * FROM Payment WHERE EntityType_DPA_ = " & entType & " AND Entity_DPA_ = " & ID
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.BOF Or rs.EOF) Then%>
        		<script language = 'vbscript'>
        			ShowMessage "At least one payment exists under this account"
        			window.self.close
        		</script>
        		<%response.end
        End If
        
        sqlStr = "SELECT Child,DeletionMessage,ParentKey FROM [_Parent_Child_Links_] WHERE (Parent = '" & DataEntity & "') AND (ChildType = " & LinkedIndependent & ")"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.BOF Or rs.EOF) Then
                Dim childRS
                Dim tableName
                
                rs.MoveFirst
                Do Until rs.EOF
                			tableName = rs.Fields("Child")
                        sqlStr = "SELECT TOP 1 * FROM [" & tableName & "] WHERE " & Entity_DPA_ & " = " & ID
                        
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

        'delete from database
        sqlStr = "DELETE FROM [" & DataEntity & "] WHERE " & Entity_DPA_ & " = " & ID
        sqlStr = SQLServerFormat(HandleQuote(sqlStr))

        conn.Execute sqlStr
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
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Delete <%=EntityName%></title>
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
 <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css"> 
 <script language="JavaScript" src="../scripts/common.js"></script>
</head>

<body>
</body>

</html>
