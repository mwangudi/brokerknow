<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Delete Bank</title>


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
        
        ID = GetBankID(ID)

        'find out whether any child records exist
        sqlStr = "SELECT Child,DeletionMessage,ParentKey FROM [_Parent_Child_Links_] WHERE (Parent = 'Bank') AND (ChildType = " & LinkedIndependent & ")"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.BOF Or rs.EOF) Then
                dim depRS
                Dim tableName
                
                rs.MoveFirst
                Do Until rs.EOF
                		tableName = rs.Fields("Child")
                        sqlStr = "SELECT TOP 1 * FROM [" & tableName & "] WHERE Bank_DPA_ = " & ID
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
        Dim childName
        Dim childRS
        
        sqlStr = "SELECT Child,DeletionMessage,ParentKey FROM [_Parent_Child_Links_] WHERE (Parent = 'Bank') AND (ChildType = " & LinkedDependent & ")"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.BOF Or rs.EOF) Then
                rs.MoveFirst
                childName = rs.Fields("Child")
                sqlStr = "SELECT * FROM [" & childName & "] WHERE Bank_DPA_ = " & ID
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
                        
                        sqlStr = "DELETE FROM [" & childName & "] WHERE Bank_DPA_ = " & ID
                        conn.Execute SQLServerFormat(HandleQuote(sqlStr))
                End If
               
                'delete from database
                sqlStr = "DELETE FROM [Bank] WHERE Bank_DPA_ = " & ID
                conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        conn.CommitTrans
        Set Conn = Nothing
        WriteDeleteCloseScript
        Response.End
   	End If
   	
   	Function GetBankID(branchID)
   		Dim getRs
   		Set getConn = GetActiveConnection("KBroker")
   	   	Set getRs = getConn.Execute("SELECT Bank_DPA_ FROM BankList WHERE BnkBranch_DPA_ = " & branchID)
   	   	If Not (getRs.EOF Or getRs.BOF) Then
   	   		GetBankID = getRs.Fields("Bank_DPA_").Value
   	   	Else
   	   		GetBankID = ""
   	   	End If	
   	End Function
%>

</body>

</html>
