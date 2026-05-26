<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Delete Security</title>


<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>

 
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

		 sqlStr = "SELECT * FROM [SecurityList]"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then
                %><p>No Securities found</p><%
                response.end
        End If
        
        rs.MoveFirst
        
%>
<form name = 'frmDeleteSecurityList' method = 'post' action = 'DeleteSecurityList.asp' >
<table border="0" width="100%">
<tr>
	<td width="34%"><b><font color="#ff3333">Name</font></b></td>
	<td width="34%"><b><font color="#ff3333">Code</font></b></td>
	<td width="34%"><b><font color="#ff3333">Market Price</font></b></td>
	<td width="34%"><b><font color="#ff3333">Transfer Fee</font></b></td>
	<td width="34%"><b><font color="#ff3333">Address</font></b></td>
	<td width="34%"><b><font color="#ff3333">Type</font></b></td>
  </tr>
<%		Do Until rs.EOF%>
                <tr>
                        <td width="34%"><b><font color="#000080"><a href = 'vbScript:ItemSelected(<%=rs.Fields("Security_DPA_")%>, "<%=rs.Fields("SecurityName")%>")'><%=rs.Fields("SecurityName")%></a></font></b></td>
                        <td width="34%"><%=rs.Fields("SecurityCode")%></td>
                        <td width="34%"><%=FormatCurrency(rs.Fields("SecurityMktPrice"),2,,,true)%></td>
                        <td width="34%"><%=FormatCurrency(rs.Fields("SecTransFeeFee"),2,,,true)%></td>
                        <td width="34%"><%=rs.Fields("SecurityAddr")%></td>
                        <td width="34%"><%=rs.Fields("OrderSecTypeDescription")%></td>
				  </tr>
                <%rs.MoveNext
        Loop
        conn.Close
        Set conn = Nothing%>
</table>
    		<input type = 'hidden' name ='ID' id = 'ID' >
    		<input type = 'hidden' name ='action' id = 'action' value="Execute">
</form>




</td>
</tr>
</table>
</div>
</div>

</body>

</html>
