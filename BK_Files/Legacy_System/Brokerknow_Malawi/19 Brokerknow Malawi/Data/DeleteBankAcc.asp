<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Delete Bank Account</title>
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
                		alert "No record specified for deletion"
                		window.self.close
                </script>
                <%response.end
        End If

        'find out whether any child records exist
        sqlStr = "SELECT Child,DeletionMessage,ParentKey FROM [_Parent_Child_Links_] WHERE (Parent = 'BankAcc') AND (ChildType = " & LinkedIndependent & ")"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.BOF Or rs.EOF) Then
                Dim childRS
                Dim tableName
                
                rs.MoveFirst
                Do Until rs.EOF
                			tableName = rs.Fields("Child")
                        sqlStr = "SELECT TOP 1 * FROM [" & tableName & "] WHERE BankAcc_DPA_ = " & ID
                        
                        Set childRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
                        If Not (childRS.BOF Or childRS.EOF) Then%>
                				<script language = 'vbscript'>
                					alert "<%=rs.Fields("DeletionMessage")%>"
                					window.self.close
                				</script>
                				<%response.end
                        End If
                        rs.MoveNext
                Loop
        End If
        
        'delete from database
        sqlStr = "DELETE FROM [BankAcc] WHERE BankAcc_DPA_ = " & ID
        conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        WriteDeleteCloseScript
        Response.End
        
   	end If

		 sqlStr = "SELECT * FROM [BankAccList]"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then
                %><p>No Bank Accounts found</p><%
                response.end
        End If
        
        rs.MoveFirst
        
%>



<form name = 'frmDeleteBankAccList' method = 'post' action = 'DeleteBankAccList.asp' >
<table border="0" width="100%">
<tr>
	<td width="34%"><b><font color="#ff3333">Client</font></b></td>
	<td width="34%"><b><font color="#ff3333">Account</font></b></td>
	

  </tr>
<%		Do Until rs.EOF%>
                <tr>
                        <td width="34%"><%=rs.Fields("BankAccClient")%>&nbsp</td>
                        <td width="34%"><b><font color="#000080"><a href = 'vbScript:ItemSelected(<%=rs.Fields("BankAcc_DPA_")%>, "<%=rs.Fields("BankAccAccount")%>")'><%=rs.Fields("BankAccAccount")%></a></font></b></td>
						
				  </tr>
                <%rs.MoveNext
        Loop
        conn.Close
        Set conn = Nothing%>
</table>
    		<input type = 'hidden' name ='ID' id = 'ID' >
    		<input type = 'hidden' name ='action' id = 'action' value="Execute">
</form>

</body>

</html>
