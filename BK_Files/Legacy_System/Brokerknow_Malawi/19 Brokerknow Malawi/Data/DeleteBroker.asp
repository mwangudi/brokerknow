<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Delete Broker</title>
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
        sqlStr = "SELECT Child,DeletionMessage,ParentKey FROM [_Parent_Child_Links_] WHERE (Parent = 'Broker') AND (ChildType = " & LinkedIndependent & ")"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.BOF Or rs.EOF) Then
                Dim childRS
                Dim tableName
                
                rs.MoveFirst
                Do Until rs.EOF
                			tableName = rs.Fields("Child")
                        sqlStr = "SELECT TOP 1 * FROM [" & tableName & "] WHERE Broker_DPA_ = " & ID
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
        
        'delete from database
        sqlStr = "DELETE FROM [Broker] WHERE Broker_DPA_ = " & ID
        conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        Set Conn = Nothing
        WriteDeleteCloseScript
        Response.End
   	end If

		 sqlStr = "SELECT * FROM [BrokerList]"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then
                %><p>No Brokers found</p><%
                response.end
        End If
        
        rs.MoveFirst
        
%>



<form name = 'frmDeleteBrokerList' method = 'post' action = 'DeleteBrokerList.asp' >
<table border="1" width="100%">
<tr>
	<td width="34%"><b><font color="#ff3333">Name</font></b></td>
	<td width="34%"><b><font color="#ff3333">Office Phone</font></b></td>
	<td width="34%"><b><font color="#ff3333">Fax</font></b></td>
	<td width="34%"><b><font color="#ff3333">Address</font></b></td>
  </tr>
<%		Do Until rs.EOF%>
                <tr>
                        <td width="34%"><b><font color="#000080"><a href = 'vbScript:ItemSelected(<%=rs.Fields("Broker_DPA_")%>, "<%=rs.Fields("BrokerName")%>")'><%=rs.Fields("BrokerName")%></a></font></b></td>
                        <td width="34%"><%=rs.Fields("BrokerOfficeTel")%>&nbsp</td>
                        <td width="34%"><%=rs.Fields("BrokerFax")%>&nbsp</td>
                        <td width="34%"><%=rs.Fields("BrokerAddr")%>&nbsp</td>
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
