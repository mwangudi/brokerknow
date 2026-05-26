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
    Set Rs = Server.CreateObject("ADODB.Recordset")
	Rs.CursorLocation = adUseClient 
		
	action = ucase(Request.Form("delAction"))
	ID=Request("ID")
	
	if action = "EXECUTE" then		  
       ID = Request.Form("ID")		
		
		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for deletion"
                		
                </script>
                <%response.end
        End If
               
        sqlStr = "SELECT * FROM MarketSector WHERE Sector_DPA_ = " & ID
               
        Set rs = Conn.Execute(sqlStr)
        
        If rs.EOF Or rs.BOF Then
                %><script language = 'vbscript'>
                		ShowMessage "No such Sector"
                		
                </script><%
                response.end
        End If        
                
        set rs=nothing
        
        'find out whether any child records exist
        sqlStr = "SELECT Child,DeletionMessage,ParentKey FROM [_Parent_Child_Links_] WHERE (Parent = 'MarketSector') AND (ChildType = " & LinkedIndependent & ")"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.BOF Or rs.EOF) Then
                Dim childRS
                Dim tableName
                
                rs.MoveFirst
                Do Until rs.EOF
                			tableName = rs.Fields("Child")
                        sqlStr = "SELECT TOP 1 * FROM [" & tableName & "] WHERE Sector_DPA_ = " & ID
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
        sqlStr = "DELETE FROM [MarketSector] WHERE Sector_DPA_ = " & ID
        conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        Set Conn = Nothing
        WriteDeleteCloseScript
        Response.End
   	end If

		 sqlStr = "SELECT * FROM [MarketSector]WHERE Sector_DPA_  = " & ID
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then
                %><p>No Sector found</p><%
                response.end
        End If        
        
%>

<form name = 'frmDeletePaymentTypes' method = 'post' action = 'DeletePaymentTypes.asp' >
<table border="1" width="100%">
<tr>
	<td width="34%"><b><font color="#ff3333">Short Description</font></b></td>	
  </tr>
<%		Do Until rs.EOF%>
                <tr>
                        <td width="34%"><b><font color="#000080"><a href = 'vbScript:ItemSelected(<%=rs.Fields("PaymentTypes_DPA_")%>, "<%=rs.Fields("Description")%>")'><%=rs.Fields("Description")%></a></font></b></td>
                        <td width="34%"><%=rs.Fields("shortDescription")%>&nbsp</td>                        
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
