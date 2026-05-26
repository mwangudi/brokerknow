<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Delete Levy</title>
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
 <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css"> 
 <script language="JavaScript" src="../scripts/common.js"></script>
 
</head>

<body Class="Dialog">

<!--#include file="../libroutines.asp"-->

<%
	const LinkedIndependent = 1
   const LinkedDependent = 2
   const DataEntity = "Levy"
	
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
        sqlStr = "SELECT Child,DeletionMessage,ParentKey FROM [_Parent_Child_Links_] WHERE (Parent = 'Levy') AND (ChildType = " & LinkedIndependent & ")"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.BOF Or rs.EOF) Then
                Dim childRS
                Dim tableName
                
                rs.MoveFirst
                Do Until rs.EOF
                			tableName = rs.Fields("Child")
                        sqlStr = "SELECT TOP 1 * FROM [" & tableName & "] WHERE Levy_DPA_ = " & ID
                        
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
       
       'remove securities under this levy
		sqlStr = "DELETE FROM LevySecurity WHERE Levy_DPA_ = " & ID
		
		conn.BeginTrans
				conn.Execute sqlStr
				
				Dim sysMaintainID
				
				sqlStr = "SELECT SystemMaintained FROM Levy WHERE Levy_DPA_ = " & ID
				Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
				If (rs.BOF Or rs.EOF) Then%>
                		<script language = 'vbscript'>
                			ShowMessage "Error encountered. Database may be corrupted."
                					
                		</script>
                		<%response.end
                End If
				
				sysMaintainID = rs.Fields("SystemMaintained")
				'determine whether the levy is in use
				sqlStr = "SELECT * FROM LevyContract WHERE SystemMaintained  = " & sysMaintainID 
				Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
				If (rs.EOF Or rs.BOF) Then
						'remove the matching account
						sqlStr = "DELETE FROM Entity WHERE LevySystemMaintained  = " & sysMaintainID 
						conn.Execute SQLServerFormat(HandleQuote(sqlStr))
						
						'delete levy
						sqlStr = "DELETE FROM Levy WHERE Levy_DPA_ = " & ID
						conn.Execute SQLServerFormat(HandleQuote(sqlStr))
				
				Else
						%>
							<script language = 'vbscript'>
								ShowMessage "The levy has been successfully deactivated. Complete deletion is not possible since there are still contracts using the levy"
												
							</script>
						<%
				End If
        conn.CommitTrans
		Set Conn = Nothing
		WriteDeleteCloseScript
        Response.End
   	end If
%>
</body>

</html>
