<html>
<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Delete Order Hold Options</title>
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
                <script language = 'JavaScript'>
                		ShowMessage("No record specified for deletion")
                		window.self.close();
                </script>
                <%response.end
        End If

		sqlStr = "SELECT OrderHoldOptions.* FROM OrderHoldOptions WHERE (OrderHoldOptionID = " & id & ") AND (UPPER(Description) = 'AWAITING MANUAL RELEASE')"
		Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		If Not (rs.BOF Or rs.EOF) Then
			%>
		    <script language = 'vbscript'>
		    	ShowMessage "Awaiting Manual Release cannot be deleted"
		    </script>
		    <%
		    Response.End
		End If
        
        'delete from database
        sqlStr = "DELETE FROM [OrderHoldOptions] WHERE OrderHoldOptionID = " & ID
        conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        WriteDeleteCloseScript
        response.end
   	end If
   	%>

</body>

</html>
