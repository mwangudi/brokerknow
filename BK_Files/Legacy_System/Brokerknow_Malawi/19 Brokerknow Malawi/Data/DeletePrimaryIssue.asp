<!--#include file="../libroutines.asp"-->
<html>
<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit <%=DataEntity%></title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/accountList.js"></SCRIPT>
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<script language="JavaScript" src="CALENDAR/calendar.js"></script>
<!--END CALENDAR -->
</head>
<%
                
	const UDLName = "KBroker"
	
	Dim UserId
	Dim conn 
	Dim sqlStr
	Dim rst
	Dim ID
	
	set rst = server.CreateObject ("ADODB.Recordset")
	
	UserId=Session("UserID")
	ID = Request.Form("ID")
    
    action = ucase(Request.Form("delAction"))
	
	if action = "EXECUTE" then
			
			If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for deletion"
                		window.self.close
                </script>
                <%response.end
        End If
        
			'Delete data
					
		    Set conn = GetActiveConnection("KBroker")
				 conn.BeginTrans		
												      
				sqlStr = "Update [PrimaryIssues] set Deleted = 1 Where PrimaryIssues_DPA_=" & ID
									
				sqlStr = SQLServerFormat(HandleQuote(sqlStr))

				conn.Execute sqlStr
				conn.CommitTrans
					
			    WritefraEnabledDialogCloseScript
		
				conn.Close
			    Set conn = Nothing
					
			    Response.End	
   	End if
	   	
%>

</html>