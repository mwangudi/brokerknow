<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add Branch</title>
 <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
</head>

<body Class="Dialog">

<!--#include file="../libroutines.asp"-->
<%
	
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
	
	action = ucase(Request.Form("action"))
	
	if action = "EXECUTE" then
		Dim name
       
       name = Request.Form("txtName")
      
       
        'validate Name
        If Trim(Name) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Name"
                		
                </script>
                <% response.end
        End If
        'validate size of Name
        If Len(name) > 100 Then%>
                <script language = 'vbscript'>
                ShowMessage "Description can only be 100 characters in length"
                
                </script>
                <% response.end
        End If
       
        'save data
         sqlStr = "INSERT INTO [Branch] (BranchDescription,Branch_DPA_) SELECT " & "'" & name & "'" & " as BranchDescription" & _
                "," & " " & "iif(isnull(max([Branch_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Branch'),max([Branch_DPA_]) + 1)" & " " & " as Branch_DPA_" & _
                " FROM [Branch]"
        Set conn = GetActiveConnection("KBroker")
        
        conn.BeginTrans
                conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
        conn.CommitTrans
        conn.Close
        Set conn = Nothing
        WritefraEnabledDialogCloseScript
        Response.End
   	end If
%>

<form name = 'frmAddBranch' method = 'post' action = 'AddBranch.asp' >
<p>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td width="18%"> Name</td>
    <td width="82%"><input type = 'text' name ='txtName' id = 'txtName' size="25"></td>
  </tr>
  <tr>
    <td colspan=2 align=right>
		<BR>
		<BR>
		<BR>
		<BR>
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save ">
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close()">
		<input type = 'hidden' name ='action' id = 'action' value="Execute">&nbsp;
     </td>
  </tr>
</table>
</form>
</body>
</HTML>

