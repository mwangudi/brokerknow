<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit Branch</title>
  <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="scripts/common.js"></SCRIPT> 
</head>

<body>
<!--#include file="../libroutines.asp"-->
<%
	
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
   Dim ID
   Dim rsEdit
	
	action = ucase(Request.Form("delAction"))
	ID = Request("ID")

		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for editing"
                		window.history.back
                </script>
                <% response.end
        End If

	if action = "EXECUTE" then
		Dim name
        
        name = Request.Form("txtName")
      
       
         'validate Name
        If Trim(Name) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Name"
                		window.history.back
                </script>
                <% response.end
        End If
        'validate size of Name
        If Len(name) > 100 Then%>
                <script language = 'vbscript'>
                ShowMessage "Description can only be 100 characters in length"
                window.history.back
                </script>
                <% response.end
        End If
        
        Set conn = GetActiveConnection("KBroker")
      

        'save data
        
        sqlStr = "UPDATE [Branch] SET BranchDescription = " & "'" & name & "'" & " WHERE Branch_DPA_  = " & ID
                
        
        conn.BeginTrans
                conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        conn.CommitTrans
        
        Set conn = Nothing
        WriteDialogCloseScript
        response.end
   	end If
%>


<form name = 'frmEditBranch' method = 'post' action = 'EditBranch.asp' >
<table border="0" width="100%" cellspacing="0" cellpadding="0">
<%
        Set conn = GetActiveConnection("KBroker")
        

        sqlStr = "SELECT BranchDescription,Branch_DPA_ FROM [Branch] WHERE Branch_DPA_  = " & ID                
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		ShowMessage "The selected Branch cannot be retrieved for editing"
                		window.history.back
                </script>
                <% response.end
        End If
     

%>
  <tr>
    <td width="17%">Name</td>
    <td width="83%"><input type = 'text' name ='txtName' id = 'txtName' size="35" value = '<%=rs.Fields("BranchDescription")%>'></td>
  </tr>
  <tr>
    <td width="17%"><input type = 'submit' name ='cmdAdd' id = 'cmdAdd' value="Edit"></td>
    <td width="83%">
    	<input type = 'hidden' name ='action' id = 'action' value="Execute">
    	<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
    </td>
  </tr>
</table>
</form>

</body>
</html>
