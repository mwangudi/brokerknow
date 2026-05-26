<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit Gender</title>
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
   Dim ID
   Dim rsEdit
	
	action = ucase(Request.Form("action"))
	ID = Request("ID")

		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for editing"
                		
                </script>
                <% response.end
        End If

	if action = "EXECUTE" then
		Dim description
        
        description = Request.Form("txtDescription")
      
       
        'validate Description
        If Trim(Description) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Description"
                		
                </script>
                <% response.end
        End If
        'validate size of Description
        If Len(Description) > 100 Then%>
                <script language = 'vbscript'>
                ShowMessage "Description can only be 100 characters in length"
                
                </script>
                <% response.end
        End If

		Set conn = GetActiveConnection("KBroker")
       
        'save data
        
        sqlStr = "UPDATE [Gender] SET GenderDescription = " & "'" & description & "'" & " WHERE Gender_DPA_  = " & ID                
        
        conn.BeginTrans
                conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        conn.CommitTrans
        
        conn.Close
        Set conn = Nothing
        WritefraEnabledDialogCloseScript
   	end If
%>

<form name = 'frmEditGender' method = 'post' action = 'EditGender.asp' >
<table border="0" width="100%" cellspacing="1" cellpadding="1">
<%
        Set conn = GetActiveConnection("KBroker")
        
        
        sqlStr = "SELECT GenderDescription,Gender_DPA_ FROM [Gender] WHERE Gender_DPA_  = " & ID
        
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		ShowMessage "The selected Gender cannot be retrieved for editing"
                		
                </script>
                <% response.end
        End If
     

%>
  <tr>
    <td width="17%"> Description</td>
    <td width="83%"><input type = 'text' name ='txtDescription' id = 'txtDescription' size="20" value = '<%=rs.Fields("GenderDescription")%>'></td>
  </tr>
  <tr>
    <td width="100%" colspan=2 align=right>
		
		<BR>
	<BR>	
	<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save ">
	&nbsp;&nbsp;
	<input type = 'button' Class=Buttons name ='cmdClose' id = 'cmdClose' value=" Cancel " OnClick="JavaScript: window.self.close();">
    	<input type = 'hidden' name ='action' id = 'action' value="Execute">
    	<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
    </td>
  </tr>
</table>
</form>


</body>

</html>
