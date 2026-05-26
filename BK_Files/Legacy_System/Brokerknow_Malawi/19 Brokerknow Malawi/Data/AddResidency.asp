<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add Residency</title>
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
       
        'save data
        sqlStr = "INSERT INTO [Residency] (ResidencyDescription,Residency_DPA_) SELECT " & "'" & description & "'" & " as ResidencyDescription" & _
                "," & " " & "iif(isnull(max([Residency_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Residency'),max([Residency_DPA_]) + 1)" & " " & " as Residency_DPA_" & _
                " FROM [Residency]"
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
<BR>
<form name = 'frmAddResidency' method = 'post' action = 'AddResidency.asp' >
<table border="0" width="100%" cellspacing="2" cellpadding="2">
  <tr>
    <td width="18%"> Description</td>
    <td width="82%"><input type = 'text' name ='txtDescription' id = 'txtDescription' size="20"></td>
  </tr>
  <tr>
    <td width="100%" colspan="2" align=right>
		<BR>
	
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save ">
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();">
		&nbsp;&nbsp;
		<input type = 'hidden' name ='action' id = 'action' value="Execute">&nbsp;
      </td>
  </tr>
 </table>
</form>



</body>

</html>
