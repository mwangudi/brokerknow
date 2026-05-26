<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add Activity Class</title>

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
		Dim clientAccess
			 
		clientAccess = cint(Request.Form("ClientAccess"))
       description = Request.Form("txtDescription")
      
       
        'validate Description
        If Trim(Description) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Description"
                		window.self.close
                </script>
                <% response.end
        End If
        'validate size of Description
        If Len(Description) > 100 Then%>
                <script language = 'vbscript'>
                ShowMessage "Description can only be 100 characters in length"
                window.self.close
                </script>
                <% response.end
        End If
       
        'save data
        sqlStr = "INSERT INTO [ActvtyClass] (ActvtyClassDescription,ClientAccess,ActvtyClass_DPA_) SELECT " & "'" & description & "'" & " as ActvtyClassDescription" & _
				"," & " " & clientAccess & " " & " as ClientAccess" & _
                "," & " " & "iif(isnull(max([ActvtyClass_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'ActvtyClass'),max([ActvtyClass_DPA_]) + 1)" & " " & " as ActvtyClass_DPA_" & _
                " FROM [ActvtyClass]"
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
<script language="javascript">
	function  UpdateClientAccess(theChk)
	{
		var holdVal = "0"; //no client access
		if (theChk.checked)
		{
			holdVal = "1";//client can access
		}
				
		document.getElementById("ClientAccess").value = holdVal;
				
		//document.frmMain.elements("ClientAccess").value = holdVal;
	}
</script>
<form name = 'frmAddActvtyClass' method = 'post' action = 'AddActvtyClass.asp' >
<BR>
<table border="0" width="100%" cellspacing="2" cellpadding="0">
  <tr>
    <td width="18%"> Description</td>
    <td width="82%">&nbsp;&nbsp;<input type = 'text' name ='txtDescription' id = 'txtDescription' size="20"></td>
  </tr>
  <tr>
    <td>Client can view</td>
    <td><input type=checkbox   value='False' name='chkClientAccess' onClick = 'UpdateClientAccess(this);'> 
      </td>
  </tr>
   <tr>
    <td width="100%" COLSPAN=2 align="right" valign=absBottom>
		<BR>
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value="Save">
		&nbsp;
		<input type = 'button' Class=Buttons name ='cmdClose' id = 'cmdClose' value=" Close " OnClick="JavaScript: window.close()">		
		&nbsp;
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		<input type = 'hidden' name ='ClientAccess' id = 'ClientAccess' value='0'>
    </td>
  </tr>
  </table>
</form>
</body>

</html>
