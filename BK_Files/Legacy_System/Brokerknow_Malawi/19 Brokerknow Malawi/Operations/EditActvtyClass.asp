<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit Activity Class</title>
 <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
<script language="javascript">
	function  UpdateClientAccess(theChk)
	{
		var holdVal = "0"; //no client access
		if (theChk.checked)
		{
			holdVal = "1";//client can access
		}
		
		document.getElementById("ClientAccess").value = holdVal;
				
		//document.frmMain.elements("CompoundStatus").value = holdVal;
	}
</script>
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
                		window.self.close
                </script>
                <% response.end
        End If

	if action = "EXECUTE" then
		Dim description
        Dim clientAccess
			 
		clientAccess = cint(Request.Form("ClientAccess"))
        description = Request.Form("txtDescription")
        cmdCancel = Request.Form("cmdCancel")
        Set conn = GetActiveConnection("KBroker")
        
        If cmdCancel <> "" Then
			
			WriteDialogCancelScript
			Set Conn = Nothing
			Response.End
        End If
      
       
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
        
        sqlStr = "UPDATE [ActvtyClass] SET ActvtyClassDescription = " & "'" & description & "'" & _
				", ClientAccess = " & " " & clientAccess & " " & _
				" WHERE ActvtyClass_DPA_  = " & ID                
        
        conn.BeginTrans
                conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        conn.CommitTrans
               
        conn.Close
        Set conn = Nothing
        WritefraEnabledDialogCloseScript
   	end If
   	
   
%>
<BR>

<form name = 'frmEditActvtyClass' id="frmMain" method = 'post' action = 'EditActvtyClass.asp' >
<table border="0" width="100%" cellspacing="0" cellpadding="0">
<%
        Set conn = GetActiveConnection("KBroker")
       
        
        sqlStr = "SELECT ActvtyClassDescription,ClientAccess,ActvtyClass_DPA_ FROM [ActvtyClass] WHERE ActvtyClass_DPA_  = " & ID
        
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		ShowMessage "The selected Activity Class cannot be retrieved for editing"
                		window.self.close
                </script>
                <% response.end
        End If
        
        

%>
  <tr>
    <td width="17%"> Description</td>
    <td width="83%">&nbsp;&nbsp;<input type = 'text' name ='txtDescription' id = 'txtDescription' size="20" value = '<%=rs.Fields("ActvtyClassDescription")%>'></td>
  </tr>
  <tr>
    <td>Client can view</td>
    <td>
     <%if cbool(Rs.Fields("ClientAccess")) then%>
		<input type=checkbox  checked value='True' name='chkClientAccess' id='chkClientAccess' onClick = 'UpdateClientAccess(this);'> 
	<%else%>
		<input type=checkbox   value='False' name='chkClientAccess' id='chkClientAccess' onClick = 'UpdateClientAccess(this);'> 
	<%end if%>
    
      </td>
  </tr>
  <tr>
     <td width="100%" COLSPAN=2 align="right" valign=absBottom>
		<BR>
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value="Save">
		&nbsp;
		<input type = 'button' Class=Buttons name ='cmdClose' id = 'cmdClose' value=" Cancel " OnClick="JavaScript: window.self.close();">		
		&nbsp;
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
		<input type = 'hidden' name ='ClientAccess' id = 'ClientAccess' value='<%= cint(Rs.Fields("ClientAccess").Value) %>'>
    </td>
  </tr>
</table>
</form>

</body>

</html>
