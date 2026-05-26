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
               
        Sqlstr = "Delete from [SystemNotification] Where SysConfig_DPA_ = " & ID
        Conn.execute sqlstr
        
        WritefraEnabledDialogCloseScript
        response.end        
        end if
        
%>

<form name = 'frm<%=DataSource%>' method = 'post' action = '<%=DataSource%>.asp' id = "frmMain">
<table border="0" width="100%" cellspacing="1" cellpadding="1">
  <tr>
    <td width="40%"> Title</td>
    <td width="60%">
    <select name ='cboTitle' id = "cboTitle">
    <% if rs.fields("Entity_DPA_") = 1 then %>
	<option value="1" selected>Email Address</option>
	<% end if%>
	</select></td>
  </tr>
  <tr>
    <td width="40%"> Description&nbsp; </td>
    <td width="60%">
	<input type = 'text' name ='txtDescription' id = 'txtDescription' value="<%=rs.Fields("Description")%>" size = "40">
	</td>
  </tr>
  
  <tr>
	  <td width="100%" colspan=3 align="center" valign=absBottom>
		<BR><BR>
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value="Save" onclick = "AllowedNavigation()">
    	<input type = 'button' Class=Buttons name ='cmdCancel' id = "cmdCancel" value="Cancel" onclick = "JavaScript: window.self.close()">
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
	</td>
  </tr>
</table>

</form>
</td>
</tr>
</table>
</div>
</div>
</body>

</html>
