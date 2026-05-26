<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Edit Path Configuration</title>

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
                <% WriteDialogRefuseOpenScript
                response.end
        End If

	if action = "EXECUTE" then
	
        
       Path = Request.Form("Path") 
       DocumentName= Request.Form("DocumentName")
             
       'validate
        If  Path ="" Then%>
                <script language = 'vbscript'>
                ShowMessage "Please specify a  Path"
                
                </script>
                <% response.end
        End If
   
        'validate 
        If instr(1,path,"\",1)=0 or instr(1,path,":",1)=0   Then%>
                <script language = 'vbscript'>
                ShowMessage "Please specify a valid Path"
                 
                </script>
                <% response.end
        End If
  
        
        Set conn = GetActiveConnection("KBroker")
        
        'save data
       sqlStr = "UPDATE [PathConfigurations] SET path = '"  & path &  "' WHERE DocumentID  = " & ID
        conn.BeginTrans
                conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        conn.CommitTrans
        
        conn.Close
        Set conn = Nothing
        WritefraEnabledDialogCloseScript
        Response.End
   	end If
   	
   	 Set conn = GetActiveConnection("KBroker")
   	
   	 sqlStr = "SELECT * FROM [PathConfigurations] WHERE DocumentID = " & ID
     Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
     
     If rsEdit.EOF Or rsEdit.BOF Then    
     %>	
		<Script Language="JavaScript">
			ShowMessage("The Path configuration cannot be retrieved for editing.")
		</Script>
      
      <%Set rsEdit = Nothing
		Set Conn = Nothing				
		WriteDialogRefuseOpenScript
		Response.End
     End If   
%>

<form name = 'frmEditTimeLimit' method = 'post' action = 'EditPathConfiguration.asp' >
<table border="0" width="100%">
	<tr>
		<td width="30%">Document Name</td>
		<td width="70%"><input type = 'text' class="Readonly" READONLY name ='txtAction' id = 'txtAction' size="20" value="<%= rsEdit.Fields("DocumentName").Value %>"></td>
	</tr>	
  <tr>
    <td width="35%">Path</td>
    <td width="70%"><input type = 'text' name ='path' id = 'path' size="20" value="<%= rsEdit.Fields("path").Value %>"></td>
  </tr>
  
 </table>


<table border=0 width="100%">  
  <tr>
     <td align=right>
		<BR>
		<BR>
		<BR>
		<BR>
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value="Save">
		&nbsp;&nbsp;
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();">		
		<input type = 'hidden' name ='action' id = 'action' value="Execute">&nbsp;
		<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
     </td>
  </tr>
</table>
</form>
</body>

</html>














