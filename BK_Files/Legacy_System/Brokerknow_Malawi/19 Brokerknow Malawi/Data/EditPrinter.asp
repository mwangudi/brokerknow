<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Edit Printer</title>

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
		Dim name
		Dim FriendlyName
		Dim Active
		Dim Description
		        
       name = Request.Form("txtName")
       FriendlyName = Request.Form("txtFriendlyName")
       Active = Request.Form("txtActive")
       Description = Request.Form("txtDescription")
       
      
        'validate Name
        If Trim(name) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Printer name"
                </script>
                <% response.end
        End If
        
        'validate Friendly Name
        If Trim(FriendlyName) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Printer friendly name"
                </script>
                <% response.end
        End If
              
        If Active <> "1" Then
			Active = "0"
		End If	        
  
        
        Set conn = GetActiveConnection("KBroker")
        
        'save data
       sqlStr = "UPDATE [Printers] SET PrinterActualName = '"  & name &  "', PrinterName = '" & FriendlyName  & "', Description = '" & Description  & "', " & _
				" Active = " & Active & " WHERE Printer_DPA_  = " & ID
        conn.BeginTrans
                conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        conn.CommitTrans
        
        conn.Close
        Set conn = Nothing
        WritefraEnabledDialogCloseScript
        Response.End
   	end If
   	
   	 Set conn = GetActiveConnection("KBroker")
   	
   	 sqlStr = "SELECT * FROM [PrinterList] WHERE Printer_DPA_ = " & ID
     Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
     
     If rsEdit.EOF Or rsEdit.BOF Then    
     %>	
		<Script Language="JavaScript">
			ShowMessage("The Printer cannot be retrieved for editing.")
		</Script>
      
      <%Set rsEdit = Nothing
		Set Conn = Nothing				
		WriteDialogRefuseOpenScript
		Response.End
     End If   
%>

<form name = 'frmEditPrinter' method = 'post' action = 'EditPrinter.asp' >
<table border="0" width="100%">
  <tr>
    <td width="30%">Actual Name</td>
    <td width="70%"><input type = 'text' name ='txtName' id = 'txtName' size="25" Value="<%= rsEdit.Fields("PrinterActualName").Value %>"></td>
  </tr>
  
  <tr>
    <td width="30%" TITLE="This is the name that is displayed to the user during printing">Friendly Name</td>
    <td width="70%"><input type = 'text' name ='txtFriendlyName' id = 'txtFriendlyName' size="25" Value="<%= rsEdit.Fields("PrinterName").Value %>"></td>
  </tr>
  <tr>
	<%If CBool(rsEdit.Fields("Active").Value) Then
		myVal = "checked"
	  Else
		myVal = ""
	  End If %>
    <td width="30%">Active</td>
		
    <td width="70%"><input type = 'checkbox' Class="BorderLess" name ='txtActive' id = 'txtActive' <%= myVal %> value='1'></td>
  </tr>
  <tr>
    <td width="30%">Description</td>
    <td width="70%"><TEXTAREA name ='txtDescription' id = 'txtDescription' rows="3"><%= rsEdit.Fields("Description").Value %></TEXTAREA></td>
  </tr>
  

</table>


<table border=0 width="100%">  
  <tr>
     <td align=right>
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














