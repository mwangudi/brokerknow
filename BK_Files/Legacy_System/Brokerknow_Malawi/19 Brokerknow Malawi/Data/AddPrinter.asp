<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Add Printer</title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">

 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>


</head>

<body Class="Dialog">


<!--#include file="../libroutines.asp"-->
<%
	
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
   Dim guidStr 
   Dim guid 
	
	action = ucase(Request.Form("action"))
	
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

         sqlStr = "INSERT INTO Printers (PrinterActualName, Printer_DPA_, Description, Active, PrinterName) SELECT " & "'" & name & "'" & " as PrinterActualName," & " " & "iif(isnull(max([Printer_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Printers'),max([Printer_DPA_]) + 1)" & " " & " as Printer_DPA_" & _
                "," & "'" & Description & "'" & " as Description, " & " " & Active & " " & " as Active," & "'" & FriendlyName & "'" & " as PrinterName FROM Printers"
        Set conn = GetActiveConnection("KBroker")
        
        'conn.BeginTrans
			conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
		'conn.CommitTrans
		Set Conn = Nothing
		WritefraEnabledDialogCloseScript
        Response.End        
   	end If
%>


<form name = 'frmAddPrinter' method = 'post' action = 'AddPrinter.asp' target="deleteFrame" OnSubmit="JavaScript: UpdateDialogHandle();">
<table border="0" width="100%">
  <tr>
    <td width="30%">Actual Name</td>
    <td width="70%"><input type = 'text' name ='txtName' id = 'txtName' size="25"></td>
  </tr>
  
  <tr>
    <td width="30%" TITLE="This is the name that is displayed to the user during printing">Friendly Name</td>
    <td width="70%"><input type = 'text' name ='txtFriendlyName' id = 'txtFriendlyName' size="25"></td>
  </tr>
  <tr>
    <td width="30%">Active</td>
    <td width="70%"><input type = 'checkbox' Class="BorderLess" name ='txtActive' id = 'txtActive' checked value='1'></td>
  </tr>
  <tr>
    <td width="30%">Description</td>
    <td width="70%"><TEXTAREA name ='txtDescription' id = 'txtDescription' rows="3"></TEXTAREA></td>
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
     </td>
  </tr>
</table>
</form>
</body>

</html>
