<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit Broker</title>
 <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 

<script language='javascript'>
		function forceSubmit()
		{
			setOpener();
			document.frmEditPaymentTypes.method='post';
			document.frmEditPaymentTypes.target='_self';
			document.frmEditPaymentTypes.submit();		
		}
		
		function setOpener()
		{
			window.self.opener = window.dialogArguments.opener;					
		}		

</script>

</head>

<body Class="Dialog" onload="setOpener()">

<!--#include file="../libroutines.asp"-->
<%
	
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
   Dim ID
   Dim rsEdit
	
	action = ucase(Request("action"))
	ID = Request("ID")

		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for editing"
                		
                </script>
                <% response.end
        End If

	if action = "EXECUTE" then
		Dim ShrDescr
		Dim Ref
       
       ShrDescr = Request.Form("txtDescription")
       Ref = Request("txtReference")
       
		Set conn = GetActiveConnection("KBroker")
		If toCancel <> "" Then
			WriteDialogCancelScript
			Set Conn = Nothing
			Response.End
		End If
		
       
        'validate Short Description
        If Trim(ShrDescr) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify a short description."
                		
                </script>
                <% 
				reloadPage(ID)
				response.end
        End If
        
        'save data
        
        sqlStr = "UPDATE PaymentTypes SET Description = " & "'" & ShrDescr & "'" & _
				",Reference = " & " " & Ref & " " & _
				" WHERE PaymentTypes_DPA_  = " & ID                
        
        conn.BeginTrans
                conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        conn.CommitTrans
       
        conn.Close
        Set conn = Nothing
		WritefraEnabledDialogCloseScript2        
        Response.End
   	end If
   	  	
%>
<form name = 'frmEditPaymentTypes' method = 'post' action = 'EditPaymentTypes.asp' >
<table border="0" width="100%" cellspacing="0" cellpadding="0">
<%
        Set conn = GetActiveConnection("KBroker")
        
        
        sqlStr = "SELECT * FROM [PaymentTypesList] WHERE PaymentTypes_DPA_  = " & ID
        
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		ShowMessage "The selected Broker cannot be retrieved for editing"
                		
                </script>
                <% response.end
        End If
        
        if instr(1,Ucase(rs("Description")),"CASH") > 0 then
        %>
        <script language = 'vbscript'>
                		ShowMessage "Cannot Edit Cash"
                		
                </script>
        <% 'response.end
        'WritefraEnabledDialogCloseScript        
        response.end
        end if
        
        if instr(1,Ucase(rs("Description")),"CHEQUE") > 0 then
        %>
        <script language = 'vbscript'>
                		ShowMessage "Cannot Edit Cheque"
                		
                </script>
        <% 'response.end
        
        'WritefraEnabledDialogCloseScript
        response.end        
        end if
%>
  <tr>
    <td width="17%"> Short Description</td>
    <td width="83%"><input type = 'text' name ='txtDescription' id = 'txtDescription' size="20" value = '<%=rs.Fields("Description")%>'></td>
  </tr>
  <tr>
    <td width="18%"> Require Reference&nbsp; </td>
    <td width="82%">
	<select name ='txtReference' id = "txtReference">
	<% if Cint(rs.Fields("Reference").value) = 1 then %>
	<option value="1" selected>Yes</option>
	<option value="0">No</option>
	<% else %>
	<option value="1">Yes</option>
	<option value="0" selected>No</option>
	<% End if%>
	</select>
	</td>
  </tr>
  
</table>
 <table border=0 cellspacing=0 cellpadding=0 align=bottom width=100%>  
  <tr>
    <td align=right>
    <BR>
		<input type = 'button' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save "  onclick = "forceSubmit()">
		&nbsp;&nbsp;
		<input type = 'button' Class=Buttons name ='cmdClose' id = 'cmdClose' value=" Cancel " OnClick="JavaScript: window.self.close();">
    	<input type = 'hidden' name ='action' id = 'action' value="Execute">
    	<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
		<input type = 'hidden' name ='buttonAction' id = 'buttonAction' value="Save">
     </td>
  </tr>
 </table>
</form>

</body>

</html>
