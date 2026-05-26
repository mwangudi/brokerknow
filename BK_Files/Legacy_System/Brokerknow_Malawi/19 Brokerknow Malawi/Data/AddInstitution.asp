<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add Institution</title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
<script language='javascript'>
	function forceSubmit()
	{
		setOpener();
		//var targetPage = window.dialogArguments.opener.document.all.item("frmMain").elements("EditPage").value;
				
		document.frmAddBroker.method='post';
		document.frmAddBroker.target='_self';
		document.frmAddBroker.submit();	
		
	}
	
	function setOpener()
	{

		window.self.opener = window.dialogArguments.opener;
				
	}
</script>
</head>

<body Class="Dialog" onLoad="setOpener()">

<!--#include file="../libroutines.asp"-->
<%
	
   Dim action
   Dim conn 
   Dim sqlStr
   Dim rs
	
	action = ucase(Request.Form("action"))
	
	if action = "EXECUTE" then
		Dim Name
		Dim code
		Dim addr
		Dim phone
		Dim fax
		Dim OpeningBal
       
     
       Name = Request.Form("txtName")
       addr = Replace(Request.Form("txtAddr"),"'","")
       phone = Request.Form("txtOfficeTel")
       fax = Request.Form("txtFax")
      
       
        'validate Name
        If Trim(Name) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Name"
                		
                </script>
                <% 
				ReloadPage(ID)
				response.end
        End If
        'validate size of Name
        If Len(Name) > 100 Then%>
                <script language = 'vbscript'>
                ShowMessage "Name can only be 100 characters in length"
                
                </script>
                <% 
				ReloadPage(ID)
				response.end
        End If
       
        'validate size of Address
        If Len(addr) > 100 Then%>
                <script language = 'vbscript'>
                ShowMessage "Address can only be 5 characters in length"
                
                </script>
                <% response.end
        End If
        'validate size of Telephone
        If Len(phone) > 50 Then%>
                <script language = 'vbscript'>
                ShowMessage "Telephone field is too long"
                
                </script>
                <%
				ReloadPage(ID)
				response.end
        End If
        'validate size of Fax
        If Len(fax) > 20 Then%>
                <script language = 'vbscript'>
                ShowMessage "Fax can only be 5 characters in length"
                
                </script>
                <% 
				ReloadPage(ID)
				response.end
        End If
        
        
        'save data
    

           sqlStr = "INSERT INTO [Institution] (InstitutionName, Address, PhoneNumber,Fax) VALUES (" & "'" & Name & "'," & "'" & addr & "'," & _
                
                 "'" & phone & "',"  & "'" & fax & "')" 
				 
				'response.write sqlstr : response.end

        Set conn = GetActiveConnection("KBroker")
        
        conn.BeginTrans
               ' sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
                conn.Execute sqlStr
        conn.CommitTrans
        conn.Close
        Set conn = Nothing
        WritefraEnabledDialogCloseScript2
        Response.End
   	end If
%>

<form name = 'frmAddBroker' method = 'post' action = 'AddInstitution.asp' >
<table border="0" cellspacing="1" cellpadding="1">
  <tr>
    <td nowrap> Name</td>
    <td nowrap><input type = 'text' name ='txtName' id = 'txtName' size="20"></td>
  </tr>
 
  <tr>
    <td nowrap> Address</td>
    <td nowrap><textarea rows=3 name ='txtAddr' id = "txtAddr"></textarea></td>
  </tr>
  <tr>
    <td nowrap> Telephone</td>
    <td nowrap><input type = 'text' name ='txtOfficeTel' id = "txtOfficeTel" size="20"></td>
  </tr>
  <tr>
    <td nowrap> Fax</td>
    <td nowrap><input type = 'text' name ='txtFax' id = "txtFax" size="20"></td>
  </tr>
  
</table>
<table border=0 cellspacing=0 cellpadding=0 align=bottom width=100%>  
  <tr>
    <td align=right>
	<b id="hide" name="hide">
    <input type = 'button' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save " onclick="forceSubmit();">
    &nbsp;&nbsp;&nbsp;</b>
    <input type = 'button' Class=Buttons name ='cmdEdit' id = 'cmdEdit' value=" Cancel " OnClick="JavaScript: window.close();">
    <input type = 'hidden' name ='action' id = 'action' value="Execute">&nbsp;
	<input type = 'hidden' name ='buttonAction' id = 'action' value="Save">
     </td>
  </tr>
 </table>
</form>

</body>

</html>
