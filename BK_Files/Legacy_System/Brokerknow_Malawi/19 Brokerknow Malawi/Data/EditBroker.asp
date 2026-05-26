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
		//var targetPage = window.dialogArguments.opener.document.all.item("frmMain").elements("EditPage").value;
				
		document.frmEditBroker.method='post';
		document.frmEditBroker.target='_self';
		document.frmEditBroker.submit();	
		
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
   Dim ID
   Dim rsEdit
	
	action = ucase(Request("action"))
	ID = Request("ID")

		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for editing"
                		
                </script>
                <% 
				
				response.end
        End If

	if action = "EXECUTE" then
		Dim Name
		Dim code
		Dim addr
		Dim phone
		Dim fax
       Dim OpeningBal
       
       OpeningBal = Request.Form("txtOpeningBal")
       Name = Request("txtName")
       code = Request("txtCode")
       addr = Request("txtAddr")
       phone = Request("txtOfficeTel")
       fax = Request("txtFax")
       toCancel = Request.Form("cmdCancel")
		Set conn = GetActiveConnection("KBroker")
		If toCancel <> "" Then
			WriteDialogCancelScript
			Set Conn = Nothing
			Response.End
		End If
		
       
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
       'validate Code
        If Trim(code) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Code"
                		
                </script>
                <% 
				ReloadPage(ID)
				response.end
        End If
        'validate size of Code
        If Len(Code) > 5 Then%>
                <script language = 'vbscript'>
                ShowMessage "Code can only be 5 characters in length"
                
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
                <% 
				ReloadPage(ID)
				response.end
        End If
        'validate size of Telephone
        If Len(phone) > 20 Then%>
                <script language = 'vbscript'>
                ShowMessage "Telephone can only be 5 characters in length"
                
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

		If Not IsNumeric(OpeningBal) Then%>
				<script language = 'vbscript'>
				ShowMessage "Opening Balance can only be numeric"						
				</script>
				<% 
				ReloadPage(ID)
				response.end
		End If
        
        'save data
        
        sqlStr = "UPDATE Broker SET BrokerCode = " & "'" & code & "'" & _
				",BrokerName = " & "'" & Name & "'" & _
				",BrokerAddr = " & "'" & addr & "'" & _
				",BrokerOfficeTel = " & "'" & phone & "'" & _
				",BrokerFax = " & "'" & fax & "'" & _
				",BrokerOpeningBal = " & " " & OpeningBal & " " & _
				" WHERE Broker_DPA_  = " & ID                
        
        conn.BeginTrans
                conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        conn.CommitTrans
       
        conn.Close
        Set conn = Nothing
		WritefraEnabledDialogCloseScript2        
        Response.End
   	end If
   	
   		
   	
%>
<form name = 'frmEditBroker' method = 'post' action = 'EditBroker.asp' >
<table border="0" width="100%" cellspacing="0" cellpadding="0">
<%
        Set conn = GetActiveConnection("KBroker")
        
        
        sqlStr = "SELECT * FROM [Broker] WHERE Broker_DPA_  = " & ID
        
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		ShowMessage "The selected Broker cannot be retrieved for editing"
                		
                </script>
                <% response.end
        End If
        
      

%>
  <tr>
    <td width="17%"> Name</td>
    <td width="83%"><input type = 'text' name ='txtName' id = 'txtName' size="20" value = '<%=rs.Fields("BrokerName")%>'></td>
  </tr>
  <tr>
    <td width="18%"> Code&nbsp; </td>
    <td width="82%"><input type = 'text' name ='txtCode' id = "txtCode" size="20" value = '<%=rs.Fields("BrokerCode")%>'></td>
  </tr>
  <tr>
    <td width="18%"> Address</td>
    <td width="82%"><textarea rows=3 name ='txtAddr' id = "txtAddr"><%=rs.Fields("BrokerAddr")%></textarea></td>
  </tr>
  <tr>
    <td width="18%"> Telephone</td>
    <td width="82%"><input type = 'text' name ='txtOfficeTel' id = "txtOfficeTel" size="20" value = '<%=rs.Fields("BrokerOfficeTel")%>'></td>
  </tr>
  <tr>
    <td width="18%"> Fax</td>
    <td width="82%"><input type = 'text' name ='txtFax' id = "txtFax" size="20" value = '<%=rs.Fields("BrokerFax")%>'></td>
  </tr>
  <tr>
    <td nowrap> Opening Balance</td>
    <td nowrap><input type = 'text' name ='txtOpeningBal' STYLE="TEXT-ALIGN: RIGHT;" id = "txtOpeningBal" size="20" value = '<%=rs.Fields("BrokerOpeningBal")%>'></td>
  </tr>
</table>
 <table border=0 cellspacing=0 cellpadding=0 align=bottom width=100%>  
  <tr>
    <td align=right>
    <BR>
		<b  name="hide" id="hide"><input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save "  onclick="forceSubmit();">
		&nbsp;&nbsp;</b>
		<input type = 'button' Class=Buttons name ='cmdClose' id = 'cmdClose' value=" Cancel " OnClick="JavaScript: window.self.close();">
    	<input type = 'hidden' name ='action' id = 'action' value="Execute">
    	<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
     </td>
  </tr>
 </table>
</form>

</body>

</html>
