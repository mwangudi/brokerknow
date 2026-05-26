<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add Broker</title>

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
       
       OpeningBal = Request.Form("txtOpeningBal")
       Name = Request.Form("txtName")
       code = Request.Form("txtCode")
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
       'validate Code
        If Trim(code) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Code"
                		
                </script>
                <%ReloadPage(ID) 
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
        
        If Not IsNumeric(OpeningBal) Then%>
				<script language = 'vbscript'>
				ShowMessage "Opening Balance can only be numeric"						
				</script>
				<% 
				ReloadPage(ID)
				response.end
		End If
				
        'save data
        sqlStr = "INSERT INTO [Broker] (BrokerCode, BrokerName, BrokerAddr, BrokerOfficeTel, BrokerFax, BrokerOpeningBal, Broker_DPA_) SELECT " & "'" & code & "'" & " as BrokerCode," & "'" & Name & "'" & " as BrokerName" & _
                "," & "'" & addr & "'" & " as BrokerAddr" & _
                "," & "'" & phone & "'" & " as BrokerOfficeTel" & _
                "," & "'" & fax & "'" & " as BrokerFax" & _
                "," & " " & OpeningBal & " " & " as BrokerOpeningBal" & _
                "," & " " & "iif(isnull(max([Broker_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Broker'),max([Broker_DPA_]) + 1)" & " " & " as Broker_DPA_" & _
                " FROM [Broker]"
        Set conn = GetActiveConnection("KBroker")
        
        conn.BeginTrans
                sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
                conn.Execute sqlStr
        conn.CommitTrans
        conn.Close
        Set conn = Nothing
        WritefraEnabledDialogCloseScript2
        Response.End
   	end If
%>

<form name = 'frmAddBroker' method = 'post' action = 'AddBroker.asp' >
<table border="0" cellspacing="1" cellpadding="1">
  <tr>
    <td nowrap> Name</td>
    <td nowrap><input type = 'text' name ='txtName' id = 'txtName' size="20"></td>
  </tr>
  <tr>
    <td nowrap> Code</td>
    <td nowrap><input type = 'text' name ='txtCode' id = "txtCode" size="20"></td>
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
  <tr>
    <td nowrap> Opening Balance</td>
    <td nowrap><input type = 'text' name ='txtOpeningBal' STYLE="TEXT-ALIGN: RIGHT;" id = "txtOpeningBal" size="20" value="0"></td>
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
