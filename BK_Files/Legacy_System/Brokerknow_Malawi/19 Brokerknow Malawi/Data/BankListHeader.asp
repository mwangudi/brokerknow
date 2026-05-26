<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit Bank</title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
 <SCRIPT language='Javascript' >
function forceSubmit()
		{
			setOpener();
			//var targetPage = window.dialogArguments.opener.document.all.item("frmMain").elements("EditPage").value;
					
			document.frmEditBank.method='post';
			document.frmEditBank.target='_self';
			document.frmEditBank.submit();		
		}
		
		function setOpener()
		{
			window.parent.opener = window.parent.dialogArguments.opener;					
		}
</script>
</head>

<body Class="Dialog" marginwidth=0 marginheight=0 margintop=0 marginleft=0>

<!--#include file="../libroutines.asp"-->



<%
	const LinkedIndependent = 1
   const LinkedDependent = 2
	
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
   Dim ID
   Dim rsEdit
	
	action = ucase(Request.Form("action"))
	ID = Request("ID")
	
	guid = Request.Form("guid")

		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		alert "No record specified for editing"
                		
                </script>
                <% response.end
        End If

	select case action 
	case "EXECUTE_HEADER"
		Dim name
		dim bnkCode
				        
       name = Request.Form("txtName")
       bnkCode=Request.Form("txtBnkCode")       
       toCancel = Request.Form("cmdCancel")
       
	    If toCancel <> "" Then
			WriteDialogCloseScript
			Response.End
	    End If                
       
        'validate Name
        If Trim(bnkCode) = "" Then%>
                <script language = 'vbscript'>
                		alert "Please specify the Bank name"
                		
                </script>
                <% response.end
        End If
        'validate size of Name
        If Len(bnkCode) > 100 Then%>
                		<script language = 'vbscript'>
                alert "Name can only be 100 characters in length"
                
                </script>
                <% response.end
        End If
       
        'validate Name
        If Trim(name) = "" Then%>
                <script language = 'vbscript'>
                		alert "Please specify the Bank name"
                		
                </script>
                <% response.end
        End If
        'validate size of Name
        If Len(Name) > 100 Then%>
                		<script language = 'vbscript'>
                alert "Name can only be 100 characters in length"
                
                </script>
                <% response.end
        End If
        
        Set conn = GetActiveConnection("KBroker")	
		 
        'save data
        sqlStr = "UPDATE [Bank] SET BankCode = " & "'" & bnkCode & "',BankName = " & "'" & name & "'" & " WHERE Bank_DPA_  = " & ID                
        
        conn.BeginTrans
                conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        conn.CommitTrans
       
        conn.Close
        Set conn = Nothing
        
   
    Case Else
		ID = GetBankID(ID)    
   	end select
   	
   	Function GetBankID(branchID)
   		Dim getRs
   		Set getConn = GetActiveConnection("KBroker")
   	   	Set getRs = getConn.Execute("SELECT Bank_DPA_ FROM BankList WHERE BnkBranch_DPA_ = " & branchID)
   	   	If Not (getRs.EOF Or getRs.BOF) Then
   	   		GetBankID = getRs.Fields("Bank_DPA_").Value
   	   	Else
   	   		GetBankID = ""
   	   	End If	
   	End Function
%>
<form name = 'frmEditBank' method = 'post' action = 'BankListHeader.asp'>
<table border="0" width="100%" cellspacing=1 cellpadding=1>
  <%
		Set conn = GetActiveConnection("KBroker")
        
        
		sqlStr = "SELECT BankCode,BankName,Bank_DPA_ FROM [Bank] WHERE Bank_DPA_  = " & ID 
        
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		alert "The selected Bank cannot be retrieved for editing"
                		
                </script>
                <% response.end
        End If       
        
        %>
  <tr>
    <td width="20%">Bank Code</td>
    <td width="80%"><input type = 'text' name ='txtBnkCode' id = 'txtBnkCode' value = '<%=rs.Fields("BankCode")%>' size="20" STYLE="WIDTH: 250px;"></td>
  </tr>  
  <tr>
    <td width="20%">Bank Name</td>
    <td width="80%"><input type = 'text' name ='txtName' id = 'txtName' value = '<%=rs.Fields("BankName")%>' size="20" STYLE="WIDTH: 250px;"></td>
  </tr>
  <tr>
    <td width="100%" colspan=2 align=right>
	<BR><BR>	
		<input type = 'button' Class=Buttons name ='cmdDelete' id = 'cmdDelete' value="Delete" OnClick="JavaScript: window.parent.frames['detail'].DoDelete();"> 
		&nbsp;&nbsp;
		<input type = 'button' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save " onClick="forceSubmit()"> 
		 &nbsp;&nbsp;
		 <input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Close " onClick="window.self.close;">       
    	<input type = 'hidden' name ='action' id = 'action' value="Execute_Header">
    	<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
    </td>
  </tr>  
</table>
</form>

</body>

</html>
