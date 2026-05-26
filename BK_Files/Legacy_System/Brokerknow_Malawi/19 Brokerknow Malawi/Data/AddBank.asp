<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add Bank</title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
 <SCRIPT language='Javascript'>
 function forceSubmit()
		{
			setOpener();
			//var targetPage = window.dialogArguments.opener.document.all.item("frmMain").elements("EditPage").value;
					
			document.frmAddBank.method='post';
			document.frmAddBank.target='_self';
			document.frmAddBank.submit();		
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
   Dim guidStr 
   Dim guid 
	
	action = ucase(Request.Form("action"))
	
	if action = "EXECUTE" then
		Dim name
		Dim bnkBranch
		Dim branchCode
		Dim swiftCode
		dim bnkCode
		        
       bnkCode=Request.Form("txtBnkCode")
       name = Request.Form("txtName")
       bnkBranch = Request.Form("txtBnkBranch")
       branchCode = Request.Form("txtBnkBranchCode")
       swiftCode = Request.Form("txtBnkBranchSwiftCode")
              
       
        'validate Bank Code
        If Trim(bnkCode) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Bank Code"
                </script>
                <% response.end
        End If
        'validate size of Bank Code
        If Len(bnkCode) > 100 Then%>
                		<script language = 'vbscript'>
						ShowMessage "Bank Code can only be 100 characters in length"
                
                </script>
                <% response.end
        End If
        
        'validate Name
        If Trim(name) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Bank name"
                </script>
                <% response.end
        End If
        'validate size of Name
        If Len(Name) > 100 Then%>
                		<script language = 'vbscript'>
						ShowMessage "Name can only be 100 characters in length"
                
                </script>
                <% response.end
        End If      		
		'validate detail info
                'validate Bank Branch
                If Trim(bnkBranch) = "" Then%>
                		<script language = 'vbscript'>
                				ShowMessage "Please specify the Bank branch"
                				
                		</script>
                		<% response.end
                End If
                'validate size of Bank Branch Name
                If Len(BnkBranch) > 100 Then%>
                		<script language = 'vbscript'>
                        ShowMessage "Bank Branch Name can only be 100 characters in length"
                        
                		</script>
                		<% response.end
                End If
                'validate size of Bank Branch Code
                If Len(branchCode) > 20 Then%>
                		<script language = 'vbscript'>
                        ShowMessage "Bank Branch Code can only be 20 characters in length"
                        
                		</script>
                		<% response.end
                End If
                'validate size of Bank Branch Swift Code
                If Len(swiftCode) > 20 Then%>
                		<script language = 'vbscript'>
                        ShowMessage "Bank Branch Swift Code can only be 20 characters in length"
                        
                		</script>
                		<% response.end
                End If
                
                
        'save header
        set guid = server.createobject("NDUtils.CGUID")
        guidStr = guid.GenerateGUID
        
        sqlStr = "INSERT INTO [Bank] (BankCode,BankName,Bank_DPA_,Bank_EIT_) SELECT " & "'" & bnkCode & "'" & " as BankCode," & "'" & name & "'" & " as BankName," & " " & "iif(isnull(max([Bank_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Bank'),max([Bank_DPA_]) + 1)" & " " & " as Bank_DPA_" & _
                "," & "'" & guidStr & "'" & " as Bank_EIT_ FROM [Bank]"
        
        Set conn = GetActiveConnection("KBroker")
        
        conn.BeginTrans
			conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
	        
			'obtain header key value
			sqlStr = "SELECT [Bank.Bank_DPA_] FROM [Bank] WHERE [Bank.Bank_EIT_] = " & "'" & guidStr & "'"
	        
			Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
			If (rs.EOF Or rs.BOF) Then%>
					<script language = 'vbscript'>
                			ShowMessage "A serious error has been encountered while saving the data. Try saving again"
                			
					</script>
					<% response.end
			End If
	        
			'save detail data
			sqlStr = "INSERT INTO [BnkBranch] (BnkBranchName,BnkBranch_DPA_,BnkBranchCode,BnkBranchSwiftCode,Bank_DPA_) SELECT " & "'" & bnkBranch & "'" & " as BnkBranchName" & _
					"," & " " & "iif(isnull(max([BnkBranch_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'BnkBranch'),max([BnkBranch_DPA_]) + 1)" & " " & " as BnkBranch_DPA_" & _
					"," & "'" & branchCode & "'" & " as BnkBranchCode" & _
					"," & "'" & swiftCode & "'" & " as BnkBranchSwiftCode" & _
					"," & " " & rs.Fields("Bank_DPA_") & " " & " as Bank_DPA_" & _
					" FROM [BnkBranch]"
	        
			conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
		conn.CommitTrans
		Set Conn = Nothing
		WritefraEnabledDialogCloseScript2
        Response.End        
   	end If
%>

<form name = 'frmAddBank' method = 'post' action = 'AddBank.asp' target="deleteFrame" OnSubmit="JavaScript: UpdateDialogHandle();">
<table border="0" width="100%">
  <tr>
    <td width="30%">Bank Code</td>
    <td width="70%"><input type = 'text' name ='txtBnkCode' id = 'txtBnkCode' size="20"></td>
  </tr>  
  <tr>
    <td width="30%">Bank Name</td>
    <td width="70%"><input type = 'text' name ='txtName' id = 'txtName' size="20"></td>
  </tr>
  <tr>
	<td width="30%"><b><font color="#000080">Branch Name</font></b></td>
    <td width="70%"><input type = 'text' name ='txtBnkBranch' id = 'txtBnkBranch' size="20"></td>
  </tr>
  <tr>
	<td width="30%"><b><font color="#000080">Branch Code</font></b></td>
    <td width="70%"><input type = 'text' name ='txtBnkBranchCode' id = 'txtBnkBranchCode' size="20"></td>
  </tr>
  <tr>
	<td width="30%"><b><font color="#000080">Swift Code</font></b></td>
    <td width="70%"><input type = 'text' name ='txtBnkBranchSwiftCode' id = 'txtBnkBranchSwiftCode' size="20"></td>
  </tr>
</table>
<table border=0 width="100%">  
  <tr>
     <td align=right>
		<BR>
		<BR>
		<BR>
		<BR>
		<input type = 'button' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value="Save" onclick="forceSubmit()">
		&nbsp;&nbsp;
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();">		
		<input type = 'hidden' name ='action' id = 'action' value="Execute">&nbsp;
     </td>
  </tr>
</table>
</form>
</body>

</html>
