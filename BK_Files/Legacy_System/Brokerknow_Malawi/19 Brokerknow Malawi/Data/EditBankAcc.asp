<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit Bank Account</title>
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
                		alert "No record specified for editing"
                		window.self.close
                </script>
                <% response.end
        End If

	if action = "EXECUTE" then
		Dim client
		Dim bnkBranch
		Dim accNum
		Dim accName
       
		client = Request.Form("cboClient")
		bnkBranch = Request.Form("cboBnkBranch")
		accNum = Request.Form("txtNumber")
		toCancel = Request.Form("cmdCancel")
		accName = Request.Form("txtAccName")
		
		Set conn = GetActiveConnection("KBroker")
		If toCancel <> "" Then
		
			WriteDialogCancelScript
			Set Conn = Nothing
			Response.End
		End If
		
        'validate Client
        If Trim(Client) = "" Then%>
                <script language = 'vbscript'>
                		alert "Please specify the Client"
                						   
                </script>
                <% response.end
        End If
        'validate Bank Branch
        If Trim(bnkBranch) = "" Then%>
                <script language = 'vbscript'>
                		alert "Please specify the Bank Branch"
                		
                </script>
                <% response.end
        End If
        'validate Account Number
        If Trim(accNum) = "" Then%>
                <script language = 'vbscript'>
                		alert "Please specify the Account Number"
                		
                </script>
                <% response.end
        End If
       'validate size of Account Number
        If Len(accNum) > 100 Then%>
                <script language = 'vbscript'>
                alert "Account Number can only be 100 characters in length"
                
                </script>
                <% response.end
        End If
        'validate size of Account Name
        If Len(accName) > 100 Then%>
                <script language = 'vbscript'>
                alert "Account Name can only be 100 characters in length"
                
                </script>
                <% response.end
        End If
        
     

        'save data
        
        sqlStr = "UPDATE [BankAcc] SET BankAccNumber = " & "'" & accNum & "'" & _
				",BankAccName = " & "'" & accName & "'" & _
				",BnkBranch_DPA_ = " & " " & bnkBranch & " " & _
				",Client_DPA_ = " & " " & client & " " & "" & _
                " WHERE BankAcc_DPA_  = " & ID
                
        
        conn.BeginTrans
                conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        conn.CommitTrans
        
        conn.Close
        Set conn = Nothing
        WritefraEnabledDialogCloseScript
        Response.End
   	end If
   	
%>

<form name = 'frmEditBankAcc' method = 'post' action = 'EditBankAcc.asp' >
<table border="0" width="100%" cellspacing="2" cellpadding="2">
  <tr>
    <td width="30%">Client</td>
    <td width="70%"><select name = 'cboClient' id = 'cboClient' size="1">
<%
        Set conn = GetActiveConnection("KBroker")
        
       
        sqlStr = "SELECT BankAccNumber,BankAccName , BankAcc.BankAcc_DPA_, BnkBranchList.BnkBranch_DPA_, ClientList.Client_DPA_ FROM [ClientList] INNER JOIN ([BnkBranchList] INNER JOIN [BankAcc] ON BnkBranchList.BnkBranch_DPA_ = BankAcc.BnkBranch_DPA_) ON ClientList.Client_DPA_ = BankAcc.Client_DPA_ WHERE BankAcc_DPA_  = " & ID
                
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		alert "The selected Bank Account cannot be retrieved for editing"
                		window.self.close
                </script>
                <% response.end
        End If
        
      
        sqlStr = "SELECT * FROM [ClientList]"
        Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rsEdit.EOF Or rsEdit.BOF) Then
                rsEdit.MoveFirst
                Do Until rsEdit.EOF
                		if rsEdit.Fields("Client_DPA_") = rs.Fields("Client_DPA_") Then%>
                			<option selected value = '<%=rsEdit.Fields("Client_DPA_")%>'><%=rsEdit.Fields("ClientName")%></option>
                		<%else%>
                        <option value = '<%=rsEdit.Fields("Client_DPA_")%>'><%=rsEdit.Fields("ClientName")%></option>
                     <%end if
						rsEdit.MoveNext
                Loop
        End If
%>

    </select></td>
  </tr>
  <tr>
    <td width="30%">Bank Branch</td>
    <td width="70%"><select name = 'cboBnkBranch' id = 'cboBnkBranch' size="1">
<%
        sqlStr = "SELECT * FROM [BnkBranchList]"
        Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rsEdit.EOF Or rsEdit.BOF) Then
                rsEdit.MoveFirst
                Do Until rsEdit.EOF
                		if rsEdit.Fields("BnkBranch_DPA_") = rs.Fields("BnkBranch_DPA_") Then%>
                			<option selected value = '<%=rsEdit.Fields("BnkBranch_DPA_")%>'><%=rsEdit.Fields("BnkBranchName")%></option>
                		<%else%>
                        <option value = '<%=rsEdit.Fields("BnkBranch_DPA_")%>'><%=rsEdit.Fields("BnkBranchName")%></option>
                     <%end if
						rsEdit.MoveNext
                Loop
        End If
%>

    </select></td>
  </tr>

  <tr>
    <td width="30%">Account Name</td>
    <td width="70%"><input type = 'text' name ='txtAccName' id = 'txtAccName' size="28" value = '<%=rs.Fields("BankAccName")%>'></td>
  </tr>

  <tr>
    <td width="30%">Account Number </td>
    <td width="70%"><input type = 'text' name ='txtNumber' id = 'txtNumber' size="28" value = '<%=rs.Fields("BankAccNumber")%>'></td>
  </tr>
  <tr>
    <td width="100%" colspan=2 align=right>
		<BR>
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save ">
		<input type = 'button' Class=Buttons name ='cmdClose' id = 'cmdClose' value=" Cancel " OnClick="JavaScript: window.self.close();">        
    	<input type = 'hidden' name ='action' id = 'action' value="Execute">
    	<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
    </td>
  </tr>
</table>
</form>


</body>

