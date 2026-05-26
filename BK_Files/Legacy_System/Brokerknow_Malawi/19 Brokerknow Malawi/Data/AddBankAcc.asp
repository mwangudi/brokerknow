<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add Bank Account </title>
<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT>
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css"> 

</head>

<body Class="Dialog">

<!--#include file="../libroutines.asp"-->

<%
	
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
	
	action = ucase(Request.Form("action"))
	
	if action = "EXECUTE" then
		Dim client
		Dim bnkBranch
		Dim accNum
		Dim accName
       
		client = Request.Form("cboClient")
		bnkBranch = Request.Form("cboBnkBranch")
		accNum = Request.Form("txtNumber")
		accName = Request.Form("txtAccName")
       
       
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
        sqlStr = "INSERT INTO [BankAcc] (BankAccNumber,BankAcc_DPA_,BankAccName,BnkBranch_DPA_,Client_DPA_) SELECT " & "'" & accNum & "'" & " as BankAccNumber" & _
                "," & " " & "iif(isnull(max([BankAcc_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'BankAcc'),max([BankAcc_DPA_]) + 1)" & " " & " as BankAcc_DPA_" & _
                "," & "'" & accName & "'" & " as BankAccName" & _
                "," & " " & bnkBranch & " " & " as BnkBranch_DPA_" & _
                "," & " " & client & " " & " as Client_DPA_" & _
                " FROM [BankAcc]"
        Set conn = GetActiveConnection("KBroker")
        
        conn.BeginTrans
                conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
        conn.CommitTrans
        conn.Close
        Set conn = Nothing
        WritefraEnabledDialogCloseScript
        Response.End
   	end If
%>

<form name = 'frmAddBankAcc' method = 'post' action = 'AddBankAcc.asp' >
<table border="0" width="100%" cellspacing="2" cellpadding="2">
  <tr>
    <td width="30%">Client</td>
    <td width="70%"><select name = 'cboClient' id = 'cboClient' size="1">
    	<option selected value = ''></option>
<%
        Set conn = GetActiveConnection("KBroker")
        
        sqlStr = "SELECT * FROM [ClientList]"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF%>
                        <option value = '<%=rs.Fields("Client_DPA_")%>'><%=rs.Fields("ClientName")%></option>
                        <%rs.MoveNext
                Loop
        End If
%>

    </select></td>
  </tr>
  <tr>
    <td width="30%">Bank Branch</td>
    <td width="70%"><select name = 'cboBnkBranch' id = 'cboBnkBranch' size="1">
    	<option selected value = ''></option>
<%
        sqlStr = "SELECT * FROM [BnkBranchList]"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF%>
                        <option value = '<%=rs.Fields("BnkBranch_DPA_")%>'><%=rs.Fields("BnkBranchName")%></option>
                        <%rs.MoveNext
                Loop
        End If
%>

    </select></td>
  </tr>
  <tr>
    <td width="30%">Account Name</td>
    <td width="70%"><input type = 'text' name ='txtAccName' id = 'txtAccName' size="28"></td>
  </tr>
  <tr>
    <td width="30%">Account Number</td>
    <td width="70%"><input type = 'text' name ='txtNumber' id = 'txtNumber' size="28"></td>
  </tr>
  <tr>
    <td width="100%" colspan=2 align=right>
		<BR>
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save ">
		&nbsp;&nbsp;
		<input type = 'button' Class=Buttons name ='cmdClose' id = 'cmdClose' value=" Cancel " OnClick="JavaScript: window.self.close();">
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
      </td>
  </tr>
  
  </table>
</form>

</body>

</html>
