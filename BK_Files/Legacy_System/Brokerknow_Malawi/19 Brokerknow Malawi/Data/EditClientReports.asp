<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit Client Reports</title>
   <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
</head>

<body Class="Dialog">

<!--#include file="../libroutines.asp"-->


<%
	UserId=Session("UserID")
	
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
   Dim ID
   Dim rsEdit
	
	action = ucase(Request.Form("action"))
	ID = Request("ID")
	
	'action = "EXECUTE"
	'ID = 904
	
	If Trim(ID) = "" Then%>
            <script language = 'vbscript'>
            		ShowMessage "No record specified for editing"
                		
            </script>
            <% response.end
    End If

	if action = "EXECUTE" then
		Dim ClientStatement
		Dim ClientContract
		Dim ClientContractCompounded
		Dim HoldingsValuation
		Dim email
		Dim gen2
		
		email = Request.Form("txtEmail")
		gen2 = Request.Form("cboGenericSetting2")
		ClientStatement = Request.Form("chkClientStatement")
		ClientContract = Request.Form("chkClientContract")
		ClientContractCompounded = Request.Form("chkClientContractCompounded")
		HoldingsValuation = Request.Form("chkHoldingsValuation")
       
		if ClientStatement = "on" then
			ClientStatement = 1
		else
			ClientStatement = 0
		end if
		
		if ClientContract = "on" then
			ClientContract = 1
		else
			ClientContract = 0
		end if
		
		if ClientContractCompounded = "on" then
			ClientContractCompounded = 1
		else
			ClientContractCompounded = 0
		end if
		
		if HoldingsValuation = "on" then
			HoldingsValuation = 1
		else
			HoldingsValuation = 0
		end if
		
		'validate size of Email
		If Len(Email) > 100 Then%>
			<script language = 'vbscript'>
			ShowMessage "Email can only be 100 characters in length"
						
			</script>
			<% response.end
		End If
		
		'validate size of Email
		If Len(gen2) = 0 Then%>
			<script language = 'vbscript'>
			ShowMessage "Please select the frequency"
						
			</script>
			<% response.end
		End If
		
		'validate Generic2
		If Trim(gen2) = "" Then
				gen2 = "NULL"
		End If
				
        Set conn = GetActiveConnection("KBroker")
       
        'save data
        sqlStr = "SELECT OtherNames + ' ' + Surname AS theName FROM Users WHERE (UserID = "& UserID &")"
		Set rst = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		If Not (rst.EOF Or rst.BOF) Then
			theName = rst("theName")
		Else
			theName = ""
		End If
							
        sqlStr = "UPDATE [SendClientReports]" & _
				" SET GenericSetting_DPA_2 = "& gen2 &"" & _
				",ClientStatements = "& ClientStatement & "" & _
				",ClientContracts = "& ClientContract & "" & _
				",ClientContractsCompounded = "& ClientContractCompounded & "" & _
				",HoldingsValuation = "& HoldingsValuation & "" & _
				",ChangedBy = '"& theName & "'" & _
				",TimeChanged = GetDate()" & _
				" WHERE ClientReportsID  = " & ID
        
        conn.BeginTrans
                conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        conn.CommitTrans
        
        sqlStr = "SELECT Client_DPA_ FROM [SendClientReports] WHERE ClientReportsID  = " & ID
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        
        If Not (rs.EOF Or rs.BOF) Then
			'sqlStr = "UPDATE [Client]" & _
			'	" SET ClientEmail = " & "'" & email & "'" & "" & _
			'	",GenericSetting_DPA_2 = " & " " & gen2 & " " & "" & _
			'	" WHERE Client_DPA_  = " & rs.Fields("Client_DPA_")
			
			sqlStr = "UPDATE [Client]" & _
				" SET ClientEmail = " & "'" & email & "'" & "" & _
				" WHERE Client_DPA_  = " & rs.Fields("Client_DPA_")
				
			conn.BeginTrans
			        conn.Execute SQLServerFormat(HandleQuote(sqlStr))
			conn.CommitTrans   
        End If
        
        conn.Close
        Set conn = Nothing
        WriteFraEnabledDialogCloseScript
        response.end
   	end If
%>



<form name = 'frmEdit' method = 'post' action = 'EditClientReports.asp' target="deleteFrame" OnSubmit="JavaScript: UpdateDialogHandle();">
<table border="0" width="100%" cellspacing="2" cellpadding="2" align="center">
<%
        Set conn = GetActiveConnection("KBroker")
       
        sqlStr = "SELECT * FROM SendClientReports" & _
        " INNER JOIN Client ON SendClientReports.Client_DPA_ = Client.Client_DPA_" & _
        " INNER JOIN GenericSetting ON SendClientReports.GenericSetting_DPA_2 = GenericSetting.GenericSetting_DPA_" & _
        " WHERE SendClientReports.ClientReportsID  = " & ID        
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		ShowMessage "The selected record cannot be retrieved for editing"
                		
                </script>
                <% response.end
        End If

		
%>
  <tr>
    <td width="30%"> Client Code</td>
    <td width="80%"> <%=rs.Fields("Client_DPA_")%></td>
  </tr>
  <tr>
    <td width="30%"> Client Name</td>
    <td width="80%"> <%=rs.Fields("ClientName")%></td>
  </tr>
  <tr>
    <td width="30%"> Client REF No</td>
    <td width="80%"> <%=rs.Fields("ClientCDSNo")%></td>
  </tr>
  <tr>
    <td width="30%"> Client Email</td>
    <td width="80%"> 
    <input type = 'text' name ='txtEmail' STYLE="WIDTH: 220px" tabIndex='13' id = 'txtEmail' size="20"  value = '<%=rs.Fields("ClientEmail")%>'></td>
  </tr>
  <tr>
    <td width="30%"> Frequency</td>
	<td width="80%">
		<select name = 'cboGenericSetting2' id = 'cboGenericSetting2' size="1">
		<!--<option selected value = ''></option>-->
		<%
		sqlStr = "SELECT * FROM [GenericSettingList] WHERE (EntityType_DPA_ = 1) AND (Generic_DPA_ = 2) Order By GenericSettingDescription"
		Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		
		If Not (rsEdit.EOF Or rsEdit.BOF) Then
			Do Until rsEdit.EOF
					if Trim(rsEdit.Fields("GenericSetting_DPA_")) = Trim(rs.Fields("GenericSetting_DPA_")) Then
						%>
						<option selected value = '<%=rsEdit.Fields("GenericSetting_DPA_")%>'><%=rsEdit.Fields("GenericSettingDescription")%></option>
						<%
					else
						%>
						<option value = '<%=rsEdit.Fields("GenericSetting_DPA_")%>'><%=rsEdit.Fields("GenericSettingDescription")%></option>
						<%
					end if
				rsEdit.MoveNext
			Loop
		End If
		%>
		</select>
	</td>
  </tr>  
  <tr>
    <td width="30%" >Client Statement</td>
    <%
	 if trim(rs.fields("ClientStatements")) = 1 then
	   ClientStatement = "Checked"
	 else
	   ClientStatement = ""
	 end if
	%>
    <td width="70%" ><input type = 'checkbox' name ='chkClientStatement' id = 'chkClientStatement' <%=ClientStatement%>></td>
  </tr>  
  
  <tr>
    <td width="30%" >Client Contracts</td>
    <%
	 if trim(rs.fields("ClientContracts")) = 1 then
	   ClientContract = "Checked"
	 else
	   ClientContract = ""
	 end if
	%>
    <td width="70%" ><input type = 'checkbox' name ='chkClientContract' id = 'chkClientContract' <%=ClientContract%>></td>
  </tr> 
  
  <tr>
    <td width="30%" >Client Contracts Compounded</td>
    <%
	 if trim(rs.fields("ClientContractsCompounded")) = 1 then
	   ClientContractCompounded = "Checked"
	 else
	   ClientContractCompounded = ""
	 end if
	%>
    <td width="70%" ><input type = 'checkbox' name ='chkClientContractCompounded' id = 'chkClientContractsCompounded' <%=ClientContractCompounded%>></td>
  </tr>  
  
  <tr>
    <td width="30%">Holdings Valuation</td>
    <%
	 if trim(rs.fields("HoldingsValuation")) = 1 then
	   HoldingsValuation = "Checked"
	 else
	   HoldingsValuation = ""
	 end if
	%>
    <td width="70%" ><input type = 'checkbox' name ='chkHoldingsValuation' id = 'chkHoldingsValuation' <%=HoldingsValuation%>></td>
  </tr>  
  
  <tr>
    <td width="100%" colspan="2" align=right>
		<BR>
		<BR>		
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save ">
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();">
		&nbsp;&nbsp;
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
    	<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
      </td>
  </tr>
</table>
</form>

</body>

</html>
