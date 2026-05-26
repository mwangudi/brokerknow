<html>

<head>

<title>Add Client Reports</title>
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT>

<script language = 'javascript'>
		var validNavigate = false;
		
		function ReleaseRecord()
		{
			if(!validNavigate)
			{
 				event.returnValue = "Please use the cancel button to close the dialog"
 			}
		}
		
		function AllowedNavigation()
		{
			validNavigate = true;
		}
		
		ExemptFromDefaultTabIndex = true;
		
		function DisplayDefaultCommission(theList)
		{
			var commissionList = document.frmMain.cboCommission
			commissionList.value = theList.options[theList.selectedIndex].DefaultCommission;
		}
</script>

</head>

<body Class="Dialog">

<!--#include file="../libroutines.asp"-->
<html>

<%
UserId=Session("UserID")
	
Dim action
Dim conn 
Dim sqlStr
Dim rs
Dim ID
Dim rsEdit
	
action = ucase(Request.Form("action"))

if action = "EXECUTE" then
	Dim ClientStatement
	Dim ClientContract
	Dim ClientContractCompounded
	Dim HoldingsValuation
	Dim client
	Dim gen2
		
	client = Request.Form("cboClient")
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
		
	If Trim(client) = "" Then%>
			<script language = 'vbscript'>
    		ShowMessage "Please select the client"
			</script>
			<% response.end
	End If
	
	If Trim(gen2) = "" Then%>
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
						
	sqlStr = "INSERT INTO [SendClientReports] (Client_DPA_, GenericSetting_DPA_2, ClientStatements, ClientContracts, ClientContractsCompounded, HoldingsValuation, ChangedBy, TimeChanged)" & _
	" VALUES (" & client & ", "& gen2 &", "& ClientStatement & ", "& ClientContract & ", "& ClientContractCompounded & ", "& HoldingsValuation & ", '"& theName & "', GetDate())"
	Set conn = GetActiveConnection("KBroker")
	
	conn.BeginTrans
			conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
    conn.CommitTrans
    
	'sqlStr = "UPDATE [Client]" & _
	'	" SET GenericSetting_DPA_2 = " & " " & gen2 & " " & "" & _
	'	" WHERE Client_DPA_  = " & client
        
	'conn.BeginTrans
	'        conn.Execute SQLServerFormat(HandleQuote(sqlStr))
	'conn.CommitTrans   
            
    conn.Close
    Set conn = Nothing
    
    WriteFraEnabledDialogCloseScript
    response.end
end If
%>

<form name = 'frmEdit' method = 'post' action = 'AddClientReports.asp'>
<table border="0" width="100%" cellspacing="2" cellpadding="2" align="center">
  <tr>
		<td width="30%"> Client</td>
		<td width="80%"><input type = 'text' name ='txtClientCode' id = 'txtClientCode' size="10" onBlur="txtval = this.value; selectItem(cboClient);"><span lang="en-gb">&nbsp;
        </span><select name = 'cboClient' id = "cboClient" size="1" 
    		onchange='UpdateCode(true,cboClient,txtClientCode)'
			onKeypress="return (dodefaultaction()==''); " 
			onKeydown="return (dodefaultaction()==''); " 
			onKeyup="return (FilterData(this,1,UpdateCode(change(cboClient,0),cboClient,txtClientCode)));" 
			onfocus="txtval = '';inputIsItemCode = 1;" 
			onblur="txtval = '';inputIsItemCode = 1;">
			<option selected SearchCode = "" SearchText = ""  value = ''></option>
			<%
			dim ClientName
			dim NameClient
			        Set conn = GetActiveConnection("KBroker")
					        
			        sqlStr = "SELECT * FROM FullClientList order by ClientName"
			        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
			        If Not (rs.EOF Or rs.BOF) Then
			                rs.MoveFirst
			                Do Until rs.EOF
			                ClientName=rs.Fields("ClientName")
			                NameClient=Mid(ClientName,1,30)
			                %>					                        
			                        <option SearchCode = "<%=rs.Fields("Client_DPA_")%>" SearchText = "<%=NameClient%>" value = '<%=rs.Fields("Client_DPA_")%>'><%=NameClient%></option>

			                        <%rs.MoveNext
			                Loop
			        End If
			%>

			    </select></td>
				
	</tr>
  
  <tr>
    <td width="30%"> Frequency</td>
	<td width="80%">
		<select name = 'cboGenericSetting2' id = 'cboGenericSetting2' size="1">
			<option selected value = ''></option>
			<%
			sqlStr = "SELECT * FROM [GenericSettingList] WHERE EntityType_DPA_ = 1 Order By GenericSettingDescription"
			Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		
			rsEdit.Filter = "Generic_DPA_ = 2"
		
			If Not (rsEdit.EOF Or rsEdit.BOF) Then
				Do Until rsEdit.EOF
						%>
						<option value = '<%=rsEdit.Fields("GenericSetting_DPA_")%>'><%=rsEdit.Fields("GenericSettingDescription")%></option>
						<%
					rsEdit.MoveNext
				Loop
			End If
			%>
		</select>
	</td>
  </tr> 
   
  <tr>
    <td width="30%" >Client Statement</td>
    <td width="70%" ><input type = 'checkbox' name ='chkClientStatement' id = 'chkClientStatement' <%=ClientStatement%>></td>
  </tr>  
  
  <tr>
    <td width="30%" >Client Contracts</td>
    <td width="70%" ><input type = 'checkbox' name ='chkClientContract' id = 'chkClientContract' <%=ClientContract%>></td>
  </tr> 
  
  <tr>
    <td width="30%" >Client Contracts Compounded</td>
    <td width="70%" ><input type = 'checkbox' name ='chkClientContractCompounded' id = 'chkClientContractsCompounded' <%=ClientContractCompounded%>></td>
  </tr>  
  
  <tr>
    <td width="30%">Holdings Valuation</td>
    <td width="70%" ><input type = 'checkbox' name ='chkHoldingsValuation' id = 'chkHoldingsValuation' <%=HoldingsValuation%>></td>
  </tr>  
  
  <tr>
    <td width="100%" colspan="2" align=right>
		<BR>
		<BR>		
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save " onclick = "AllowedNavigation()">
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
