<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add Agent</title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>


<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<script language="JavaScript" src="CALENDAR/calendar.js"></script>

<!--END CALENDAR -->

<script >
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
		function hideButton()
		{
			document.getElementById('hide').style.display='none';		 
		}

</script>

</head>

<body Class="Dialog">
<!--#include virtual="libroutines.asp"-->
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<%
	
	Dim action
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim UserId
	
	action = ucase(Request.Form("action"))
	UserId=Session("UserID")
	
	if action = "EXECUTE" then
		Dim buttonAction
		Dim reloadRequired
		
		reloadRequired = false
		buttonAction = Trim(Ucase(Request.Form("cmdAdd")))
		if buttonAction = "SAVE" then
				Dim branch
				Dim commission
				Dim residency
				Dim gender
				Dim addr
				Dim bDate
				Dim cell
				Dim email
				Dim contact
				Dim office
				Dim home
				Dim fax
				Dim idPass
				Dim photo
				Dim signature
				Dim gen1
				Dim gen2
				Dim gen3
				Dim OpeningBal
       
				OpeningBal = Request.Form("txtOpeningBal")		        
				branch = Request.Form("cboBranch")
				commission = Request.Form("cboCommission")
				gender = Request.Form("cboGender")
				residency = Request.Form("cboResidency")
				name = Request.Form("txtName")
				addr = Request.Form("txtAddr")
				bDate = Request.Form("txtBDate")
				cell = Request.Form("txtCellTel")
				contact = Request.Form("txtContact")
				email = Request.Form("txtEmail")
				office = Request.Form("txtOfficeTel")
				home = Request.Form("txtHomeTel")
				fax = Request.Form("txtFax")
				idPass = Request.Form("txtIDPass")
				photo = Request.Form("txtPhoto")
				signature = Request.Form("txtSignature")
				gen1 = Request.Form("cboGenericSetting1")
				gen2 = Request.Form("cboGenericSetting2")
				gen3 = Request.Form("cboGenericSetting3")

				'validate Branch
				If Trim(branch) = "" Then%>
						<script language = 'vbscript'>
                				ShowMessage "Please specify the Branch"
                				
						</script>
						<% response.end
				End If
				'validate Commission Types
				If Trim(commission) = "" Then%>
						<script language = 'vbscript'>
                				ShowMessage "Please specify the Commission "
                				
						</script>
						<% response.end
				End If
				'validate Gender
				If Trim(Gender) = "" Then
						Gender = "NULL"
				End If
				'validate Residency
				If Trim(Residency) = "" Then%>
						<script language = 'vbscript'>
                				ShowMessage "Please specify the Residency"
                				
						</script>
						<% response.end
				End If
				'validate Name
				If Trim(Name) = "" Then%>
						<script language = 'vbscript'>
                				ShowMessage "Please specify the Name"
                				
						</script>
						<% response.end
				End If
				'validate size of Address
				If Len(Addr) > 100 Then%>
						<script language = 'vbscript'>
						ShowMessage "Address can only be 100 characters in length"
						
						</script>
						<% response.end
				End If
				'validate size of Cell Phone
				If Len(Cell) > 100 Then%>
						<script language = 'vbscript'>
						ShowMessage "Cell Phone can only be 100 characters in length"
						
						</script>
						<% response.end
				End If
				'validate size of Contact Name
				If Len(Contact) > 100 Then%>
						<script language = 'vbscript'>
						ShowMessage "Contact Name can only be 100 characters in length"
						
						</script>
						<% response.end
				End If
				'validate size of Email
				If Len(Email) > 100 Then%>
						<script language = 'vbscript'>
						ShowMessage "Email can only be 100 characters in length"
						
						</script>
						<% response.end
				End If
				'validate size of Fax
				If Len(Fax) > 100 Then%>
						<script language = 'vbscript'>
						ShowMessage "Fax can only be 100 characters in length"
						
						</script>
						<% response.end
				End If
				'validate Generic1
				If Trim(gen1) = "" Then
						gen1 = "NULL"
				End If
				'validate Generic2
				If Trim(gen2) = "" Then
						gen2 = "NULL"
				End If
				'validate Generic3
				If Trim(gen3) = "" Then
						gen3 = "NULL"
				End If
				'validate size of Home Phone
				If Len(home) > 100 Then%>
						<script language = 'vbscript'>
						ShowMessage "Home Phone can only be 100 characters in length"
						
						</script>
						<% response.end
				End If
				
				'validate size of ID or Passport
				If Len(IDPass) > 100 Then%>
						<script language = 'vbscript'>
						ShowMessage "ID or Passport can only be 100 characters in length"
						
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
				'validate size of Office Phone
				If Len(Office) > 100 Then%>
						<script language = 'vbscript'>
						ShowMessage "Office Phone can only be 100 characters in length"
						
						</script>
						<% response.end
				End If
				
				If Not IsNumeric(OpeningBal) Then%>
						<script language = 'vbscript'>
						ShowMessage "Opening Balance can only be numeric"						
						</script>
						<% response.end
				End If
				
				If Photo <> "" Then
					photoNew = Photo
				Else
					photoNew = "/Data/Photos/_blank.jpg"
				End If
				
				If signature <> "" Then
					sigNew = signature
				Else
					sigNew = "/Data/Photos/_blank.jpg"
				End If
				
				'save data
				sqlStr = "INSERT INTO [Agent] (AgentAddr,AgentBDate,AgentCellTel,AgentContact,AgentEmail" & _
						",AgentFax,GenericSetting_DPA_,GenericSetting_DPA_2,GenericSetting_DPA_3,AgentHomeTel,AgentIDPass,AgentName,AgentOfficeTel" & _
						",AgentPhoto,AgentSignature,Agent_DPA_,Branch_DPA_,Commission_DPA_,Gender_DPA_,Residency_DPA_,AgentOpeningBal,ChangedBy" & _
						") SELECT " & "'" & addr & "'" & " as AgentAddr," & "#" & FormatDate(bDate) & "#" & " as AgentBDate" & _
						"," & "'" & cell & "'" & " as AgentCellTel" & _
						"," & "'" & contact & "'" & " as AgentContact" & _
						"," & "'" & email & "'" & " as AgentEmail," & "'" & fax & "'" & " as AgentFax" & _
						"," & " " & gen1 & " " & " as GenericSetting_DPA_" & _
						"," & " " & gen2 & " " & " as GenericSetting_DPA_2" & _
						"," & " " & gen3 & " " & " as GenericSetting_DPA_3" & _
						"," & "'" & home & "'" & " as AgentHomeTel" & _
						"," & "'" & idPass & "'" & " as AgentIDPass," & "'" & name & "'" & " as AgentName" & _
						"," & "'" & office & "'" & " as AgentOfficeTel" & _
						"," & "'" & photoNew & "'" & " as AgentPhoto,'" & sigNew & "' as AgentSignature," & " " & "iif(isnull(max([Agent_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Agent'),max([Agent_DPA_]) + 1)" & " " & " as Agent_DPA_" & _
						"," & " " & branch & " " & " as Branch_DPA_" & _
						"," & " " & commission & " " & " as Commission_DPA_" & _
						"," & " " & gender & " " & " as Gender_DPA_" & _
						"," & " " & residency & " " & " as Residency_DPA_" & _
						"," & " " & OpeningBal & " " & " as AgentOpeningBal" & _
						"," & " " & UserId & " " & " as ChangedBy" & _
						" FROM [Agent]"
				
				Set conn = GetActiveConnection("KBroker")
		        	        
				conn.BeginTrans
						conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
				conn.CommitTrans
				conn.Close
				Set conn = Nothing
				WritefraEnabledDialogCloseScript
				Response.End
			end if
			Dim clientCode
        
			clientCode = "var validNavigate = true;" & chr(13)
			clientCode = clientCode & "window.parent.close()" & chr(13)%>
			<script>
					<%=clientCode%>
			</script>
			<%
			response.End
   	end If
	Set conn = GetActiveConnection("KBroker")
%>


<form name = 'frmAddAgent' method = 'post' action = 'AddAgent.asp' id = "frmMain">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
	<tr>
		<td width="17%">Name</td>
		<td width="33%"><input type = 'text' name ='txtName' id = 'txtName' size="25"></td>
		<td width="17%">ID/Passport</td>
		<td width="33%"><input type = 'text' name ='txtIDPass' id = 'txtIDPass' size="28"></td>
	</tr>
	<tr>
		<td width="17%">Contact Name</td>
		<td width="33%"><input type = 'text' name ='txtContact' id = 'txtContact' size="25"></td>
		<td width="17%">Date of Birth</td>
		<td width="33%">
		<SCRIPT language="JavaScript">
			var cal=new ctlSpiffyCalendarBox("cal", "frmAddAgent", "txtBDate","cmdDate","<%= FormatDate(Date) %>",1);
			cal.writeControl();
		</SCRIPT>
		</td>
	</tr>
	<tr>
		<td width="17%">Gender</td>
		<td width="33%">
		<select name = 'cboGender' id = 'cboGender' size="1">
			<option selected value = ''></option>
			<%
		    sqlStr = "SELECT * FROM [GenderList]"
		    Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		    If Not (rs.EOF Or rs.BOF) Then
		            rs.MoveFirst
		            Do Until rs.EOF%>
							<option value = '<%=rs.Fields("Gender_DPA_")%>'><%=rs.Fields("GenderGender")%></option>
		                    <%rs.MoveNext
		            Loop
		    End If
			%>
		</select>
		</td>
		<td width="17%">Cell Phone</td>
		<td width="33%"><input type = 'text' name ='txtCellTel' id = 'txtCellTel' size="25"></td>
	</tr>
	<tr>
		<td width="17%">Office Phone</td>
		<td width="33%"><input type = 'text' name ='txtOfficeTel' id = 'txtOfficeTel' size="25"></td>
		<td width="17%">Home Phone</td>
		<td width="33%"><input type = 'text' name ='txtHomeTel' id = 'txtHomeTel' size="25"></td>
	</tr>
	<tr>
		<td width="17%">E-mail</td>
		<td width="33%"><input type = 'text' name ='txtEmail' id = 'txtEmail' size="25"></td>
		<td width="17%">Fax</td>
		<td width="33%"><input type = 'text' name ='txtFax' id = 'txtFax' size="25"></td>
	</tr>
	<tr>
		<td width="17%">Branch</td>
		<td width="33%">
		<select name = 'cboBranch' id = 'cboBranch' size="1">
			<option selected value = ''></option>
			<%
		    sqlStr = "SELECT * FROM [BranchList]"
		    Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		    If Not (rs.EOF Or rs.BOF) Then
		            rs.MoveFirst
		            Do Until rs.EOF
		                    if cbool(rs.Fields("DefaultSelection")) then%>
									<option selected value = '<%=rs.Fields("Branch_DPA_")%>'><%=rs.Fields("BranchName")%></option>
		                    <%else%>
									<option value = '<%=rs.Fields("Branch_DPA_")%>'><%=rs.Fields("BranchName")%></option>
		                    <%end if
		                    rs.MoveNext
		            Loop
		    End If
			%>
		</select>
		</td>
		<td width="17%">Commission</td>
		<td width="33%">
		<select name = 'cboCommission' id = 'cboCommission' size="1">
			<option selected value = ''></option>
			<%
		    sqlStr = "SELECT * FROM [CommissionList]"
		    Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		    If Not (rs.EOF Or rs.BOF) Then
		            rs.MoveFirst
		            Do Until rs.EOF
		                    if cbool(rs.Fields("DefaultSelection")) then%>
									<option selected value = '<%=rs.Fields("Commission_DPA_")%>'><%=rs.Fields("CommissionRate")%></option>
		                    <%else%>
									<option value = '<%=rs.Fields("Commission_DPA_")%>'><%=rs.Fields("CommissionRate")%></option>
		                    <%end if
		                    rs.MoveNext
		            Loop
		    End If
			%>
		</select>
		</td>
	</tr>
	<tr>
		<td width="17%">Residency</td>
		<td width="33%">
		<select name = 'cboResidency' id = 'cboResidency' size="1">
			<option selected value = ''></option>
			<%
		    sqlStr = "SELECT * FROM [ResidencyList]"
		    Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		    If Not (rs.EOF Or rs.BOF) Then
		            rs.MoveFirst
		            Do Until rs.EOF
		                    if cbool(rs.Fields("DefaultSelection")) then%>
									<option selected value = '<%=rs.Fields("Residency_DPA_")%>'><%=rs.Fields("ResidencyName")%></option>
		                    <%else%>
									<option value = '<%=rs.Fields("Residency_DPA_")%>'><%=rs.Fields("ResidencyName")%></option>
		                    <%end if
		                    rs.MoveNext
		            Loop
		    End If
			%>
		</select>
		</td>
		<td width="12%">Generic 1</td>
		    <td width="29%"><select name = 'cboGenericSetting1' id = 'cboGenericSetting1' size="1">
							<option value = ''></option>
		<%
		        Dim genericRS
		        sqlStr = "SELECT * FROM [GenericSettingList] WHERE EntityType_DPA_ = 2 Order By GenericSettingDescription"
		        Set genericRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		        genericRS.Filter = "Generic_DPA_ = 1"
		        If Not (genericRS.EOF Or genericRS.BOF) Then
		                Do Until genericRS.EOF%>
								<option value = '<%=genericRS.Fields("GenericSetting_DPA_")%>'><%=genericRS.Fields("GenericSettingDescription")%></option>
		                        <%genericRS.MoveNext
		                Loop
		        End If
		%>

		    </select></td>
	</tr>
	<tr>
		<td width="17%">Opening Balance</td>
		<td width="33%"><input type = 'text' name ='txtOpeningBal' STYLE="TEXT-ALIGN: RIGHT;" id = "txtOpeningBal" size="20" value="0"></td>
		<td width="12%">Generic 2</td>
		    <td width="29%"><select name = 'cboGenericSetting2' id = 'cboGenericSetting2' size="1">
							<option value = ''></option>
		<%
		        genericRS.Filter = "Generic_DPA_ = 2"
		        If Not (genericRS.EOF Or genericRS.BOF) Then
		                Do Until genericRS.EOF%>
								<option value = '<%=genericRS.Fields("GenericSetting_DPA_")%>'><%=genericRS.Fields("GenericSettingDescription")%></option>
		                        <%genericRS.MoveNext
		                Loop
		        End If
		%>

		    </select></td>
	</tr>
	<tr>
		<td width="17%">Address</td>
		<td width="33%"><textarea rows=3 name ='txtAddr' id = 'txtAddr' size="35" cols="26"></textarea></td>
		<td width="12%" Valign="top">Generic 3</td>
    <td width="29%" Valign="top"><select name = 'cboGenericSetting3' id = 'cboGenericSetting3' size="1">
					<option value = ''></option>
<%
        genericRS.Filter = "Generic_DPA_ = 3"
        If Not (genericRS.EOF Or genericRS.BOF) Then
                Do Until genericRS.EOF%>
						<option value = '<%=genericRS.Fields("GenericSetting_DPA_")%>'><%=genericRS.Fields("GenericSettingDescription")%></option>
                        <%genericRS.MoveNext
                Loop
        End If
%>

    </select></td>
	</tr>
</table>

<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
	<td colspan="4">
		<table width="100%" border="0">
			<tr>
  				<td width="50%">Photo</td>
				<td width="50%">Signature</td>
			</tr>
			<tr>
  				<td width="50%">
					<input type="hidden" name="txtPhoto">
					<IFRAME FRAMEBORDER=0 SCROLLING=NO SRC="upload.asp?filetext=txtPhoto" width="170px" height="170px" tabIndex='-1'></IFRAME>
				</td>
  				<td width="50%">
					<input type="hidden" name="txtSignature">
					<IFRAME FRAMEBORDER=0 SCROLLING=NO SRC="upload.asp?filetext=txtSignature" width="170px" height="170px" tabIndex='-1'></IFRAME>
				</td>
			</tr>
		</table>
	</td>
  </tr>
  <tr>
    <td width="100%" colspan=4 align=right>
		<b id="hide" name="hide"> <input type = 'submit' Class=Buttons name ='cmdAdd' id = "cmdAdd" value=" Save " onclick = "AllowedNavigation();hideButton()"></b>
   		<input type = 'button' Class=Buttons name ='cmdClose' id = 'cmdClose' value=" Cancel " OnClick="JavaScript: window.self.close();">
		<input type = 'hidden' name ='action' id = "action" value="Execute">
	</td>
  </tr>
</table>

</form>
</body>