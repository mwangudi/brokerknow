<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit Agent</title>
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
 <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css"> 
 <script language="JavaScript" src="../scripts/common.js"></script>


  <!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="../CALENDAR/calendar.css">

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
		
		function DoInit() {
			if (window.dialogArguments != null){
				window.name = "editWindow";
			}
    
			else{
				window.close();
			}    
    
	}

	function GetEditAction(selID){
		if (selID!="")
			{
				//give the popup a handle to the list window for refresh purpose
				window.self.opener = window.dialogArguments.opener;
				
				//prepare list window for submit
				var targetPage = window.dialogArguments.opener.document.all.item("frmMain").elements("EditPage").value
				
				//window.dialogArguments.opener.document.all.item("frmMain").action = targetPage;
				window.dialogArguments.opener.document.all.item("frmMain").target = self;
				//document.write (window.dialogArguments.opener.document.all.item("frmMain").outerHTML);
				window.dialogArguments.opener.document.all.item("frmMain").submit();
				//window.self.location.replace(ActionWin.location);
			}
		else{window.close();}	
	 }
</script>

</head>

<body Class="Dialog" OnLoad="JavaScript: DoInit();">
<!--#include file="../libroutines.asp"-->

<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<%
	
	Dim action
	Dim conn 
    Dim sqlStr
    Dim rs
    Dim ID
    Dim rsEdit
	Dim UserId
	
	action = ucase(Request.Form("action"))
	ID = Request("ID")
	UserId=Session("UserID")
	
		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for editing"
                		
                </script>
                <% response.end
        End If

	if action = "EXECUTE" then
		Dim buttonAction
		Dim reloadRequired
		
		reloadRequired = false
		buttonAction = Ucase(Request.Form("cmdSave"))
		
		if buttonAction = "CANCEL" then
				
		else
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
				
				Set conn = GetActiveConnection("KBroker")
				
				'save data
				If Photo <> "" Then
					photoNew = photo
				Else
					photoNew = "/Data/Photos/_blank.jpg"
				End If
				
				'save data
				If signature <> "" Then
					sigNew = signature
				Else
					sigNew = "/Data/Photos/_blank.jpg"
				End If
				
				sqlStr = "UPDATE [Agent] SET AgentAddr = " & "'" & addr & "'" & ",AgentBDate = " & "#" & FormatDate(bDate) & "#" & "" & _
						",AgentCellTel = " & "'" & cell & "'" & ",AgentContact = " & "'" & contact & "'" & "" & _
						",AgentEmail = " & "'" & email & "'" & ",AgentFax = " & "'" & fax & "'" & "" & _
						",GenericSetting_DPA_ = " & " " & gen1 & " " & ",GenericSetting_DPA_2 = " & " " & gen2 & " " & "" & _
						",GenericSetting_DPA_3 = " & " " & gen3 & " " & "" & _
						",AgentHomeTel = " & "'" & home & "'" & ",AgentIDPass = " & "'" & idPass & "'" & "" & _
						",AgentName = " & "'" & name & "'" & ",AgentOfficeTel = " & "'" & office & "'" & "" & _
						",AgentPhoto = " & "'" & photoNew & "'" & ",AgentSignature = '" & sigNew & "',Branch_DPA_ = " & " " & branch & " " & ",Commission_DPA_ = " & " " & commission & " " & "" & _
						",Gender_DPA_ = " & " " & gender & " " & ",Residency_DPA_ = " & " " & residency & " " & "" & _
						",AgentOpeningBal = " & " " & OpeningBal & " " & "" & _
                        ",ChangedBy = " & " " & UserId & " " & "" & _
						",TimeChanged = " & "'" & Now() & "'" & "" & _
						" WHERE Agent_DPA_  = " & ID
		                
				conn.BeginTrans
						conn.Execute SQLServerFormat(HandleQuote(sqlStr))
				conn.CommitTrans
				
				conn.Close
				Set conn = Nothing
				WritefraEnabledDialogCloseScript
				Response.End
        end if
        Dim clientCode
        
        clientCode = "var validNavigate = true;" & chr(13)
		clientCode = clientCode & "window.self.close();" & chr(13)%>
		<script>
			<%=clientCode%>
		</script>
		<%
		response.End
   	end If

    Set conn = GetActiveConnection("KBroker")
    Dim timeNow
    timeNow = Timer
        
    sqlStr = "SELECT AgentAddr,AgentBDate,AgentOpeningBal,AgentCellTel,AgentContact,AgentEmail" & _
            ",AgentFax,GenericSetting_DPA_,GenericSetting_DPA_2,GenericSetting_DPA_3,AgentHomeTel,AgentIDPass,AgentName,AgentOfficeTel" & _
            ",AgentPhoto,AgentSignature,Agent_DPA_,Agent.Branch_DPA_,Agent.Commission_DPA_,Agent.Gender_DPA_,Agent.Residency_DPA_ " & _
            " FROM  GenderList RIGHT OUTER JOIN " & _
            "        ResidencyList INNER JOIN " & _
            "        CommissionList INNER JOIN " & _
            "        BranchList INNER JOIN " & _
            "        Agent ON BranchList.Branch_DPA_ = Agent.Branch_DPA_ ON CommissionList.Commission_DPA_ = Agent.Commission_DPA_ ON " & _ 
            "        ResidencyList.Residency_DPA_ = Agent.Residency_DPA_ ON GenderList.Gender_DPA_ = Agent.Gender_DPA_ " & _
            "        WHERE Agent.Agent_DPA_  = " & ID
    sqlStr = SQLServerFormat(HandleQuote(sqlStr))
        
    Set rs = conn.Execute(sqlStr)
    Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
    If rs.EOF Or rs.BOF Then%>
            <script language = 'vbscript'>
            		ShowMessage "The selected Agent cannot be retrieved for editing"
            </script>
            <% response.end
    End If
%>

<form name = 'frmEditAgent' method = 'post' action = 'EditAgent.asp' >
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
	<tr>
		<td width="17%">Name</td>
		<td width="33%">
        <input type = 'text' name ='txtName' id = 'txtName' value = '<%=rs.Fields("AgentName")%>' size="20"></td>
		<td width="17%">ID/Passport</td>
		<td width="33%"><input type = 'text' name ='txtIDPass' id = 'txtIDPass' size="20" value = '<%=rs.Fields("AgentIDPass")%>'></td>
	</tr>
	<!--CALENDAR -->
	<script language="JavaScript" src="../CALENDAR/calendar.js"></script>
	<SCRIPT language="JavaScript">
		var cal=new ctlSpiffyCalendarBox("cal", "frmEditAgent", "txtBDate","cmdDate","<%= FormatDate(rs.Fields("AgentBDate")) %>",1);
	</SCRIPT>
	<!--END CALENDAR -->
	<tr>
		<td width="17%">Contact Name</td>
		<td width="33%"><input type = 'text' name ='txtContact' id = 'txtContact' size="20" value = '<%=rs.Fields("AgentContact")%>'></td>
		<td width="17%">Date of Birth</td>
		<td width="33%"><SCRIPT language="JavaScript">cal.writeControl();</SCRIPT></td>
	</tr>
	<tr>
		<td width="17%">Gender</td>
		<td width="33%">
		<select name = 'cboGender' id = 'cboGender' size="1">
    	<option selected value = ''></option>
		<%
        sqlStr = "SELECT * FROM [GenderList]"
        Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rsEdit.EOF Or rsEdit.BOF) Then
            rsEdit.MoveFirst
            Do Until rsEdit.EOF
               if rsEdit.Fields("Gender_DPA_") = rs.Fields("Gender_DPA_") Then%>
               	<option selected value = '<%=rsEdit.Fields("Gender_DPA_")%>'><%=rsEdit.Fields("GenderGender")%></option>
				<%else%>
				<option value = '<%=rsEdit.Fields("Gender_DPA_")%>'><%=rsEdit.Fields("GenderGender")%></option>
				<%end if
				rsEdit.MoveNext
            Loop
    End If
		%>
		</select>
		</td>
		<td width="17%">Cell Phone</td>
		<td width="33%"><input type = 'text' name ='txtCellTel' id = 'txtCellTel' size="20" value = '<%=rs.Fields("AgentCellTel")%>'></td>
	</tr>
	<tr>
		<td width="17%">Office Phone</td>
		<td width="33%"><input type = 'text' name ='txtOfficeTel' id = 'txtOfficeTel' size="20" value = '<%=rs.Fields("AgentOfficeTel")%>'></td>
		<td width="17%">Home Phone</td>
		<td width="33%"><input type = 'text' name ='txtHomeTel' id = 'txtHomeTel' size="20" value = '<%=rs.Fields("AgentHomeTel")%>'></td>
	</tr>
	<tr>
		<td width="17%">E-mail</td>
		<td width="33%"><input type = 'text' name ='txtEmail' id = 'txtEmail' size="20" value = '<%=rs.Fields("AgentEmail")%>'></td>
		<td width="17%">Fax</td>
		<td width="33%"><input type = 'text' name ='txtFax' id = 'txtFax' size="20" value = '<%=rs.Fields("AgentFax")%>'></td>
	</tr>
	<tr>
		<td width="17%">Branch</td>
		<td width="33%">
		<select name = 'cboBranch' id = 'cboBranch' size="1">
		<%
        sqlStr = "SELECT * FROM [BranchList]"
        Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rsEdit.EOF Or rsEdit.BOF) Then
            rsEdit.MoveFirst
            Do Until rsEdit.EOF
            	   if rsEdit.Fields("Branch_DPA_") = rs.Fields("Branch_DPA_") Then%>
            	   	<option selected value = '<%=rsEdit.Fields("Branch_DPA_")%>'><%=rsEdit.Fields("BranchName")%></option>
            	   <%else%>
                   <option value = '<%=rsEdit.Fields("Branch_DPA_")%>'><%=rsEdit.Fields("BranchName")%></option>
                <%end if
				   rsEdit.MoveNext
            Loop
        End If
		%>
		</select>
		</td>
		<td width="17%">Commission</td>
		<td width="33%">
		<select name = 'cboCommission' id = 'cboCommission' size="1">
		<%
        sqlStr = "SELECT * FROM [CommissionList]"
        Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rsEdit.EOF Or rsEdit.BOF) Then
                rsEdit.MoveFirst
                Do Until rsEdit.EOF
                		if rsEdit.Fields("Commission_DPA_") = rs.Fields("Commission_DPA_") Then%>
                			<option selected value = '<%=rsEdit.Fields("Commission_DPA_")%>'><%=rsEdit.Fields("CommissionRate")%></option>
                		<%else%>
                        <option value = '<%=rsEdit.Fields("Commission_DPA_")%>'><%=rsEdit.Fields("CommissionRate")%></option>
                     <%end if
						rsEdit.MoveNext
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
		  Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		  If Not (rsEdit.EOF Or rsEdit.BOF) Then
		          rsEdit.MoveFirst
		          Do Until rsEdit.EOF
		          		if rsEdit.Fields("Residency_DPA_") = rs.Fields("Residency_DPA_") Then%>
		          			<option selected value = '<%=rsEdit.Fields("Residency_DPA_")%>'><%=rsEdit.Fields("ResidencyName")%></option>
		          		<%else%>
		                  <option value = '<%=rsEdit.Fields("Residency_DPA_")%>'><%=rsEdit.Fields("ResidencyName")%></option>
		               <%end if
		  				rsEdit.MoveNext
		          Loop
		  End If
		  %>
		</select>
		</td>
		<td width="22%">Generic 1</td>
		    <td width="22%"><select name = 'cboGenericSetting1' id = 'cboGenericSetting1' size="1">
								<option selected value = ''></option>
		<%
		        sqlStr = "SELECT * FROM [GenericSettingList] WHERE EntityType_DPA_ = 2 Order By GenericSettingDescription"
		        Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		        rsEdit.Filter = "Generic_DPA_ = 1"
		        If Not (rsEdit.EOF Or rsEdit.BOF) Then
		                Do Until rsEdit.EOF
		                		if rsEdit.Fields("GenericSetting_DPA_") = rs.Fields("GenericSetting_DPA_") Then%>
		                			<option selected value = '<%=rsEdit.Fields("GenericSetting_DPA_")%>'><%=rsEdit.Fields("GenericSettingDescription")%></option>
		                		<%else%>
		                        <option value = '<%=rsEdit.Fields("GenericSetting_DPA_")%>'><%=rsEdit.Fields("GenericSettingDescription")%></option>
		                     <%end if
								rsEdit.MoveNext
		                Loop
		        End If
		 %>

		    </select></td>
	</tr>
	<tr>
		<td width="17%">Opening Balance</td>
		<td width="33%"><input type = 'text' name ='txtOpeningBal' STYLE="TEXT-ALIGN: RIGHT;" id = "txtOpeningBal" size="20" value = '<%=rs.Fields("AgentOpeningBal")%>'></td>
		<td width="22%">Generic 2</td>
		    <td width="22%"><select name = 'cboGenericSetting2' id = 'cboGenericSetting2' size="1">
							<option selected value = ''></option>
		<%
		        rsEdit.Filter = "Generic_DPA_ = 2"
		        If Not (rsEdit.EOF Or rsEdit.BOF) Then
		                Do Until rsEdit.EOF
		                		if rsEdit.Fields("GenericSetting_DPA_") = rs.Fields("GenericSetting_DPA_2") Then%>
		                			<option selected value = '<%=rsEdit.Fields("GenericSetting_DPA_")%>'><%=rsEdit.Fields("GenericSettingDescription")%></option>
		                		<%else%>
		                        <option value = '<%=rsEdit.Fields("GenericSetting_DPA_")%>'><%=rsEdit.Fields("GenericSettingDescription")%></option>
		                     <%end if
								rsEdit.MoveNext
		                Loop
		        End If
		 %>

		    </select></td>
	</tr>
	<tr>
		<td width="17%">Address</td>
		<td width="33%"><textarea rows=3 name ='txtAddr' id = 'txtAddr' size="35" cols="26"><%=rs.Fields("AgentAddr")%></textarea></td>
		<td width="22%" Valign="top">Generic 3</td>
		    <td width="22%" Valign="top"><select name = 'cboGenericSetting3' id = 'cboGenericSetting3' size="1">
							<option selected value = ''></option>
		<%
		        rsEdit.Filter = "Generic_DPA_ = 3"
		        If Not (rsEdit.EOF Or rsEdit.BOF) Then                
		                Do Until rsEdit.EOF
		                		if rsEdit.Fields("GenericSetting_DPA_") = rs.Fields("GenericSetting_DPA_3") Then%>
		                			<option selected value = '<%=rsEdit.Fields("GenericSetting_DPA_")%>'><%=rsEdit.Fields("GenericSettingDescription")%></option>
		                		<%else%>
		                        <option value = '<%=rsEdit.Fields("GenericSetting_DPA_")%>'><%=rsEdit.Fields("GenericSettingDescription")%></option>
		                     <%end if
								rsEdit.MoveNext
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
					<input type="hidden" name="txtPhoto" value="<%= rs.Fields("AgentPhoto") %>">
					<IFRAME FRAMEBORDER=0 SCROLLING=NO SRC="upload.asp?filetext=txtPhoto" width="170px" height="170px" tabIndex='-1'></IFRAME>
				</td>
  				<td width="50%">
					<input type="hidden" name="txtSignature" value="<%= rs.Fields("AgentSignature") %>">
					<IFRAME FRAMEBORDER=0 SCROLLING=NO SRC="upload.asp?filetext=txtSignature" width="170px" height="170px" tabIndex='-1'></IFRAME>
				</td>
			</tr>
		</table>
	</td>
  </tr>
	<tr>
    <td width="100%" colspan="4" align=right>
		<BR>
		<input type = 'submit' Class=Buttons name ='cmdSave' id = 'cmdSave' value=" Save " onclick = "AllowedNavigation()">
		<input type = 'button' Class=Buttons name ='cmdClose' id = 'cmdClose' value=" Cancel " OnClick="JavaScript: window.self.close();">
    	<input type = 'hidden' name ='action' id = "action" value="Execute">
    	<input type = 'hidden' name ='ID' id = "ID" value="<%=ID%>">
	
    </td>
	  </tr>
</table>
</form>
</body>