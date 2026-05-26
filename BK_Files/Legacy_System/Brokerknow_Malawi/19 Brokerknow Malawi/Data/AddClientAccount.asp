<!--#include virtual="libroutines.asp"-->
<%
	
	Dim action
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim rsUnique
	Dim guidStr 
	Dim guid
	
	UserId=Session("UserID")
	
	action = ucase(Request.Form("action"))		
	
	if action = "EXECUTE" then
		Dim buttonAction
		Dim reloadRequired		

		reloadRequired = false
		buttonAction = Trim(Ucase(Request.Form("cmdAdd")))
		
						
		if buttonAction = "SAVE" then						
				Dim agent
				Dim classType
				Dim owner
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
				Dim vip
				Dim gen1
				Dim gen2
				Dim gen3
				Dim pAddr
				dim comment
				Dim bnkBranch
				Dim accNum
				Dim accName
				Dim bankAccSpecified
				Dim OpeningBal
				Dim UniqueSql
				Dim IsCustodian
				bankAccSpecified = false
				
				bnkBranch = Request.Form("cboBnkBranch")
				accNum = Request.Form("txtNumber")
				accName = Request.Form("txtAccName")
				comment = Request.Form("txtComment")
				agent = Request.Form("cboAgent") 
				classType = Request.Form("cboClass")
				owner = Request.Form("cboOwner")
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
				vip = Request.Form("cboVIP")
				gen1 = Request.Form("txtGeneric1")
				gen2 = Request.Form("txtGeneric2")
				pAddr = Request.Form("txtPAddr")
				OpeningBal = Request.Form("txtOpeningBal")
				CreditLimit = Request.Form("txtCreditLimit") 
				CDSNo = Request.Form("txtCDSNo")
				gen1 = Request.Form("cboGenericSetting1")
				gen2 = Request.Form("cboGenericSetting2")
				gen3 = Request.Form("cboGenericSetting3") 
				IsCustodian = Request.Form("custodianStatus") 

				
				'validate CDS
			
				If Trim(CDSNo) = "" Then%>
						<script language = 'vbscript'>
                				ShowMessage "Please specify the CDS Number"
                				
						</script>
						<% response.end
				End If				
		       
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

				'validate Agent
				If Trim(Agent) = "" Then
						Agent = "NULL"
				End If
				'validate Client Class
				If Trim(classType) = "" Then%>
						<script language = 'vbscript'>
                				ShowMessage "Please specify the Class"
                				
						</script>
						<% response.end
				End If
				'validate Owner
				If Trim(Owner) = "" Then'% >
						'<script language = 'vbscript'>
                				'ShowMessage "Please specify the Owner"
                				
						'</script>
						'< % response.end
						Owner = "NULL"
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
				If Len(Addr) > 500 Then%>
						<script language = 'vbscript'>
						ShowMessage "Address can only be 500 characters in length"
						
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
				
				If Not IsNumeric(OpeningBal) Then%>
						<script language = 'vbscript'>
						ShowMessage "Opening Balance can only be numeric"						
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
				'validate size of Physical Address
				If Len(pAddr) > 500 Then%>
						<script language = 'vbscript'>
						ShowMessage "Physical Address can only be 500 characters in length"
						
						</script>
						<% response.end
				End If
				'validate size of Home Phone
				If Len(home) > 100 Then%>
						<script language = 'vbscript'>
						ShowMessage "Home Phone can only be 100 characters in length"
						
						</script>
						<% response.end
				End If
				
				'Make sure user enters ID No/Passport No/Cert
				if Len(Trim(idPass)) = 0 then
					%>
					<script language = 'vbscript'>
						ShowMessage "You must fill the ID No/Passport/Cert field"
					</script>
					<%
					Response.End
				end if
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
				
				'check for bank account specification
				'validate Bank Branch
				 If Trim(bnkBranch) <> "" Then
						 bankAccSpecified = true
						 'validate Account Number
						 If Trim(accNum) = "" Then%>
						         <script language = 'vbscript'>
						         		ShowMessage "Please specify the Bank Account Number"
						         		
						         </script>
						         <% response.end
						 End If
						'validate size of Account Number
						 If Len(accNum) > 100 Then%>
						         <script language = 'vbscript'>
						         ShowMessage "Account Number can only be 100 characters in length"
						         
						         </script>
						         <% response.end
						 End If
						 'validate size of Account Name
						 If Len(accName) > 100 Then%>
						         <script language = 'vbscript'>
						         ShowMessage "Account Name can only be 100 characters in length"
						         
						         </script>
						         <% response.end
						 End If	
				End If
			
				'save data
				If Photo <> "" Then
					photoNew = photo
				Else
					'default photo
					photoNew = "/Data/Photos/_blank.jpg"		
				End If
				
				If signature <> "" Then
					sigNew = signature
				Else
					'default photo
					sigNew = "/Data/Photos/_blank.jpg"		
				End If
				
				'Enforce ID and CDS uniqueness
				Set conn = GetActiveConnection("KBroker")
				UniqueSql="Select ClientIDPass,ClientCDSNo From Client"
				
				Set rsUnique = Server.CreateObject("ADODB.Recordset")
				rsUnique.CursorLocation = adUseClient 
				set rsUnique=Conn.execute(UniqueSql)
				
				if not(rsUnique.EOF and rsUnique.BOF) then
					Do while rsUnique.EOF=false
						if(Trim(idPass)=rsUnique("ClientIDPass")) then%>
						<script language = 'vbscript'>
						ShowMessage "Client ID Or PassPort Number is unique"						
						</script>
						<% response.end
						
						end if
						if(Trim(CDSNo)=rsUnique("ClientCDSNo")) then%>
						<script language = 'vbscript'>
						ShowMessage "Client CDS Number is unique"						
						</script>
						<% response.end
						
						end if
						
					rsUnique.MoveNext 
					loop
				end if
				
				'Save Data to the system
				set guid = server.createobject("NDUtils.CGUID")
				guidStr = guid.GenerateGUID	
						
				sqlStr = "INSERT INTO [Client] (ClientAddr,ClientBDate,ClientCellTel,ClientContact,ClientEmail" & _
						",ClientFax,ClientComment,GenericSetting_DPA_,GenericSetting_DPA_2,GenericSetting_DPA_3,ClientPAddr" & _
						",ClientHomeTel,ClientIDPass,ClientName,ClientOfficeTel,ClientPhoto,ClientSignature" & _
						",ClientVIP,Client_DPA_,Agent_DPA_,Branch_DPA_,Class_DPA_,Commission_DPA_,Gender_DPA_" & _
						",ClientOpeningBal, Owner_DPA_,Residency_DPA_,Client_EIT_, CreditLimit, ClientCDSNo,ChangedBy,IsCustodian) SELECT " & "'" & addr & "'" & " as ClientAddr," & "#" & FormatDate(bDate) & "#" & " as ClientBDate" & _
						"," & "'" & cell & "'" & " as ClientCellTel" & _
						"," & "'" & contact & "'" & " as ClientContact" & _
						"," & "'" & email & "'" & " as ClientEmail," & "'" & fax & "'" & " as ClientFax" & _
						"," & "'" & comment & "'" & " as ClientComment" & _
						"," & " " & gen1 & " " & " as GenericSetting_DPA_" & _
						"," & " " & gen2 & " " & " as GenericSetting_DPA_2" & _
						"," & " " & gen3 & " " & " as GenericSetting_DPA_3" & _
						"," & "'" & pAddr & "'" & " as ClientPAddr" & _
						"," & "'" & home & "'" & " as ClientHomeTel" & _
						"," & "'" & idPass & "'" & " as ClientIDPass," & "'" & name & "'" & " as ClientName" & _
						"," & "'" & Office & "'" & " as ClientOfficeTel" & _
						"," & "'" & photoNew & "'" & " as ClientPhoto, '" & sigNew & "' as ClientSignature," & " " & vip & " " & " as ClientVIP" & _ 
						"," & " " & "iif(isnull(max([Client_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Client'),max([Client_DPA_]) + 1)" & " " & " as Client_DPA_" & _
						"," & " " & agent & " " & " as Agent_DPA_" & _
						"," & " " & branch & " " & " as Branch_DPA_" & _
						"," & " " & classType & " " & " as Class_DPA_" & _
						"," & " " & commission & " " & " as Commission_DPA_" & _
						"," & " " & gender & " " & " as Gender_DPA_" & _
						"," & " " & OpeningBal & " " & " as ClientOpeningBal" & _						
						"," & " " & owner & " " & " as Owner_DPA_" & _
						"," & " " & residency & " " & " as Residency_DPA_" & _
						"," & "'" & guidStr & "'" & " as Client_EIT_" & _
						"," & "" & CreditLimit & "" & " as CreditLimit" & _						
						"," & "'" & CDSNo & "'" & " as ClientCDSNo" & _
						"," & "" & UserId & "" & " as ChangedBy" & _
						"," & "" & Iscustodian & "" & " as IsCustodian" & _
						" FROM [Client]"				
				
				
				conn.BeginTrans
				SQL = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
				
						conn.Execute SQL
						
						if bankAccSpecified then
								'obtain header key value
								sqlStr = "SELECT [Client_DPA_] FROM [Client] WHERE [Client_EIT_] = " & "'" & guidStr & "'"
	        
								Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
								If (rs.EOF Or rs.BOF) Then%>
										<script language = 'vbscript'>
								    			ShowMessage "A serious error has been encountered while saving the data. Try saving again"
								    			
										</script>
										<% response.end
								End If
								
								'save detail data
								sqlStr = "INSERT INTO [BankAcc] (BankAccNumber,BankAcc_DPA_,BankAccName,BnkBranch_DPA_,Client_DPA_) SELECT " & "'" & accNum & "'" & " as BankAccNumber" & _
										"," & " " & "iif(isnull(max([BankAcc_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'BankAcc'),max([BankAcc_DPA_]) + 1)" & " " & " as BankAcc_DPA_" & _
										"," & "'" & accName & "'" & " as BankAccName" & _
										"," & " " & bnkBranch & " " & " as BnkBranch_DPA_" & _
										"," & " " & rs.Fields("Client_DPA_") & " " & " as Client_DPA_" & _
										" FROM [BankAcc]"
	        
								conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
						end if
				conn.CommitTrans
				conn.Close
				Set conn = Nothing
				WritefraEnabledDialogCloseScript
				Response.End
			end if
			Dim clientCode
        
			clientCode = "var validNavigate = true;" & chr(13)
			%>
			<script>
				<%=clientCode%>
			</script>
			<%
			response.End
   	end If
    Set conn = GetActiveConnection("KBroker")
%>

<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add Client</title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<script language="JavaScript" src="CALENDAR/calendar.js"></script>
<!--END CALENDAR -->

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
		function  UpdateCustodianStatus(theChk)
		{
			var holdVal = "0"; //Custodian
			if (theChk.checked)
			{
				holdVal = "1";//Custodian
			}
				
			document.frmMain.elements("custodianStatus").value = holdVal;
		}
</script>
</head>

<body Class="Dialog">
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>

<SCRIPT language="JavaScript">
	var cal=new ctlSpiffyCalendarBox("cal", "frmAddClient", "txtBDate","cmdDate","<%= FormatDate(Date) %>",1);
</SCRIPT>
<form name = 'frmAddClient' method = 'post' action = 'AddClientAccount.asp' id = "frmMain">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="441">
	<tr>
		<td width="16%" height="27">CDS No</td>
		<td width="84%" colspan="3" height="27"><input type = 'text' name ='txtCDSNo' tabIndex='0' id = 'txtCDSNo' size="20"></td>
	</tr>
	<tr>
		<td width="16%" height="27">Name</td>
		<td width="30%" height="27">
        <input type = 'text' name ='txtName' tabIndex='1' id = 'txtName' size="20"></td>
		<td width="23%" align="right" height="27">ID No/Passport/Cert&nbsp;&nbsp;</td>
		<td width="31%" height="27"><input type = 'text' name ='txtIDPass' tabIndex='12' id = 'txtIDPass' size="20"></td>
	</tr>
	<tr>
		<td width="16%" height="27">Date of Birth</td>
		<td width="30%" height="27"><SCRIPT language="JavaScript">cal.writeControl();  document.all.item('txtBDate').tabIndex='2'</SCRIPT></td>
		<td width="23%" align="right" height="27">Gender&nbsp;&nbsp;</td>
		<td width="31%" height="27">
		<select name = 'cboGender' tabIndex='13' id = 'cboGender' size="1">
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
	</tr>
	<tr>
		<td width="16%" height="27">Contact Name</td>
		<td width="30%" height="27"><input type = 'text' name ='txtContact' STYLE="WIDTH: 220px" tabIndex='3' id = 'txtContact' size="20"></td>
		<td width="23%" align="right" height="27">E-mail&nbsp;&nbsp;</td>
		<td width="31%" height="27"><input type = 'text' name ='txtEmail' STYLE="WIDTH: 220px" tabIndex='14' id = 'txtEmail' size="20"></td>
	</tr>
	<tr>
		<td width="16%" height="27">Cell Phone</td>
		<td width="30%" height="27"><input type = 'text' name ='txtCellTel' tabIndex='4' id = 'txtCellTel' size="20"></td>
		<td width="23%" align="right" height="27">Home Phone&nbsp;&nbsp;</td>
		<td width="31%" height="27"><input type = 'text' name ='txtHomeTel' tabIndex='15' id = 'txtHomeTel' size="20"></td>
	</tr>
	<tr>
		<td width="16%" height="27">Office Phone</td>
		<td width="30%" height="27"><input type = 'text' name ='txtOfficeTel' tabIndex='5' id = 'txtOfficeTel' size="20"></td>
		<td width="23%" align="right" height="27">Fax&nbsp;&nbsp;</td>
		<td width="31%" height="27"><input type = 'text' name ='txtFax' tabIndex='16' id = 'txtFax' size="20"></td>
	</tr>
	<tr>
		<td width="16%" height="27">Branch</td>
		<td width="30%" height="27">
		<select name = 'cboBranch' tabIndex='6' id = 'cboBranch' size="1">
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
		<td width="23%" align="right" height="27">Class&nbsp;&nbsp;</td>
		<td width="31%" height="27">
		<select name = 'cboClass' tabIndex='17' id = 'cboClass' size="1" onchange='javascript:DisplayDefaultCommission(this);'>
    		<option selected value = ''></option>
			<%
			sqlStr = "SELECT * FROM [ClassList]"
			Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
			If Not (rs.EOF Or rs.BOF) Then
			        rs.MoveFirst
			        Do Until rs.EOF
			                if cbool(rs.Fields("DefaultSelection")) then%>
									<option selected DefaultCommission = '<%=rs.Fields("Commission_DPA_")%>' value = '<%=rs.Fields("Class_DPA_")%>'><%=rs.Fields("ClassClass")%></option>
			                <%else%>
									<option DefaultCommission = '<%=rs.Fields("Commission_DPA_")%>' value = '<%=rs.Fields("Class_DPA_")%>'><%=rs.Fields("ClassClass")%></option>
			                <%end if
			                rs.MoveNext
			        Loop
			End If
			%>
		</select>
		</td>
	</tr>
	<tr>
		<td width="16%" height="27">Commission</td>
		<td width="30%" height="27">
		<select name = 'cboCommission' tabIndex='7' id = 'cboCommission' size="1" style="50px">
			<option selected value = '0'></option>
			<%
		    sqlStr = "SELECT * FROM [CommissionList]"
		    Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		    If Not (rs.EOF Or rs.BOF) Then
		            rs.MoveFirst
		            Do Until rs.EOF
		                    if cbool(rs.Fields("DefaultSelection")) then%>
									<option selected value = '<%=rs.Fields("Commission_DPA_")%>'><%=rs.Fields("CommissionDisplay")%></option>
		                    <%else%>
									<option value = '<%=rs.Fields("Commission_DPA_")%>'><%=rs.Fields("CommissionDisplay")%></option>
		                    <%end if
		                    rs.MoveNext
		            Loop
		    End If
			%>
		</select>
		</td>
		<td width="23%" align="right" height="27">Residency&nbsp;&nbsp;</td>
		<td width="31%" height="27">
		<select name = 'cboResidency' tabIndex='18' id = 'cboResidency' size="1">
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
	</tr>
	<tr>
		<td width="16%" height="27">Agent</td>
		<td width="30%" height="27">
		<select name = 'cboAgent' tabIndex='8' id = 'cboAgent' size="1">
			<option selected value = ''></option>
			<%
		    Set conn = GetActiveConnection("KBroker")
		    
		    sqlStr = "SELECT * FROM [AgentList]"
		    Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		    If Not (rs.EOF Or rs.BOF) Then
		            rs.MoveFirst
		            Do Until rs.EOF%>
							<option value = '<%=rs.Fields("Agent_DPA_")%>'><%=rs.Fields("AgentName")%></option>
							
		                    <%rs.MoveNext
		            Loop
		    End If
			%>
		</select>
		</td>
		<td width="23%" align="right" height="27">Account Manager&nbsp;&nbsp;</td>
		<td width="31%" height="27">
		<select name = 'cboOwner' tabIndex='19' id = 'cboOwner' size="1">
			<option selected value = ''></option>
			<%
		    sqlStr = "SELECT * FROM [OwnerList]"
		    Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		    If Not (rs.EOF Or rs.BOF) Then
		            rs.MoveFirst
		            Do Until rs.EOF%>
		                    <option value = '<%=rs.Fields("Owner_DPA_")%>'><%=rs.Fields("OwnerName")%></option>
		                    <%rs.MoveNext
		            Loop
		    End If
			%>
		</select>
		</td>
	</tr>
	<tr>
		<td width="16%" height="27">Credit Limit</td>
		<td width="30%" height="27">
		<input type = 'text' name ='txtCreditLimit' tabIndex='9' STYLE="TEXT-ALIGN: RIGHT; WIDTH: 100PX" id = 'txtCreditLimit' size="20" value="0">
		<select style="display: none" name = 'cboVIP' tabIndex='17' id = 'cboVIP' size="1">
    		<option value = '1'>Yes</option>
    		<option value = '0' selected>No</option>
		 </select>
		 </td>
		<td width="23%" align="right" height="27">Opening Balance&nbsp;&nbsp;</td>
		<td width="31%" height="27"><input type = 'text' name ='txtOpeningBal' tabIndex='20' STYLE="TEXT-ALIGN: RIGHT; WIDTH: 100PX" id = 'txtOpeningBal' size="20" value="0"></td>
	</tr>
	<tr>
		<td width="16%" height="46">Notes/Comments</td>
		<td width="30%" height="46"><TEXTAREA name ='txtComment' tabIndex='10' id = 'txtComment' rows='2' cols="25"></TEXTAREA></td>
		<td width="54%" colspan="2" height="46">&nbsp;</td>
	</tr>
	<tr>
		<td width="16%" height="67">Address</td>
		<td width="30%" height="67"><textarea rows=3 name ='txtAddr' tabIndex='11' id = 'txtAddr' size="35" cols="33"></textarea></td>
		<td width="23%" align="right" height="67">Physical Address&nbsp;&nbsp;</td>
		<td width="31%" height="67"><textarea rows=3 name ='txtPAddr' tabIndex='21' id = 'txtPAddr' size="35" cols="25"></textarea></td>
	</tr>
	<tr>
	<td width="20%"  valign="top" height="27">Generic1</td>
    <td width="20%" height="27"><select name = 'cboGenericSetting1' id = 'cboGenericSetting1' size="1">
					<option value = ''></option>
<%
        Dim genericRS
        sqlStr = "SELECT * FROM [GenericSettingList] WHERE EntityType_DPA_ = 1 Order By GenericSettingDescription"
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
    <td width="20%"  align=right valign="top" height="27">Generic3&nbsp;&nbsp;&nbsp;</td>
    <td width="20%" height="27"><select name = 'cboGenericSetting3' id = 'cboGenericSetting3' size="1">
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
  <tr>
	<td width="20%"  valign="top" height="31">Generic2</td>
    <td width="20%" height="31"><select name = 'cboGenericSetting2' id = 'cboGenericSetting2' size="1">
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
    <td width="20%"  valign="top" height="31" align="right">CDA&nbsp;&nbsp;&nbsp;&nbsp;</td>
    <td width="20%" height="27"><input type=checkbox   value='False' name='chkCustodian' onClick = 'UpdateCustodianStatus(this);'></td>
  </tr>
</table>
<table border="0" width="877" cellspacing="0" cellpadding="0" height="217">
  <tr>
	<td colspan="4" width="708" height="216">
		<table width="117%" border="0" height="16">
			<tr>
  				<td width="20%" height="15">Photo</td>
				<td width="20%" height="15">Signature</td>
			</tr>
			<tr>
  				<td width="20%" height="180">
					<input type="hidden" name="txtPhoto">
					<IFRAME FRAMEBORDER=0 SCROLLING=NO SRC="upload.asp?filetext=txtPhoto" width="170px" height="170px" tabIndex='-1'></IFRAME>
				</td>
  				<td width="20%" height="180">
					<input type="hidden" name="txtSignature">
					<IFRAME FRAMEBORDER=0 SCROLLING=NO SRC="upload.asp?filetext=txtSignature" width="170px" height="170px" tabIndex='-1'></IFRAME>
				</td>
			
  				<td width="94%" height="180">
    <p>
    <b>Bank Account Details</b>
    <table border=0 style="border: 1px ridge #000080; border-collapse:collapse" WIDTH="75%" height="106" bordercolor="#111111" cellpadding="0" cellspacing="0">	
	
    <td width="30%" height="32">Bank Branch</td>
    <td width="70%" height="32"><select name = 'cboBnkBranch' id = 'cboBnkBranch' size="1" tabIndex='24'>
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
    <td width="30%" height="32">Account Name</td>
    <td width="70%" height="32"><input type = 'text' name ='txtAccName' id = 'txtAccName' size="25" tabIndex='24'></td>
  </tr>
  <tr>
    <td width="30%" height="30">Account Number</td>
    <td width="70%" height="30"><input type = 'text' name ='txtNumber' id = 'txtNumber' size="26" tabIndex='25'></td>
  </tr>
  </table>
                    <p>&nbsp;</td>
  				<td width="94%" height="180">
					&nbsp;</td>
			</tr>
			</table>
	</td>
	<td width="169" align="right" valign=Bottom height="216" >
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save " tabIndex='25' onclick = "AllowedNavigation()"><input type = 'button' Class=Buttons name ='cmdAdd' id = "cmdCancel" value=" Cancel " tabIndex='26' onclick = "JavaScript: window.self.close()">&nbsp;
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		
	</td>
  </tr>
  <tr>
    <td width="104" height="1"></td>
    <td width="139" height="1"><input type = 'hidden' name ='txtGeneric1' id = 'txtGeneric1' size="20"></td>
    <td width="139" height="1"><input type = 'hidden' name ='custodianStatus' id = 'custodianStatus' size="20" value='0'></td>
    <td width="209" align=right height="1"></td>
    <td width="256" height="1"><input type = 'hidden' name ='txtGeneric2' id = 'txtGeneric2' size="20"></td>
  </tr>   
</table>
</tr>

</form>
</body>

</html>