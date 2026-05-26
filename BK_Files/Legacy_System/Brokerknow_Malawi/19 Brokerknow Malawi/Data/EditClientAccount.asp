<!--#include file="../libroutines.asp"-->
<%
	const LinkedIndependent = 1
	const LinkedDependent = 2
	
	Dim UserId
	Dim action
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim ID
	Dim rsEdit
	Dim bankAccExists
	

	UserId=Session("UserID")
	action = ucase(Request.Form("action"))
	ID = Request("ID")

		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		window.self.ShowMessage "No record specified for editing"
                		
                		window.self.close
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
				Dim BankAcc_DPA_
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
				bankAccExists = cbool(Request.Form("bankAccExists"))
				BankAcc_DPA_ = Request.Form("BankAcc_DPA_")
				CreditLimit = Request.Form("txtCreditLimit") 
				CDSNo = Request.Form("txtCDSNo")
				gen1 = Request.Form("cboGenericSetting1")
				gen2 = Request.Form("cboGenericSetting2")
				gen3 = Request.Form("cboGenericSetting3") 
				IsCustodian = Request.Form("custodianStatus") 
				
				'validate CDS
			
				If Trim(CDSNo) = "" Then%>
						<script language = 'vbscript'>
                				ShowMessage "Please specify the REF Number"
                				
						</script>
						<% response.end
				End If				
		       
				
				If Not IsNumeric(OpeningBal) Then%>
						<script language = 'vbscript'>
						ShowMessage "Opening Balance can only numeric"						
						</script>
						<% response.end
				End If
			    
				'validate Branch
				If Trim(branch) = "" Then%>
						<script language = 'vbscript'>
                				window.self.ShowMessage "Please specify the Branch"
                				
						</script>
						<% response.end
				End If
				'validate Commission Types
				If Trim(commission) = "" Then%>
						<script language = 'vbscript'>
                				window.self.ShowMessage "Please specify the Commission "
                				
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
                				window.self.ShowMessage "Please specify the Residency"
                				
						</script>
						<% response.end
				End If
				'validate Name
				If Trim(Name) = "" Then%>
						<script language = 'vbscript'>
                				window.self.ShowMessage "Please specify the Name"
                				
						</script>
						<% response.end
				End If
				'validate size of Address
				If Len(Addr) > 500 Then%>
						<script language = 'vbscript'>
						window.self.ShowMessage "Address can only be 100 characters in length"
						
						</script>
						<% response.end
				End If
				'validate size of Cell Phone
				If Len(Cell) > 100 Then%>
						<script language = 'vbscript'>
						window.self.ShowMessage "Cell Phone can only be 100 characters in length"
						
						</script>
						<% response.end
				End If
				'validate size of Contact Name
				If Len(Contact) > 100 Then%>
						<script language = 'vbscript'>
						window.self.ShowMessage "Contact Name can only be 100 characters in length"
						
						</script>
						<% response.end
				End If
				'validate size of Email
				If Len(Email) > 100 Then%>
						<script language = 'vbscript'>
						window.self.ShowMessage "Email can only be 100 characters in length"
						
						</script>
						<% response.end
				End If
				'validate size of Fax
				If Len(Fax) > 100 Then%>
						<script language = 'vbscript'>
						window.self.ShowMessage "Fax can only be 100 characters in length"
						
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
						window.self.ShowMessage "Home Phone can only be 100 characters in length"
						
						</script>
						<% response.end
				End If
				
				'Make sure user enters ID No/Passport No/Cert
				if Len(Trim(IDPass)) = 0 then
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
						window.self.ShowMessage "ID or Passport can only be 100 characters in length"
						
						</script>
						<% response.end
				End If
				'validate size of Name
				If Len(Name) > 100 Then%>
						<script language = 'vbscript'>
						window.self.ShowMessage "Name can only be 100 characters in length"
						
						</script>
						<% response.end
				End If
				'validate size of Office Phone
				If Len(Office) > 100 Then%>
						<script language = 'vbscript'>
						window.self.ShowMessage "Office Phone can only be 100 characters in length"
						
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
			    
				Set conn = GetActiveConnection("KBroker")
				
				TimeChanged=Now()
				'save data
				sqlStr = "UPDATE [Client] SET ClientAddr = " & "'" & addr & "'" & ",ClientBDate = " & "#" & FormatDate(bDate) & "#" & "" & _
						",ClientCellTel = " & "'" & cell & "'" & ",ClientContact = " & "'" & contact & "'" & "" & _
						",ClientEmail = " & "'" & email & "'" & ",ClientFax = " & "'" & fax & "'" & "" & _
						",GenericSetting_DPA_ = " & " " & gen1 & " " & _
						",GenericSetting_DPA_2 = " & " " & gen2 & " " & "" & _
						",GenericSetting_DPA_3 = " & " " & gen3 & " " & "" & _
						",ClientPAddr = " & "'" & pAddr & "'" & "" & _
						",ClientComment = " & "'" & comment & "'" & "" & _
						",ClientHomeTel = " & "'" & home & "'" & ",ClientIDPass = " & "'" & idPass & "'" & "" & _
						",ClientName = " & "'" & name & "'" & ",ClientOfficeTel = " & "'" & Office & "'" & "" & _
						",ClientPhoto = " & "'" & photoNew & "'" & "" & _ 
						",ClientSignature = " & "'" & sigNew & "'" & "" & _ 
						",ClientVIP = " & " " & vip & " " & "" & _						
						",IsCustodian = " & " " & IsCustodian & " " & "" & _						
						",ClientCDSNo = " & " '" & CDSNo & "' " & "" & _
						",ChangedBy = " & " " & UserId & " " & "" & _
						",TimeChanged = " & " '" & TimeChanged & "' " & "" & _
						",CreditLimit = " & " " & CreditLimit & " " & "" & _
						",Agent_DPA_ = " & " " & agent  & " " & ",Branch_DPA_ = " & " " & branch & " " & "" & _
						",Class_DPA_ = " & " " & classType & " " & ",Commission_DPA_ = " & " " & commission & " " & "" & _
						",Gender_DPA_ = " & " " & gender & " " & ",ClientOpeningBal = " & CDbl(OpeningBal) & ", Owner_DPA_ = " & " " & owner & " " & "" & _
						",Residency_DPA_ = " & " " & residency & " " & " WHERE Client_DPA_  = " & ID								  					
					
				conn.BeginTrans
						conn.Execute SQLServerFormat(HandleQuote(sqlStr))
						
						if (not(bankAccExists)) and (bankAccSpecified) then
							'new account
							sqlStr = "INSERT INTO [BankAcc] (BankAccNumber,BankAcc_DPA_,BankAccName,BnkBranch_DPA_,Client_DPA_) SELECT " & "'" & accNum & "'" & " as BankAccNumber" & _
										"," & " " & "iif(isnull(max([BankAcc_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'BankAcc'),max([BankAcc_DPA_]) + 1)" & " " & " as BankAcc_DPA_" & _
										"," & "'" & accName & "'" & " as BankAccName" & _
										"," & " " & bnkBranch & " " & " as BnkBranch_DPA_" & _
										"," & " " & ID & " " & " as Client_DPA_" & _
										" FROM [BankAcc]"
						elseif (not(bankAccExists)) and (not(bankAccSpecified)) then
							'no change
							sqlStr = ""
						elseif (bankAccExists) and (not(bankAccSpecified)) then
							'remove account
							'find out whether any child records exist
							sqlStr = "SELECT Child,DeletionMessage,ParentKey FROM [_Parent_Child_Links_] WHERE (Parent = 'BankAcc') AND (ChildType = " & LinkedIndependent & ")"
							Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
							If Not (rs.BOF Or rs.EOF) Then
							        Dim childRS
							        Dim tableName
							        
							        rs.MoveFirst
							        Do Until rs.EOF
							        		tableName = rs.Fields("Child")
							                sqlStr = "SELECT TOP 1 * FROM [" & tableName & "] WHERE BankAcc_DPA_ = " & BankAcc_DPA_
							                
							                Set childRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
							                If Not (childRS.BOF Or childRS.EOF) Then%>
							        				<script language = 'vbscript'>
							        					ShowMessage "<%=rs.Fields("DeletionMessage")%>"
							        					window.self.close
							        				</script>
							        				<%response.end
							                End If
							                rs.MoveNext
							        Loop
							End If
        
							'delete from database
							sqlStr = "DELETE FROM [BankAcc] WHERE BankAcc_DPA_ = " & BankAcc_DPA_
						elseif (bankAccExists) and (bankAccSpecified) then
							'update account
							sqlStr = "UPDATE [BankAcc] SET BankAccNumber = " & "'" & accNum & "'" & _
									",BankAccName = " & "'" & accName & "'" & _
									",BnkBranch_DPA_ = " & " " & bnkBranch & " " & _
									",Client_DPA_ = " & " " & ID & " " & "" & _
									" WHERE BankAcc_DPA_  = " & BankAcc_DPA_
						end if
						
						if sqlStr <> "" then
								conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
						end if 
				conn.CommitTrans
				
				Set conn = Nothing
				WritefraEnabledDialogCloseScript
				Response.End
        end if
        Dim clientCode
        
        clientCode = "var validNavigate = true;" & chr(13)
		clientCode = clientCode & "window.parent.close();" & chr(13)%>
		<script>
			<%=clientCode%>
		</script>
		<%
		response.End
   	end If
%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<META HTTP-EQUIV="Expires" CONTENT="0">

<title>Edit Client</title>

  <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
 <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css"> 
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT>
 
  <!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">

<script language="javascript">
		var validNavigate = false;
		ExemptFromDefaultTabIndex = true;
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
	function drawPicLayout(){
		try{
			window.parent.frames[0].execScript("drawPicture()", "JavaScript");
		}
		catch(e){}
		
		
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
			//alert(document.frmMain.elements("custodianStatus").value);
		}

</script>



</head>

<body Class="Dialog" OnLoad="JavaScript: DoInit();">

<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<form name = 'frmEditClient' method = 'post' action = 'EditClientAccount.asp' id = "frmMain" target="_self">

<%
 Set conn = GetActiveConnection("KBroker")
      
		sqlStr = "SELECT ClientCDSNo, CreditLimit,  Client.ClientAddr, Client.ClientBDate, Client.ClientCellTel, Client.ClientContact, Client.ClientEmail, Client.ClientComment, Client.ClientFax,  " & _
                  "    Client.GenericSetting_DPA_,Client.GenericSetting_DPA_2, Client.GenericSetting_DPA_3, Client.ClientPAddr, Client.ClientHomeTel, Client.ClientIDPass, Client.ClientName, Client.ClientOfficeTel,  " & _
                  "    Client.ClientPhoto, Client.ClientVIP, Client.Client_DPA_, AgentList.Agent_DPA_, BranchList.Branch_DPA_, Client.Class_DPA_, Client.Commission_DPA_,  " & _
                  "    Client.Gender_DPA_, Client.ClientSignature, Client.Owner_DPA_, Client.Residency_DPA_,Client.IsCustodian, Client.ClientOpeningBal, ISNULL(BankAcc.BnkBranch_DPA_, 0)  " & _
                  "    AS BnkBranch_DPA_, ISNULL(BankAcc.BankAccName, '') AS BankAccName, ISNULL(BankAcc.BankAccNumber, '') AS BankAccNumber,  " & _
                  "    ISNULL(BankAcc.BankAcc_DPA_, 0) AS BankAcc_DPA_  " & _
				"	FROM         OwnerList RIGHT OUTER JOIN " & _
                 "	     GenderList RIGHT OUTER JOIN " & _
                 "	     ResidencyList INNER JOIN " & _
                 "	     CommissionList INNER JOIN " & _
                 "	     ClassList INNER JOIN " & _
                 "	     BranchList INNER JOIN " & _
                  "	    Client LEFT OUTER JOIN " & _
                 "	     AgentList ON AgentList.Agent_DPA_ = Client.Agent_DPA_ ON BranchList.Branch_DPA_ = Client.Branch_DPA_ ON  " & _
                 "	     ClassList.Class_DPA_ = Client.Class_DPA_ ON CommissionList.Commission_DPA_ = Client.Commission_DPA_ ON  " & _
                 "	     ResidencyList.Residency_DPA_ = Client.Residency_DPA_ ON GenderList.Gender_DPA_ = Client.Gender_DPA_ ON  " & _
                 "	     OwnerList.Owner_DPA_ = Client.Owner_DPA_ LEFT OUTER JOIN " & _
                 "	     BankAcc ON Client.Client_DPA_ = BankAcc.Client_DPA_ " & _
				"	WHERE     Client.Client_DPA_ = " & ID
				
        
        sqlStr = SQLServerFormat(HandleQuote(sqlStr))
        
        Set rs = conn.Execute(sqlStr)
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		window.self.ShowMessage "The selected Client cannot be retrieved for editing"
                		
                </script>
                <% response.end
        End If
%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
	<tr>
		<td width="16%">REF No</td>
		<td width="30%"><input type = 'text' name ='txtCDSNo' tabIndex='0' id = 'txtCDSNo' size="20" value="<%=rs.Fields("ClientCDSNo")%>"></td>
		<td width="23%" align="right">Client Code&nbsp;&nbsp;</td>
		<td width="31%"><%=rs.Fields("Client_DPA_")%></td>
	</tr>
	<tr>
		<td width="16%">Name</td>
		<td width="30%">
        <input type = 'text' name ='txtName' STYLE="WIDTH: 220px" tabIndex='1' id = 'txtName' value = '<%=rs.Fields("ClientName")%>' size="20"></td>
		<td width="23%" align="right">ID No/Passport/Cert&nbsp;&nbsp;</td>
		<td width="31%"><input type = 'text' name ='txtIDPass' tabIndex='12' id = 'txtIDPass' size="20" value = '<%=rs.Fields("ClientIDPass")%>'></td>
	</tr>
	<tr>
		<td width="16%">Date of Birth</td>
		<td width="30%">
		<script language="JavaScript" src="../CALENDAR/calendar.js"></script>
		<SCRIPT language="JavaScript">
			var cal=new ctlSpiffyCalendarBox("cal", "frmEditClient", "txtBDate","cmdDate","<%= FormatDate(rs.Fields("ClientBDate")) %>",1);		
			cal.writeControl();  
			document.all.item('txtBDate').tabIndex='2'
		</SCRIPT>
		</td>
		<td width="23%" align="right">Gender&nbsp;&nbsp;</td>
		<td width="31%">
		<select name = 'cboGender' tabIndex='13' id = 'cboGender' size="1">
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
	</tr>
	<tr>
		<td width="16%">Contact Name</td>
		<td width="30%"><input type = 'text' name ='txtContact' STYLE="WIDTH: 220px" tabIndex='3' id = 'txtContact' size="20" value = '<%=rs.Fields("ClientContact")%>'></td>
		<td width="23%" align="right">E-mail&nbsp;&nbsp;</td>
		<td width="31%"><input type = 'text' name ='txtEmail' STYLE="WIDTH: 220px" tabIndex='14' id = 'txtEmail' size="20"  value = '<%=rs.Fields("ClientEmail")%>'></td>
	</tr>
	<tr>
		<td width="16%">Cell Phone</td>
		<td width="30%"><input type = 'text' name ='txtCellTel' tabIndex='4' id = 'txtCellTel' size="20" value = '<%=rs.Fields("ClientCellTel")%>'></td>
		<td width="23%" align="right">Home Phone&nbsp;&nbsp;</td>
		<td width="31%"><input type = 'text' name ='txtHomeTel' tabIndex='15' id = 'txtHomeTel' size="20"  value = '<%=rs.Fields("ClientHomeTel")%>'></td>
	</tr>
	<tr>
		<td width="16%">Office Phone</td>
		<td width="30%"><input type = 'text' name ='txtOfficeTel' tabIndex='5' id = 'txtOfficeTel' size="20"  value = '<%=rs.Fields("ClientOfficeTel")%>'></td>
		<td width="23%" align="right">Fax&nbsp;&nbsp;</td>
		<td width="31%"><input type = 'text' name ='txtFax' tabIndex='16' id = 'txtFax' size="20"  value = '<%=rs.Fields("ClientFax")%>'></td>
	</tr>
	<tr>
		<td width="16%">Branch</td>
		<td width="30%">
		<select name = 'cboBranch' tabIndex='6' id = 'cboBranch' size="1">
			<%
		    Photo = Rs.Fields("ClientPhoto")
		    If IsNull(Photo) Then
				Photo = "/Data/Photos/_blank.jpg"
			ElseIf Photo = "" Then
				Photo = "/Data/Photos/_blank.jpg"
			End If	
		    
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
		<td width="23%" align="right">Class&nbsp;&nbsp;</td>
		<td width="31%">
		<select name = 'cboClass' tabIndex='17' id = 'cboClass' size="1" onchange='javascript:DisplayDefaultCommission(this);'>
			<%
		    sqlStr = "SELECT * FROM [ClassList]"
		    Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		    If Not (rsEdit.EOF Or rsEdit.BOF) Then
		            rsEdit.MoveFirst
		            Do Until rsEdit.EOF
		            		if rsEdit.Fields("Class_DPA_") = rs.Fields("Class_DPA_") Then%>
		            			<option selected DefaultCommission = '<%=rsEdit.Fields("Commission_DPA_")%>' value = '<%=rsEdit.Fields("Class_DPA_")%>'><%=rsEdit.Fields("ClassClass")%></option>
		            		<%else%>
		                    <option DefaultCommission = '<%=rsEdit.Fields("Commission_DPA_")%>' value = '<%=rsEdit.Fields("Class_DPA_")%>'><%=rsEdit.Fields("ClassClass")%></option>
		                 <%end if
							rsEdit.MoveNext
		            Loop
		    End If
			%>
		</select>
		</td>
	</tr>
	<tr>
		<td width="16%">Commission</td>
		<td width="30%">
		<select name = 'cboCommission' tabIndex='7' id = 'cboCommission' size="1">
		<%
		    sqlStr = "SELECT * FROM [CommissionList]"
		    Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		    If Not (rsEdit.EOF Or rsEdit.BOF) Then
		            rsEdit.MoveFirst
		            Do Until rsEdit.EOF
		            		if rsEdit.Fields("Commission_DPA_") = rs.Fields("Commission_DPA_") Then%>
		            			<option selected value = '<%=rsEdit.Fields("Commission_DPA_")%>'><%=rsEdit.Fields("CommissionDisplay")%></option>
		            		<%else%>
		                    <option value = '<%=rsEdit.Fields("Commission_DPA_")%>'><%=rsEdit.Fields("CommissionDisplay")%></option>
		                 <%end if
							rsEdit.MoveNext
		            Loop
		    End If
		%>
		</select>
		</td>
		<td width="23%" align="right">Residency&nbsp;&nbsp;</td>
		<td width="31%">
		<select name = 'cboResidency' tabIndex='18' id = 'cboResidency' size="1">
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
	</tr>
	<tr>
		<td width="16%">Agent</td>
		<td width="30%">
		<select name = 'cboAgent' tabIndex='8' id = 'cboAgent' size="1">
			<option value = ''></option>
			<%
		    sqlStr = "SELECT * FROM [AgentList]"
		    Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		    If Not (rsEdit.EOF Or rsEdit.BOF) Then
		            rsEdit.MoveFirst
		            Do Until rsEdit.EOF
		            		if rsEdit.Fields("Agent_DPA_") = rs.Fields("Agent_DPA_") Then%>
		            			<option selected value = '<%=rsEdit.Fields("Agent_DPA_")%>'><%=rsEdit.Fields("AgentName")%></option>
		            		<%else%>
		                    <option value = '<%=rsEdit.Fields("Agent_DPA_")%>'><%=rsEdit.Fields("AgentName")%></option>
		                 <%end if
							rsEdit.MoveNext
		            Loop
		    End If
			%>
		</select>
		</td>
		<td width="23%" align="right">Account Manager&nbsp;&nbsp;</td>
		<td width="31%">
		<select name = 'cboOwner' tabIndex='19' id = 'cboOwner' size="1">
			<option value = ''></option>
			 <%
		    sqlStr = "SELECT * FROM [OwnerList]"
		    Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		    If Not (rsEdit.EOF Or rsEdit.BOF) Then
		            rsEdit.MoveFirst
		            Do Until rsEdit.EOF
		            		if rsEdit.Fields("Owner_DPA_") = rs.Fields("Owner_DPA_") Then%>
		            			<option selected value = '<%=rsEdit.Fields("Owner_DPA_")%>'><%=rsEdit.Fields("OwnerName")%></option>
		            		<%else%>
		                    <option value = '<%=rsEdit.Fields("Owner_DPA_")%>'><%=rsEdit.Fields("OwnerName")%></option>
		                 <%end if
							rsEdit.MoveNext
		            Loop
		    End If
			%>
		</select>
		</td>
	</tr>
	<tr>
		<td width="16%">Credit Limit</td>
		<td width="30%">
		<input type = 'text' name ='txtCreditLimit' tabIndex='9' STYLE="TEXT-ALIGN: RIGHT; WIDTH: 100PX" id = 'txtCreditLimit' size="20" value="<%=rs.Fields("CreditLimit")%>" >
		<select style="display: none" name = 'cboVIP' tabIndex='19' id = 'cboVIP' size="1">
			<%
			Dim default
			Dim other
			Dim valDefault
			Dim valOther
		    
		    if rs.Fields("ClientVIP") then
		    		default = "Yes"
		    		valDefault = "1"
		    		valOther = "0"
		    		other = "No"
		    else
		    		default = "No"
		    		valDefault = "0"
		    		valOther = "1"
		    		other = "Yes"
		    end if%>
		   <option selected value = '<%=valDefault%>'><%=default%></option>
		   <option value = '<%=valOther%>'><%=other%></option>
		  </select>
		 </td>
		<td width="23%" align="right">Opening Balance&nbsp;&nbsp;</td>
		<td width="31%"><input type = 'text' name ='txtOpeningBal' tabIndex='20' STYLE="TEXT-ALIGN: RIGHT; WIDTH: 100PX" id = 'txtOpeningBal' size="20"  value = '<%= FormatNum(rs.Fields("ClientOpeningBal"))%>'></td>
	</tr>
	<tr>
		<td width="16%">Notes/Comments</td>
		<td width="30%"><TEXTAREA name ='txtComment' tabIndex='10' id = 'txtComment' rows='2' cols="25"  value = "<%=rs.Fields("ClientComment")%>"></TEXTAREA></td>
		<td width="54%" colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td width="16%">Address</td>
		<td width="30%"><textarea rows=3 name ='txtAddr' tabIndex='11' id = 'txtAddr' size="35" cols="33"><%=rs.Fields("ClientAddr")%></textarea></td>
		<td width="23%" align="right">Physical Address&nbsp;&nbsp;</td>
		<td width="31%"><textarea rows=3 name ='txtPAddr' tabIndex='21' id = 'txtPAddr' size="35" cols="25"><%=rs.Fields("ClientPAddr")%></textarea></td>
	</tr>
	<tr>
	<td width="20%"  valign="top">Generic1</td>
    <td width="20%"><select name = 'cboGenericSetting1' id = 'cboGenericSetting1' size="1">
						<option selected value = ''></option>
<%
        sqlStr = "SELECT * FROM [GenericSettingList] WHERE EntityType_DPA_ = 1 Order By GenericSettingDescription"
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
    <td width="20%"  align=right valign="top">Generic3&nbsp;&nbsp;&nbsp;</td>
    <td width="20%"><select name = 'cboGenericSetting3' id = 'cboGenericSetting3' size="1">
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
  <tr>
	<td width="20%"  valign="top">Generic2</td>
    <td width="20%"><select name = 'cboGenericSetting2' id = 'cboGenericSetting2' size="1">
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
    <td width="20%"  valign="top" height="31" align="right">CDA&nbsp;&nbsp;&nbsp;&nbsp;</td>
    <td width="20%" height="27"><%if cbool(Rs.Fields("IsCustodian")) then%>
		<input type=checkbox  checked value='True' name='chkCustodian' id='chkCustodian' onClick = 'UpdateCustodianStatus(this);'> 
	<%else%>
		<input type=checkbox   value='False' name='chkCustodian' id='chkCustodian' onClick = 'UpdateCustodianStatus(this);'> 
	<%end if%></td>      
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
  				<INPUT name="txtPhoto" ID="txtPhoto" type="hidden" value="<%= rs.Fields("ClientPhoto") %>">
				<IFRAME FRAMEBORDER=0 SCROLLING=NO SRC="upload.asp?filetext=txtPhoto" width="170px" height="170px" tabIndex="-1"></IFRAME>		    
			</td>
			<td width="50%">
  				<INPUT name="txtSignature" ID="txtSignature" type="hidden" value="<%= rs.Fields("ClientSignature") %>">
				<IFRAME FRAMEBORDER=0 SCROLLING=NO SRC="upload.asp?filetext=txtSignature" width="170px" height="170px" tabIndex="-1"></IFRAME>		    
			</td>				
			</tr>
		</table>
	</td>
  </tr>
  <tr>
    <td colspan=4>
    <p>
    <b>Bank Account Details</b>
    <table border=0 style="border: 1px ridge #000080" WIDTH="100%">	
    <td width="30%">Bank Branch</td>
    <td width="70%"><select name = 'cboBnkBranch' id = 'cboBnkBranch' size="1" tabIndex='24'>
		<%
        bankAccExists = false
        
        sqlStr = "SELECT * FROM [BnkBranchList]"
        Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rsEdit.EOF Or rsEdit.BOF) Then
                rsEdit.MoveFirst
                Do Until rsEdit.EOF
                		if rsEdit.Fields("BnkBranch_DPA_") = rs.Fields("BnkBranch_DPA_") Then%>
                			<option selected value = '<%=rsEdit.Fields("BnkBranch_DPA_")%>'><%=rsEdit.Fields("BnkBranchName")%></option>
                		<%	
                			bankAccExists = true
                		else%>
                        <option value = '<%=rsEdit.Fields("BnkBranch_DPA_")%>'><%=rsEdit.Fields("BnkBranchName")%></option>
                     <%end if
						rsEdit.MoveNext
                Loop
                
                if not(bankAccExists) then%>
						<option selected value = ''></option>
                <%else%>
						<option value = ''></option>
                <%end if
        End If
	%>
    </select></td>
  </tr>
  <tr>
    <td width="15%">Account Name</td>
    <td width="85%"><input type = 'text' name ='txtAccName' id = 'txtAccName' size="25" tabIndex='24' value = '<%=rs.Fields("BankAccName")%>'></td>
  </tr>
  <tr>
    <td width="15%">Account Number </td>
    <td width="85%"><input type = 'text' name ='txtNumber' id = 'txtNumber' size="28" tabIndex='25' value = '<%=rs.Fields("BankAccNumber")%>'></td>
  </tr>
  </table>
	
  </td>  
  </tr>  
   
</table>

<table align=right>
	<tr>
	
	  <td width="100%" align="right" valign=absBottom >
		<input type = 'submit' Class=Buttons name ='cmdSave' id = 'cmdSave' value=" Save " tabIndex='25' onclick = "AllowedNavigation()">
    	<input type = 'button' Class=Buttons name ='cmdAdd' id = "cmdCancel" value=" Cancel " tabIndex='26' onclick = "JavaScript: window.self.close()">
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		<td width="139" height="1"><input type = 'hidden' name ='custodianStatus' id = 'custodianStatus' size="20" value='0'></td>
    	<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
    	<input type = 'hidden' name ='BankAcc_DPA_' id = 'BankAcc_DPA_' value="<%=rs.Fields("BankAcc_DPA_")%>">
    	<input type = 'hidden' name ='bankAccExists' id = 'bankAccExists' value="<%=bankAccExists%>">
		<input type = 'hidden' name ='handlePhotoRefresh' id = 'handlePhotoRefresh' value="1">
		
	</td>
  
	</tr>	
</table>

</form>
</body>

</html>