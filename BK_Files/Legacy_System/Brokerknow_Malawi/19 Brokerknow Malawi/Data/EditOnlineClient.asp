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
	
	Set rs = CreateObject("ADODB.Recordset")   						        
   rs.CursorLocation = adUseClient
	
   CompanyContacts=""
   	
   Set conn = GetActiveConnection("KBroker")	   
	
	'Response.Write(Session("Company"))
	sqlStr="Select * From CompanyInfo"
	set rs=conn.execute(sqlStr)
	CompanyContacts=CompanyContacts & " CompanyName: " & rs("CompanyName") & "<p>"
	CompanyContacts=CompanyContacts & ", Address: " & rs("Address") & " (" & rs("PostalCode") & ")" & "<p>"
	CompanyContacts=CompanyContacts & ", Phone Number: " & rs("PhoneNumber") & "<p>"
	CompanyContacts=CompanyContacts & ", Fax Number: " & rs("FaxNumber") & "<p>"
	CompanyContacts=CompanyContacts & ", Email: Admin@dyerandblair.com" & rs("FaxNumber") & "<p>"
	CompanyContacts=CompanyContacts & ", City: " & rs("City") & "; Country: " & rs("Country") & "<p>"

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
                				ShowMessage "Please specify the REF Number"
                				
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
						ShowMessage "Client REF Number is unique"						
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
						conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
						
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
								
								ClientDPA=rs.Fields("Client_DPA_")
								
								'save detail data
								sqlStr = "INSERT INTO [BankAcc] (BankAccNumber,BankAcc_DPA_,BankAccName,BnkBranch_DPA_,Client_DPA_) SELECT " & "'" & accNum & "'" & " as BankAccNumber" & _
										"," & " " & "iif(isnull(max([BankAcc_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'BankAcc'),max([BankAcc_DPA_]) + 1)" & " " & " as BankAcc_DPA_" & _
										"," & "'" & accName & "'" & " as BankAccName" & _
										"," & " " & bnkBranch & " " & " as BnkBranch_DPA_" & _
										"," & " " & rs.Fields("Client_DPA_") & " " & " as Client_DPA_" & _
										" FROM [BankAcc]"
	        
								conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
						end if
						Conn.Execute("Update OnlineClient Set Deleted=1 where Client_DPA_=" & ID)
						Conn.Execute("Update Users set Client_DPA_=" & ClientDPA & " where Client_DPA_=" & ID)
				conn.CommitTrans
				
				Subject = "Client Approval"
				Body = "Your application details has been approved by dyer and blair company and you ." & "<p>"
				Body = Body & "have been registered as one of its clients : " & "<p>"
				Body = Body & "Your login details will be sent to you later via mail: " & OrderID & "<p>"
				Body = Body & "For further information contact us : " & CompanyContacts & "<p>"		
				
				if(Isnull(ClientEmail) or ClientEmail="") then
				else			
				SendMail ClientEmail, Subject, Body			
				end if				

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
	

	sub SendMail(toRecipient, subject, bodyText)
	cc = ""
	bcc = ""
	
	Set Conn = GetActiveConnection("KBroker")
	SQL = "SELECT * FROM MAilConfigList"
	Set Rs = Conn.Execute(SQL)
	
	If (Rs.EOF Or Rs.BOF) Then%>
		<Script Language="JavaScript">
			alert("The mail configurations have not been set.");
		</Script>
		<%Response.End 
	End If
				
	Const cdoSchema = "http://schemas.microsoft.com/cdo/configuration/" 
	Set objMsg = CreateObject("CDO.Message") 
	
	With objMsg
	
		.Configuration.Fields.Item(cdoSchema & "sendusing") = Rs.Fields("SendUsingMethod").Value 
		.Configuration.Fields.Item(cdoSchema & "smtpserver") = Rs.Fields("SMTPServer").Value 
		.Configuration.Fields.Item(cdoSchema & "smtpserverport") = Rs.Fields("SMTPServerPort").Value 
		.Configuration.Fields.Update 		 
		.Subject = subject
		.Sender = Rs.Fields("SendDisplayName").Value 
		.To = toRecipient
		If cc <> "" Then
			.CC = cc
		End If
		
		If bcc <> "" Then
			.BCC = bcc
		End If
		
		If bodyText <> "" Then
			.TextBody = bodyText
		End If
		
		.HTMLBody = bodyText	' "file://" & pathTo
		
		On Error Resume Next
		.Send 
		
		If Err.number > 0 Then%>
			<Script Language="JavaScript">
				alert("The page could not be sent. An unexpected error occured: <%= Err.description %>")
			</Script>
			<%Response.End 
		End If
	End With
	
	Set Rs = Nothing
	Set Conn = Nothing
	Set objMsg = Nothing
	Set Fso = Nothing%>
		<Script Language="JavaScript">
			alert("The page was sent successfully.")		
		</Script>
	<%	
end sub

    Set conn = GetActiveConnection("KBroker")
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
		}

</script>



</head>

<body Class="Dialog" OnLoad="JavaScript: DoInit();">

<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<form name = 'frmEditClient' method = 'post' action = 'EditOnlineClient.asp' id = "frmMain" target="_self">

<%
 Set conn = GetActiveConnection("KBroker")
      
		sqlStr = "SELECT OnlineClients.* From OnlineClients WHERE  OnlineClients.Client_DPA_ = " & ID
				
        'Response.write(sqlStr)
		'Response.end

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
		<input type = 'text' name ='txtCreditLimit' tabIndex='9' STYLE="TEXT-ALIGN: RIGHT; WIDTH: 100PX" id = 'txtCreditLimit' readonly size="20" value="<%=rs.Fields("CreditLimit")%>" >
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
		<td width="31%"><input type = 'text' name ='txtOpeningBal' tabIndex='20' STYLE="TEXT-ALIGN: RIGHT; WIDTH: 100PX" id = 'txtOpeningBal' size="20"  readonly value = '<%= FormatNum(rs.Fields("ClientOpeningBal"))%>'></td>
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
    <td width="20%"  valign="top" height="31" align="right">CDA</td>
    <td width="20%" height="27"><%if cbool(Rs.Fields("IsCustodian")) then
		custodian=1
		%>
		<input type=checkbox  checked value='True' name='chkCustodian' id='chkCustodian' onClick = 'UpdateCustodianStatus(this);'> 
	<%else
		custodian=0	
	%>
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
	  <td width="169" align="right" valign=Bottom height="50" >
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
	<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
  </tr>	
	</td>
  
	</tr>	
</table>

</form>
</body>

</html>