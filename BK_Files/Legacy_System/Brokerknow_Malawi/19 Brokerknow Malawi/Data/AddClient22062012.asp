<!--#include virtual="libroutines.asp"-->
<%
	
	Dim action
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim guidStr 
	Dim guid
	
	action = ucase(Request.Form("action"))
	
	if action = "EXECUTE" then
		Dim buttonAction
		Dim reloadRequired
		
		reloadRequired = false
		buttonAction = Request("buttonAction")

		if Trim(Ucase(buttonAction ))= "SAVE" then
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
				Dim vip
				Dim gen1
				Dim gen2
				Dim gen3
				Dim pAddr
				dim comment
				
				Dim bnk
				Dim bnkBranch
				Dim accNum
				Dim accName
				
				Dim bankAccSpecified
				Dim bankAccSpecified2
				Dim bankAccSpecified3
				
				Dim IsCustodian
				Dim UseContactNameInPortfolioReports
				
				bankAccSpecified = false
				bankAccSpecified2 = false
				bankAccSpecified3 = false
				
				''1
				bnk = Request.Form("cboBank")
				bnkBranch = Request.Form("txtBnkBranch")
				accNum = Request.Form("txtNumber")
				accName = Request.Form("txtAccName")
				
				''2
				bnk2 = Request.Form("cboBank2")
				bnkBranch2 = Request.Form("txtBnkBranch2")
				accNum2 = Request.Form("txtNumber2")
				accName2 = Request.Form("txtAccName2")
				
				''3
				bnk3 = Request.Form("cboBank3")
				bnkBranch3 = Request.Form("txtBnkBranch3")
				accNum3 = Request.Form("txtNumber3")
				accName3 = Request.Form("txtAccName3")
				
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
				gen1 = Request.Form("cboGenericSetting1")
				gen2 = Request.Form("cboGenericSetting2")
				gen3 = Request.Form("cboGenericSetting3") 
				IsCustodian = Request.Form("custodianStatus") 
				smsContract = Request.Form("smsContract") 
				smsDebit = Request.Form("smsDebit") 
				'CDSNo = Request.Form("txtCDSNumber")
				UseContactNameInPortfolioReports = Request.Form("chkcontact")
				
				if trim(smsDebit)="" then smsDebit=0
				if trim(smsContract)= "" then smsContract=0
				if trim(UseContactNameInPortfolioReports) = "" then UseContactNameInPortfolioReports=0
				
				'validate CDS 
				'If Trim(CDSNo) = "" Then%>
					<script language = 'vbscript'>
                				'ShowMessage "Please specify CSD Number"
                				
						</script>
						<% 'response.end
				'End If				
		    	'validate Branch
				If Trim(branch) = "" Then%>
						<script language = 'vbscript'>
                				ShowMessage "Please specify the Branch"                				
						</script>
						<% 
						'ReloadNewPage
						response.end
				End If
				'validate Commission Types
				If Trim(commission) = "" Then%>
						<script language = 'vbscript'>
                				ShowMessage "Please specify the Commission "
                				
						</script>
						
						<% 
						ReloadNewPage
						response.end
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
						<% 
						ReloadNewPage
						response.end
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
						<% 
						ReloadNewPage
						response.end
				End If
				'validate Name
				If Trim(Name) = "" Then%>
						
						<script language = 'javascript'>
                				alert('Please specify the Name ');               				
						</script>
						<% 
						ReloadNewPage
						response.end
				End If
				'ensure that at least one phone no is entered
				If Trim(Len(cell)) <= 3 Then
					If Trim(office) = "" Then
						If Trim(home) = "" Then%>
						<script language = 'vbscript'>
                				ShowMessage "You must specify at least one telephone number"
                				
						</script>
						<% 
						ReloadPage(ID)
						response.end
						End If
					End If
				End If
				'validate size of Address
				If Len(Addr) > 500 Then%>
						<script language = 'vbscript'>
						ShowMessage "Address can only be 500 characters in length"
						
						</script>
						<% ReloadPage(ID)
						response.end
				End If
				
				'validate size of Cell Phone
				If Len(Cell) > 100 Then%>
						<script language = 'vbscript'>
						ShowMessage "Cell Phone can only be 100 characters in length"
						
						</script>
						<% ReloadPage(ID)
						response.end
				End If

				if not isnumeric(replace(replace(trim(Cell),"+","")," " ,""))  then %>
						<script language = 'vbscript'>
						ShowMessage "Cell Phone number should contain numeric numbers only"
						
						</script>
						<% ReloadPage(ID)
						response.end

				end if 
				'validate size of Contact Name
				If Len(Contact) > 100 Then%>
						<script language = 'vbscript'>
						ShowMessage "Contact Name can only be 100 characters in length"
						
						</script>
						<% ReloadPage(ID)
						response.end
				End If
				'validate size of Email
				If Len(Email) > 100 Then%>
						<script language = 'vbscript'>
						ShowMessage "Email can only be 100 characters in length"
						
						</script>
						<% ReloadPage(ID)
						response.end
				End If

				If trim(email)<>"" then
					if instr(email,"@") = 0  or instr(email,".") = 0 then%>
						<script language = 'vbscript'>
						ShowMessage "Please enter a valid Email address"
						
						</script>
						<% ReloadPage(ID)
						response.end

					end if 
				end if
				
				If Not IsNumeric(OpeningBal) Then%>
						<script language = 'vbscript'>
						ShowMessage "Opening Balance can only be numeric"						
						</script>
						<% ReloadPage(ID)
						response.end
				End If
				
				'validate size of Fax
				If Len(Fax) > 100 Then%>
						<script language = 'vbscript'>
						ShowMessage "Fax can only be 100 characters in length"
						
						</script>
						<%ReloadPage(ID)
						response.end
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
						<% ReloadPage(ID)
						response.end
				End If
				'validate size of Home Phone
				If Len(home) > 100 Then%>
						<script language = 'vbscript'>
						ShowMessage "Home Phone can only be 100 characters in length"
						
						</script>
						<% ReloadPage(ID)
						response.end
				End If
				
				if trim(home)<>"" then
					if  not  isnumeric(replace(replace(trim(home),"+","")," " ,"")) then %>
							<script language = 'vbscript'>
							ShowMessage "Home Phone should contain numeric numbers only"
							
							</script>
							<% ReloadPage(ID)
							response.end

					end if 
				end if
				'validate size of ID or Passport
				If Len(IDPass) > 100 Then%>
						<script language = 'vbscript'>
						ShowMessage "ID or Passport can only be 100 characters in length"
						
						</script>
						<% ReloadPage(ID)
						response.end
				End If
				'validate size of Name
				If Len(Name) > 100 Then%>
						<script language = 'vbscript'>
						ShowMessage "Name can only be 100 characters in length"
						
						</script>
						<% ReloadPage(ID)
						response.end
				End If
				'validate size of Office Phone
				If Len(Office) > 100 Then%>
						<script language = 'vbscript'>
						ShowMessage "Office Phone can only be 100 characters in length"
						
						</script>
						<% ReloadPage(ID)
						response.end
				End If

				''1				
				'check for bank account specification
				'validate Bank Branch
				If Trim(bnk) <> "" Then
				If Trim(bnkBranch) <> "" Then
					bankAccSpecified = true
					
					'validate Account Number
					If Trim(accNum) = "" Then
						%>
						<script language = 'vbscript'>
							ShowMessage "Please specify the Bank Account Number"
						</script>
						<% 
						ReloadPage(ID)
						response.end
					End If
					
					'validate size of Account Number
					If Len(accNum) > 100 Then
						%>
						<script language = 'vbscript'>
							ShowMessage "Account Number can only be 100 characters in length"
						</script>
						<% 
						ReloadPage(ID)
						response.end
					End If
					
					'validate size of Account Name
					If Len(accName) > 100 Then
						%>
						<script language = 'vbscript'>
							ShowMessage "Account Name can only be 100 characters in length"
						</script>
						<%
						ReloadPage(ID)
						response.end
					End If	
				End If
				End If
				
				''2				
				'check for bank account specification
				'validate Bank Branch
				If Trim(bnk2) <> "" Then
				If Trim(bnkBranch2) <> "" Then
					bankAccSpecified2 = true
					
					'validate Account Number
					If Trim(accNum2) = "" Then
						%>
						<script language = 'vbscript'>
							ShowMessage "Please specify the Bank Account Number"
						</script>
						<% 
						ReloadPage(ID)
						response.end
					End If
					
					'validate size of Account Number
					If Len(accNum2) > 100 Then
						%>
						<script language = 'vbscript'>
							ShowMessage "Account Number can only be 100 characters in length"
						</script>
						<% 
						ReloadPage(ID)
						response.end
					End If
					
					'validate size of Account Name
					If Len(accName2) > 100 Then
						%>
						<script language = 'vbscript'>
							ShowMessage "Account Name can only be 100 characters in length"
						</script>
						<%
						ReloadPage(ID)
						response.end
					End If	
				End If
				End If
				
				''3			
				'check for bank account specification
				'validate Bank Branch
				If Trim(bnk3) <> "" Then
				If Trim(bnkBranch3) <> "" Then
					bankAccSpecified3 = true
					
					'validate Account Number
					If Trim(accNum3) = "" Then
						%>
						<script language = 'vbscript'>
							ShowMessage "Please specify the Bank Account Number"
						</script>
						<% 
						ReloadPage(ID)
						response.end
					End If
					
					'validate size of Account Number
					If Len(accNum3) > 100 Then
						%>
						<script language = 'vbscript'>
							ShowMessage "Account Number can only be 100 characters in length"
						</script>
						<% 
						ReloadPage(ID)
						response.end
					End If
					
					'validate size of Account Name
					If Len(accName3) > 100 Then
						%>
						<script language = 'vbscript'>
							ShowMessage "Account Name can only be 100 characters in length"
						</script>
						<%
						ReloadPage(ID)
						response.end
					End If	
				End If
				end if
				
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
				
				If (commission = 0) Or (Len(commission)=0) Or (IsNull(commission)) Then
					commission = 1
				End If
				
					'Enforce ID and CDS uniqueness
				Set conn = GetActiveConnection("KBroker")
				UniqueSql="Select ClientIDPass,ClientCDSNo From Client" & _
				" WHERE ltrim(rtriM(ClientIDPass)) like '" & Trim(replace(idPass,"'","''")) & "'"

				Set rsUnique = Server.CreateObject("ADODB.Recordset")
				rsUnique.CursorLocation = adUseClient 
				conn.BeginTrans
					set rsUnique=Conn.execute(UniqueSql)
				conn.CommitTrans

			
				if not(rsUnique.EOF and rsUnique.BOF) then
					
						if(Trim(idPass)=Trim(rsUnique("ClientIDPass")) ) then%>
						
						   <script language = 'vbscript'>
						      ShowMessage "Client ID Or PassPort Number must be unique"						
						   </script>
						   <% 
						   set rsUnique = nothing
						   Set conn=nothing
						   ReloadPage(ID)
						response.end
						
						end if
				end if

			'UniqueSql =	"SELECT     RTRIM(LTRIM(ClientCDSNo)) AS CDS " & _
			'		 " FROM         Client " & _
			'		 " WHERE     (RTRIM(LTRIM(ClientCDSNo)) LIKE '" &  trim(CDSNo) & "') and (deleted<>1)"

			'response.write UniqueSql:response.end

			'Set rsUnique2 = Server.CreateObject("ADODB.Recordset")
			'rsUnique2.CursorLocation = adUseClient 
			'conn.BeginTrans
			'	set rsUnique2=Conn.execute(UniqueSql)
			'conn.CommitTrans
			'if not(rsUnique2.EOF and rsUnique2.BOF) then
			'		if(Trim(CDSNo)=Trim(rsUnique2("CDS")) ) then%>
		
					   <script language = 'vbscript'>
						  'ShowMessage "CSD Number Must be Unique"						
					   </script>
					   <% 
					   'set rsUnique = nothing
					   'Set conn=nothing
					   'ReloadPage(ID)
					'	response.end
				   'end if
				
			'end if 
				
				
				set guid = server.createobject("NDUtils.CGUID")
				guidStr = guid.GenerateGUID	
						
				sqlStr = "INSERT INTO [Client] (ClientAddr,ClientBDate,ClientCellTel,ClientContact,ClientEmail " & _
						",ClientFax,ClientComment,GenericSetting_DPA_,GenericSetting_DPA_2,GenericSetting_DPA_3,ClientPAddr" & _
						",ClientHomeTel,ClientIDPass,ClientName,ClientOfficeTel,ClientPhoto,ClientSignature" & _
						",ClientVIP,Client_DPA_,Agent_DPA_,Branch_DPA_,Class_DPA_,Commission_DPA_,Gender_DPA_" & _
						",ClientOpeningBal, Owner_DPA_,Residency_DPA_,Client_EIT_, CreditLimit, IsCustodian,updateOnDebt,updateOnContract,UseContactNameInPortfolioReports,ChangedBY) SELECT " & "'" & replace(addr ,"'","''") & "'" & " as ClientAddr," & "#" & FormatDate(bDate) & "#" & " as ClientBDate" & _
						"," & "'" & replace(cell,"'","''") & "'" & " as ClientCellTel" & _
						"," & "'" & replace(contact,"'","''") & "'" & " as ClientContact" & _
						"," & "'" & replace(email,"'","''") & "'" & " as ClientEmail," & "'" & fax & "'" & " as ClientFax" & _
						"," & "'" & replace(comment,"'","''") & "'" & " as ClientComment" & _
						"," & " " & gen1 & " " & " as GenericSetting_DPA_" & _
						"," & " " & gen2 & " " & " as GenericSetting_DPA_2" & _
						"," & " " & gen3 & " " & " as GenericSetting_DPA_3" & _
						"," & "'" & replace(pAddr,"'","''") & "'" & " as ClientPAddr" & _
						"," & "'" & home & "'" & " as ClientHomeTel" & _
						"," & "'" & replace(idPass,"'","''") & "'" & " as ClientIDPass," & "'" & replace(name,"'","''") & "'" & " as ClientName" & _
						"," & "'" & Office & "'" & " as ClientOfficeTel" & _
						"," & "'" & photoNew & "'" & " as ClientPhoto," & "'" & sigNew & "'" & " as ClientSignature" & _ 
						"," & "'" & vip & "'" & " as ClientVIP" & _
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
						"," & "" & Iscustodian & "" & " as IsCustodian" & _
						"," & "" & smsDebit & "" & " as updateOnDebt" & _
						"," & "" & smsContract & "" & " as updateOnContract" & _
						"," & "" & UseContactNameInPortfolioReports & "" & " as UseContactNameInPortfolioReports" & _
						"," & "" & session("UserID") & "" & " as ChangedBy" & _
						" FROM [Client]"

						'Response.Write sqlStr
						'response.end
						
				Set conn = GetActiveConnection("KBroker")
				
				conn.BeginTrans
					
					conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
						
					'obtain header key value
					sqlStr = "SELECT [Client_DPA_] FROM [Client] WHERE [Client_EIT_] = " & "'" & guidStr & "'"
	        
					Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
					If (rs.EOF Or rs.BOF) Then
						%>
						<script language = 'vbscript'>
					    	ShowMessage "A serious error has been encountered while saving the data. Try saving again"
						</script>
						<%
						response.end
					End If
					
					''1
					if bankAccSpecified then
						'save detail data
						sqlStr = "INSERT INTO BankAcc (BankAccNumber,BankAccName,BnkBranch,Client_DPA_,BankNo,Bank_DPA_)" & _
						" VALUES('" & accNum & "', '" & replace(accName,"'","''") & "', '" & replace(bnkBranch,"'","''") & "', " & rs.Fields("Client_DPA_") & ", 1,"& bnk &")"
								
						'Response.Write sqlstr
						'Response.End 
								
						conn.Execute sqlStr
					end if
						
					''2
					if bankAccSpecified2 then
						'save detail data
						sqlStr = "INSERT INTO BankAcc (BankAccNumber,BankAccName,BnkBranch,Client_DPA_, BankNo,Bank_DPA_)" & _
						" VALUES('" & accNum2 & "', '" & replace(accName2,"'","''") & "', '" & replace(bnkBranch2,"'","''") & "', " & rs.Fields("Client_DPA_") & ", 2,"& bnk2 &")"

						conn.Execute sqlStr
					end if
						
					''3
					if bankAccSpecified3 then
						'save detail data
						sqlStr = "INSERT INTO BankAcc (BankAccNumber,BankAccName,BnkBranch,Client_DPA_, BankNo,Bank_DPA_)" & _
						" VALUES('" & accNum3 & "', '" & replace(accName3,"'","''") & "', '" & replace(bnkBranch3,"'","''") & "', " & rs.Fields("Client_DPA_") & ", 3,"& bnk3 &")"
								
						conn.Execute sqlStr
					end if
						
				conn.CommitTrans
				conn.Close
				
				Set conn = Nothing
				
				WritefraEnabledDialogCloseScript2
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
%>

<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
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
			hideButton();
			forceSubmit();

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
		function hideButton()
		{
		  document.getElementById('hide').style.display='none';
		 
		}
		function unhideButton()
		{
		 document.getElementById('hide').style.display='';
		 
		}

		function forceSubmit()
		{
			//give the popup a handle to the list window for refresh purpose
			//alert(window.self.name);
			//alert(window.dialogArguments.opener.location);
			//alert(this.opener.location);
			//this.opener = window.parent.dialogArguments.opener;
			window.self.opener = window.dialogArguments.opener;
			//alert();
			//document.getElementById('hiddenAdd').value='Save';	
			//this.document.all.item("frmAddClient").target = "_self";
			//this.document.all.item("frmAddClient").method = "post";
			//alert();
				//document.write (window.dialogArguments.opener.document.all.item("frmMain").outerHTML);
				//this.document.all.item("frmAddClient").submit();
				//window.self.location.replace(ActionWin.location);
			
			document.frmAddClient.target="_self";
			document.frmAddClient.submit();	
			
		}
	function setOpener()
	{
		window.self.opener = window.dialogArguments.opener;
	}

		
</script>
</head>

<body Class="Dialog" >
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>

<SCRIPT language="JavaScript">
	var cal=new ctlSpiffyCalendarBox("cal", "frmAddClient", "txtBDate","cmdDate","<%= FormatDate(Date) %>",1);
</SCRIPT>
<form name = 'frmMain' method = 'post' action = 'AddClient.asp' id = "frmAddClient" target="_self">
<table border="0" width="60%" cellspacing="0" cellpadding="0">
  
	<tr>
		<td width="10%" nowrap>&nbsp;</td>
		<td width="10%" nowrap>&nbsp;</td>
		<td width="10%" nowrap align=right>Photo</td>
		<td width="10%" nowrap rowspan="8">
			<input type="hidden" name="txtPhoto">
			<IFRAME FRAMEBORDER=0 SCROLLING=NO SRC="upload.asp?filetext=txtPhoto" width="170px" height="170px" tabIndex='-1'></IFRAME>
		</td>
	</tr>	 
	<tr style ="display:none;">
		<td width="10%" nowrap>CSD Number</td>
		<td><input type="text" name="txtCDSNumber" id="txtCDSNumber"></td>
	</tr>  
	<tr>
		<td width="10%" nowrap>Branch</td>
		<td><select name = 'cboBranch' tabIndex='1' id = 'cboBranch' size="1">
		<option selected value = ''></option>
			<%
			Set conn = GetActiveConnection("KBroker")
				        
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
		</select></td>
		<td width="10%" nowrap>&nbsp;</td>
	</tr>
	
	<tr>
		<td width="10%" nowrap>Class</td>
		<td width="10%" nowrap><select name = 'cboClass' tabIndex='2' id = 'cboClass' size="1" onchange='javascript:DisplayDefaultCommission(this);'>
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
		</select></td>
		<td width="10%" nowrap>&nbsp;</td>
	</tr>
  
	<tr>
		<td width="10%" nowrap>Commission</td>
		<td width="10%" nowrap><select style="width:200px" name = 'cboCommission' tabIndex='3' id = 'cboCommission' size="1">
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
		</select></td>
		<td width="10%" nowrap>&nbsp;</td>
	</tr>
  
	<tr>
		<td width="10%" nowrap>Residency</td>
		<td width="10%" nowrap><select name = 'cboResidency' tabIndex='5' id = 'cboResidency' size="1">
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
		</select></td>
		<td width="10%" nowrap>&nbsp;</td>
	</tr>
  
	<tr>
		<td width="10%" nowrap>Gender</td>
		<td width="10%" nowrap><select name = 'cboGender' tabIndex='4' id = 'cboGender' size="1">
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
		</select></td>
		<td width="10%" nowrap>&nbsp;</td>
	</tr>
  
	<tr>
		<td width="10%" nowrap>Agent</td>
		<td width="10%" nowrap><select name = 'cboAgent' tabIndex='6' id = 'cboAgent' size="1">
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
		</select></td>
		<td width="10%" nowrap>&nbsp;</td>
	</tr>
  
	<tr>
		<td width="10%" nowrap>Account Manager</td>
		<td width="10%" nowrap><select name = 'cboOwner' tabIndex='7' id = 'cboOwner' size="1">
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
		</select></td>
		<td width="10%" nowrap>&nbsp;</td>
	</tr>
  
	<tr>
		<td width="10%" nowrap>Name</td>
		<td width="10%" nowrap><input type = 'text' name ='txtName' STYLE="WIDTH: 220px" tabIndex='8' id = 'txtName'></td>
		<td width="10%" nowrap align=right>ID/Passport</td>
		<td width="10%" nowrap><input type = 'text' name ='txtIDPass' tabIndex='17' id = 'txtIDPass' size="20"></td>
	</tr>
	
	<tr>
		<td width="10%" nowrap>Contact Name</td>
		<td width="10%" nowrap><input type = 'text' name ='txtContact' STYLE="WIDTH: 220px" tabIndex='9' id = 'txtContact' size="20"></td>
		<td width="10%" nowrap align=right>Date of Birth</td>
		<td width="10%" nowrap><SCRIPT language="JavaScript">cal.writeControl();  document.all.item('txtBDate').tabIndex='18'</SCRIPT></td>
	</tr>
  
	<tr>
		<td width="10%" nowrap>Cell Phone</td>
		<td width="10%" nowrap><input type = 'text' name ='txtCellTel' tabIndex='10' id = 'txtCellTel' size="20" value=""></td>
		<td width="10%" nowrap align=right>Credit Limit&nbsp;&nbsp;&nbsp;</td>
		<td width="10%" nowrap>
		<input type = 'text' name ='txtCreditLimit' tabIndex='19' STYLE="TEXT-ALIGN: RIGHT; WIDTH: 100PX" id = 'txtCreditLimit' size="20" value="0">
		<select style="display: none" name = 'cboVIP' tabIndex='19' id = 'cboVIP' size="1">
		<option value = '1'>Yes</option>
		<option value = '0' selected>No</option>
		</select></td>
	</tr>
	  
	<tr>
		<td width="10%" nowrap>Home Phone</td>
		<td width="10%" nowrap><input type = 'text' name ='txtHomeTel' tabIndex='11' id = 'txtHomeTel' size="20"></td>
		<td width="10%" nowrap align=right>Opening Balance&nbsp;&nbsp;&nbsp;</td>
		<td width="10%" nowrap><input class=readonly readonly=true  name ='txtOpeningBal' tabIndex='20' STYLE="TEXT-ALIGN: RIGHT; WIDTH: 100PX" id = 'txtOpeningBal' size="20" value="0"></td>
	</tr>
	  
	<tr>
		<td width="10%" nowrap>Office Phone</td>
		<td width="10%" nowrap><input type = 'text' name ='txtOfficeTel' tabIndex='12' id = 'txtOfficeTel' size="20"></td>
		<td width="10%" nowrap align=right>Notes/Comments&nbsp;&nbsp;&nbsp;</td>
		<td width="10%" nowrap rowspan=2><TEXTAREA name ='txtComment' tabIndex='21' id = 'txtComment' rows='2' cols="25"></TEXTAREA></td>
	</tr>
	
	<tr>
		<td width="10%" nowrap>Email</td>
		<td width="10%" nowrap><input type = 'text' name ='txtEmail' STYLE="WIDTH: 220px" tabIndex='13' id = 'txtEmail' size="20"></td>
		<td width="10%" nowrap>&nbsp;</td>
	</tr>  
	  
	<tr>
		<td width="10%" nowrap valign="top">Address</td>
		<td width="10%" nowrap><textarea rows=3 name ='txtAddr' tabIndex='15' id = 'txtAddr' size="35" cols="33"></textarea></td>
		<td width="10%" nowrap align=right valign="top">Physical Address &nbsp;&nbsp;&nbsp;</td>
		<td width="10%" nowrap><textarea rows=3 name ='txtPAddr' tabIndex='22' id = 'txtPAddr' size="35" cols="25"></textarea></td>
	</tr>
	  
	<tr>
		<td width="10%" nowrap>Fax</td>
		<td width="10%" nowrap><input type = 'text' name ='txtFax' tabIndex='14' id = 'txtFax' size="20">&nbsp;&nbsp;CDA &nbsp;<input type=checkbox   value='False' name='chkCustodian' onClick = 'UpdateCustodianStatus(this);'></td>
		<td width="10%" nowrap align=right>Signature</td>
		<td width="10%" nowrap rowspan="6">
		<input type="hidden" name="txtSignature">
		<IFRAME FRAMEBORDER=0 SCROLLING=NO SRC="upload.asp?filetext=txtSignature" width="170px" height="170px" tabIndex='-1'></IFRAME>
		</td>
	</tr>
  
	<tr>
	  <td width="10%" nowrap>SMS Contract&nbsp;<input type="checkbox" name="SmsContract" value="1"></td>
	  <td width="10%" nowrap>SMS Debit&nbsp;<input type="checkbox" name="SmsDebit" value="1"></td>
	  <td width="10%" nowrap>&nbsp;</td>
	</tr>	
  
	<tr>
		<td width="10%" nowrap valign="top">Generic1</td>
		<td width="10%" nowrap><select name = 'cboGenericSetting1' id = 'cboGenericSetting1' size="1">
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
		<td width="10%" nowrap>&nbsp;</td>
	</tr>	
  
	<tr>
		<td width="10%" nowrap valign="top">Generic2</td>
		<td width="10%" nowrap><select name = 'cboGenericSetting2' id = 'cboGenericSetting2' size="1">
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
		<td width="10%" nowrap>&nbsp;</td>
	</tr>
  
	<tr>
		<td width="10%" nowrap valign="top">Generic3</td>
		<td width="10%" nowrap><select name = 'cboGenericSetting3' id = 'cboGenericSetting3' size="1">
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
		<td width="10%" nowrap>&nbsp;</td>
	</tr>
  
	<tr>
		<td width="10%" nowrap height="50">Use contact name<br>in Portfolio Reports</td>
		<td width="10%" nowrap><input type="checkbox" name="chkcontact" id="chkcontact" value="1"></td>
		<td width="10%" nowrap>&nbsp;</td>
	</tr>
  
	<tr>
		<!--BANK DETAILS-->
		<td colspan=4>
			<b>Bank Account Details</b>
			<table border=0 style="border: 1px ridge #000080" WIDTH="50%">
				<tr>
					<td>
						<!--BANK 1-->
						<table border=0 style="border: 1px ridge #000080" WIDTH="100%">	
							<tr>
								<td width="30%">Bank</td>
								<td width="70%"><select name = 'cboBank' id = 'cboBank' size="1" tabIndex='23'>
								<option selected value = ''></option>
								<%
								sqlStr = "SELECT * FROM [Bank]"
								Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
								If Not (rs.EOF Or rs.BOF) Then
									rs.MoveFirst
									Do Until rs.EOF
											%>
											<option value = '<%=rs.Fields("Bank_DPA_")%>'><%=rs.Fields("BankName")%></option>
											<%
										rs.MoveNext
									Loop
								End If
								%>
								</select></td>
							</tr>
							<tr>
								<td width="30%">Branch</td>
								<td width="70%"><input type = 'text' name ='txtBnkBranch' id = 'txtBnkBranch' size="20" tabIndex='24'></td>
							</tr>
							<tr>
								<td width="30%">Account Name</td>
								<td width="70%"><input type = 'text' name ='txtAccName' id = 'txtAccName' size="20" tabIndex='25'></td>
							</tr>
							<tr>
								<td width="30%">Account Number</td>
								<td width="70%" valign="top"><input type = 'text' name ='txtNumber' id = 'txtNumber' size="20" tabIndex='26'></td>
							</tr>
						</table>
					</td>
				</tr>  
		
				<tr>
					<td>
						<!--BANK 2-->
						<table border=0 style="border: 1px ridge #000080" WIDTH="100%">	
							<tr>
								<td width="30%">Bank</td>
								<td width="70%"><select name = 'cboBank2' id = 'cboBank2' size="1" tabIndex='27'>
								<option selected value = ''></option>
								<%
								sqlStr = "SELECT * FROM [Bank]"
								Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
								If Not (rs.EOF Or rs.BOF) Then
									rs.MoveFirst
									Do Until rs.EOF
											%>
											<option value = '<%=rs.Fields("Bank_DPA_")%>'><%=rs.Fields("BankName")%></option>
											<%
										rs.MoveNext
									Loop
								End If
								%>
								</select></td>
							</tr>
							<tr>
								<td width="30%">Branch</td>
								<td width="70%"><input type = 'text' name ='txtBnkBranch2' id = 'txtBnkBranch2' size="20" tabIndex='28'></td>
							</tr>
							<tr>
								<td width="30%">Account Name</td>
								<td width="70%"><input type = 'text' name ='txtAccName2' id = 'txtAccName2' size="20" tabIndex='29'></td>
							</tr>
							<tr>
								<td width="30%">Account Number</td>
								<td width="70%" valign="top"><input type = 'text' name ='txtNumber2' id = 'txtNumber2' size="20" tabIndex='30'></td>
							</tr>
						</table>
					</td>
				</tr>  
		
				<tr>
					<td>
						<!--BANK 3-->
						<table border=0 style="border: 1px ridge #000080" WIDTH="100%">	
							<tr>
								<td width="30%">Bank</td>
								<td width="70%"><select name = 'cboBank3' id = 'cboBank3' size="1" tabIndex='31'>
								<option selected value = ''></option>
								<%
								sqlStr = "SELECT * FROM [Bank]"
								Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
								If Not (rs.EOF Or rs.BOF) Then
									rs.MoveFirst
									Do Until rs.EOF
											%>
											<option value = '<%=rs.Fields("Bank_DPA_")%>'><%=rs.Fields("BankName")%></option>
											<%
										rs.MoveNext
									Loop
								End If
								%>
								</select></td>
							</tr>
							<tr>
								<td width="30%">Branch</td>
								<td width="70%"><input type = 'text' name ='txtBnkBranch3' id = 'txtBnkBranch3' size="20" tabIndex='32'></td>
							</tr>
							<tr>
								<td width="30%">Account Name</td>
								<td width="70%"><input type = 'text' name ='txtAccName3' id = 'txtAccName3' size="20" tabIndex='33'></td>
							</tr>
							<tr>
								<td width="30%">Account Number</td>
								<td width="70%" valign="top"><input type = 'text' name ='txtNumber3' id = 'txtNumber3' size="20" tabIndex='34'></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>

</table>

<table align=right>
	<tr>
	
	  <td width="100%" align="right" valign=absBottom >
		<b id="hide" name="hide"> <input type = 'button' Class='Buttons' name ='cmdAdd' id = 'cmdAdd' value="Save" tabIndex='35' onclick = "AllowedNavigation();"></b>
    	<input type = 'button' Class="Buttons" name ='cmdAdd' id = "cmdCancel" value=" Cancel " tabIndex='36' onclick = "JavaScript: window.self.close()">
    	<input type = 'hidden' name ='custodianStatus' id = 'custodianStatus' size="20" value='0'>
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		<input type = 'hidden' name ='buttonAction' id = 'action' value="Save">
		
	</td>
  
	</tr>	
</table>

</form>
</body>

</html>
