<!--#include file="../libroutines.asp"-->
<%
	const LinkedIndependent = 1
	const LinkedDependent = 2
	
	Dim action
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim ID
	Dim rsEdit
	Dim bankAccExists
	Dim bankAccExists2
	Dim bankAccExists3
	

	
	action = ucase(Request.Form("action"))
	ID = Request("ID")

'ID=2783
'action = "EXECUTE"              

	If Trim(ID) = "" Then
		%>
		<script language = 'vbscript'>
				window.self.ShowMessage "No record specified for editing"
				window.self.close
		</script>
		<%
		response.end
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
				
				Dim OnlineRegistration
				Dim IsCustodian
				Dim UseContactNameInPortfolioReports

				Dim hiddenID
				Dim hiddenCDS
				
				bankAccSpecified = false
				bankAccSpecified2 = false
				bankAccSpecified3 = false
				
				''1
				bnk = Request.Form("cboBank")
				bnkBranch = trim(replace(Request.Form("txtBnkBranch"),"'",""))
				accNum = trim(replace(Request.Form("txtNumber"),"'",""))
				accName = trim(replace(Request.Form("txtAccName"),"'",""))
				
				''2
				bnk2 = Request.Form("cboBank2")
				bnkBranch2 = trim(replace(Request.Form("txtBnkBranch2"),"'",""))
				accNum2 = trim(replace(Request.Form("txtNumber2"),"'",""))
				accName2 = trim(replace(Request.Form("txtAccName2"),"'",""))
				
				''3
				bnk3 = Request.Form("cboBank3")
				bnkBranch3 = trim(replace(Request.Form("txtBnkBranch3"),"'",""))
				accNum3 = trim(replace(Request.Form("txtNumber3"),"'",""))
				accName3 = trim(replace(Request.Form("txtAccName3"),"'",""))
				
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
				Signature = Request.Form("txtSignature")
				vip = Request.Form("cboVIP")
				gen1 = Request.Form("txtGeneric1")
				gen2 = Request.Form("txtGeneric2")
				pAddr = Request.Form("txtPAddr")
				OpeningBal = Request.Form("txtOpeningBal")
				CreditLimit = Request.Form("txtCreditLimit") 
				'CDSNo = Request.Form("txtCDSNo")
				institution = Request.Form("cboGenericSetting1")
				gen2 = Request.Form("cboGenericSetting2")
				gen3 = Request.Form("cboGenericSetting3") 
				OnlineRegistration  = Cint(Request.Form("OnlineRegistration")) 
				IsCustodian = Request.Form("custodianStatus") 
				'CDSNumber = Request.Form("txtCDSNumber")
				hiddenID = trim(Request.Form("hiddenID"))
				hiddenCDS = trim(Request.Form("hiddenCSD"))

				'response.write hiddenCDS & ": " & CDSNumber: response.end
				bankAccExists = Request.Form("bankAccExists") 
				bankAccExists2 = Request.Form("bankAccExists2") 
				bankAccExists3 = Request.Form("bankAccExists3") 
				
				smsContract = Request.Form("smsContract") 
				smsDebit = Request.Form("smsDebit") 
				UseContactNameInPortfolioReports = Request.Form("chkcontact")
				
				if trim(smsDebit)="" then smsDebit=0
				if trim(smsContract)= "" then smsContract=0
				if trim(UseContactNameInPortfolioReports) = "" then UseContactNameInPortfolioReports = 0
				
				'validate CDS
				If Trim(CDSNo) = "" Then
				End If				
				'validate Branch
				If Trim(branch) = "" Then%>
						<script language = 'vbscript'>
                				ShowMessage "Please specify the Branch"
						</script>
						<% 
						ReloadPage(ID)
						response.end
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
						<% 
						ReloadPage(ID)
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
						ReloadPage(ID)
						response.end
				End If
				'validate Name
				If Trim(Name) = "" Then%>
						<script language = 'vbscript'>
                				ShowMessage "Please specify the Name"
						</script>
						<% 
						ReloadPage(ID) 
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
						<% 
						ReloadPage(ID)
						response.end
				End If
				'validate size of Cell Phone
				If Len(Cell) > 100 Then%>
						<script language = 'vbscript'>
						ShowMessage "Cell Phone can only be 100 characters in length"
						</script>
						<% response.end
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
						<% 
						ReloadPage(ID)
						response.end
				End If
				'validate size of Email
				If Len(Email) > 100 Then%>
						<script language = 'vbscript'>
						ShowMessage "Email can only be 100 characters in length"
						</script>
						<% 
						ReloadPage(ID)
						response.end
				End If
				
				If trim(Email) = "" or len(trim(Email))=0  Then%>
						<script language = 'vbscript'>
						ShowMessage "Please Enter an Email Address"
						</script>
						<% 
						ReloadPage(ID)
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
						<% 
						ReloadPage(ID)
						response.end
				End If
				
				'validate size of Fax
				If Len(Fax) > 100 Then%>
						<script language = 'vbscript'>
						ShowMessage "Fax can only be 100 characters in length"
						</script>
						<% 
						ReloadPage(ID)
						response.end
				End If
				'validate Generic1
				If Trim(institution) = "" Then
						institution = "NULL"
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
						<% 
						ReloadPage(ID)
						response.end
				End If
				'validate size of Home Phone
				If Len(home) > 100 Then%>
						<script language = 'vbscript'>
						ShowMessage "Home Phone can only be 100 characters in length"
						</script>
						<% 
						ReloadPage(ID)
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
						<% 
						ReloadPage(ID)
						response.end
				End If

				If trim(IDPass) = "" or len(trim(IDPass))= 0 Then%>
						<script language = 'vbscript'>
						ShowMessage "Please Specify ID/Passport"
						</script>
						<% 
						ReloadPage(ID)
						response.end
				End If
				'validate size of Name
				If Len(Name) > 100 Then%>
						<script language = 'vbscript'>
						ShowMessage "Name can only be 100 characters in length"
						</script>
						<% 
						ReloadPage(ID)
						response.end
				End If
				'validate size of Office Phone
				If Len(Office) > 100 Then%>
						<script language = 'vbscript'>
						ShowMessage "Office Phone can only be 100 characters in length"
						</script>
						<% 
						ReloadPage(ID)
						response.end
				End If
				if trim(Office)<>"" then
					if not isnumeric(replace(replace(trim(Office),"+","")," " ,""))  then %>
							<script language = 'vbscript'>
							ShowMessage "Office Phone number should contain numeric numbers only"
							
							</script>
							<% ReloadPage(ID)
							response.end

					end if 
				end if
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
				end if
				
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
				end if
				
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

				'*************************************************
					'Enforce ID and CDS uniqueness
				'*************************************************
				Set conn = GetActiveConnection("KBroker")
				'*** CDS uniqueness**************
				'if hiddenCDS <> trim(CDSNumber) then
					'response.write hiddenCDS & " : " & CDSNumber & vbcrlf					
					'response.write "Hapa"
					'response.end
					'UniqueSql =	"SELECT     RTRIM(LTRIM(ClientCDSNo)) AS CDS " & _
					'	 " FROM         Client " & _
					'	 " WHERE     (RTRIM(LTRIM(ClientCDSNo)) LIKE '" &  trim(CDSNumber) & "') and (deleted<>1)"

					'response.write UniqueSql:response.end

					'Set rsUnique2 = Server.CreateObject("ADODB.Recordset")
					'rsUnique2.CursorLocation = adUseClient 
					'conn.BeginTrans
					'	set rsUnique2=Conn.execute(UniqueSql)
					'conn.CommitTrans
					'if not(rsUnique2.EOF or rsUnique2.BOF) then
					'		if(Trim(CDSNumber)=Trim(rsUnique2("CDS"))) then%>				
							   <script language = 'vbscript'>
					'			  ShowMessage "CSD Number Must be Unique"
								  'window.history.go(-1)
							   </script>
							   <% 
					'		   set rsUnique = nothing
					'		   Set conn=nothing
					'		   ReloadPage(ID)
					'		   response.end
					'	   end if				
					'end if 

				'end if'Condition
				'response.end
				
				'*** ID uniqueness**************
				
				if hiddenID <> trim(IDPASS) then
					UniqueSql="Select ClientIDPass,ClientCDSNo From Client" & _
					" WHERE ltrim(rtriM(ClientIDPass)) like '" & Trim(idPass) & "'"
					'response.write UniqueSql :response.end
					Set rsUnique = Server.CreateObject("ADODB.Recordset")
					rsUnique.CursorLocation = adUseClient 
					conn.BeginTrans
						set rsUnique=Conn.execute(UniqueSql)
					conn.CommitTrans

				
					if not(rsUnique.EOF or rsUnique.BOF) then						
							if(Trim(idPass)=Trim(rsUnique("ClientIDPass")) ) then%>							
							   <script language = 'vbscript'>
								  ShowMessage "Client ID Or PassPort Number must be unique"
								   'window.history.go(-1)
							   </script>
							   <% 
							   set rsUnique = nothing
							   Set conn=nothing
							   ReloadPage(ID)
							   response.end							
							end if
					end if
				end if 
				


			'****************************************
				
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


				
				
				'Confirm Online clients as Registered Clients
				 if OnlineRegistration = "" then OnlineRegistration = 0 'Default value
				
				Set conn = GetActiveConnection("KBroker")
				
				If (commission = 0) Or (Len(commission)=0) Or (IsNull(commission)) Then
					commission = 1
				End If


				
				
				'save data
				sqlStr = "UPDATE [Client] SET ClientAddr = " & "'" &  replace(trim(addr),"'","''") & "'" & ",ClientBDate = " & "'" & FormatDate(bDate) & "'" & "" & _
						",ClientCellTel = " & "'" & replace(cell,"'","''") & "'" & ",ClientContact = " & "'" & replace(contact,"'","''") & "'" & "" & _
						",ClientEmail = " & "'" & replace(email,"'","''") & "'" & ",ClientFax = " & "'" & replace(fax,"'","") & "'" & "" & _
						",Institution_DPA_ = " & " " & institution & " " & _
						",GenericSetting_DPA_2 = " & " " & gen2 & " " & "" & _
						",GenericSetting_DPA_3 = " & " " & gen3 & " " & "" & _
						",ClientPAddr = " & "'" & replace(pAddr,"'","''") & "'" & "" & _
						",ClientComment = " & "'" & replace(comment,"'","''") & "'" & "" & _
						",ClientHomeTel = " & "'" & replace(home,"'","''" )& "'" & ",ClientIDPass = " & "'" & replace(idPass,"'","") & "'" & "" & _
						",ClientName = " & "'" & replace(name,"'","''") & "'" & ",ClientOfficeTel = " & "'" & replace(Office,"'","") & "'" & "" & _
						",ClientPhoto = " & "'" & photoNew & "'" & "" & _ 
						",ClientSignature = " & "'" & sigNew & "'" & "" & _ 
						",IsCustodian = " & " " & IsCustodian & " " & "" & _
						",updateOnDebt = " & " " & smsDebit & " " & "" & _
						",updateOnContract = " & " " & smsContract & " " & "" & _
						",OnlineRegistration = " & " " & OnlineRegistration & " " & "" & _
						",ChangedBy = " & " " & session("UserID") & " " & "" & _
						",TimeChanged = " & " GetDate() " & "" & _ 
						",ClientVIP = " & " " & vip & " " & "" & _
						",CreditLimit = " & " " & CreditLimit & " " & "" & _
						",UseContactNameInPortfolioReports = " & " " & UseContactNameInPortfolioReports & " " & "" & _
						",Agent_DPA_ = " & " " & agent  & " " & ",Branch_DPA_ = " & " " & branch & " " & "" & _
						",Class_DPA_ = " & " " & classType & " " & ",Commission_DPA_ = " & " " & commission & " " & "" & _
						",Gender_DPA_ = " & " " & gender & " " & ",ClientOpeningBal = " & CDbl(OpeningBal) & ", Owner_DPA_ = " & " " & owner & " " & "" & _
						",Residency_DPA_ = " & " " & residency & " " & " WHERE Client_DPA_  = " & ID
				
				'response.write handlequote(addr)
				'response.end
				conn.BeginTrans
				
					conn.Execute SQLServerFormat(sqlStr)
					
					''1						
					if (not(bankAccExists)) and (bankAccSpecified) then
						'new account
						sqlStr = "INSERT INTO BankAcc (BankAccNumber,BankAccName,BnkBranch,Client_DPA_,BankNo,Bank_DPA_)" & _
						" VALUES('" & accNum & "', '" & accName & "', '" & bnkBranch & "', " & ID & ", 1,"& bnk &")"
					elseif (not(bankAccExists)) and (not(bankAccSpecified)) then
						'no change
						sqlStr = ""
					elseif (bankAccExists) and (not(bankAccSpecified)) then
						'remove account
						sqlStr = "DELETE FROM [BankAcc] WHERE Client_DPA_="& ID &" AND BankNo=1"
					elseif (bankAccExists) and (bankAccSpecified) then
						'update account
						sqlStr = "UPDATE [BankAcc] SET BankAccNumber = " & "'" & replace(accNum,"'","''")  & "'" & _
							",BankAccName = " & "'" & replace(accName,"'","''") & "'" & _
							",BnkBranch = " & "'" & replace(bnkBranch,"'","''")  & "'" & _
							",Bank_DPA_ = " & " " & bnk & " " & _
							" WHERE Client_DPA_="& ID &" AND BankNo=1"
					end if
							'response.write SQLServerFormat(sqlStr):response.end					
					if sqlStr <> "" then
						conn.Execute SQLServerFormat(sqlStr)
					end if 
					
					''2	
					if (not(bankAccExists2)) and (bankAccSpecified2) then
						'new account
						sqlStr = "INSERT INTO BankAcc (BankAccNumber,BankAccName,BnkBranch,Client_DPA_,BankNo,Bank_DPA_)" & _
						" VALUES('" & replace(accNum2 ,"'","''")& "', '" & replace(accName2,"'","''") & "', '" & replace( bnkBranch2,"'","''") & "', " & ID & ", 2,"& bnk2 &")"
					elseif (not(bankAccExists2)) and (not(bankAccSpecified2)) then
						'no change
						sqlStr = ""
					elseif (bankAccExists2) and (not(bankAccSpecified2)) then
						'delete from database
						sqlStr = "DELETE FROM [BankAcc] WHERE Client_DPA_="& ID &" AND BankNo=2"
					elseif (bankAccExists2) and (bankAccSpecified2) then
						'update account
						sqlStr = "UPDATE [BankAcc] SET BankAccNumber = " & "'" & replace(accNum2,"'","''") & "'" & _
							",BankAccName = " & "'" & replace(accName2,"'","''") & "'" & _
							",BnkBranch = " & "'" & replace(bnkBranch2,"'","''")& "'" & _
							",Bank_DPA_ = " & " " & bnk2 & " " & _
							" WHERE Client_DPA_="& ID &" AND BankNo=2"
					end if
												
					if sqlStr <> "" then
						conn.Execute SQLServerFormat(sqlStr)
					end if 
					
					''3						
					if (not(bankAccExists3)) and (bankAccSpecified3) then
						'new account
						sqlStr = "INSERT INTO BankAcc (BankAccNumber,BankAccName,BnkBranch,Client_DPA_,BankNo,Bank_DPA_)" & _
						" VALUES('" & replace(accNum3,"'","''") & "', '" & replace(accName3,"'","''") & "', '" & replace(bnkBranch3,"'","''") & "', " & ID & ", 3,"& bnk3 &")"
					elseif (not(bankAccExists3)) and (not(bankAccSpecified3)) then
						'no change
						sqlStr = ""
					elseif (bankAccExists3) and (not(bankAccSpecified3)) then
						'remove account
						'delete from database
						sqlStr = "DELETE FROM [BankAcc] WHERE Client_DPA_="& ID &" AND BankNo=3"
					elseif (bankAccExists3) and (bankAccSpecified3) then
						'update account
						sqlStr = "UPDATE [BankAcc] SET BankAccNumber = " & "'" & replace(accNum3,"'","''") & "'" & _
							",BankAccName = " & "'" & replace(accName3,"'","''") & "'" & _
							",BnkBranch = " & "'" & replace(bnkBranch3,"'","''") & "'" & _
							",Bank_DPA_ = " & " " & bnk3 & " " & _
							" WHERE Client_DPA_="& ID &" AND BankNo=3"
					end if
											
					if sqlStr <> "" then
						conn.Execute SQLServerFormat(sqlStr)
					end if 
				conn.CommitTrans
				
				Set conn = Nothing
				 WritefraEnabledDialogCloseScript2
				'WritefraEnabledDialogCloseScript
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
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
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
		
		function DoInit() 
		{
			if (window.dialogArguments != null)
			{
				window.name = "editWindow";
			}
    
			else
			{
				window.close();
			}    
    
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
		
	function drawPicLayout(){
		try
		{
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
</script>
<script language='javascript'>
	function forceSubmit()
	{
		//alert(window.parent.opener.name);
		//Set window.parent.dialogArguments.opener.parent.frames("footer").editDocOpener = window.self;
		//document.frmEditClient.target="_self";
		//document.frmEditClient.submit();
		window.self.opener = window.dialogArguments.opener;
		document.frmEditClient.target="_self";
		document.frmEditClient.submit();
		/*		
		//prepare list window for submit
		var targetPage = window.dialogArguments.opener.document.all.item("frmMain").elements("EditPage").value
		
		//window.dialogArguments.opener.document.all.item("frmMain").action = targetPage;
		window.dialogArguments.opener.document.all.item("frmMain").target = self;
		//document.write (window.dialogArguments.opener.document.all.item("frmMain").outerHTML);
		window.dialogArguments.opener.document.all.item("frmMain").submit();*/
		
	}

</script>



</head>

<body Class="Dialog" OnLoad="JavaScript: DoInit();">

<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<form name = 'frmEditClient' method = 'post' action = 'EditClient.asp' id = "frmMain" target="_self">

<%
 

 Dim CredlimitFlag
CredlimitFlag = GetUserGroup(Session("UserID")) 			
		
Function GetUserGroup(varUserID)
'This Function is basicaly used to determine the 
'Usergroup of the current user. We are interested 
'with the managaement level users because the credit limit should
'be editable by guys in mgt level only	
	Set conn = GetActiveConnection("KBroker")
	sqlStr = "SELECT * FROM UsersGroups WHERE (UserID = " & varUserID & ") AND (LOWER(rtrim(ltrim(GroupName))) = 'management')"
	Set Rs = conn.Execute(sqlStr)
	 
	if Rs.eof or Rs.Bof then
		GetUserGroup = "ReadOnly"  'Meaning that the credit limit is disabled for this dude coz he is not in mgt level  
		set conn = Nothing
		set Rs = Nothing
	else 
		GetUserGroup = "" 'can be allowed to edit coz he exists in mgt level 
		set conn = Nothing
		set Rs = Nothing		
	end if
End Function
		Set conn = GetActiveConnection("KBroker")
      
		sqlStr = " SELECT UseContactNameInPortfolioReports, ClientCDSNo,Client.Institution_DPA_,InstitutionName, CreditLimit,  Client.ClientAddr, Client.ClientBDate, Client.ClientCellTel, Client.ClientContact, Client.ClientEmail, Client.ClientComment, Client.ClientFax,  " & _
                  "    Client.GenericSetting_DPA_,Client.GenericSetting_DPA_2, Client.GenericSetting_DPA_3, Client.ClientPAddr, Client.ClientHomeTel, Client.ClientIDPass, Client.ClientName, Client.ClientOfficeTel ,updateOnDebt,updateOnContract,  " & _
                  "    Client.ClientPhoto, Client.ClientSignature, Client.ClientVIP, Client.Client_DPA_, AgentList.Agent_DPA_, BranchList.Branch_DPA_, Client.Class_DPA_, Client.Commission_DPA_,  " & _
                  "    cast(Client.OnlineRegistration as Numeric) as OnlineRegistration, Client.Gender_DPA_, Client.Owner_DPA_, Client.Residency_DPA_, Client.ClientOpeningBal, IsCustodian as IsCustodian  " & _
				"	FROM         OwnerList RIGHT OUTER JOIN " & _
                 "	     GenderList RIGHT OUTER JOIN " & _
                 "	     ResidencyList INNER JOIN " & _
                 "	     CommissionList INNER JOIN " & _
                 "	     ClassList INNER JOIN " & _
				 "       InstitutionList RIGHT OUTER JOIN " & _
                 "	     BranchList INNER JOIN " & _
                  "	    Client LEFT OUTER JOIN " & _

                 "	 AgentList ON AgentList.Agent_DPA_ = Client.Agent_DPA_ ON BranchList.Branch_DPA_ = Client.Branch_DPA_ ON  " & _
				 " InstitutionList.Institution_DPA_ = Client.Institution_DPA_  ON " & _
                 "	     ClassList.Class_DPA_ = Client.Class_DPA_ ON CommissionList.Commission_DPA_ = Client.Commission_DPA_ ON  " & _
                 "	     ResidencyList.Residency_DPA_ = Client.Residency_DPA_ ON GenderList.Gender_DPA_ = Client.Gender_DPA_ ON  " & _
                 "	     OwnerList.Owner_DPA_ = Client.Owner_DPA_" & _
				"	WHERE     Client.Client_DPA_ = " & ID
			
			'Response.Write sqlStr
			'Response.End 	
        sqlStr = SQLServerFormat(HandleQuote(sqlStr))
                
        Set rs = conn.Execute(sqlStr)
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		window.self.ShowMessage "The selected Client cannot be retrieved for editing"
                		
                </script>
                <% response.end
        End If
  
        Photo = Rs.Fields("ClientPhoto")
        If IsNull(Photo) Then
			Photo = "/Data/Photos/_blank.jpg"
		ElseIf Photo = "" Then
			Photo = "/Data/Photos/_blank.jpg"
		End If	
        
         Signature = Rs.Fields("ClientSignature")
         
        If IsNull(Signature) or Signature = "" Then
			Signature = "/Data/Photos/_blank.jpg"
		End If	

%>
<table border="0" width="60%" cellspacing="0" cellpadding="0">

	<tr>
		<td width="10%" nowrap>&nbsp;</td>
		<td width="10%" nowrap>&nbsp;</td>
		<td width="10%" nowrap align="right"><b>Client Code:</b>&nbsp;&nbsp;</td>
		<td width="10%" nowrap>&nbsp;<b><%=rs.Fields("Client_DPA_")%></b></td>
	</tr>
	<tr style="display:none">
		<td width="10%" nowrap>CSD Number</td>
		<td><input type="text" name="txtCDSNumber" id="txtCDSNumber" value="<%=rs.Fields("ClientCDSNo")%>">
		<input type="hidden" name="hiddenCSD" id="hiddenCSD" value="<%=rs.Fields("ClientCDSNo")%>">
		</td>
	</tr>    
	<tr>
		<td width="10%" nowrap>Branch</td>
		<td width="10%" nowrap><select name = 'cboBranch' tabIndex='1' id = 'cboBranch' size="1">
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
		</select></td>
		<td width="10%" nowrap>&nbsp;</td>
		<td rowspan="8">
		<input type="hidden" name="txtPhoto" value="<%= Photo %>">
		<IFRAME FRAMEBORDER=0 SCROLLING=NO SRC="upload.asp?filetext=txtPhoto" width="170px" height="170px" tabIndex='-1'></IFRAME>
		</td>
	</tr>
	
  <tr>
    <td width="20%">Class</td>
    <td width="90%" colspan="2"><select name = 'cboClass' tabIndex='2' id = 'cboClass' size="1" onchange='javascript:DisplayDefaultCommission(this);'>
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

    </select></td>
  </tr>
  <tr>
    <td width="20%">Commission</td>
    <td width="90%" colspan="2"><select name = 'cboCommission' tabIndex='3' id = 'cboCommission' size="1">
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

    </select></td>
  </tr>
  
  <tr>
    <td width="20%">Residency</td>
    <td width="90%" colspan="2"><select name = 'cboResidency' tabIndex='5' id = 'cboResidency' size="1">
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
    </select></td>
  </tr>
  <tr>
    <td width="20%">Gender</td>
    <td width="90%" colspan="2"><select name = 'cboGender' tabIndex='4' id = 'cboGender' size="1">
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

    </select></td>
  </tr>
  <tr>
    <td width="20%">Agent</td>
    <td width="90%" colspan="2"><select name = 'cboAgent' tabIndex='6' id = 'cboAgent' size="1">
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


    </select></td>
  </tr>
  
  <tr>
    <td width="27%" nowrap>Account Manager</td>
    <td width="90%" colspan="2"><select name = 'cboOwner' tabIndex='7' id = 'cboOwner' size="1">
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
    </select></td>
  </tr>
  <tr>
    <td width="20%">Name</td>
    <td width="20%"><input type = 'text' name ='txtName' STYLE="WIDTH: 220px" tabIndex='8' id = 'txtName' value = "<%=rs.Fields("ClientName")%>"></td>
    <td width="30%" align=right valign=top>Photo&nbsp;&nbsp;&nbsp;</td>
    
  </tr>
  <tr>
    <td width="20%">Contact Name</td>
    <td width="20%"><input type = 'text' name ='txtContact' STYLE="WIDTH: 220px" tabIndex='9' id = 'txtContact' size="20" value = "<%=rs.Fields("ClientContact")%>"></td>
    <td width="30%" align=right>ID/Passport&nbsp;&nbsp;&nbsp;</td>
    <td width="30%"><input type = 'text' name ='txtIDPass' tabIndex='17' id = 'txtIDPass' size="20" value = '<%=rs.Fields("ClientIDPass")%>'><input type="hidden" name="hiddenID" value='<%=rs.Fields("ClientIDPass")%>'></td>
  </tr>
  <tr>
    <td width="20%">Cell Phone</td>
    <td width="20%"><input type = 'text' name ='txtCellTel' tabIndex='10' id = 'txtCellTel' size="20" value = '<%=rs.Fields("ClientCellTel")%>'></td>
    <td width="30%" align=right>Date of Birth&nbsp;&nbsp;&nbsp;</td>    
    <td width="36%">
		<script language="JavaScript" src="../CALENDAR/calendar.js"></script>
		<SCRIPT language="JavaScript">
			var cal=new ctlSpiffyCalendarBox("cal", "frmEditClient", "txtBDate","cmdDate","<%= FormatDate(rs.Fields("ClientBDate")) %>",1);		
			cal.writeControl();  
			document.all.item('txtBDate').tabIndex='18'
			
		</SCRIPT>
	</td>
  </tr>
  <tr>
    <td width="20%">Home Phone</td>
    <td width="20%"><input type = 'text' name ='txtHomeTel' tabIndex='11' id = 'txtHomeTel' size="20"  value = '<%=rs.Fields("ClientHomeTel")%>'></td>
    <td width="30%" align=right>Credit Limit&nbsp;&nbsp;&nbsp;</td>
    <td width="36%">
    <input type = 'text' <%=CredlimitFlag %> name ='txtCreditLimit' tabIndex='19' STYLE="TEXT-ALIGN: RIGHT; WIDTH: 100PX" id = 'txtCreditLimit' size="20" value="<%=rs.Fields("CreditLimit")%>" >
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
      </select></td>
  </tr>
  <tr>
    <td width="20%">Office Phone</td>
    <td width="20%"><input type = 'text' name ='txtOfficeTel' tabIndex='12' id = 'txtOfficeTel' size="20"  value = '<%=rs.Fields("ClientOfficeTel")%>'></td>
    <td width="30%" align=right nowrap>Opening Balance&nbsp;&nbsp;&nbsp;</td>
    <td width="36%"><input class=readonly readonly=true name ='txtOpeningBal' tabIndex='20' STYLE="TEXT-ALIGN: RIGHT; WIDTH: 100PX" id = 'txtOpeningBal' size="20"  value = '<%= FormatNum(rs.Fields("ClientOpeningBal"))%>'></td>
  </tr>
  <tr>
    <td width="20%">Email</td>
    <td width="20%"><input type = 'text' name ='txtEmail' STYLE="WIDTH: 220px" tabIndex='13' id = 'txtEmail' size="20"  value = '<%=rs.Fields("ClientEmail")%>'></td>
    <td width="30%" align=right>Notes/Comments&nbsp;&nbsp;&nbsp;</td>
    <td width="36%" rowspan=2><TEXTAREA name ='txtComment' tabIndex='21' id = 'txtComment' rows='2' cols="25"  value = ""><%=rs.Fields("ClientComment")%></TEXTAREA></td>
  </tr>  
  
  <tr>
    <td width="20%">Fax</td>
    <td width="20%"><input type = 'text' name ='txtFax' tabIndex='14' id = 'txtFax' size="20"  value = '<%=rs.Fields("ClientFax")%>'>&nbsp;&nbsp;CDA &nbsp; 
    <%if cbool(Rs.Fields("IsCustodian")) then
		custodian=1
		%>
		<input type=checkbox  checked value='True' name='chkCustodian' id='chkCustodian' onClick = 'UpdateCustodianStatus(this);'> 
	<%else
		custodian=0	
	%>
		<input type=checkbox   value='False' name='chkCustodian' id='chkCustodian' onClick = 'UpdateCustodianStatus(this);'> 
	<%end if%></td>
    <td width="30%" align=right></td>
    <td width="36%">&nbsp;</td>
  </tr>
  <tr>
    <td width="20%" valign="top">Address</td>
    <td width="20%" ><textarea rows=3 name ='txtAddr' tabIndex='15' id = 'txtAddr' size="35" cols="33"><%=rs.Fields("ClientAddr")%></textarea></td>
    <td width="30%" align=right valign="top" nowrap>Physical Address &nbsp;&nbsp;</td>
    <td width="36%"><textarea rows=3 name ='txtPAddr' tabIndex='22' id = 'txtPAddr' size="35" cols="25"><%=rs.Fields("ClientPAddr")%></textarea></td>
  </tr>	
  
  <tr>
	<td width="20%" nowrap>SMS Contract&nbsp;<input type="checkbox" name="SmsContract" value="1" <%if cbool(rs("updateOnContract")) = true then response.write "checked"%>></td>
	<td width="20%" nowrap>SMS Debit&nbsp;<input type="checkbox" name="SmsDebit" value="1" <%if cbool(rs("updateOnDebt")) = true then response.write "checked"%>></td>
	<td width="36%" align=right>Signature&nbsp;</td>
	<td rowspan="5">
		<input type="hidden" name="txtSignature" value="<%= Signature %>">
		<IFRAME FRAMEBORDER=0 SCROLLING=NO SRC="upload.asp?filetext=txtSignature" width="170px" height="170px" tabIndex='-1'></IFRAME>
	</td>
  </tr>	
  <tr>
	<td width="20%"  valign="top">Institution</td>
    <td width="20%" colspan="2"><select name = 'cboGenericSetting1' id = 'cboGenericSetting1' size="1">
						<option selected value = ''></option>
<%

           Dim genRS
		sqlStr = "SELECT * FROM [InstitutionList]"
		Set genRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))


        sqlStr = "SELECT * FROM [GenericSettingList] WHERE EntityType_DPA_ = 1 Order By GenericSettingDescription"
        Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        'rsEdit.Filter = "Generic_DPA_ = 1"

       


        If Not (genRS.EOF Or genRS.BOF) Then
                Do Until genRS.EOF
                		if genRS.Fields("Institution_DPA_") = rs.Fields("Institution_DPA_") Then%>
                			<option selected value = '<%=genRS.Fields("Institution_DPA_")%>'><%=genRS.Fields("InstitutionName")%></option>
                		<%else%>
                        <option value = '<%=genRS.Fields("Institution_DPA_")%>'><%=genRS.Fields("InstitutionName")%></option>
                     <%end if
						genRS.MoveNext
                Loop
        End If
 %>

    </select></td>
    
  </tr>
   
  <tr>
	<td width="20%"  valign="top">Generic2</td>
    <td width="20%" colspan="2"><select name = 'cboGenericSetting2' id = 'cboGenericSetting2' size="1">
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
	<td width="20%"  align=left valign="top">Generic3&nbsp;&nbsp;&nbsp;</td>
    <td width="20%" colspan="2"><select name = 'cboGenericSetting3' id = 'cboGenericSetting3' size="1">
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
	<td width="20%" height="50" nowrap>Use contact name<br>in Portfolio Reports</td>
	<td width="20%" colspan="2" nowrap>
	<%If Rs("UseContactNameInPortfolioReports") = True Then%>
		<input type="checkbox" name="chkcontact" id="chkcontact" value="1" checked></td>
	<%Else%>
		<input type="checkbox" name="chkcontact" id="chkcontact" value="1"></td>
	<%End If%>
</tr>
  
  <%
	sqlStr = "SELECT ISNULL(BankAcc.BnkBranch, '') AS BnkBranch, ISNULL(BankAcc.BankAccName, '') AS BankAccName, ISNULL(BankAcc.BankAccNumber, " & _
		" '') AS BankAccNumber, ISNULL(BankAcc.BankAcc_DPA_, 0) AS BankAcc_DPA_, BankAcc.BankNo, ISNULL(BankAcc.Bank_DPA_, '') AS Bank_DPA_" & _
		" FROM Client INNER JOIN" & _
		" BankAcc ON Client.Client_DPA_ = BankAcc.Client_DPA_" & _
		" WHERE (Client.Client_DPA_ = "& ID &")"
	sqlStr = SQLServerFormat(HandleQuote(sqlStr))
	  'Response.Write sqlstr
	  'Response.End               
	Set rs2 = conn.Execute(sqlStr)

	If Not (rs2.EOF Or rs2.BOF) Then
		Do Until rs2.EOF
			
			If (rs2("BankNo")=1) Then
				Bank1 = rs2("Bank_DPA_")
				BnkBranch1 = rs2("BnkBranch")
				BankAccName1 = rs2("BankAccName")
				BankAccNumber1 = rs2("BankAccNumber")
			End If
			
			If (rs2("BankNo")=2) Then
				Bank2 = rs2("Bank_DPA_")
				BnkBranch2 = rs2("BnkBranch")
				BankAccName2 = rs2("BankAccName")
				BankAccNumber2 = rs2("BankAccNumber")
			End If
			
			If (rs2("BankNo")=3) Then
				Bank3 = rs2("Bank_DPA_")
				BnkBranch3 = rs2("BnkBranch")
				BankAccName3 = rs2("BankAccName")
				BankAccNumber3 = rs2("BankAccNumber")
			End If
			
			rs2.MoveNext
		Loop
	End If
  %>
<tr>
	<td colspan=2>
		<b>Bank Account Details</b>
		<table border=0 style="border: 1px ridge #000080" WIDTH="100%">
			<tr>
				<td>
					<!--BANK 1-->
					<table border=0 style="border: 1px ridge #000080" WIDTH="100%">	
						<tr>
							<td width="30%">Bank</td>
							<td width="70%"><select name = 'cboBank' id = 'cboBank' size="1" tabIndex='23'>
							<%
							bankAccExists = false
					        
							sqlStr = "SELECT * FROM [Bank]"
							Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
							If Not (rsEdit.EOF Or rsEdit.BOF) Then
								rsEdit.MoveFirst
								Do Until rsEdit.EOF
									if rsEdit.Fields("Bank_DPA_") = Bank1 Then
										%>
										<option selected value = '<%=rsEdit.Fields("Bank_DPA_")%>'><%=rsEdit.Fields("BankName")%></option>
										<%	
										bankAccExists = true
									else
										%>
										<option value = '<%=rsEdit.Fields("Bank_DPA_")%>'><%=rsEdit.Fields("BankName")%></option>
										<%
									end if
									rsEdit.MoveNext
								Loop
					                
								if not(bankAccExists) then
									%>
									<option selected value = ''></option>
									<%
								else
									%>
									<option value = ''></option>
									<%
								end if
							End If
							%>
							</select></td>
						</tr>
						
						<tr>
							<td width="30%">Branch</td>
							<td width="70%"><input type = 'text' name ='txtBnkBranch' id = 'txtBnkBranch' size="20" tabIndex='24' value="<%=BnkBranch1%>"></td>
						</tr>
						
						<tr>
							<td width="30%">Account Name</td>
							<td width="70%"><input type = 'text' name ='txtAccName' id = 'txtAccName' size="28" tabIndex='24' value = '<%=BankAccName1%>'></td>
						</tr>
						
						<tr>
							<td width="30%">Account Number </td>
							<td width="70%"><input type = 'text' name ='txtNumber' id = 'txtNumber' size="28" tabIndex='25' value = '<%=BankAccNumber1%>'></td>
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
							<td width="70%"><select name = 'cboBank2' id = 'cboBank2' size="1" tabIndex='23'>
							<%
							bankAccExists2 = false
					        
							sqlStr = "SELECT * FROM [Bank]"
							Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
							If Not (rsEdit.EOF Or rsEdit.BOF) Then
								rsEdit.MoveFirst
								Do Until rsEdit.EOF
									if rsEdit.Fields("Bank_DPA_") = Bank2 Then
										%>
										<option selected value = '<%=rsEdit.Fields("Bank_DPA_")%>'><%=rsEdit.Fields("BankName")%></option>
										<%	
										bankAccExists2 = true
									else
										%>
										<option value = '<%=rsEdit.Fields("Bank_DPA_")%>'><%=rsEdit.Fields("BankName")%></option>
										<%
									end if
									rsEdit.MoveNext
								Loop
					                
								if not(bankAccExists2) then
									%>
									<option selected value = ''></option>
									<%
								else
									%>
									<option value = ''></option>
									<%
								end if
							End If
							%>
							</select></td>
						</tr>
						
						<tr>
							<td width="30%">Branch</td>
							<td width="70%"><input type = 'text' name ='txtBnkBranch2' id = 'txtBnkBranch2' size="20" tabIndex='24' value="<%=BnkBranch2%>"></td>
						</tr>
  
						<tr>
							<td width="30%">Account Name</td>
							<td width="70%"><input type = 'text' name ='txtAccName2' id = 'txtAccName2' size="28" tabIndex='24' value = '<%=BankAccName2%>'></td>
						</tr>
						
						<tr>
							<td width="30%">Account Number </td>
							<td width="70%"><input type = 'text' name ='txtNumber2' id = 'txtNumber2' size="28" tabIndex='25' value = '<%=BankAccNumber2%>'></td>
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
							<td width="70%"><select name = 'cboBank3' id = 'cboBank3' size="1" tabIndex='23'>
							<%
							bankAccExists3 = false
					        
							sqlStr = "SELECT * FROM [Bank]"
							Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
							If Not (rsEdit.EOF Or rsEdit.BOF) Then
								rsEdit.MoveFirst
								Do Until rsEdit.EOF
									if rsEdit.Fields("Bank_DPA_") = Bank3 Then
										%>
										<option selected value = '<%=rsEdit.Fields("Bank_DPA_")%>'><%=rsEdit.Fields("BankName")%></option>
										<%	
										bankAccExists3 = true
									else
										%>
										<option value = '<%=rsEdit.Fields("Bank_DPA_")%>'><%=rsEdit.Fields("BankName")%></option>
										<%
									end if
									rsEdit.MoveNext
								Loop
					                
								if not(bankAccExists3) then
									%>
									<option selected value = ''></option>
									<%
								else
									%>
									<option value = ''></option>
									<%
								end if
							End If
							%>
							</select></td>
						</tr>
						
						<tr>
							<td width="30%">Branch</td>
							<td width="70%"><input type = 'text' name ='txtBnkBranch3' id = 'txtBnkBranch3' size="20" tabIndex='24' value="<%=BnkBranch3%>"></td>
						</tr>
  
						<tr>
							<td width="30%">Account Name</td>
							<td width="70%"><input type = 'text' name ='txtAccName3' id = 'txtAccName3' size="28" tabIndex='24' value = '<%=BankAccName3%>'></td>
						</tr>
						
						<tr>
							<td width="30%">Account Number </td>
							<td width="70%"><input type = 'text' name ='txtNumber3' id = 'txtNumber3' size="28" tabIndex='25' value = '<%=BankAccNumber3%>'></td>
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
		<input type = 'button' Class="Buttons" name ='cmdSave' id = 'cmdSave' value=" Save " tabIndex='25' onclick = "AllowedNavigation();forceSubmit()">
    	<input type = 'button' Class=Buttons name ='cmdAdd' id = "cmdCancel" value=" Cancel " tabIndex='26' onclick = "JavaScript: window.self.close()">
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
    	<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
    	<input type = 'hidden' name ='bankAccExists' id = 'bankAccExists' value="<%=bankAccExists%>">
    	<input type = 'hidden' name ='bankAccExists2' id = 'bankAccExists2' value="<%=bankAccExists2%>">
    	<input type = 'hidden' name ='bankAccExists3' id = 'bankAccExists3' value="<%=bankAccExists3%>">
    	<input type = 'hidden' name ='OnlineRegistration' id = 'OnlineRegistration' value="<%=rs.Fields("OnlineRegistration")%>">
		<input type = 'hidden' name ='handlePhotoRefresh' id = 'handlePhotoRefresh' value="1">
		<input type = 'hidden' name ='custodianStatus' id = 'custodianStatus' size="20" value='<%=custodian%>'>
	</td>
  
	</tr>	
</table>

</form>
</body>

</html>














