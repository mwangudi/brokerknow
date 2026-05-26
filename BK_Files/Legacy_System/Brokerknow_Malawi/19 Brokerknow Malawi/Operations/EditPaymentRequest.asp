<!--#include virtual="libroutines.asp"-->
<%
	const LinkedIndependent = 1
	const LinkedDependent = 2
   
	const UDLName = "KBroker"
	const DataSource = "EditPaymentRequest"
	const DataEntity = "PaymentRequest"
	const DataEntityPlural = "PaymentRequests"
	const ActionFolder = "Operations"
	
	Dim UserId
	Dim action
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim ID
	dim currentEntityType
	dim currentPaymentType
	
	UserId=Session("UserID")
	
	action = ucase(Request.Form("action"))
	ID = Request("ID")
	currentEntityType = 0

'ID = 1

	If Trim(ID) = "" Then
		%>
		<script language = 'vbscript'>
			ShowMessage "No record specified for editing"
		</script>
		<% response.end
	End If
        
	select case  action 
		case"EXECUTE" 
			Dim buttonAction
			Dim reloadRequired
		
			reloadRequired = false
			buttonAction = Trim(Ucase(Request.Form("cmdAdd")))
			if buttonAction = "SAVE" then
				Dim client
				Dim account
				Dim PaymentType
				Dim PDate
				Dim amount
				Dim bank
				Dim reference
				Dim narrative
				Dim status
								
				client = Request.Form("cboClient") 
				account = Request.Form("cboAccountToUse")
				PaymentType = Request.Form("cboPaymentTypes")
				PDate = Trim(Request.Form("txtDate"))
				PDate = PDate & " " & Time
				amount = Request.Form("txtAmount")
				bank = Request.Form("cboBank")
				reference = Request.Form("txtRef")
				narrative = Request.Form("txtNarrative")
			    status = Request.Form("CboStatus")
			    
				If Trim(client) = "" Then%>
				        <script language = 'vbscript'>
				        		ShowMessage "Please specify the client"
				        </script>
				        <% response.end
				End If
				'validate Account
				If Trim(Account) = "" Then%>
				        <script language = 'vbscript'>
				        		ShowMessage "Please specify the Account to use"
				        </script>
				        <% response.end
				End If
				'validate Entity
				If Trim(PaymentType) = "" Then%>
				        <script language = 'vbscript'>
				        		ShowMessage "Please specify the Payment Types."
				        </script>
				        <% response.end
				End If
				'validate Amount
				If Trim(Amount) = "" Then%>
				   <script language = 'vbscript'>
				        	ShowMessage "Please specify the Amount "
				   </script>
				   <% response.end
				End If
				'ensure Amount is numeric
				If (Amount <> "") And (Not IsNumeric(Amount)) Then%>
				   <script language = 'vbscript'>
				   	ShowMessage "Order Detail Estimated Amount must be numeric"
				   </script>
				   <% response.end
				End If
				'validate Bank
				If Trim(Bank) = "" Then%>
				        <script language = 'vbscript'>
				        		ShowMessage "Please specify the Bank"
				        </script>
				        <% response.end
				End If
				'validate size of Reference 
				If Len(Reference) > 20 Then%>
				        <script language = 'vbscript'>
				        ShowMessage "Reference can only be 20 characters in length"
				        </script>
				        <% response.end
				End If
				'validate size of Narrative 
				If Len(narrative) > 200 Then%>
				        <script language = 'vbscript'>
				        ShowMessage "Narrative can only be 200 characters in length"
				        </script>
				        <% response.end
				End If
					
				Set conn = GetActiveConnection("KBroker")
					
				conn.BeginTrans
					sqlStr = "UPDATE PaymentRequest" & _
						" SET BankAccount_DPA_ = " & bank & _
						", PaymentAmount = " & ccur(amount) & _
						", PaymentPDate = #" & FormatDate(pDate) & "#" & _ 
						", PaymentReference = '" & reference & "'" & _
						", PaymentTypes_DPA_ = " & PaymentType & "" & _
						", PaymentNarrative = '" & narrative & "'" & _
						", AccountToUse = '" & account & "'" & _
						", Entity_DPA_ = " & client & " " & _
						", ChangedBy = " & UserId & " " & _
						", TimeChanged = '" & FormatDateandTime(now()) & "' " & _
						", Status = '" & Status & "' " & _
						" WHERE PaymentRequest_DPA_= " & ID
					sqlStr = SQLServerFormat(HandleQuote(sqlStr))
					
					conn.Execute sqlStr
							
					'save data					
					sqlStr = "INSERT INTO [Payment] ( PaymentAmount, PaymentPDate, BankAccount_DPA_,ChangedBy,TimeChanged, " & _
					"PayType_DPA_, EntityType_DPA_, Payment_DPA_, PaymentReference, PaymentTypes_DPA_, PaymentNarrative, Entity_DPA_, Voucher_DPA_,ClientVoucher_DPA_, ChequeCollection) " & _
					"SELECT " & " " & CCur(amount) & " " & " as PaymentAmount" & _
					"," & "#" & FormatDate(PDate) & "#" & " as PaymentPDate" & _
					"," & " " & bank & " " & " as BankAccount_DPA_" & _
					"," & " " & UserId & " " & " as ChangedBy" & _
					"," & "GetDate() as TimeChanged" & _
					"," & "2 as PayType_DPA_" & _
					"," & "1" & " as EntityType_DPA_" & _
					"," & " " & "iif(isnull(max([Payment_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Payment'),max([Payment_DPA_]) + 1)" & " " & " as Payment_DPA_" & _
					"," & "'" & reference & "'" & " as PaymentReference" & _
					"," & " " & PaymentType & " " & " as PaymentTypes_DPA_" & _
					"," & "'" & narrative & "'" & " as PaymentNarrative" & _
					"," & " " & client & " " & " as Entity_DPA_" & _
					"," & "null" & " as Voucher_DPA_" & _
					"," & "null" & " as ClientVoucher_DPA_" & _
					"," & "null" & " as ChequeCollection" & _
					" FROM [Payment]"
					sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))

					conn.Execute sqlStr
					
					conn.execute ("Exec ClientBalanceProcedure " & client)		
									
					conn.CommitTrans
				conn.Close
				
				Set conn = Nothing
				WritefraEnabledDialogCloseScript
				Response.End
			end if
			
			if instr(1,buttonAction,"PRINT") > 0 then
				%>
				<SCRIPT LANGUAGE="JAVASCRIPT">
					//window.parent.parent.frames['maininfo'].location.reload();
				</SCRIPT>
				<%
				WriteDialogRelocateScript "PaymentRequestForm.asp?ID=" & ID
			else
			    WritefraEnabledDialogCloseScript
			End if
				
			Dim clientCode
        
			clientCode = "var validNavigate = true;" & chr(13)
			%>
			<script>
				<%=clientCode%>
			</script>
			<%
			response.End
   	end select
   	
   	Set conn = GetActiveConnection("KBroker")
	sqlStr = "SELECT * FROM PaymentRequest WHERE PaymentRequest_DPA_=" & ID
        
    Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
    If rsEdit.EOF Or rsEdit.BOF Then
		%>
        <script language = 'vbscript'>
        		window.self.ShowMessage "The selected <%=DataEntity%> cannot be retrieved for editing"
        </script>
        <%
        response.end
    End If
	%>

<html>

<head>
<title>Edit <%=DataEntity%></title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<SCRIPT language=Javascript src="../scripts/accountList.js"></SCRIPT>
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<script language="JavaScript" src="CALENDAR/calendar.js"></script>
<!--END CALENDAR -->

<script language = "javascript">
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
</script>
</head>

<body Class="Dialog">
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<SCRIPT language="JavaScript">
	var cal=new ctlSpiffyCalendarBox("cal", "frm<%=DataSource%>", "txtDate","cmdDate","<%=  FormatDate(rsEdit.Fields("PaymentPDate")) %>",1);
</SCRIPT>

<form name = 'frm<%=DataSource%>' method = 'post' action = '<%=DataSource%>.asp' id = "frmMain">
<table border="0" width="100%" cellspacing="1" cellpadding="1">
	<tr>
		<td width="10%">Client</td>
		<td width="20%"><input type = 'text' name ='txtClientCode' id = 'txtClientCode' size="10" onBlur="txtval = this.value; selectItem(cboClient);"> 
		&nbsp;&nbsp;&nbsp;
		<select name = "cboClient" id = "cboClient" size="1" 
    		onchange = "UpdateCode(true,cboClient,txtClientCode); frmMain.txtFavour.value=this.options[this.selectedIndex].id"
			onKeypress = "return (dodefaultaction()==''); "  
			onKeydown = "return (dodefaultaction()==''); " 
			onKeyup = "change(cboClient,0);"
			onfocus = "txtval = '';inputIsItemCode = 0;" 
			onblur = "txtval = '';inputIsItemCode = 0;">
			<option selected SearchCode = "" SearchText = ""  value = ""></option>
			<%
				Set conn = GetActiveConnection("KBroker")
						        
				sqlStr = "SELECT * FROM Client ORDER BY ClientName"
				Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
				If Not (rs.EOF Or rs.BOF) Then
				        rs.MoveFirst
				        Do Until rs.EOF
							If rsEdit("Entity_DPA_") = rs.Fields("Client_DPA_") Then
								ClientName = rs.Fields("ClientName")
								%>
								<option selected SearchCode = "<%=rs.Fields("Client_DPA_")%>" SearchText = "<%=rs.Fields("ClientName")%>" id = '<%=rs.Fields("ClientName")%>' value = '<%=rs.Fields("Client_DPA_")%>'><%=mid(rs.Fields("ClientName"),1,30)%></option>
								<%
							else
								%>
								<option SearchCode = "<%=rs.Fields("Client_DPA_")%>" SearchText = "<%=rs.Fields("ClientName")%>" id = '<%=rs.Fields("ClientName")%>' value = '<%=rs.Fields("Client_DPA_")%>'><%=mid(rs.Fields("ClientName"),1,30)%></option>
								<%
				            end if
				            rs.MoveNext
				        Loop
				End If
			%>
		</select>
		</td>
		<td width="10%"> </td>
	</tr>
	
	<tr>
		<td width="10%">Account</td>    
		<td nowrap width="20%">
			<select name='cboAccountToUse' id="cboAccountToUse" size="1">
				<option value="ClientName">Use Client Name</option>
				<option value="Account1">Use Account #1</option>
				<option value="Account2">Use Account #2</option>
				<option value="Account3">Use Account #3</option>
			</select>
		</td>
		<td width="10%"> </td>
		<input type="hidden" name="hidAccountToUse" id="hidAccountToUse" value="<%=trim(ucase(rsEdit("AccountToUse")))%>">
		<script language="vbscript">
			dim i
			
			for i = 0 to document.all.item("cboAccountToUse").length-1
				if trim(ucase(document.all.item("cboAccountToUse").options(i).value)) = document.all.item("hidAccountToUse").value then
					document.all.item("cboAccountToUse").options(i).selected = true
				end if
			next
		</script>
	</tr>  
	
	<tr>
		<td width="10%">In Favour Of</td>
		<td width="20%"><input type = 'text' name ='txtFavour' id = "txtFavour" value = '<%=ClientName%>' size="20" readonly></td>
		<td width="10%"> </td>
	</tr>
			
	<tr>
	<td width="10%">Payment Type</td>
	<td width="20%">
		<select name = 'cboPaymentTypes' id = 'cboPaymentTypes' size="1">
			<%
			Set conn = GetActiveConnection("KBroker")
			sqlStr = "SELECT * FROM [PaymentTypes]  Order By Description ASC"
			Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
			
			If Not (rs.EOF Or rs.BOF) Then
				rs.MoveFirst
				Do Until rs.EOF
					If rsEdit("PaymentTypes_DPA_") = rs.Fields("PaymentTypes_DPA_") Then
						%>
						<option Selected value = '<%=rs.Fields("PaymentTypes_DPA_")%>' SearchText = "<%=rs.Fields("Description")%>"><%=rs.Fields("Description")%></option>
						<%
					else
						%>
						<option value = '<%=rs.Fields("PaymentTypes_DPA_")%>' SearchText = "<%=rs.Fields("Description")%>"><%=rs.Fields("Description")%></option>
						<%      
					end if 
					rs.MoveNext
				Loop
			End If
			%>
		</select>
	</td>
	<td width="10%"> </td>
	</tr>
		
	<tr>
		<td width="15%">Date</td>
		<td width="54%"><SCRIPT language="JavaScript">cal.writeControl();</SCRIPT></td>
	</tr>

	<tr>
		<td width="10%">Amount</td>
		<td width="20%"><input  STYLE="text-align:right" type = 'text' name ='txtAmount' id = "txtAmount" value = '<%=rsEdit.Fields("PaymentAmount")%>' size="20"></td>
		<td width="10%"> </td>
	</tr>
	
	<tr>
		<td width="10%">Bank</td>
		<td width="20%"><select name = 'cboBank' id = 'cboBank' size="1"  
		onKeypress="return (dodefaultaction()==''); " 
		onKeydown="return (dodefaultaction()==''); " 
		onKeyup="return (change(cboBank));" 
		onfocus="txtval = '';inputIsItemCode = 1;" 
		onblur="txtval = '';inputIsItemCode = 1;">
		<option selected SearchCode = "0" SearchText = ""  value = ''></option>
		<%
		sqlStr = "SELECT * FROM [BankAccountList] Order By AccountName"
		Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		
		If Not (rs.EOF Or rs.BOF) Then
			rs.MoveFirst
			Do Until rs.EOF
				If rs("Account_DPA_") = rsEdit("BankAccount_DPA_") Then
				%>
					<option selected SearchCode = "<%=rs.Fields("AccountCode")%>" SearchText = "<%=rs.Fields("AccountName")%>" value = '<%=rs.Fields("Account_DPA_")%>'><%=rs.Fields("AccountNameEx")%></option>
				<%Else%>
					<option SearchCode = "<%=rs.Fields("AccountCode")%>" SearchText = "<%=rs.Fields("AccountName")%>" value = '<%=rs.Fields("Account_DPA_")%>'><%=rs.Fields("AccountNameEx")%></option>
				<%End If%>
				<%
				rs.MoveNext
			Loop
		End If
		%>
		</select></td>
		<td width="10%"> </td>
	</tr>
	
	<tr>
		<td width="10%">Reference</td>
		<td width="20%"><input type = 'text' name ='txtRef' id = 'txtRef' size="20" value="<%=rsEdit.Fields("PaymentReference")%>"></td>
		<td width="10%"> </td>
	</tr>
	
	<tr>
		<td width="10%">Narrative</td>
		<td width="20%"><textarea name ='txtNarrative' id = 'txtNarrative' rows="3" cols="20" ><%=rsEdit.Fields("PaymentNarrative")%></textarea></td>
		<td width="10%"> </td>
	</tr>

	<tr>
		<td width="10%">Status</td>
		<td width="20%">
			<select name='cboStatus' id="cboStatus" size="1">
				<option value="Approved">Approved</option>
				<option value="Cancelled">Cancelled</option>
				<option value="Pending">Pending</option>
			</select>
		</td>
		<td width="10%"> </td>
		<input type="hidden" name="hidStatus" id="hidStatus" value="<%=trim(ucase(rsEdit("Status")))%>">
		<script language="vbscript">
			dim i
			
			for i = 0 to document.all.item("cboStatus").length-1
				if trim(ucase(document.all.item("cboStatus").options(i).value)) = document.all.item("hidStatus").value then
					document.all.item("cboStatus").options(i).selected = true
				end if
			next
		</script>
	</tr>
	
	<tr>
		<td width="10%"> </td>
		<td width="100%" colspan=2 valign=absBottom>
		<BR><BR>
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save " onclick = "AllowedNavigation()">
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdPrint' value=" Print " onclick = "AllowedNavigation()">
		<input type = 'button' Class=Buttons name ='cmdCancel' id = "cmdCancel" value=" Cancel " onclick = "JavaScript: window.self.close()">
		
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		<input type = 'hidden' name ='ID' id = 'ID' value="<%= rsEdit.Fields("PaymentRequest_DPA_").Value %>">
		<input type = 'hidden' name ='CurrentEntityTypeID' id = 'CurrentEntityTypeID'>
		</td>
	</tr>
</table>

</form>
</body>

</html>