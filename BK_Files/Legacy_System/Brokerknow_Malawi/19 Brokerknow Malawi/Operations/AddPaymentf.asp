<!--#include virtual="libroutines.asp"-->
<%
	const UDLName = "KBroker"
	const DataSource = "AddPaymentf"
	const DataEntity = "Payment"
	const DataEntityPlural = "Payments"
	const ActionFolder = "Operations"
	
	Dim UserId
	Dim action
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim ID
	dim currentEntityType
	
	UserId=Session("UserID")
	
	action = ucase(Request.Form("action"))
	ID = Request.Form("ID")
	currentEntityType = 5 'Default is client
	
	select case action
		case "EXECUTE" 
			Dim buttonAction
			Dim reloadRequired
		
			reloadRequired = false
			buttonAction = Trim(Ucase(Request.Form("cmdAdd")))
			if buttonAction = "SAVE" then
					Dim bank
					Dim amount
					Dim account
					Dim reference
					Dim PDate
					Dim narrative
					Dim entity
					Dim voucherType
					Dim ContractsSel
					dim ChequeStatus
					Dim PaymentType
					Dim Ref 'PaymentTypes
					Dim BatchNo
			      
			        voucherType = Request.Form("txtVoucherType")
			        ContractsSel = Request.Form ("ContractsSel")
					entity = Request.Form("cboEntity") 
					bank = Request.Form("cboBank")
					account = Request.Form("cboAccount")
					amount = Request.Form("txtTotal")
					reference = Request.Form("txtRef")
					PDate = Trim(Request.Form("txtDate"))
					PDate = PDate & " " & Time
					narrative = Request.Form("txtNarrative")
					ChequeStatus = Request.Form("ChequeCollection")
			        PaymentType = Request.Form("cboPaymentTypes")
			        BatchNo = Request.Form("cboBatch")
			        
					'validate Entity
					If Trim(PaymentType) = "" Then%>
					         <script language = 'vbscript'>
					         		ShowMessage "Please specify the Payment Types."
					         		
					         </script>
					         <% response.end
					 End If
					 
					 If Trim(entity) = "" Then%>
					         <script language = 'vbscript'>
					         		ShowMessage "Please specify the Entity"
					         		
					         </script>
					         <% response.end
					 End If
					 
					 'validate Account
					 If Trim(Account) = "" Then%>
					         <script language = 'vbscript'>
					         		ShowMessage "Please specify the Account"
					         		
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
					
					'Validate: If Payment Type requires a reference, then The reference No field is mandatory
							Set conn = GetActiveConnection("KBroker")
							
							sqlStr = "SELECT * FROM [PaymentTypes] WHERE [PaymentTypes_DPA_] = " & PaymentType 
							Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
							
							If (rs.EOF Or rs.BOF) Then
							%>
							 <script language = 'vbscript'>
									ShowMessage "Problems encountered while processing Receipt Payment Type."
											         			
								</script>
							<% response.end

							else
							    Ref = rs("Reference")
								if (Cint(Ref) = 1 And reference = "") then%>
								
							       <script language = 'vbscript'>
									   ShowMessage "The Receipt Payment Type selected requires a reference."
											         			
								    </script>
							      <% response.end
							   End if 
							End If
							
					if trim(UCase(ChequeStatus)) = "ON" then
						ChequeStatus = "AWAITING COLLECTION"
					else
						ChequeStatus = "NOT COLLECTED"
					end if
					
					Dim clientVoucher
					Dim brokerVoucher
					
					Set conn = GetActiveConnection("KBroker")
					
					conn.BeginTrans
							'create voucher
							
							If (Trim(ContractsSel) <> "") Then
									set guid = server.createobject("NDUtils.CGUID")
									guidStr = guid.GenerateGUID
									if (voucherType = 1) then
											sqlStr = "INSERT INTO [ClientVoucher] (ClientVoucher_DPA_, VoucherDate, Voucher_EIT_) SELECT " & " " & "iif(isnull(max([ClientVoucher_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'ClientVoucher'),max([ClientVoucher_DPA_]) + 1)" & " " & " as ClientVoucher_DPA_" & _
													"       ," & "#" & FormatDate(PDate) & "#" & " as VoucherDate" & _
													"       ," & "'" & guidStr & "'" & " as Voucher_EIT_" & _
													"        FROM [ClientVoucher]"
											conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
											
											'obtain voucher number
											sqlStr = "SELECT [ClientVoucher_DPA_] FROM [ClientVoucher] WHERE [Voucher_EIT_] = " & "'" & guidStr & "'"
							
											Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
											If (rs.EOF Or rs.BOF) Then%>
											         	<script language = 'vbscript'>
											         			ShowMessage "A serious error has been encountered while saving the data. Try saving again"
											         			
											         	</script>
											         	<% response.end
											End If
											
											clientVoucher = rs.fields("ClientVoucher_DPA_").value
											brokerVoucher = "Null"
											
											'add contracts to voucher
											sqlStr = "UPDATE Contract SET ClientVoucher_DPA_ = " & clientVoucher & _
													", ContractClientVouchered = 1 WHERE Contract_DPA_ IN (" & ContractsSel &  ")"
							
											conn.Execute SQLServerFormat(HandleQuote(sqlStr))
									elseif (voucherType = 3) then
											sqlStr = "INSERT INTO [Voucher] (Voucher_DPA_, VoucherDate, Voucher_EIT_) SELECT " & " " & "iif(isnull(max([Voucher_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Voucher'),max([Voucher_DPA_]) + 1)" & " " & " as Voucher_DPA_" & _
													"       ," & "#" & FormatDate(PDate) & "#" & " as VoucherDate" & _
													"       ," & "'" & guidStr & "'" & " as Voucher_EIT_" & _
													"        FROM [Voucher]"
											conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
											
											'obtain voucher number
											sqlStr = "SELECT [Voucher_DPA_] FROM [Voucher] WHERE [Voucher_EIT_] = " & "'" & guidStr & "'"
				
											Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
											If (rs.EOF Or rs.BOF) Then%>
											         	<script language = 'vbscript'>
											         			ShowMessage "A serious error has been encountered while saving the data. Try saving again"
											         			
											         	</script>
											         	<% response.end
											End If
											clientVoucher = "Null"
											brokerVoucher = rs.fields("Voucher_DPA_").value
											
											'add contracts to voucher
											sqlStr = "UPDATE Contract SET Voucher_DPA_ = " & brokerVoucher & _
													", ContractVouchered = 1 WHERE Contract_DPA_ IN (" & ContractsSel &  ")"
				
											conn.Execute SQLServerFormat(HandleQuote(sqlStr))
									end if
							else
									clientVoucher = "Null"
									brokerVoucher = "Null"
							End If
					
	
							'save data					
							sqlStr = "INSERT INTO [Payment] ( PaymentAmount, PaymentPDate, BankAccount_DPA_,ChangedBy,TimeChanged, " & _
									"PayType_DPA_, EntityType_DPA_, Payment_DPA_, PaymentReference, PaymentTypes_DPA_, PaymentNarrative, Entity_DPA_, Voucher_DPA_,ClientVoucher_DPA_, ChequeCollection,Batch_No) " & _
									"SELECT " & " " & CCur(amount) & " " & " as PaymentAmount," & "#" & FormatDate(PDate) & "#" & " as PaymentPDate" & _
									"," & " " & bank & " " & " as BankAccount_DPA_" & _
									"," & " " & UserId & " " & " as ChangedBy" & _
									"," & "GetDate() as TimeChanged" & _
									"," & "2 as PayType_DPA_" & _
									"," & " " & entity & " " & " as EntityType_DPA_" & _
									"," & " " & "iif(isnull(max([Payment_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Payment'),max([Payment_DPA_]) + 1)" & " " & " as Payment_DPA_" & _
									"," & "'" & reference & "'" & " as PaymentReference" & _
									"," & " " & PaymentType & " " & " as PaymentTypes_DPA_" & _
									"," & "'" & narrative & "'" & " as PaymentNarrative" & _
									"," & " " & account & " " & " as Entity_DPA_" & _
									"," & " " & brokerVoucher & " " & " as Voucher_DPA_" & _
									"," & " " & clientVoucher & " " & " as ClientVoucher_DPA_" & _
									"," & " '" & ChequeStatus & "' " & " as ChequeCollection" & _
									"," & " " & BatchNo & " " & " as Batch_No " & _
									" FROM [Payment]"
							
							sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))

							conn.Execute sqlStr
							
							if (voucherType = 1) then
									sqlStr = "UPDATE [ClientVoucher] SET VoucherPaid = 1 WHERE ClientVoucher_DPA_ = "  & clientVoucher
							elseif (voucherType = 3) then
									sqlStr = "UPDATE [Voucher] SET VoucherPaid = 1 WHERE Voucher_DPA_ = "  & brokerVoucher
							else
									sqlStr = ""
							end if
							
							if sqlStr <> "" then
									conn.Execute SQLServerFormat(HandleQuote(sqlStr))
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
		case "FETCH_ACCOUNTS"
				currentEntityType = cint(ID)
   	end select
%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add <%=DataEntity%></title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/accountList.js"></SCRIPT>
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<script language="JavaScript" src="CALENDAR/calendar.js"></script>
<!--END CALENDAR -->

<script language='vbscript'>

					function EntitySelected(itemID)
 							frm<%=DataSource%>.elements("ID").value = itemID
 							frm<%=DataSource%>.elements("action").value = "Fetch_Accounts"
 							frm<%=DataSource%>.submit
 							
					end function
</script>
<script language='javascript'>
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
		
		var currentEntityType = <%=currentEntityType%>
		function setFilter(obj,entityType,retVal)
		{
			var IDName = "";
			var frameName = "";
			var framePageName = "";
			var voucherParamName = "";
			var saleType = "";
			var selValue = "";
			
			switch (entityType)
			{
			case 1:
				IDName = "Client_DPA_";
				frameName = "fraInnerSelects";
				framePageName = "inner_select_vouchers";
				voucherParamName = "ContractClientVouchered";
				saleType = "1";
				selValue = obj.value;
				
				break;
			case 3:
				IDName = "BrokerCode";
				frameName = "fraInnerBrokerSelects";
				framePageName = "inner_select_voucherBroker";
				voucherParamName = "ContractVouchered";
				saleType = "0";
				selValue = obj.options[obj.selectedIndex].SearchCode;
				break;
			default:
				return retVal;
			}
			
			var fra = document.getElementById(frameName);
			
			var pageTo = framePageName + '.asp?' + voucherParamName + '=0&OrderTypeSale=' + saleType + '&' + IDName + '=' + selValue; 
			
			if(selValue.length == 0)
			{
				pageTo = framePageName + '.asp';
			}
			
			fra.src = pageTo;
			return retVal;
		}
		
		function FetchAccounts(theList)
		{
			var i = 0;
			var entity = theList.value;
			var toList = document.frmMain.cboAccount;
			
			currentEntityType = entity;
			frm = document.frmMain;				
			xmlhttp = createXMLHTTPObj();
			
			url="GetList.asp?ID="+entity+"&action=GetAccountList";
			xmlhttp.open("GET",url,true);
			xmlhttp.onreadystatechange=function() {
				if (xmlhttp.readyState==4) {
				returnStr = xmlhttp.responseText;
				returnStr = getBodyHTML(returnStr);
				
				var secList = "<select name = '" + toList.name + "' id = '" + toList.id + "' size='1' ";
				secList += "onChange='setFilter(this," + currentEntityType + ");' " ;
				secList += "onKeypress='return (dodefaultaction()==\"\"); ' "  ;
				secList += "onKeydown='return (dodefaultaction()==\"\");' " ; 
				secList += "onKeyup='return (change(" + toList.name + "));' " ; 
				secList += "onfocus='txtval = \"\";inputIsItemCode = 1;' "  ;
				secList += "onblur='txtval = \"\";inputIsItemCode = 1;'>" ;
				secList += returnStr ;
				secList += "</select>";
				
				toList.outerHTML = secList;															
				}
				}
			xmlhttp.setRequestHeader('Accept','message/x-jl-formresult');
			xmlhttp.send(); 
			
			var fra = document.getElementById('fraInnerSelects');		
			var pageTo = 'inner_select_vouchers.asp';  
			fra.src = pageTo;
			fra = document.getElementById('fraInnerBrokerSelects');		
			pageTo = 'inner_select_voucherBroker.asp';
			fra.src = pageTo;
			
			totalContractAmt = 0;
			document.all.item("txtTotal").value = totalContractAmt;
			document.all.item("ContractsSel").value = "";	
				
			
			if (entity==3){
				document.getElementById('brokerVoucherRow').style.display = '';
				document.getElementById('clientVoucherRow').style.display = 'none';
				document.getElementById('txtVoucherType').value = currentEntityType;
			} 
			
			else{
				 if (entity==1){
					document.getElementById('brokerVoucherRow').style.display = 'none';
					document.getElementById('clientVoucherRow').style.display = '';
					document.getElementById('txtVoucherType').value = currentEntityType;
				}			
				else {
					document.getElementById('brokerVoucherRow').style.display = 'none';
					document.getElementById('clientVoucherRow').style.display = 'none';
					document.getElementById('txtVoucherType').value = '0';
				}
			}	
		
			
		}
		
		var totalContractAmt = 0;
	
		function UpdateVoucherAmount(inAmount, inAction){	
				
				 totalContractAmt = parseFloat(document.all.item("txtTotal").value)
				
			switch(inAction){
				case 0:
					totalContractAmt = totalContractAmt - parseFloat(inAmount)			
					break;
				default:
					totalContractAmt = totalContractAmt + parseFloat(inAmount)				
					break;
						
			}
			
			document.all.item("txtTotal").value = FormatNum(totalContractAmt); 
			
		}

function EnableCheqCollection (obj) 
	{
	  var cheque = obj[obj.selectedIndex].SearchText;
	   
	 if (cheque.toLowerCase()=='cheque')
	  {
	   document.all.item("ChequeCollection").disabled = false;
	  }
	  else
	   {
	  document.all.item("ChequeCollection").disabled = true;
	   }
	
	}
	
function FormatPrice()
	{
	 document.all.item("txtTotal").value= formatNum(document.all.item("txtTotal").value);
	}
	
</script>
</head>

<body Class="Dialog">
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<SCRIPT language="JavaScript">
	var cal=new ctlSpiffyCalendarBox("cal", "frm<%=DataSource%>", "txtDate","cmdDate","<%= FormatDate(Date) %>",1);
</SCRIPT>

<form name = 'frm<%=DataSource%>' method = 'post' action = '<%=DataSource%>.asp' id = "frmMain">
<table border="0" width="100%" cellspacing="1" cellpadding="1">
 <tr>
    <td width="15%">Payment Type</td>
    <td width="54%">
    <select name = 'cboPaymentTypes' id = 'cboPaymentTypes' size="1" onchange='EnableCheqCollection(this)'>
    	
<%      ChkDisable = "disabled" 'Default: Cheque collection diabled a Payment Type Cheque is found

		Set conn = GetActiveConnection("KBroker")
        sqlStr = "SELECT * FROM [PaymentTypes]  Order By Description ASC"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF
			if Ucase(trim(rs.Fields("Description"))) = "CHEQUE" then
			 ChkDisable = ""
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

    </select></td>
    <td width="31%">

	</td>
  </tr>
  <tr>
    <td width="15%">Date</td>
    <td width="54%"><SCRIPT language="JavaScript">cal.writeControl();</SCRIPT></td>
    
  </tr>
  <tr>
    <td width="15%">Entity</td>
    <td width="54%"><select name = 'cboEntity' id = 'cboEntity' size="1" onchange='FetchAccounts(this)'>
    	
<%
		
		'Response.End
		
		Set conn = GetActiveConnection("KBroker")
        sqlStr = "SELECT * FROM [FullEntityTypeList] Order By EntityTypeName"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF
						if Cint(rs.Fields("EntityType_DPA_").value) = Cint(currentEntityType) then%>
								<option selected value = '<%=rs.Fields("EntityType_DPA_")%>'><%=rs.Fields("EntityTypeName")%></option>
                        <%else%>
								<option value = '<%=rs.Fields("EntityType_DPA_")%>'><%=rs.Fields("EntityTypeName")%></option>
                        <%end if
                        rs.MoveNext
                Loop
        End If
%>

    </select></td>
    <td width="31%">

	</td>
  </tr>
 
   <tr>
    <td width="15%">Account</td>    
    <td nowrap>
    <input type = 'text' name ='txtClientCode' id = 'txtClientCode' size="10" onBlur="txtval = this.value; selectItem(cboAccount); setFilter(this,<%=currentEntityType%>,UpdateCode(true,cboAccount,txtClientCode));">
    <select name = 'cboAccount' id = "cboAccount" size="1" 
	onchange='setFilter(this,<%=currentEntityType%>,UpdateCode(true,cboAccount,txtClientCode))'	
	onKeypress="return (dodefaultaction()==''); " 
	onKeydown="return (dodefaultaction()==''); " 
	onKeyup="return (change(cboAccount));" 
	onfocus="txtval = '';inputIsItemCode = 1;" 
	onLoad="setFilter(this,<%=currentEntityType%>,UpdateCode(true,cboAccount,txtClientCode))"
	onblur="txtval = '';inputIsItemCode = 1;">
	<option selected SearchCode = "" SearchText = ""  value = ''></option>
 
<%
		Dim displayField
		
        'if currentEntityType = 1 then 'client data
				displayField = "EntityName"
		'else
		'		displayField = "EntityNameEx"
		'end if
		
        sqlStr = "SELECT * FROM [CompleteEntityList] WHERE EntityType_DPA_ =" & currentEntityType & " Order By EntityName"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF
                AccountName=rs.Fields(displayField)
                %>
                        <option SearchCode = "<%=rs.Fields("EntityCode")%>" SearchText = "<%=rs.Fields("EntityName")%>" value = '<%=rs.Fields("Entity_DPA_")%>'><%=AccountName%></option>
                        <%rs.MoveNext
                Loop
        End If
%>

    </select></td>
    <td width="31%">
	</td>
  </tr>  
		  
		<tr> 
			<td width="15%" align="right">&nbsp;</td>
			<td width="85%" colspan="2">
            <input type="checkbox" name="ChequeCollection" value="ON" <%=ChkDisable%>>&nbsp;&nbsp;&nbsp;Cheque Collection</td>
		</tr>
		
		<tr>
    <td width="15%">Batch No</td>
    <td width="54%"><select name = 'cboBatch' id = 'cboBatch' size="1">
    	
<%
		
		'Response.End
		
		Set conn = GetActiveConnection("KBroker")
        sqlStr = "SELECT DISTINCT Batch_No FROM Offerings WHERE (NOT (Batch_No IS NULL)) AND (Batch_No NOT IN " & _
                 " (SELECT DISTINCT isnull(Batch_no, 0) FROM Payment)) and deleted=0 "
                 
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF
						%>
							<option value = '<%=rs.Fields("Batch_No")%>'><%=rs.Fields("Batch_no")%></option>
                        <%
                        rs.MoveNext
                Loop
        End If
%>

    </select></td>
    <td width="31%">

	</td>
  </tr>
  
		<tr>
		  <td width="15%">Amount</td>
		  <td width="54%"><input  STYLE="text-align:right" type = 'text' name ='txtTotal' id = "txtTotal" value = '0' size="20" onchange="JavaScript: FormatPrice()"></td>
		  <td width="31%">
				<input type="hidden" name="ContractsSel" id="ContractsSel" value="">
			</td>
		</tr>
		<tr>
		  <td width="15%">Bank</td>
		  <td width="54%"><select name = 'cboBank' id = 'cboBank' size="1"  
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
                Do Until rs.EOF%>
                        <option SearchCode = "<%=rs.Fields("AccountCode")%>" SearchText = "<%=rs.Fields("AccountName")%>" value = '<%=rs.Fields("Account_DPA_")%>'><%=rs.Fields("AccountNameEx")%></option>
                        <%rs.MoveNext
                Loop
        End If
%>

    </select></td>
    <td width="31%">

	</td>
  </tr>
  <tr>
    <td width="15%">Reference</td>
    <td width="54%"><input type = 'text' name ='txtRef' id = 'txtRef' size="20"></td>
    <td width="31%">

	</td>
  </tr>
  <tr>
    <td width="15%">Narrative</td>
    <td width="54%">
    <textarea name ='txtNarrative' id = 'txtNarrative' rows="3" cols="20" ></textarea></td>
    <td width="31%">

	</td>
  </tr>
  <tr>
	  <td width="100%" colspan=3 align="right" valign=absBottom>
		<BR><BR>
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value="Save" onclick = "AllowedNavigation()">
    	<input type = 'button' Class=Buttons name ='cmdCancel' id = "cmdCancel" value="Cancel" onclick = "JavaScript: window.self.close()">
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		<input type = 'hidden' name ='ID' id = 'ID'>
	</td>
  </tr>
</table>

</form>
</body>

</html>