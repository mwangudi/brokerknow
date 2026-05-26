<!--#include virtual="libroutines.asp"-->
<%
	const LinkedIndependent = 1
	const LinkedDependent = 2
	
	const UDLName = "KBroker"
	const DataSource = "EditReceipt"
	const DataEntity = "Receipt"
	const DataEntityPlural = "Receipts"
	const ActionFolder = "Operations"
	
	Dim UserId
	Dim action
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim rsEdit
	Dim guid
	Dim guidStr
	Dim ID
	dim currentEntityType
	dim currentPaymentType
	
	action = ucase(Request.Form("action"))
	ID = Request("ID")
	currentEntityType = 0
	currentPaymentType = 0
	UserId=Session("UserID")
	
		If Trim(ID) = "" Then%>
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
			
			if instr(1,buttonAction,"SAVE") > 0 then
					Dim amount
					Dim reference
					Dim narrative
					Dim custOrder
					Dim bank
					Dim account
					Dim PDate
					Dim entity
					Dim voucherType
					Dim ContractsSel
					Dim PaymentType
					Dim Ref 'PaymentTypes
			        
					voucherType = Request.Form("txtVoucherType")
			        ContractsSel = Request.Form ("ContractsSel")
					amount = Request.Form("txtTotal")
					reference = Request.Form("txtRef")
					narrative = Request.Form("txtNarrative")
					custOrder = Request.Form("cboOrder")
					entity = Request.Form("cboEntity") 
					bank = Request.Form("cboBank")
					account = Request.Form("cboAccount")
					PDate = Trim(Request.Form("txtDate"))
					PaymentType = Request.Form("cboPaymentTypes")
					
					'validate Date

					Dim num
					num=DateDiff("d",Date,formatdate(Pdate))
					If ( num>0) Then%>
					         <script language = 'vbscript'>
					         		ShowMessage "Please specify a valid Date "					         		
					         </script>
					         <% response.end
					End If

					PDate = PDate & " " & Time
			       
					 
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
							
					custOrder = iif(custOrder = 0,"Null",custOrder)
					
					'save data
					Dim receiptVoucher
					Set conn = GetActiveConnection("KBroker")
					
					conn.BeginTrans
							'deal with vouchers
							sqlStr = "SELECT BrokerReceiptVoucher_DPA_ FROM [Payment] WHERE [Payment_DPA_] = " & ID
							Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
							If (rs.EOF Or rs.BOF) Then%>
							         	<script language = 'vbscript'>
							         			ShowMessage "A serious error has been encountered while saving the data. Try saving again"
											         			
							         	</script>
							         	<% response.end
							End If
							
							receiptVoucher = rs.fields("BrokerReceiptVoucher_DPA_").value
							
							if not(isnull(receiptVoucher)) then
									'handle broker voucher
									'find out whether any child records exist
									sqlStr = "SELECT Child,DeletionMessage,ParentKey FROM [_Parent_Child_Links_] WHERE (Parent = 'BrokerReceiptVoucher') AND (ChildType = " & LinkedIndependent & ")"
									Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
									If Not (rs.BOF Or rs.EOF) Then
									        rs.MoveFirst
									        Do Until rs.EOF
									        		tableName = rs.Fields("Child")
									                sqlStr = "SELECT TOP 1 * FROM [" & tableName & "] WHERE BrokerReceiptVoucher_DPA_ = " & receiptVoucher
									                Set depRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
									                If Not (depRS.BOF Or depRS.EOF) Then%>
									        				<script language = 'vbscript'>
									        					ShowMessage <%=rs.Fields("DeletionMessage")%>
									        					
									        				</script>
									        				<%response.end
									                End If
									                rs.MoveNext
									        Loop
									End If
									If (Trim(ContractsSel) = "") Then
											'voucher removed
											'remove contracts 
											sqlStr = "UPDATE Contract SET BrokerReceiptVoucher_DPA_ = NULL" & _
													", BrokerReceiptVouchered = 0 WHERE BrokerReceiptVoucher_DPA_ =" & receiptVoucher
											conn.Execute SQLServerFormat(HandleQuote(sqlStr))
											'delete from database
											sqlStr = "DELETE FROM [Voucher] WHERE Voucher_DPA_ = " & receiptVoucher
											conn.Execute SQLServerFormat(HandleQuote(sqlStr))
											
											receiptVoucher = "Null"
											voucherType = 0
									else
											'voucher modified
											sqlStr = "UPDATE [BrokerReceiptVoucher] SET VoucherDate = #" & FormatDate(PDate) & "#" & _
													"        WHERE BrokerReceiptVoucher_DPA_ = " & receiptVoucher
											conn.Execute SQLServerFormat(HandleQuote(sqlStr))
														
											'remove previous contracts 
											sqlStr = "UPDATE Contract SET BrokerReceiptVoucher_DPA_ = NULL" & _
													", BrokerReceiptVouchered = 0 WHERE BrokerReceiptVoucher_DPA_ =" & receiptVoucher
														
											conn.Execute SQLServerFormat(HandleQuote(sqlStr))
														
											'add contracts to voucher
											sqlStr = "UPDATE Contract SET BrokerReceiptVoucher_DPA_ = " & receiptVoucher & _
													", BrokerReceiptVouchered = 1 WHERE Contract_DPA_ IN (" & ContractsSel &  ")"
														
											conn.Execute SQLServerFormat(HandleQuote(sqlStr))
									end if
							else
									'check for new vouchers
									If (Trim(ContractsSel) <> "") Then
											set guid = server.createobject("NDUtils.CGUID")
											guidStr = guid.GenerateGUID
											if (voucherType = 3) then
													sqlStr = "INSERT INTO [BrokerReceiptVoucher] (BrokerReceiptVoucher_DPA_, VoucherDate, BrokerReceiptVoucher_EIT_) SELECT " & " " & "iif(isnull(max([BrokerReceiptVoucher_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'BrokerReceiptVoucher'),max([BrokerReceiptVoucher_DPA_]) + 1)" & " " & " as BrokerReceiptVoucher_DPA_" & _
															"       ," & "#" & FormatDate(PDate) & "#" & " as VoucherDate" & _
															"       ," & "'" & guidStr & "'" & " as BrokerReceiptVoucher_EIT_" & _
															"        FROM [BrokerReceiptVoucher]"
													conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
													
													'obtain voucher number
													sqlStr = "SELECT [BrokerReceiptVoucher_DPA_] FROM [BrokerReceiptVoucher] WHERE [BrokerReceiptVoucher_EIT_] = " & "'" & guidStr & "'"
				
													Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
													If (rs.EOF Or rs.BOF) Then%>
													         	<script language = 'vbscript'>
													         			ShowMessage "A serious error has been encountered while saving the data. Try saving again"
													         			
													         	</script>
													         	<% response.end
													End If
													receiptVoucher = rs.fields("BrokerReceiptVoucher_DPA_").value
													
													'add contracts to voucher
													sqlStr = "UPDATE Contract SET BrokerReceiptVoucher_DPA_ = " & receiptVoucher & _
															", BrokerReceiptVouchered = 1 WHERE Contract_DPA_ IN (" & ContractsSel &  ")"
				
													conn.Execute SQLServerFormat(HandleQuote(sqlStr))
											end if
									else
											receiptVoucher = "Null"
									End If
							end if
							
							'update receipt info	
							sqlStr = "UPDATE Payment SET BankAccount_DPA_ = " & bank & _
									", PaymentAmount = " & ccur(amount) & _
									", PaymentPDate = #" & FormatDate(pDate) & "#" & _ 
									", PaymentReference = '" & reference & "'" & _
									", PaymentTypes_DPA_ = '" & PaymentType & "'" & _
									", PaymentNarrative = '" & narrative & "'" & _
									", TimeChanged = '" & FormatDate(Now()) & "'" & _
									", EntityType_DPA_ = " & entity & " " & _
									", Entity_DPA_ = " & account & " " & _
									", ChangedBy = " & UserId & " " & _
									", Order_DPA_ = " & custOrder & " " & _
									", BrokerReceiptVoucher_DPA_ = " & receiptVoucher & " " & _
									" WHERE Payment_DPA_= " & ID
							sqlStr = SQLServerFormat(HandleQuote(sqlStr))
							conn.Execute sqlStr
							
							if(Cint(entity)=1) then
								'conn.execute ("Exec ClientTotalProcedure " & account)							
								conn.execute ("Exec ClientBalanceProcedure " & account)		
							end if
							
					conn.CommitTrans		
					
					if instr(1,buttonAction,"PRINT") > 0 then
						%>
						<SCRIPT LANGUAGE="JAVASCRIPT">
							window.parent.parent.frames['maininfo'].location.reload();
						</SCRIPT>
						<%
						WriteDialogRelocateScript "ReceiptForm.asp?ID=" & ID
					else
					    WritefraEnabledDialogCloseScript
					End if
					
				    conn.Close
					Set conn = Nothing
					
					Response.End
				end if
				
				if instr(1,buttonAction,"PRINT") > 0 then
					%>
					<SCRIPT LANGUAGE="JAVASCRIPT">
						window.parent.parent.frames['maininfo'].location.reload();
					</SCRIPT>
					<%
					WriteDialogRelocateScript "ReceiptForm.asp?ID=" & ID
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
		case "FETCH_ACCOUNTS"
				currentEntityType = cint(Request.Form("CurrentEntityTypeID"))
   	end select
   	
 
Set conn = GetActiveConnection("KBroker")
sqlStr = "SELECT * FROM Payment WHERE Payment_DPA_=" & ID
        
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		window.self.ShowMessage "The selected <%=DataEntity%> cannot be retrieved for editing"
                		
                </script>
                <% response.end
        End If
%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit <%=DataEntity%></title>

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
 							frm<%=DataSource%>.elements("CurrentEntityTypeID").value = itemID
 							frm<%=DataSource%>.elements("action").value = "Fetch_Accounts"
 							frm<%=DataSource%>.submit
 							
					end function
</script>
<script language = "javascript" >
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
		var pageJustLoaded = "1";
		
		
		//==========BEGIN REMOVE OPTION/S FROM DROP-DOWN FUNCTION ON THE FLY=====
		function RemoveOptions(Field){		   
		   if (Field.length==0) return;
		  
		   for (loop=Field.length - 1; loop >= 0; loop--) {
		       var GoneOption = Field.options[loop]		  
		       Field.remove(GoneOption.index);		        
		       }
		   
		 }

		//==============END REMOVE OPTION/S FUNCTION====================
		
		
		
		//var totalContractAmt = 0;
		<%if  isnull(rs.Fields("BrokerReceiptVoucher_DPA_")) then%>
				var totalContractAmt = 0;
		<%else%>
				var totalContractAmt = <%=rs.Fields("PaymentAmount")%>;
		<%end if%>
	
		function UpdateVoucherAmount(inAmount, inAction){			

				
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
		
	function evaluateEntity(Val, Entity)
	{
      //Enable Printing if entity is Broker or Client
	  if (Val == 1 || Val == 3){
	    document.getElementById("cmdPrint").style.display = ""
	  }
	  else {
	    document.getElementById("cmdPrint").style.display = "none"
	  }
	  
	  //Furher Processing
	  FetchAccounts(Entity)
	}	
 function FormatPrice()
	{
	 document.all.item("txtTotal").value= formatNum(document.all.item("txtTotal").value);
	}

</script>


</head>

<body Class="Dialog" >
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<SCRIPT language="JavaScript">
	var cal=new ctlSpiffyCalendarBox("cal", "frm<%=DataSource%>", "txtDate","cmdDate","<%=  FormatDate(rs.Fields("PaymentPDate")) %>",1);
</SCRIPT>

<form name = 'frm<%=DataSource%>' method = 'post' action = '<%=DataSource%>Test.asp' id = "frmMain">
<table border="0" width="100%" cellspacing="1" cellpadding="1">
  <tr>
    <td width="30%">Receipt No</td>
    <td width="54%"><input readonly = 'true' class=readonly  type = 'text' name ='txtReceipt' id = 'txtReceipt' size="20" value = '<%=rs.Fields("PaymentReceiptNo")%>'></td>
    <td width="31%">

	</td>
  </tr>
  <tr>
    <td width="15%">Payment Type</td>
    <td width="54%"><select name = 'cboPaymentTypes' id = 'cboPaymentTypes' size="1" >
    <%
  		
		if currentPaymentType = 0 then
			currentPaymentType = rs.Fields("PaymentTypes_DPA_")
		end if
		
        sqlStr = "SELECT * FROM [PaymentTypes] Order By Description"
        Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rsEdit.EOF Or rsEdit.BOF) Then
                rsEdit.MoveFirst
                Do Until rsEdit.EOF
                		if rsEdit.Fields("PaymentTypes_DPA_") = currentPaymentType Then%>
                			<option selected value = '<%=rsEdit.Fields("PaymentTypes_DPA_")%>'><%=rsEdit.Fields("Description")%></option>
                		<%else%>
                        <option value = '<%=rsEdit.Fields("PaymentTypes_DPA_")%>'><%=rsEdit.Fields("Description")%></option>
                     <%end if
						rsEdit.MoveNext
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
    <td width="35%"><select name = 'cboEntity' id = 'cboEntity' size="1" onchange='evaluateEntity(this.value,this)'>
<%		
		if currentEntityType = 0 then
			currentEntityType = rs.Fields("EntityType_DPA_")
		end if
        sqlStr = "SELECT * FROM [FullEntityTypeList] Order By EntityTypeName"
        Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rsEdit.EOF Or rsEdit.BOF) Then
                rsEdit.MoveFirst
                Do Until rsEdit.EOF
                		if rsEdit.Fields("EntityType_DPA_") = currentEntityType Then%>
                			<option selected value = '<%=rsEdit.Fields("EntityType_DPA_")%>'><%=rsEdit.Fields("EntityTypeName")%></option>
                		<%else%>
                        <option value = '<%=rsEdit.Fields("EntityType_DPA_")%>'><%=rsEdit.Fields("EntityTypeName")%></option>
                     <%end if
						rsEdit.MoveNext
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
    <input type = 'text' name ='txtClientCode' id = 'txtClientCode' size="10" onBlur="txtval = this.value; if(pageJustLoaded == '0'){selectItem(cboAccount); FilterData(cboAccount,<%=currentEntityType%>,true)} else {pageJustLoaded = '0'}">
    <select name = 'cboAccount' id = 'cboAccount' size="1"
    onchange='FilterData(this,<%=currentEntityType%>,UpdateCode(true,cboAccount,txtClientCode))'     
	onKeypress="return (dodefaultaction()==''); " 
	onKeydown="return (dodefaultaction()==''); " 
	onKeyup="return (FilterData(this,<%=currentEntityType%>,UpdateCode(change(cboAccount,0),cboAccount,txtClientCode)));" 
	onfocus="txtval = '';inputIsItemCode = 1;" 
	onblur="txtval = '';inputIsItemCode = 1;">
	
	<option selected SearchCode = "" SearchText = ""  value = ''></option>
 
<%
		Dim entityCode
        Dim xFilter
        Dim displayField
		
		displayField = "EntityName"
        
        sqlStr = "SELECT Entity_DPA_,LTRIM(RTRIM(EntityName)) as EntityName,EntityCode FROM [CompleteEntityList] WHERE EntityType_DPA_ =" & currentEntityType & " Order By EntityName"
        Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rsEdit.EOF Or rsEdit.BOF) Then
                rsEdit.MoveFirst
                Do Until rsEdit.EOF
                AccountName=Mid(rsEdit.Fields(displayField),1,30)
                		if rsEdit.Fields("Entity_DPA_") = rs.Fields("Entity_DPA_") Then
                			entityCode = rsEdit.Fields("EntityCode")%>
                			<option selected SearchCode = "<%=rsEdit.Fields("EntityCode")%>" SearchText = "<%=rsEdit.Fields("EntityName")%>" value = '<%=rsEdit.Fields("Entity_DPA_")%>'><%=AccountName%></option>
                		<%else%>
							<option SearchCode = "<%=rsEdit.Fields("EntityCode")%>" SearchText = "<%=rsEdit.Fields("EntityName")%>" value = '<%=rsEdit.Fields("Entity_DPA_")%>'><%=AccountName%></option>
                     <%end if
						rsEdit.MoveNext
                Loop
        End If
%>

    </select></td>
    <td width="31%">

	</td>
  </tr>
  <tr id="orderRow" style="display: none">
			    <td width="15%">Order</td>
			    <td width="35%">
			    <select name = 'cboOrderBag' id = "cboOrderBag" size="1" style="display:none">
			<%
			        sqlStr = "SELECT * FROM [OrderListPlain]"
			        Set rsEdit = CreateObject("ADODB.Recordset")
			        rsEdit.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic 
			        If Not (rsEdit.EOF Or rsEdit.BOF) Then
			                rsEdit.MoveFirst
			                Do Until rsEdit.EOF%>
			                        <option value = '<%=rsEdit.Fields("Order_DPA_")%>' ClientTag = '<%=rsEdit.Fields("Client_DPA_")%>'><%=rsEdit.Fields("Order_DPA_")%></option>
			                        <%rsEdit.MoveNext
			                Loop
			        End If
			%>
			    </select>
			    <select name = 'cboOrder' id = "cboOrder" size="1">
					<option selected value = '0'></option>
					
					<%
					rsEdit.Filter = "Client_DPA_ = " & rs.Fields("Entity_DPA_")
					If Not (rsEdit.EOF Or rsEdit.BOF) Then
			                rsEdit.MoveFirst
			                Do Until rsEdit.EOF
			                        if rsEdit.Fields("Order_DPA_") = rs.Fields("Order_DPA_") Then%>
                							<option  selected value = '<%=rsEdit.Fields("Order_DPA_")%>' ClientTag = '<%=rsEdit.Fields("Client_DPA_")%>'><%=rsEdit.Fields("Order_DPA_")%></option>
                					<%else%>
											<option value = '<%=rsEdit.Fields("Order_DPA_")%>' ClientTag = '<%=rsEdit.Fields("Client_DPA_")%>'><%=rsEdit.Fields("Order_DPA_")%></option>
									<%end if
			                        rsEdit.MoveNext
			                Loop
			        End If %>
			    
			    </select></td>
			    <td width="31%">

				</td>
			  </tr>
			  
			  <script language="javascript">
					UpdateCode(true,document.all.item("cboAccount"),document.all.item("txtClientCode"));
			  </script>
  <tr id="brokerVoucherRow" style="display: none">
		<td width="15%" valign="top">Contracts</td>
		<td width="54%">
		<%if (currentEntityType = 3) then
				if isnull(rs.Fields("BrokerReceiptVoucher_DPA_")) then
						xFilter = "(BrokerReceiptVouchered = 0)"
				else
						xFilter = "(BrokerReceiptVouchered = 0 OR BrokerReceiptVoucher_DPA_ = " & rs.Fields("BrokerReceiptVoucher_DPA_") & ")"
				end if%>
				
				<iframe id="fraInnerBrokerSelects" name="fraInnerBrokerSelects" width="400px" height="200px" src="inner_select_voucherBroker.asp?OrderTypeSale=1&BrokerCode=<%= entityCode %>&XTra=<%=xFilter%>"></iframe>
		<%else%>
				<iframe id="fraInnerBrokerSelects" name="fraInnerBrokerSelects" width="400px" height="200px" src="inner_select_voucherBroker.asp"></iframe>
		<%end if%>
				<BR>
		</td>
		<td width="31%">
			<input type = 'hidden' name ='txtVoucherType' id = 'txtVoucherType' value="<%=currentEntityType%>">
		</td>
</tr>
		<script language="javascript">
					currentEntityType = <%=currentEntityType%>
					
					if (currentEntityType == 3){
						document.getElementById('brokerVoucherRow').style.display = '';
						document.getElementById('orderRow').style.display = 'none';
					} 
					else{
						 if (currentEntityType == 1){
							document.getElementById('brokerVoucherRow').style.display = 'none';
							document.getElementById('orderRow').style.display = '';
							document.getElementById('txtVoucherType').value = '0';
						}			
						else {
							document.getElementById('brokerVoucherRow').style.display = 'none';
							document.getElementById('orderRow').style.display = 'none';
							document.getElementById('txtVoucherType').value = '0';
						}
					}
		</script>
  <tr>
    <td width="15%">Amount</td>
    <td width="54%"><input type = 'text' name ='txtTotal' id = 'txtTotal'  STYLE="text-align:right" size="20" value="<%=FormatNum(rs.Fields("PaymentAmount"))%>" onchange="JavaScript: FormatPrice()"></td>
    <td width="31%">
	<%
			Dim selContracts
			Dim contRS
			
			if  Not(isnull(rs.Fields("Voucher_DPA_"))) then
					'get broker contracts
					sqlStr = "SELECT Contract_DPA_ FROM  Contract WHERE (Voucher_DPA_ = " & rs.Fields("Voucher_DPA_") & ")"
					Set contRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
					If not(contRS.EOF Or contRS.BOF) Then
							do until contRS.EOF
									if trim(selContracts) = "" then
											selContracts = selContracts & contRS.Fields("Contract_DPA_")
									else
											selContracts = selContracts & "," & contRS.Fields("Contract_DPA_")
									end if
									contRS.MoveNext
							loop
					End If
			else
					'no vouchers
					selContracts = ""
			end if
	%>
	<input type="hidden" name="ContractsSel" id="ContractsSel" value="<%=selContracts%>">
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
        Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rsEdit.EOF Or rsEdit.BOF) Then
                rsEdit.MoveFirst
                Do Until rsEdit.EOF
                		if rsEdit.Fields("Account_DPA_") = rs.Fields("BankAccount_DPA_") Then%>
                			<option selected SearchCode = "<%=rsEdit.Fields("AccountCode")%>" SearchText = "<%=rsEdit.Fields("AccountName")%>" value = '<%=rsEdit.Fields("Account_DPA_")%>'><%=rsEdit.Fields("AccountNameEx")%></option>
                		<%else%>
							<option SearchCode = "<%=rsEdit.Fields("AccountCode")%>" SearchText = "<%=rsEdit.Fields("AccountName")%>" value = '<%=rsEdit.Fields("Account_DPA_")%>'><%=rsEdit.Fields("AccountNameEx")%></option>
                     <%end if
						rsEdit.MoveNext
                Loop
        End If
%>

    </select></td>
    <td width="31%">

	</td>
  </tr>
  
  <tr>
    <td width="15%">Reference</td>
    <td width="54%"><input type = 'text' name ='txtRef' id = 'txtRef' size="20" value = '<%=rs.Fields("PaymentReference")%>'></td>
    <td width="31%">

	</td>
  </tr>
  <tr>
    <td width="15%">Narrative</td>
    <td width="54%">
    <textarea name ='txtNarrative' id = 'txtNarrative' rows="1" cols="20" ><%=rs.Fields("PaymentNarrative")%></textarea></td>
    <td width="31%">

	</td>
  </tr>
  <tr>
	  <td width="100%" colspan=3 align="center" valign=absBottom>
		<BR><BR>
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdPrint' value=" Save & Print " onclick = "AllowedNavigation()">
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value="Save" onclick = "AllowedNavigation()">
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAddPrint' value="Print" onclick = "AllowedNavigation()">
    	<input type = 'button' Class=Buttons name ='cmdCancel' id = "cmdCancel" value="Cancel" onclick = "JavaScript: window.self.close()">
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		<input type = 'hidden' name ='ID' id = 'ID' value="<%= Rs.Fields("Payment_DPA_").Value %>">
		<input type = 'hidden' name ='CurrentEntityTypeID' id = 'CurrentEntityTypeID'>
	</td>
  </tr>
</table>

</form>
</body>

</html>