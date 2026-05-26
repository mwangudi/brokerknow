<!--#include virtual="libroutines.asp"-->
<%
	const LinkedIndependent = 1
	const LinkedDependent = 2
   
	const UDLName = "KBroker"
	const DataSource = "EditPayment"
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
	dim currentPaymentType
	
	UserId=Session("UserID")
	
	action = ucase(Request.Form("action"))
	ID = Request("ID")
	currentEntityType = 0

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
			buttonAction = Trim(Ucase(Request.Form("ButtonAction")))
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
					Dim PaymentType
					Dim Ref 'PaymentTypes
			        
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
			        
					'validate Entity
					If Trim(PaymentType) = "" Then%>
					         <script language = 'vbscript'>
					         		ShowMessage "Please specify the Payment Types."
					         		
					         </script>
					         <% 
							  ReloadPage(ID)
							 response.end
					 End If
					 
					 If Trim(entity) = "" Then%>
					         <script language = 'vbscript'>
					         		ShowMessage "Please specify the Entity"
					         		
					         </script>
					         <% 
							  ReloadPage(ID)
							 response.end
					 End If
					 
					 'validate Account
					 If Trim(Account) = "" Then%>
					         <script language = 'vbscript'>
					         		ShowMessage "Please specify the Account"
					         		
					         </script>
					         <% 
							  ReloadPage(ID)
							 response.end
					 End If
					 'validate Bank
					 If Trim(Bank) = "" Then%>
					         <script language = 'vbscript'>
					         		ShowMessage "Please specify the Bank"
					         		
					         </script>
					         <% 
							  ReloadPage(ID)
							 response.end
					 End If
					'validate Amount
					If Trim(Amount) = "" Then%>
					    <script language = 'vbscript'>
					         	ShowMessage "Please specify the Amount "
					         	
					    </script>
					    <% 
						 ReloadPage(ID)
						response.end
					End If
					'ensure Amount is numeric
					If (Amount <> "") And (Not IsNumeric(Amount)) Then%>
					    <script language = 'vbscript'>
							ShowMessage "Order Detail Estimated Amount must be numeric"
							
					    </script>
					    <% 
						 ReloadPage(ID)
						response.end
					End If
					
					'validate size of Reference 
					 If Len(Reference) > 20 Then%>
					         <script language = 'vbscript'>
					         ShowMessage "Reference can only be 20 characters in length"
					         
					         </script>
					         <% 
							  ReloadPage(ID)
							 response.end
					 End If
					'validate size of Narrative 
					 If Len(narrative) > 200 Then%>
					         <script language = 'vbscript'>
					         ShowMessage "Narrative can only be 200 characters in length"
					         
					         </script>
					         <% 
							  ReloadPage(ID)
							 response.end
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
							<% 
							 ReloadPage(ID)
							response.end

							else
							    Ref = rs("Reference")
								if (Cint(Ref) = 1 And reference = "") then%>
								
							       <script language = 'vbscript'>
									   ShowMessage "The Receipt Payment Type selected requires a reference."
											         			
								    </script>
							      <% 
								   ReloadPage(ID)
								  response.end
							   End if 
							End If
							
					if trim(UCase(ChequeStatus)) = "ON" then
						ChequeStatus = "AWAITING COLLECTION"
					else
						ChequeStatus = "NOT COLLECTED"
					end if
					
					Dim clientVoucher
					Dim brokerVoucher
					Dim depRS
					
					Set conn = GetActiveConnection("KBroker")
					
					conn.BeginTrans
					
							'deal with vouchers
							sqlStr = "SELECT ClientVoucher_DPA_, Voucher_DPA_ FROM [Payment] WHERE [Payment_DPA_] = " & ID
							Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
							If (rs.EOF Or rs.BOF) Then%>
							         	<script language = 'vbscript'>
							         			ShowMessage "A serious error has been encountered while saving the data. Try saving again"
											         			
							         	</script>
							         	<% 
										 ReloadPage(ID)
										response.end
							End If
							
							clientVoucher = rs.fields("ClientVoucher_DPA_").value
							brokerVoucher = rs.fields("Voucher_DPA_").value
							
							if not(isnull(clientVoucher)) then
									brokerVoucher = "Null"
									'handle client voucher
									'find out whether any child records exist
									sqlStr = "SELECT Child,DeletionMessage,ParentKey FROM [_Parent_Child_Links_] WHERE (Parent = 'ClientVoucher') AND (ChildType = " & LinkedIndependent & ")"
									Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
									If Not (rs.BOF Or rs.EOF) Then
									        rs.MoveFirst
									        Do Until rs.EOF
									        		tableName = rs.Fields("Child")
									                sqlStr = "SELECT TOP 1 * FROM [" & tableName & "] WHERE ClientVoucher_DPA_ = " & clientVoucher
									                Set depRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
									                If Not (depRS.BOF Or depRS.EOF) Then%>
									        				<script language = 'vbscript'>
									        					ShowMessage <%=rs.Fields("DeletionMessage")%>
											        					
									        				</script>
									        				<%
															 ReloadPage(ID)
															response.end
									                End If
									                rs.MoveNext
									        Loop
									End If
									
									If (Trim(ContractsSel) = "") Then
											'voucher removed
											'remove contracts 
											sqlStr = "UPDATE Contract SET ClientVoucher_DPA_ = NULL" & _
													", ContractClientVouchered = 0 WHERE ClientVoucher_DPA_ =" & clientVoucher
											conn.Execute SQLServerFormat(HandleQuote(sqlStr))
											
											'delete from database
											sqlStr = "DELETE FROM [ClientVoucher] WHERE ClientVoucher_DPA_ = " & clientVoucher
											conn.Execute SQLServerFormat(HandleQuote(sqlStr))
											
											clientVoucher = "Null"
											voucherType = 0
									else
											'voucher modified
											sqlStr = "UPDATE [ClientVoucher] SET VoucherDate = #" & FormatDate(PDate) & "#" & _
													"        WHERE ClientVoucher_DPA_ = " & clientVoucher
											conn.Execute SQLServerFormat(HandleQuote(sqlStr))
														
											'remove previous contracts 
											sqlStr = "UPDATE Contract SET ClientVoucher_DPA_ = NULL" & _
													", ContractClientVouchered = 0 WHERE ClientVoucher_DPA_ =" & clientVoucher
														
											conn.Execute SQLServerFormat(HandleQuote(sqlStr))
														
											'add contracts to voucher
											sqlStr = "UPDATE Contract SET ClientVoucher_DPA_ = " & clientVoucher & _
													", ContractClientVouchered = 1 WHERE Contract_DPA_ IN (" & ContractsSel &  ")"
														
											conn.Execute SQLServerFormat(HandleQuote(sqlStr))
									end if
							elseif not(isnull(brokerVoucher)) then
									clientVoucher = "Null"
									'handle broker voucher
									'find out whether any child records exist
									sqlStr = "SELECT Child,DeletionMessage,ParentKey FROM [_Parent_Child_Links_] WHERE (Parent = 'Voucher') AND (ChildType = " & LinkedIndependent & ")"
									Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
									If Not (rs.BOF Or rs.EOF) Then
									        rs.MoveFirst
									        Do Until rs.EOF
									        		tableName = rs.Fields("Child")
									                sqlStr = "SELECT TOP 1 * FROM [" & tableName & "] WHERE Voucher_DPA_ = " & brokerVoucher
									                Set depRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
									                If Not (depRS.BOF Or depRS.EOF) Then%>
									        				<script language = 'vbscript'>
									        					ShowMessage <%=rs.Fields("DeletionMessage")%>
									        					
									        				</script>
									        				<%
															 ReloadPage(ID)
															response.end
									                End If
									                rs.MoveNext
									        Loop
									End If
									If (Trim(ContractsSel) = "") Then
											'voucher removed
											'remove contracts 
											sqlStr = "UPDATE Contract SET Voucher_DPA_ = NULL" & _
													", ContractVouchered = 0 WHERE Voucher_DPA_ =" & brokerVoucher
											conn.Execute SQLServerFormat(HandleQuote(sqlStr))
											'delete from database
											sqlStr = "DELETE FROM [Voucher] WHERE Voucher_DPA_ = " & brokerVoucher
											conn.Execute SQLServerFormat(HandleQuote(sqlStr))
											
											brokerVoucher = "Null"
											voucherType = 0
									else
											'voucher modified
											sqlStr = "UPDATE [Voucher] SET VoucherDate = #" & FormatDate(PDate) & "#" & _
													"        WHERE Voucher_DPA_ = " & brokerVoucher
											conn.Execute SQLServerFormat(HandleQuote(sqlStr))
														
											'remove previous contracts 
											sqlStr = "UPDATE Contract SET Voucher_DPA_ = NULL" & _
													", ContractVouchered = 0 WHERE Voucher_DPA_ =" & brokerVoucher
														
											conn.Execute SQLServerFormat(HandleQuote(sqlStr))
														
											'add contracts to voucher
											sqlStr = "UPDATE Contract SET Voucher_DPA_ = " & brokerVoucher & _
													", ContractVouchered = 1 WHERE Contract_DPA_ IN (" & ContractsSel &  ")"
														
											conn.Execute SQLServerFormat(HandleQuote(sqlStr))
									end if
							else
									'check for new vouchers
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
													         	<% 
																 ReloadPage(ID)
																response.end
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
							end if
							
							'update payment info	
							sqlStr = "UPDATE Payment SET BankAccount_DPA_ = " & bank & _
									", PaymentAmount = " & ccur(amount) & _
									", PaymentPDate = #" & FormatDate(pDate) & "#" & _ 
									", PaymentReference = '" & reference & "'" & _
									", PaymentTypes_DPA_ = '" & PaymentType & "'" & _
									", PaymentNarrative = '" & narrative & "'" & _
									", EntityType_DPA_ = " & entity & " " & _
									", Entity_DPA_ = " & account & " " & _
									", ChangedBy = " & UserId & " " & _
									", TimeChanged = '" & FormatDateandTime(now()) & "' " & _
									", Voucher_DPA_ = " & brokerVoucher & " " & _
									", ClientVoucher_DPA_ = " & clientVoucher & " " & _
									", ChequeCollection = '" & ChequeStatus & "' " & _
									" WHERE Payment_DPA_= " & ID
									
							sqlStr = SQLServerFormat(HandleQuote(sqlStr))
							
							Dim tableName
							
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
					
					if(Cint(entity)=1) then
						'conn.execute ("Exec ClientTotalProcedure " & account)							
						conn.execute ("Exec ClientBalanceProcedure " & account)		
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
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
 <SCRIPT language=Javascript src="../scripts/accountList.js"></SCRIPT>
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
			forceSubmit();
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
		
		<%if  (isnull(rs.Fields("Voucher_DPA_"))) and (isnull(rs.Fields("ClientVoucher_DPA_"))) then%>
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

	function forceSubmit()
	{
		setOpener();
		//var targetPage = window.dialogArguments.opener.document.all.item("frmMain").elements("EditPage").value;
				
		document.frm<%=DataSource%>.method='post';
		document.frm<%=DataSource%>.target='_self';
		document.frm<%=DataSource%>.submit();	
		
	}
	
	function setOpener()
	{
		window.self.opener = window.dialogArguments.opener;				
	}
	

	function hideButton()
	{
	 document.getElementById('hide').style.display='none';
	}
</script>
</head>

<body Class="Dialog"  onLoad="setOpener()">
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<SCRIPT language="JavaScript">
	var cal=new ctlSpiffyCalendarBox("cal", "frm<%=DataSource%>", "txtDate","cmdDate","<%=  FormatDate(rs.Fields("PaymentPDate")) %>",1);
</SCRIPT>

<form name = 'frm<%=DataSource%>' method = 'post' action = '<%=DataSource%>.asp' id = "frmMain">
<table border="0" width="100%" cellspacing="1" cellpadding="1">
  <tr>
    <td width="15%">Payment Type</td>
    <td width="54%"><select name = 'cboPaymentTypes' id = 'cboPaymentTypes' size="1" onchange='EnableCheqCollection(this)'>
    <%
  		
		if currentPaymentType = 0 then
			currentPaymentType = rs.Fields("PaymentTypes_DPA_")
		end if
		
		ChkDisable = "disabled" 'Default: Cheque collection diabled unless Payment Type is Check
		
        sqlStr = "SELECT * FROM [PaymentTypes] Order By Description"
        Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rsEdit.EOF Or rsEdit.BOF) Then
                rsEdit.MoveFirst
                Do Until rsEdit.EOF
                		if rsEdit.Fields("PaymentTypes_DPA_") = currentPaymentType Then
                		   
                		   if Ucase(trim(rsEdit.Fields("Description"))) = "CHEQUE" then ChkDisable = ""  'Enable Cheque collection
                		%>
                			<option selected SearchText = '<%=rsEdit.Fields("Description")%>' value = '<%=rsEdit.Fields("PaymentTypes_DPA_")%>'><%=rsEdit.Fields("Description")%></option>
                		<%else%>
                        <option  SearchText = '<%=rsEdit.Fields("Description")%>' value = '<%=rsEdit.Fields("PaymentTypes_DPA_")%>'><%=rsEdit.Fields("Description")%></option>
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
    <td width="35%"><select name = 'cboEntity' id = 'cboEntity' size="1" onchange='FetchAccounts(this)'>
<%		
		if currentEntityType = 0 then
			currentEntityType = rs.Fields("EntityType_DPA_")
		end if
		
		if isnull(rs.Fields("EntityType_DPA_")) then currentEntityType = 0 else currentEntityType = rs.Fields("EntityType_DPA_")
		
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
    <input type = 'text' name ='txtClientCode' id = 'txtClientCode' size="10" onBlur="txtval = this.value; selectItem(cboAccount); setFilter(this,<%=currentEntityType%>,UpdateCode(true,cboAccount,txtClientCode));">
    <select name = 'cboAccount' id = "cboAccount" size="1" 
	onchange='setFilter(this,<%=currentEntityType%>,UpdateCode(true,cboAccount,txtClientCode))'	
	onKeypress="return (dodefaultaction()==''); " 
	onKeydown="return (dodefaultaction()==''); " 
	onKeyup="return (change(cboAccount));" 
	onfocus="txtval = '';inputIsItemCode = 1;" 
	onLoad="setFilter(this,<%=currentEntityType%>,UpdateCode(true,cboAccount,txtClientCode))"
	onblur="txtval = '';inputIsItemCode = 1;">
	
	<option selected SearchCode = '0' SearchText = '' value = ''></option>
 
<%
        Dim entityCode
        Dim xFilter
        
        sqlStr = "SELECT * FROM [CompleteEntityList] WHERE EntityType_DPA_ =" & currentEntityType & " Order By EntityName"
        Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rsEdit.EOF Or rsEdit.BOF) Then
                rsEdit.MoveFirst
                Do Until rsEdit.EOF
                AccountName=rsEdit.Fields("EntityNameEx")
                		if rsEdit.Fields("Entity_DPA_") = rs.Fields("Entity_DPA_") Then
                			entityCode = rsEdit.Fields("EntityCode")%>
                			<option selected SearchCode = '<%=rsEdit.Fields("EntityCode")%>' SearchText = '<%=rsEdit.Fields("EntityName")%>' value = '<%=rsEdit.Fields("Entity_DPA_")%>'><%=AccountName%></option>
                		<%else%>
							<option SearchCode = '<%=rsEdit.Fields("EntityCode")%>' SearchText = '<%=rsEdit.Fields("EntityName")%>' value = '<%=rsEdit.Fields("Entity_DPA_")%>'><%=AccountName%></option>
                     <%end if
						rsEdit.MoveNext
                Loop
        End If
%>
    </select></td>
    <td width="31%">

	</td>
  </tr>
	<tr id="brokerVoucherRow" style="display: none">
		    <td width="15%" valign="top">Contracts</td>
		    <td width="54%">
		    <%if (currentEntityType = 3) then
					if isnull(rs.Fields("Voucher_DPA_")) then
							xFilter = "(ContractVouchered = 0)"
					else
							xFilter = "(ContractVouchered = 0 OR Voucher_DPA_ = " & rs.Fields("Voucher_DPA_") & ")"
					end if%>
					<iframe id="fraInnerBrokerSelects" name="fraInnerBrokerSelects" width="400px" height="200px" src="inner_select_voucherBroker.asp?OrderTypeSale=0&BrokerCode=<%= entityCode %>&XTra=<%=xFilter%>"></iframe>
			<%else%>
					<iframe id="fraInnerBrokerSelects" name="fraInnerBrokerSelects" width="400px" height="200px" src="inner_select_voucherBroker.asp"></iframe>
			<%end if%>
					<BR>
			</td>
		    <td width="31%">
				<input type = 'hidden' name ='txtVoucherType' id = 'txtVoucherType' value="<%=currentEntityType%>">
			</td>
		  </tr>
		  
		  <tr  id="clientVoucherRow">
		    <td width="15%" valign="top">Contracts</td>
		    <td width="54%">
			<%if (currentEntityType = 1) then
			
					if isnull(rs.Fields("ClientVoucher_DPA_")) then
							xFilter = "(ContractClientVouchered = 0)"
					else
							xFilter = "(ContractClientVouchered = 0 OR ClientVoucher_DPA_ = " & rs.Fields("ClientVoucher_DPA_") & ")"
					end if%>
					
					<iframe id="fraInnerSelects" name="fraInnerSelects" width="400px" height="200px" src="inner_select_vouchers.asp?OrderTypeSale=1&Client_DPA_=<%= entityCode %>&XTra=<%=xFilter%>"></iframe>
			<%else%>
					<iframe id="fraInnerSelects" name="fraInnerSelects" width="400px" height="200px" src="inner_select_vouchers.asp"></iframe>
			<%end if%>
						<BR>
		    </td>
		    <td width="31%">
				
			</td>
		  </tr>
		<%
		
		
		if trim(UCase(rs.Fields("ChequeCollection"))) = "AWAITING COLLECTION" then
			ChequeStatus = " checked"					
		elseif trim(UCase(rs.Fields("ChequeCollection"))) = "COLLECTED" then			
			ChequeStatus = " "
		end if
		
		'if Ucase(trim(rs.Fields("PaymentTypes_DPA_"))) <> "CHEQUE" then
		'	ChkDisable = "disabled"
		'else
		'	ChkDisable = " "
		'end if
		%>
		<tr>
			<td width="15%" align="right">&nbsp;</td>
			<td width="85%" colspan="2">
            <input type="checkbox" name="ChequeCollection" <%=ChkDisable%><%=ChequeStatus%> value="ON">&nbsp;&nbsp;&nbsp;Cheque Collection</td>
		</tr>
		<script language="javascript">
			currentEntityType = <%=currentEntityType%>

			if (currentEntityType == 1)
			{
				document.getElementById('brokerVoucherRow').style.display = 'none'
				document.getElementById('clientVoucherRow').style.display = ''   
			}
			else
			{
				if (currentEntityType == 3)
				{
					document.getElementById('brokerVoucherRow').style.display = ''
					document.getElementById('clientVoucherRow').style.display = 'none'   
				}
				else
				{
					document.getElementById('brokerVoucherRow').style.display = 'none'
					document.getElementById('clientVoucherRow').style.display = 'none'   
				}
			}
			
		</script>		  
	
  <tr>
    <td width="15%">Amount</td>
    <td width="54%"><input STYLE="text-align:right" type = 'text' name ='txtTotal' id = 'txtTotal' size="20" value='<%=FormatNum(rs.Fields("PaymentAmount"))%>' onchange="JavaScript: FormatPrice()"></td>
    <td width="31%">
    <%	Dim selContracts
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
	elseif Not(isnull(rs.Fields("ClientVoucher_DPA_"))) then
			'get client contracts
			sqlStr = "SELECT Contract_DPA_ FROM  Contract WHERE (ClientVoucher_DPA_ = " & rs.Fields("ClientVoucher_DPA_") & ")"
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
	end if%>
			
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
    <td width="54%"><input type = 'text' name ='txtRef' id = 'txtRef' size="20" value='<%=rs.Fields("PaymentReference")%>'></td>
    <td width="31%">

	</td>
  </tr>
  <tr>
    <td width="15%">Narrative</td>
    <td width="54%">
    <textarea name ='txtNarrative' id = 'txtNarrative' rows="1" cols="20" > <%=rs.Fields("PaymentNarrative")%></textarea></td>
    <td width="31%">

	</td>
  </tr>
  <tr>
	  <td width="100%" colspan=3 align="right" valign=absBottom>
		<BR><BR>
		<b  name="hide" id="hide"><input type = 'button' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value="Save" onclick = "AllowedNavigation(); hideButton();"></b>
    	<input type = 'button' Class=Buttons name ='cmdCancel' id = "cmdCancel" value="Cancel" onclick = "JavaScript: window.self.close()">
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		<input type = 'hidden' name ='ID' id = 'ID' value="<%= Rs.Fields("Payment_DPA_").Value %>">
		<input type = 'hidden' name ='CurrentEntityTypeID' id = 'CurrentEntityTypeID'>
		<input type = 'hidden' name ='buttonAction' id = 'action' value="Save">
	</td>
  </tr>
</table>

</form>
</body>

</html>