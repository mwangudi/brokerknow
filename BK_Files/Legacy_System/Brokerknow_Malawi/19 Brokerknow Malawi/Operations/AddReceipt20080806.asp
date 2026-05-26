<!--#include virtual="libroutines.asp"-->
<%
	const UDLName = "KBroker"
	const DataSource = "AddReceipt"
	const DataEntity = "Receipt"
	const DataEntityPlural = "Receipts"
	const ActionFolder = "Operations"
	
	Dim UserId
	Dim action
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim guid
	Dim guidStr
	Dim ID
	dim currentEntityType
	
	UserId=Session("UserID")
	
	action = ucase(Request.Form("action"))
	ID = Request.Form("ID")
	currentEntityType = 1
	
	select case action
		case "EXECUTE"
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
			        'response.end
						

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
					custOrder = iif(custOrder = 0,"Null",custOrder)
					    
					
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
						
					'save data
					Dim receiptVoucher
					
					Set conn = GetActiveConnection("KBroker")
					conn.BeginTrans		
							'create voucher
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
							
							set guid = server.createobject("NDUtils.CGUID")
							guidStr = guid.GenerateGUID
									      
							sqlStr = "INSERT INTO [Payment] ( PaymentAmount, PaymentPDate, BankAccount_DPA_, " & _
									"Payment_EIT_,PayType_DPA_, EntityType_DPA_, Payment_DPA_, PaymentReference, PaymentTypes_DPA_, PaymentNarrative, Entity_DPA_, Order_DPA_,ChangedBy,TimeChanged,CreatedBy,TimeCreated,BrokerReceiptVoucher_DPA_,PaymentReceiptNo) " & _
									"SELECT " & " " & ccur(amount) & " " & " as PaymentAmount," & "#" & FormatDate(PDate) & "#" & " as PaymentPDate" & _
									"," & " " & bank & " " & " as BankAccount_DPA_" & _
									"," & "'" & guidStr & "'" & " as Payment_EIT_" & "," & "1 as PayType_DPA_" & _
									"," & " " & entity & " " & " as EntityType_DPA_" & _
									"," & " " & "iif(isnull(max([Payment_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Payment'),max([Payment_DPA_]) + 1)" & " " & " as Payment_DPA_" & _
                                    "," & "'" & reference & "'" & " as PaymentReference" & _
									"," & "'" & PaymentType & "'" & " as PaymentTypes_DPA_" & _
									"," & "'" & narrative & "'" & " as PaymentNarrative" & _
									"," & " " & account & " " & " as Entity_DPA_" & _
									"," & " " & custOrder & " " & " as Order_DPA_" & _
									"," & " " & UserId & " " & " as ChangedBy" & _
									"," & "getdate() " & " as TimeChanged" & _
									"," & " " & UserId & " " & " as CreatedBy" & _
									"," & "getdate() " & " as TimeCreated" & _
									"," & " " & receiptVoucher & " " & " as BrokerReceiptVoucher_DPA_" & _
									"," & " " & "(SELECT IIf(IsNull(Max([PaymentReceiptNo])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Receipt'),max([PaymentReceiptNo]) + 1) AS PaymentReceiptNo FROM Payment WHERE PayType_DPA_ = 1)" & " " & " as PaymentReceiptNo" & _
									" FROM [Payment]"
									
							sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))														
							conn.Execute sqlStr
							
							if(Cint(entity)=1) then
								%>
									<SCRIPT LANGUAGE="JAVASCRIPT">
										//alert('hapa');
									</SCRIPT>
								<%
								conn.execute ("Exec ClientTotalProcedure " & account)							
								conn.execute ("Exec ClientStatementProcBrief " & account)
								'conn.execute ("Exec ClientBalanceProcedure " & account)
								conn.execute(" Exec UpdateIndividualClientBalance '"& account &"', '"& ccur(amount) &"'")
								
							end if	
							
					conn.CommitTrans
					
					'Retrieve Record ID
					sqlStr = "Select Payment_DPA_ From Payment Where Payment_EIT_ Like '%" & guidStr & "%'"
					Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
					If (rs.EOF Or rs.BOF) Then%>
					         	<script language = 'vbscript'>
					         			ShowMessage "An error has been encountered while saving the order. Try editing the Order if you wish to add more entries"
					         			
					         	</script>
					         	<% response.end
					End If
					
					if instr(1,buttonAction,"PRINT") > 0 then
							%>
							<SCRIPT LANGUAGE="JAVASCRIPT">
								window.parent.parent.frames['maininfo'].location.reload();
							</SCRIPT>
							<%
						WriteDialogRelocateScript "ReceiptForm.asp?ID=" & rs.Fields("Payment_DPA_")
					else
					    WritefraEnabledDialogCloseScript
					End if
					
				    conn.Close
					Set conn = Nothing
					
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
		
				
		//==========BEGIN REMOVE OPTION/S FROM DROP-DOWN FUNCTION ON THE FLY=====
		function RemoveOptions(Field){		   
		   if (Field.length==0) return;
		  
		   for (loop=Field.length - 1; loop >= 0; loop--) {
		       var GoneOption = Field.options[loop]		  
		       Field.remove(GoneOption.index);		        
		       }
		   
		 }

		//==============END REMOVE OPTION/S FUNCTION====================

		var currentEntityType = <%=currentEntityType%>
		
		
		var totalContractAmt = 0;
	
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
	  if (Val == 1 || Val == 3 || Val == 5){
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
	function hideButton()
	{
	 document.getElementById('hide').style.display='none';
	}

	
		function ClearFields(element)
		{
		   if (element == 'txtClientCode')
		   {
			document.frmAddReceipt.elements("txtClientCode").value = '';
			document.frmAddReceipt.elements("txtcdsno").value = 'CDS No.';
			document.frmAddReceipt.elements("txtclientname").value = 'Client Name';
			return;
		   }
		   if (element == 'txtcdsno')
		   {
			document.frmAddReceipt.elements("txtClientCode").value = 'Code';
			document.frmAddReceipt.elements("txtcdsno").value = '';
			document.frmAddReceipt.elements("txtclientname").value = 'Client Name';
			return;
		   }
		   if (element == 'txtclientname')
		   {
		    document.frmAddReceipt.elements("txtclientname").value = '';
			document.frmAddReceipt.elements("txtClientCode").value = 'Code';
			document.frmAddReceipt.elements("txtcdsno").value = 'CDS No.';
			return;
		   }		
		   
		}

		function updatefields(selectedclient)
		{
		 document.frmAddReceipt.elements("txtclientname").value = '';
		 document.frmAddReceipt.elements("txtcdsno").value = '';
		 document.frmAddReceipt.elements("txtClientCode").value = selectedclient.value;
		 //alert(selectedclient.value);
			LoadMyClient();
		}

		function LoadMyClient()
		{
			var clientcode = document.frmAddReceipt.elements("txtClientCode").value
			var clientcds = document.frmAddReceipt.elements("txtCdsNo").value
			var clientcobo = document.getElementById("cboAccount");	
			var entitycobo = document.getElementById("cboEntity");
			var guidstr = Math.random();
			//alert(entitycobo.value);
			if (entitycobo.value != 1) 
			 {
				return;
			 }
			xmlhttp = createXMLHTTPObj();
				
			url="GetList.asp?clientcode="+clientcode+"&cdsno="+clientcds+"&clientname=&action=SLoadClient&guidstr="+guidstr;
			
			xmlhttp.open("GET",url,true);

			xmlhttp.onreadystatechange=function() 
			{
				if (xmlhttp.readyState==4) 
				{
					returnStr = xmlhttp.responseText;
					returnStr = getBodyHTML(returnStr);
									
					myArray = returnStr.split("<->");
					//alert(myArray);
									
					document.frmAddReceipt.elements("txtclientname").value = myArray[7];
					document.frmAddReceipt.elements("txtClientCode").value = myArray[5];
					document.frmAddReceipt.elements("txtCdsNo").value = myArray[9]; 
				}
		   }
				 
		xmlhttp.setRequestHeader('Accept','message/x-jl-formresult');
		xmlhttp.send();
		}

		function LoadClient(accountno, element)
		{
		 var clientcode = document.frmAddReceipt.elements("txtClientCode").value;
		 var clientcds = document.frmAddReceipt.elements("txtcdsno").value;
		 var clientname = document.frmAddReceipt.elements("txtclientname").value;
		 var clientcobo = document.getElementById("cboAccount");		 
		 var entitycobo = document.getElementById("cboEntity");
		 //alert(entitycobo.value)
	     var x_clientname;
		
		 if (entitycobo.value != 1) 
		 {
		 	return;
		 }
		 
		 if (element == 'txtClientCode')
		 {
			clientcds = '';
			clientname = '';
			
			if (clientcode == '')
			{
			document.frmAddReceipt.elements("txtClientCode").value = 'Code'
			clientcobo.length = 1;
			clientcobo[0].text = 'Load Account';
			clientcobo[0].value = '';
			loadIframe(0);
			return;
			}
			
		 }
		 else if (element == 'txtcdsno')
		 {
			clientcode = '';
			clientname = '';

			if (clientcds == '')
			{
			document.frmAddReceipt.elements("txtcdsno").value = 'CDS No.'
			clientcobo.length = 1;
			clientcobo[0].text = 'Load Account';
			clientcobo[0].value = '';
			loadIframe(0);
			return;
			}
						
		 }
		 else if (element == 'txtclientname')
		 {
			clientcode = '';
			clientcds = '';

			if (clientname == '')
			{
			document.frmAddReceipt.elements("txtclientname").value = 'Client Name';
			clientcobo.length = 1;
			clientcobo[0].text = 'Load Account';
			clientcobo[0].value = '';
			loadIframe(0);
			return;
			}
			
		 }

				xmlhttp = createXMLHTTPObj();
				
				url="GetList.asp?clientcode="+clientcode+"&cdsno="+clientcds+"&clientname="+clientname+"&action=SLoadClient";
				
				//alert(url);

				xmlhttp.open("GET",url,true);

				xmlhttp.onreadystatechange=function() 
				  {
							if (xmlhttp.readyState==4) 
							{
								returnStr = xmlhttp.responseText;
								returnStr = getBodyHTML(returnStr);
								
								//alert(returnStr);

								myArray = returnStr.split("<->");
								
								x_clientname = myArray[7];

								if (x_clientname.length > 12) 
								{
									x_clientname = x_clientname.substring(0,16)  + '...';
								}
								
								//document.getElementById("cboAccount").options.length = 0;
								clientcobo.length = 1;
								if (element != 'txtclientname')
								{
									
									document.frmAddReceipt.elements("txtclientname").value = x_clientname;
									document.frmAddReceipt.elements("txtClientCode").value = myArray[5];
									document.frmAddReceipt.elements("txtCdsNo").value = myArray[9]; 
									
									clientcobo[0].Credit = myArray[0];
									clientcobo[0].CurrentBal = myArray[1];
									clientcobo[0].Agent = myArray[2];
									clientcobo[0].Owner = myArray[3];
									clientcobo[0].AgentID = myArray[4];
									clientcobo[0].SearchCode = myArray[5];
									clientcobo[0].OrderContact = myArray[6];
									clientcobo[0].SearchText = myArray[7];
									clientcobo[0].OwnerID = myArray[8];
									clientcobo[0].SearchCDS = myArray[9];
									clientcobo[0].IsCustodian = myArray[10];
									
									clientcobo[0].text = myArray[7];
									clientcobo[0].value = myArray[5];

								}
								else
								{
									var myArrayx;
									var myArrayz;
									
									//alert(returnStr);
									myArrayx = returnStr.split("|");
									myArrayxsize = myArrayx.length - 1;
									
									//alert(myArrayxsize);

									for (i=myArrayxsize; i>=0; i--)
									{
										
										myArrayz = myArrayx[i].split("<->");

										//alert(myArrayz)
										
										document.frmAddReceipt.elements("txtClientCode").value = '';
										document.frmAddReceipt.elements("txtcdsno").value = '';
										document.frmAddReceipt.elements("txtclientname").value = '';
										document.frmAddReceipt.elements("txtClientCode").value = myArray[5];
										document.frmAddReceipt.elements("txtCdsNo").value = myArray[9]; 
										document.frmAddReceipt.elements("txtClientname").value = myArray[7];
										
										document.getElementById("cboAccount").options[i] = new Option(myArrayz[7],myArrayz[5],myArrayz[6],myArrayz[10],myArrayz[4],myArrayz[8],myArrayz[3],myArrayz[3],myArrayz[0],myArrayz[1],myArrayz[5],myArrayz[7],myArrayz[9]);
																				
									}
									
								}
							}
					}
				 
				 xmlhttp.setRequestHeader('Accept','message/x-jl-formresult');
				 xmlhttp.send();
		}

	function evaluateEntity(Val, Entity)
	{
		var clientcobo = document.getElementById("cboAccount");		 
      //Enable Printing if entity is Broker or Client
	  if (Val == 1 || Val == 3 || Val == 5){
	    document.getElementById("cmdPrint").style.display = "";
	  }
	  else {
	    document.getElementById("cmdPrint").style.display = "none";
	  }
	  
	  if (Val != 1) 
	  {
			document.frmAddReceipt.elements("txtClientCode").disabled  = true;
			document.frmAddReceipt.elements("txtcdsno").disabled  = true;
			document.frmAddReceipt.elements("txtclientname").disabled  = true;
	  }
	  else
	  {
	  		document.frmAddReceipt.elements("txtClientCode").disabled  = false;
			document.frmAddReceipt.elements("txtcdsno").disabled  = false;
			document.frmAddReceipt.elements("txtclientname").disabled  = false;
	  }
	  
	  if (Val == 1) 
	  {
	 		clientcobo.length = 1;
			clientcobo[0].text = 'Load Account';
			clientcobo[0].value = '';
			
			document.frmAddReceipt.elements("txtClientCode").value = 'Code';
			document.frmAddReceipt.elements("txtcdsno").value = 'CDS No.';
			document.frmAddReceipt.elements("txtclientname").value = 'Client Name';
			
			return;
	  }
	  //Furher Processing
	  FetchAccounts(Entity)
	}

	function FetchAccounts(theList)
		{
			var i = 0;
			var entity = theList.value;
			var toList = document.frmAddReceipt.cboAccount;
			
			currentEntityType = entity;
			frm = document.frmAddReceipt;				
			xmlhttp = createXMLHTTPObj();
			
			url="GetList.asp?ID="+entity+"&action=GetAccountList";
			xmlhttp.open("GET",url,true);
			xmlhttp.onreadystatechange=function() {
				if (xmlhttp.readyState==4) {
				returnStr = xmlhttp.responseText;
				returnStr = getBodyHTML(returnStr);
				
				var secList = "<select name = '" + toList.name + "' id = '" + toList.id + "' size='1' ";
				secList += "onChange='updatefields(this)'>"; 
				/*secList += "onChange='setFilter(this," + currentEntityType + ");' " ;
				secList += "onKeypress='return (dodefaultaction()==\"\"); ' "  ;
				secList += "onKeydown='return (dodefaultaction()==\"\");' " ; 
				secList += "onKeyup='return (change(" + toList.name + "));' " ; 
				secList += "onfocus='txtval = \"\";inputIsItemCode = 1;' "  ;
				secList += "onblur='txtval = \"\";inputIsItemCode = 1;'>" ;*/
				secList += returnStr ;
				secList += "</select>";
				
				toList.outerHTML = secList;															
				}
				}
			xmlhttp.setRequestHeader('Accept','message/x-jl-formresult');
			xmlhttp.send(); 
			
			/*var fra = document.getElementById('fraInnerSelects');		
			var pageTo = 'inner_select_vouchers.asp';  
			fra.src = pageTo;
			fra = document.getElementById('fraInnerBrokerSelects');		
			pageTo = 'inner_select_voucherBroker.asp';
			fra.src = pageTo;*/
			
			totalContractAmt = 0;
			document.all.item("txtTotal").value = totalContractAmt;
			document.all.item("ContractsSel").value = "";	
				
			
			if (entity==3){
				document.getElementById('brokerVoucherRow').style.display = '';
				//document.getElementById('clientVoucherRow').style.display = 'none';
				//document.getElementById('txtVoucherType').value = currentEntityType;
			} 
			
			else{
				 if (entity==1){
					document.getElementById('brokerVoucherRow').style.display = 'none';
					//document.getElementById('clientVoucherRow').style.display = '';
					//document.getElementById('txtVoucherType').value = currentEntityType;
				}			
				else {
					document.getElementById('brokerVoucherRow').style.display = 'none';
					//document.getElementById('clientVoucherRow').style.display = 'none';
					//document.getElementById('txtVoucherType').value = '0';
				}
			}	
			
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
    <td width="54%"><select name = 'cboPaymentTypes' id = 'cboPaymentTypes' size="1" >
    	
<%
		Set conn = GetActiveConnection("KBroker")
        sqlStr = "SELECT * FROM [PaymentTypes]  Order By Description ASC"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF
                 if(Cint(rs.Fields("PaymentTypes_DPA_"))=4) then
			%>
			   <option Selected value = '<%=rs.Fields("PaymentTypes_DPA_")%>'><%=rs.Fields("Description")%></option>
             	<%
			else
			%>
			   <option value = '<%=rs.Fields("PaymentTypes_DPA_")%>'><%=rs.Fields("Description")%></option>
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
    <td width="54%"><select name = 'cboEntity' id = 'cboEntity' size="1" onchange='evaluateEntity(this.value,this)'>
    	
<%
		Set conn = GetActiveConnection("KBroker")
        sqlStr = "SELECT * FROM [FullEntityTypeList] Order By EntityTypeName"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF
						if rs.Fields("EntityType_DPA_").value = currentEntityType then%>
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
<input type = 'text' name ='txtClientCode' id = 'txtClientCode' size="5" value = 'Code' onBlur="txtval = this.value;LoadClient(cboAccount, this.name);"  onClick  = "ClearFields(this.name)">
	&nbsp;
	<input type = 'text' name ='txtcdsno' id = 'txtcdsno' size="10" value = 'CDS NO.' onBlur="txtval = this.value;LoadClient(cboAccount, this.name);" onClick  = "ClearFields(this.name)">

	<input type = 'text' name ='txtclientname' id = 'txtclientname' size="15" value = 'Client Name' onBlur="txtval = this.value;LoadClient(cboAccount, this.name);" onClick  = "ClearFields(this.name)">

	<select name = 'cboAccount' id = "cboAccount" size="1" onChange ="updatefields(this)">		
		<option SearchCode = "" SearchText = "" value = '' >Load Account</option>
	</select> </td>
    <td width="31%">

	</td>
  </tr>
  
		 <tr id="orderRow">
		    <td width="15%">Order</td>
		    <td width="54%">
		    <select name = 'cboOrder' id = "cboOrder" size="1">
		    <option selected value = '0'></option></select>
		    <select name = 'cboOrderBag' id = "cboOrderBag" size="1" style="display:none">
		<%
		        sqlStr = "SELECT * FROM [OrderListPlain]"
		        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		        If Not (rs.EOF Or rs.BOF) Then
		                rs.MoveFirst
		                Do Until rs.EOF%>
		                        <option value = '<%=rs.Fields("Order_DPA_")%>' ClientTag = '<%=rs.Fields("Client_DPA_")%>'><%=rs.Fields("Order_DPA_")%></option>
		                        <%rs.MoveNext
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
					<iframe id="fraInnerBrokerSelects" name="fraInnerBrokerSelects" width="400px" height="200px" src="inner_select_voucherBroker.asp"></iframe>
						<BR>
			</td>
		    <td width="31%">
				<input type = 'hidden' name ='txtVoucherType' id = 'txtVoucherType'value="0">
				<input type="hidden" name="ContractsSel" id="ContractsSel" value="">
			</td>
		  </tr>
  <tr>
    <td width="15%">Amount</td>
    <td width="54%"><input type = 'text' name ='txtTotal' id = 'txtTotal'  STYLE="text-align:right" size="20" value="" onchange="JavaScript: FormatPrice()"></td>
    <td width="31%">

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
<%
        sqlStr = "SELECT * FROM [BankAccountList] Order By AccountName"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF
                
                        
                			if(trim(rs.Fields("AccountCode"))="12500") then
                			%>
                			<option selected SearchCode = "<%=rs.Fields("AccountCode")%>" SearchText = "<%=rs.Fields("AccountName")%>" value = '<%=rs.Fields("Account_DPA_")%>'><%=rs.Fields("AccountNameEx")%></option>
		                <%
                        end if
                        %>
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
    <textarea name ='txtNarrative' id = 'txtNarrative' rows="1" cols="20" ></textarea></td>
    <td width="31%">

	</td>
  </tr>
  <tr>
	  <td width="100%" colspan=2 align="center" valign=absBottom nowrap>
		<BR>
	<b id="hide" name="Hide">
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdPrint' value=" Save & Print " onclick = "javascript: AllowedNavigation();hideButton();">
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value="Save" onclick = "javascript: AllowedNavigation(); hideButton();"></b>
    	<input type = 'button' Class=Buttons name ='cmdCancel' id = "cmdCancel" value="Cancel" onclick = "JavaScript: window.self.close();">
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		<input type = 'hidden' name ='ID' id = 'ID'>
	</td>
  </tr>
</table>

</form>
</body>

</html>