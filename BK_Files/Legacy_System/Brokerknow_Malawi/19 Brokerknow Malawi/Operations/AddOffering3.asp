<%
const UDLName = "KBroker"
const DataSource = "AddOffering"
const DataEntity = "Offering"
const DataEntityPlural = "Offerings"
const ActionFolder = "Operations"
%>
<html>
<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add Security</title>
 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>

<!--CALENDAR -->
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
 <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>

<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css"> 
<script language="JavaScript" src="CALENDAR/calendar.js"></script>

<!--END CALENDAR -->

<script language="vbscript">
	Function UpdatePayable()
		 
		 dim price
		 dim acrights
		 dim payable
		 dim credit
		 dim credittxt
		 
		 price = 0
		 acrights = 0
		 payable = 0
		 credit = 0
		 credittxt = ""
		 
		 price=document.frmMain.elements("txtprice").value
		 acrights=document.frmMain.elements("txtAlloted").value
		 credittxt = document.frmMain.elements("txtAvailableCredit").value
		 credit = replace(credittxt,",","")	 
		 		   
		   if(CCur(acrights) < 500 ) then
		   ShowMessage "Applicable Quantity Should be atleast 500"		   
		   end if
		   
		   'ShowMessage Cdbl(acrights) mod 100
		   if((Cdbl(acrights) mod 100) <> 0 ) then
		   ShowMessage "The Applied Quantity should in multiples of Hundreds"
		   end if
		   		   		   
		   if (price<>0 and acrights<>0) then		   
		   payable=price*acrights
		   end if
			
		   document.frmMain.elements("txtPayable").value=payable  
		
			IF ENABLED THEN
				if(CCur(payable)>CCur(credit)) then
				ShowMessage "You have insuficient balance to complete this transaction"
				end if
			END IF
			   
		End Function			
		
</script>

<script language="javascript">

		//Update client Balances and Credits
		function UpdateBalances()
		{
		 client = document.frmMain.elements("cboclient")
		
		 document.frmMain.elements("txtAvailableCredit").value = client[client.selectedIndex].Credit;
		 document.frmMain.elements("txtCurrentBal").value = client[client.selectedIndex].CurrentBal;		 
		}

		function  UpdateImmobilised(theChk)
		{
			if (theChk.checked)
			{
				document.frmMain.elements("txtImmobilised").value = "1";
			}
			else
			{
				document.frmMain.elements("txtImmobilised").value = "0";
			}
				
			
		}
		
		function  UpdateCanTrade(theChk)
		{
			if (theChk.checked)
			{
				document.frmMain.elements("txtCanTrade").value = "1";
			}
			else
			{
				document.frmMain.elements("txtCanTrade").value = "0";
			}
				
			
		}
		
		//This function verifies whether the payable amount is less or equal to amount to pay.
		function VerifyAmount(txtamount)
		{
		var payable;
		var paidamount;

		payable = document.frmMain.elements("txtpayable").value;
		paidamount = txtamount.value;

			if(payable>paidamount)
			{
			alert('The amount you have entered is less than '+ payable +' the payable amount')
			}
		}

		function UpdateAction(theaction)
		{
			if(theaction=="previous")
			{
			document.frmMain.elements("action").value='Previous';
			}
			else
			{
			document.frmMain.elements("action").value='Print';
			}

			//alert(document.frmMain.elements("action").value)
		} 
		
		//Gets the price of the offerings and updates the price text field
		function UpdatePrice(drpOfferings)
		{
		var price;
		var securityname;

		price = drpOfferings[drpOfferings.selectedIndex].SearchPrice; 
		
		securityname = drpOfferings[drpOfferings.selectedIndex].text

		document.frmMain.elements("txtprice").value = price ;
		document.frmMain.elements("securityname").value = securityname ;
		}
		
</script>
</head>

<body Class="Dialog">

<!--#include file="../libroutines.asp"-->
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>

<%
	 
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
   Dim guidStr 
   Dim guid 
   dim first
   'first=0
   'WriteDialogRelocateScript "AddOffering.asp" 	
	
	action = Request.Form("action")
	first=Request.QueryString("first")
	
	'Response.Write(action)
	'Response.End 
	
	UserId=Session("UserID")

	Dim clientName
					Dim narrative
					Dim offering
					Dim bank
					Dim price
					Dim PDate
					Dim AlRights
					Dim AcRights
					Dim ChkNo
                    Dim ChkAmt
					Dim Renouncee 'PaymentTypes
			        dim idno
			        dim entity
			        dim AvailableCredits
			        
			        bank=1
			        
			        entity=1
					palno = Request.Form("txtpalno")
			        clientName = Request.Form ("cboClient")
					price = Request.Form("txtprice")
					offering = Request.Form("cboOfferings")
					narrative = Request.Form("txtNarrative")
					AlRights = Request.Form("txtAlloted")
					AcRights = Request.Form("txtAccepted") 
					bank = Request.Form("cboBank")
					Renouncee = Request.Form("txtRenouncee")
					PDate = Trim(Request.Form("txtDate"))					
					idno=Request.Form("txtidno")
					ChkNo=Request.Form("txtChkNo")
					ChkAmt=Request.Form("txtChkAmt")					
					payable = Request.Form("txtpayable")
					PaymentTypes = Request.Form("cboPaymentTypes")
					securityname = Request.Form("securityname")
	if(PDate="") then
	PDate=Date
	End if
	
	if(Cint(first)=1) then	
	action=Request.QueryString("action")	
	palno=Request.QueryString("palno")			
	clientName = Request.QueryString ("clientName")
	price = Request.QueryString("price")
	offering = Request.QueryString("offering")
	narrative = Request.QueryString("narrative")
	AlRights = Request.QueryString("AlRights")
	AcRights = Request.QueryString("AcRights") 
	bank = Request.QueryString("bank")
	Renouncee = Request.QueryString("Renouncee")
	PDate = Request.QueryString("pdate")					
	idno=Request.QueryString("idno")
	ChkNo=Request.QueryString("ChkNo")
	ChkAmt=Request.QueryString("ChkAmt")
	payable = Request.Querystring("payable")
	PaymentTypes = Request.Querystring("PaymentTypes")
	securityname = Request.Querystring("securityname")
	end if
	
	if(action="Next") then
	if(Cint(first)<>1) then
	WriteDialogRelocateScript "AddOffering3.asp?action=" & action & "&first=1&palno=" & palno &"&clientName=" & clientName & _	
	                          "&price=" & price & "&offering=" & offering &"&narrative=" & narrative & "&AlRights=" & AlRights & _
	                          "&AcRights=" & AcRights & "&bank=" & bank & "&Renouncee=" & Renouncee & "&pdate=" & PDate & "&idno=" & idno & _
	                          "&ChkNo=" & ChkNo & "&ChkAmt=" & ChkAmt & _
							  "&payable=" & payable & _
							  "&PaymentTypes=" & PaymentTypes & _
							  "&securityname=" & securityname
	end if
	end if
	
	if(action="Previous") then
	if(Cint(first)<>1) then
	WriteDialogRelocateScript "AddOffering3.asp?first=1&palno=" & palno &"&clientName=" & clientName & _	
	                          "&price=" & price & "&offering=" & offering &"&narrative=" & narrative & "&AlRights=" & AlRights & _
	                          "&AcRights=" & AcRights & "&bank=" & bank & "&Renouncee=" & Renouncee & "&pdate=" & PDate & "&idno=" & idno & _
	                          "&ChkNo=" & ChkNo & "&ChkAmt=" & ChkAmt & _
							  "&payable=" & payable & _
							  "&PaymentTypes=" & PaymentTypes & _
							  "&securityname=" & securityname
	end if
	
	end if
	
	action = ucase(action)		
	
	if(trim(price)<>"" and trim(AlRights)<>"") then
	payable=price*AlRights
	else
	payable=0
	end if
	
	select case action
	case "EXECUTE" 
	   Dim palno
					entity=1
					palno = Request.Form("txtpalno")
			        clientName = Request.Form ("cboClient")
					price = Request.Form("txtprice")
					offering = Request.Form("cboOfferings")
					narrative = Request.Form("txtNarrative")
					AlRights = Request.Form("txtAlloted")
					AcRights = Request.Form("txtAccepted") 
					bank = Request.Form("cboBank")
					Renouncee = Request.Form("txtRenouncee")
					PDate = Trim(Request.Form("txtDate"))					
					idno=Request.Form("txtidno")
					ChkNo=Request.Form("txtChkNo")
					ChkAmt=Request.Form("txtChkAmt")					
					PaymentTypes = Request.Form("cboPaymentTypes")
					AvailableCredit = Request.Form("txtAvailableCredit")
					 
					PDate = PDate & " " & Time			        
			             
                    'validate Entity
					 If Trim(palno) = "" Then%>
					         <script language = 'vbscript'>
					         		ShowMessage "Please type the PAL No."
					         		
					         </script>
					         <% response.end
					 End If
					 
					 If Trim(clientName) = "" Then%>
					         <script language = 'vbscript'>
					         		ShowMessage "Please Specify the client Name"
					         		
					         </script>
					         <% response.end
					 End If	 
					 
					 'validate Account
					 If Trim(offering) = "" Then%>
					         <script language = 'vbscript'>
					         		ShowMessage "Please select the Offering"					         		
					         </script>
					         <% response.end
					 End If
					 					 
					'Check whether price has been entered
					If Trim(price) = "" Then%>
					    <script language = 'vbscript'>
					         	ShowMessage "Please Enter the Price "
					         	
					    </script>
					    <% response.end
					End If
					
					'ensure price is numeric
					If (price <> "") And (Not IsNumeric(price)) Then%>
					    <script language = 'vbscript'>
							ShowMessage "price Amount must be numeric"
							
					    </script>
					    <% response.end
					End If
					
					'validate Rights
					 If Len(AlRights) = "" Then%>
					         <script language = 'vbscript'>
					         ShowMessage "Please enter the Alloted Rights"
					         
					         </script>
					         <% response.end
					 End If
										
					'ensure Quantity is numeric
					If (Alrights <> "") And (Not IsNumeric(Alrights)) Then%>
					    <script language = 'vbscript'>
							ShowMessage "Alloted Rights must be numeric"
							
					    </script>
					    <% response.end
					End If					
					
					'The quantity should be atleast 500
					if(CCur(Alrights) < 500 ) then
					%>
					<script language="VBScript">
					ShowMessage "Applicable Quantity Should be atleast 500"		   
					</script>
					<%
					Response.End 
					end if
		   
					'The quantity should 
					if((Cdbl(Alrights) mod 100) <> 0 ) then
					%>
					<script language="VBScript">
					ShowMessage "The Applied Quantity should in multiples of Hundreds"
					</script>
					<%
					Response.End 
					end if
		   
					payable = price*Alrights
					
					IF ENABLED THEN
						if(CCur(payable)>CCur(AvailableCredit)) then
						%>
						<script language="javascript">
						alert('This amount Should be atleast equal to the available credit which is <%=AvailableCredit%>');
						</script>
						<%
						Response.end
						end if
					END IF
					
					'save data
					Dim Paymentid					
					Dim PaymentTypes
					
					'PaymentTypes=4
					
					Paymentid = "NULL"
					
					Set conn = GetActiveConnection("KBroker")
					conn.BeginTrans		
							'Save payment
																
							sqlStr = "INSERT INTO [Offerings] ( Offering_DPA_,PAL_No,Client_DPA_, " & _
									"Offering,Offering_Price,Alloted_Rights,Receipt,ChangedBy) " & _
									" SELECT " & " " & "iif(isnull(max([Offering_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Offerings'),max([Offering_DPA_]) + 1)" & " " & " as Offering_DPA_" & _
									"," & "'" & palno & "'" & " as PAL_No" & _
									"," & "" & ClientName & "" & " as Client_DPA_" & _
									"," & " " & offering & " " & " as Offering" & _						
                                    "," & " " & Price & "" & " as Offering_Price" & _
									"," & " " & AlRights & " " & " as Alloted_Rights" & _				
									"," & " " & Paymentid & " " & " as Receipt" & _
									"," & " " & UserId & " " & " as ChangedBy" & _
									" FROM [Offerings]"			
													

							sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))							
														
							'Response.Write sqlStr
							'Response.End
																		                                                     
							conn.Execute sqlStr												
														
							'conn.execute ("Exec ClientTotalProcedure " & ClientName)							
							'conn.execute ("Exec ClientBalanceProcedure " & ClientName)
											
					conn.CommitTrans
					conn.Close
					Set conn = Nothing
					%>
					<script language="vbscript">
						window.location.href="AddOffering3.asp"
					</script>
					<%
					Response.End	
   	case "PRINT"
	
	'Dim palno
					entity=1
					palno = Request.Form("txtpalno")
			        clientName = Request.Form ("cboClient")
					price = Request.Form("txtprice")
					offering = Request.Form("cboOfferings")
					narrative = Request.Form("txtNarrative")
					AlRights = Request.Form("txtAlloted")
					AcRights = Request.Form("txtAccepted") 
					bank = Request.Form("cboBank")
					Renouncee = Request.Form("txtRenouncee")
					PDate = Trim(Request.Form("txtDate"))					
					idno=Request.Form("txtidno")
					ChkNo=Request.Form("txtChkNo")
					ChkAmt=Request.Form("txtChkAmt")					
					PaymentTypes = Request.Form("cboPaymentTypes")

					PDate = PDate & " " & Time			        
			             
                    'validate Entity
					 If Trim(palno) = "" Then%>
					         <script language = 'vbscript'>
					         		ShowMessage "Please type the PAL No."
					         		
					         </script>
					         <% response.end
					 End If
					 
					 If Trim(clientName) = "" Then%>
					         <script language = 'vbscript'>
					         		ShowMessage "Please Specify the client Name"
					         		
					         </script>
					         <% response.end
					 End If	 
					 
					 'validate Account
					 If Trim(offering) = "" Then%>
					         <script language = 'vbscript'>
					         		ShowMessage "Please select the Offering"					         		
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
					
					'Check whether price has been entered
					If Trim(price) = "" Then%>
					    <script language = 'vbscript'>
					         	ShowMessage "Please Enter the Price "
					         	
					    </script>
					    <% response.end
					End If
					
					'ensure price is numeric
					If (price <> "") And (Not IsNumeric(price)) Then%>
					    <script language = 'vbscript'>
							ShowMessage "price Amount must be numeric"
							
					    </script>
					    <% response.end
					End If
					
					'validate Rights
					 If Len(AlRights) = "" Then%>
					         <script language = 'vbscript'>
					         ShowMessage "Please enter the Alloted Rights"
					         
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
					
					'ensure Quantity is numeric
					If (Alrights <> "") And (Not IsNumeric(Alrights)) Then%>
					    <script language = 'vbscript'>
							ShowMessage "Alloted Rights must be numeric"
							
					    </script>
					    <% response.end
					End If					
					
					payable = price*Alrights
					
					if(CCur(payable)>CCur(ChkAmt)) then
					%>
					<script language="javascript">
					alert('This amount Should be atleast equal to the payable amount which is <%=payable%>');
					</script>
					<%
					Response.end
					end if
					
					'save data
					'Dim Paymentid					
					'Dim PaymentTypes
					
					'PaymentTypes=4
					
					Set conn = GetActiveConnection("KBroker")
					conn.BeginTrans		
							'Save payment
												
									set guid = server.createobject("NDUtils.CGUID")
									guidStr =guid.GenerateGUID
																							
							sqlStr = "INSERT INTO [Payment] ( PaymentAmount, PaymentPDate, BankAccount_DPA_, " & _
									"Payment_EIT_,PayType_DPA_, EntityType_DPA_, Payment_DPA_, PaymentReference,PaymentNarrative, Entity_DPA_,PaymentTypes_DPA_,ChangedBy,PaymentReceiptNo) " & _
									"SELECT " & " " & ccur(ChkAmt) & " " & " as PaymentAmount," & "#" & PDate & "#" & " as PaymentPDate" & _
									"," & " " & bank & " " & " as BankAccount_DPA_" & _
									"," & "'" & guidStr & "'" & " as Payment_EIT_" & "," & "1 as PayType_DPA_" & _
									"," & " " & entity & " " & " as EntityType_DPA_" & _
									"," & " " & "iif(isnull(max([Payment_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Payment'),max([Payment_DPA_]) + 1)" & " " & " as Payment_DPA_" & _
                                    "," & "'" & ChkNo & "'" & " as PaymentReference" & _									
									"," & "'" & narrative & "'" & " as PaymentNarrative" & _
									"," & " " & ClientName & " " & " as Entity_DPA_" & _
									"," & " " & PaymentTypes & " " & " as PaymentTypes_DPA_" & _
									"," & " " & UserId & " " & " as ChangedBy" & _										
									"," & " " & "(SELECT IIf(IsNull(Max([PaymentReceiptNo])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Receipt'),max([PaymentReceiptNo]) + 1) AS PaymentReceiptNo FROM Payment WHERE PayType_DPA_ = 1)" & " " & " as PaymentReceiptNo" & _
									" FROM [Payment]"
								
											'conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr)).
											
													
											sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))									        
											
											conn.Execute sqlStr
																						
											'conn.CommitTrans											
											
											'obtain Payment id
											sqlStr = "SELECT [Payment_DPA_] FROM [Payment] WHERE [Payment_EIT_] = " & "'" & guidStr & "'"
				
											Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
											
											If (rs.EOF Or rs.BOF) Then%>
											         	<script language = 'vbscript'>
											         			ShowMessage "A serious error has been encountered while saving the data. Try saving again"
											         			
											         	</script>
											         	<% response.end
											End If
											Paymentid = rs.fields("payment_DPA_").value
											
											'save to offering file										
											
				
							sqlStr = "INSERT INTO [Offerings] ( Offering_DPA_,PAL_No,Client_DPA_, " & _
									"Offering,Offering_Price,Alloted_Rights,Receipt,ChangedBy) " & _
									" SELECT " & " " & "iif(isnull(max([Offering_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Offerings'),max([Offering_DPA_]) + 1)" & " " & " as Offering_DPA_" & _
									"," & "'" & palno & "'" & " as PAL_No" & _
									"," & "" & ClientName & "" & " as Client_DPA_" & _
									"," & " " & offering & " " & " as Offering" & _						
                                    "," & " " & Price & "" & " as Offering_Price" & _
									"," & " " & AlRights & " " & " as Alloted_Rights" & _				
									"," & " " & Paymentid & " " & " as Receipt" & _
									"," & " " & UserId & " " & " as ChangedBy" & _
									" FROM [Offerings]"			
													

							sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))							
							                                                     
							conn.Execute sqlStr					
							
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
										
						%>
						<SCRIPT LANGUAGE="JAVASCRIPT">
							window.parent.parent.frames['maininfo'].location.reload();
						</SCRIPT>
						<%
						WriteDialogRelocateScript "ReceiptForm.asp?ID=" & rs.Fields("Payment_DPA_")
					
					conn.Close
					Set conn = Nothing
					WritefraEnabledDialogCloseScript
					Response.End
	case "NEXT"
   'Dim conn 
   'Dim sqlStr
   'Dim rs
   'Dim guidStr 
   'Dim guid 
   
   'if(Cint(first)<>1) then
   'WriteDialogRelocateScript "AddOffering.asp?action=Next"  
   'first=1
   'else   
   
   if(PaymentTypes="") then
	PaymentTypes=4
   end if

   Set conn = GetActiveConnection("KBroker")   
   	%>
   	<SCRIPT language="JavaScript">
	var cal=new ctlSpiffyCalendarBox("cal", "frm<%=DataSource%>", "txtDate","cmdDate","<%= FormatDate(PDate) %>",1);
	</SCRIPT>   	
   	<form name = 'frmAddOffering' method = 'post' id="frmMain" action = "AddOffering3.asp" >
<table border="0" width="100%" cellpadding=2 cellspacing=2>
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
                 if(Cint(rs.Fields("PaymentTypes_DPA_"))=PaymentTypes) then
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
	</td>
  </tr>
 
              <tr>
                <td width="35%">Date Of Payment</td>
				<td width="65%"><SCRIPT language="JavaScript">cal.writeControl();</SCRIPT></td>   
  
              </tr>              
              <tr>
                <td width="35%">Amount</td>
				<td width="65%"><input type = 'text' name ='txtChkAmt' id = 'txtChkAmt' size="20" value='<%=ChkAmt%>' onblur='VerifyAmount(this)'></td>   
              </tr>
              <tr>              
     <tr>
    <td width="15%">Bank</td>
    <td width="54%"><select name = 'cboBank' id = 'cboBank' size="1" 
			onKeypress="return (dodefaultaction()==''); " 
			onKeydown="return (dodefaultaction()==''); " 
			onKeyup="return (change(cboBank));" 
			onfocus="txtval = '';inputIsItemCode = 1;" 
			onblur="txtval = '';inputIsItemCode = 1;">    	
<%
        sqlStr = "SELECT distinct * FROM [BankAccountList] Order By AccountName"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
				
		
		'Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))	

		'set rsEdit=nothing

        If Not (rs.EOF Or rs.BOF) Then				
                rs.MoveFirst
                Do Until rs.EOF  
						
							if(rs.Fields("Account_DPA_")=10036) then
							%>
							<option Selected SearchCode = "<%=rs.Fields("AccountCode")%>" SearchText = "<%=rs.Fields("AccountName")%>" value = '<%=rs.Fields("Account_DPA_")%>'><%=rs.Fields("AccountNameEx")%></option>
							<%
							else
							%>
							<option SearchCode = "<%=rs.Fields("AccountCode")%>" SearchText = "<%=rs.Fields("AccountName")%>" value = '<%=rs.Fields("Account_DPA_")%>'><%=rs.Fields("AccountNameEx")%></option>
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
    <td width="15%">Reference</td>
    <td width="54%"><input type = 'text' name ='txtchkNo' id = 'txtchkNo' size="20" value='<%=ChkNo%>'></td>    
  </tr>
  <%
  if(trim(narrative)="") then
  narrative =Alrights & " " & securityname & "@" & price
  end if
  %>

 <tr>
    <td width="35%">Narrative</td>
    <td><textarea name ='txtNarrative' id = 'txtNarrative' cols="40" rows="5"><%=narrative%></textarea></td>   	
  </tr>      
  <tr>
     <td width="100%" colspan="2" align=right>
		<BR>
		<BR>
		<BR>
		<input type = 'Submit' Class=Buttons name ='cmdPrevious' id = 'cmdCancel' value=" Previous " OnClick="JavaScript: UpdateAction('previous');">
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdPrint' value=" Save & Print " OnClick="JavaScript: UpdateAction('print');">
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save ">		
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();">		
		&nbsp;&nbsp;
		
		<input type = 'hidden' name ='action' id = 'action' value="Execute">&nbsp;
		<input type = 'hidden' name ='txtpalno' id = 'txtpalno' value=<%=palno%>>
		<input type = 'hidden' name ='cboClient' id = 'cboClient' value='<%=clientName%>'>
		<input type = 'hidden' name ='cboOfferings' id = 'cboOfferings' value='<%=offering%>'>
		<input type = 'hidden' name ='txtidno' id = 'txtidno' value='<%=idno%>'>
		<input type = 'hidden' name ='txtprice' id = 'txtprice' value='<%=price%>'>
		<input type = 'hidden' name ='txtAlloted' id = 'txtAlloted' value='<%=Alrights%>'>		
		<input type = 'hidden' name ='txtAccepted' id = 'txtAccepted' value='<%=AcRights%>'>
		<input type = 'hidden' name ='txtpayable' id = 'txtpayable' value='<%=payable%>'>
		<input type = 'hidden' name ='securityname' id = 'securityname' value='<%=securityname%>'>
      </td>
  </tr>
</table>
</form>


</body>

</html>

   	<%   	
   	Case ""   	
%>

<form name = 'frmAddSecurity' method = 'post' id="frmMain" action = "AddOffering3.asp" >
	<table border="0" width="100%" cellpadding=2 cellspacing=2>
		<tr>
		   <td width="20%">PAL NO</td>
		   <td width="80%">
				<input type="text" name="txtpalno" id="txtName" size="25" value='<%=palno%>'>
			</td>
		</tr>
        <tr>
           <td>Client</td>
		   <td>&nbsp;
				<input type = 'text' name ='txtClientCode' id = 'txtClientCode' size="10" onBlur="txtval = this.value; selectItem(cboClient);UpdateCodes(true,cboClient,txtCdsNo);UpdateBalances();">&nbsp;
				<input type = 'text' name ='txtCdsNo' id = 'txtCdsNo' size="16" onBlur="txtval = this.value; selectItems(cboClient);UpdateCode(true,cboClient,txtClientCode);UpdateBalances();">&nbsp;
					<select name = 'cboClient' id = 'cbocboClient' size="1" 
					onKeypress="return (dodefaultaction()==''); " 
					onKeydown="return (dodefaultaction()==''); " 
					onKeyup="return (UpdateCode(change(cboClient,0),cboClient,txtClientCode));UpdateBalances();" 
					onChange="UpdateCode(true,cboClient,txtClientCode);UpdateCodes(true,cboClient,txtCdsNo);UpdateBalances();"
					onfocus="txtval = '';inputIsItemCode = 1;" 
					onblur="txtval = '';inputIsItemCode = 1;">
    	
					<%
		
					Dim thisClientName
					Dim NameClient

					Set conn = GetActiveConnection("KBroker")
				    sqlStr = "SELECT LTRIM(RTRIM(ClientName)) as ClientName,Client_DPA_,ClientCDSNo FROM [client]  where Deleted=0 Order By ClientName ASC"
				    
				    sqlStr = " SELECT TOP 100 PERCENT ISNULL(dbo.ClientBalances.CurrentBal, 0) + ISNULL(dbo.Client.CreditLimit, 0) - ISNULL(dbo.ClientTotal.Total, 0) AS AvailableCredit,   " & _
							"         ISNULL(dbo.ClientBalances.CurrentBal, 0) AS CurrentBal," & _
							"         dbo.Client.Client_DPA_, dbo.Client.ClientName," & _
							"         dbo.Client.ClientCDSNo, dbo.ClientTotal.Total  " & _
							"  FROM dbo.Client LEFT OUTER JOIN  " & _
							"       dbo.ClientTotal ON dbo.Client.Client_DPA_ = dbo.ClientTotal.Client_DPA_ LEFT OUTER JOIN  " & _
							"       dbo.ClientBalances ON dbo.Client.Client_DPA_ = dbo.ClientBalances.client_DPA_ " & _							
							"  WHERE     (dbo.Client.Deleted = 0)" & _
							"      ORDER BY LTRIM(RTRIM(dbo.Client.ClientName)) "
					
					'and Client.Client_DPA_ not in (Select Client_DPA_ from offerings where deleted=0)
					
				    Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
				    If Not (rs.EOF Or rs.BOF) Then
				            rs.MoveFirst
				            Do Until rs.EOF

							thisClientName=rs.Fields("ClientName")
							
							NameClient=Mid(thisClientName,1,30)
							
							if trim(rs.Fields("Client_DPA_")) = trim(clientName) Then%>
				            			<option Credit="<%=FormatNumEx(rs.Fields("AvailableCredit"),2)%>" CurrentBal="<%=FormatNumEx(rs.Fields("CurrentBal"),2)%>" SearchCode = "<%=rs.Fields("Client_DPA_")%>" SearchText = "<%=rs.Fields("ClientName")%>" SearchCds = "<%=rs.Fields("ClientCDSNo")%>"  selected value = '<%=rs.Fields("Client_DPA_")%>'><%=NameClient%></option>
				            		<%else%>                   						
						   			<option Credit="<%=FormatNumEx(rs.Fields("AvailableCredit"),2)%>" CurrentBal="<%=FormatNumEx(rs.Fields("CurrentBal"),2)%>" SearchCode = "<%=rs.Fields("Client_DPA_")%>" SearchText = "<%=rs.Fields("ClientName")%>" SearchCds = "<%=rs.Fields("ClientCDSNo")%>" value = '<%=rs.Fields("Client_DPA_")%>'><%=NameClient%></option>
				         <%  end if         
				                rs.MoveNext                
				            Loop
				    End If
				%>
				</select>
			</td>

     </tr>
     <tr>
		<td>&nbsp;</td>
		<td>
			<table>
				<tr>
					<td align="center">Client Balance</td>
					<td align="center">Available Credit</td>
				</tr>     
				<tr>
					<td><input type = 'text' name ='txtCurrentBal' id = 'txtCurrentBal' readonly class="readonlyex" size="15"></td>
					<td><input type = 'text' name ='txtAvailableCredit' id = 'txtAvailableCredit' readonly class="readonlyex" size="15"></td>
				</tr>
			</table>
		</td>	         
     <tr>
		<td nowrap>Offering Name</td>
		<td>
			<select name = 'cboOfferings' id = 'cboOfferings' size="1" 
				onKeyup="UpdatePrice(this);" onChange="UpdatePrice(this);" onblur="UpdatePrice(this);">
    	
				<% 
				Set conn = GetActiveConnection("KBroker")
				sqlStr = "SELECT * FROM [SecurityListOfferings] Order By SecurityName ASC"
				Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
				If Not (rs.EOF Or rs.BOF) Then
				        rs.MoveFirst
				        Do Until rs.EOF
						if trim(rs.Fields("Security_DPA_")) = trim(offering) Then%>
				        			<option selected SearchPrice = "<%=rs.Fields("SecurityMktPrice")%>" value = '<%=rs.Fields("Security_DPA_")%>'><%=rs.Fields("SecurityName")%></option>
				        		<%else%>                   						
					   				<option SearchPrice = "<%=rs.Fields("SecurityMktPrice")%>" value = '<%=rs.Fields("Security_DPA_")%>'><%=rs.Fields("SecurityName")%></option>
				     <%  end if 
							
							if(price="") then
							price = rs("SecurityMktPrice")
							end if					

				            rs.MoveNext
				        Loop
				End If
					%>

			</select>
	</td>     
              </tr>              
                            
     <tr>
	<td>Price</td>
	<td><input type = 'text' name ='txtprice' id = 'txtprice' size="20" value='<%=price%>' onchange='UpdatePayable()' readonly class="readonlyex"></td>
	</td>   	
    </tr>          
 <tr>
    <td nowrap>Quantity applied</td>
	<td><input type = 'text' name ='txtAlloted' id = 'txtAlloted' size="20" value='<%=AlRights%>' onchange='UpdatePayable()'></td>
	</td>
  </tr>     
 
  <tr>
  <td>Payable</td>
  <td><input type = 'text' name ='txtpayable' id = 'txtpayable' size="25" value='<%=payable%>' readonly class="readonlyex"></td>
  </tr>
 
  <tr>
     <td width="100%" colspan="2" align=right>
		<BR>
		<BR>
		<BR>
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save ">
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();">
		&nbsp;&nbsp;
		<%
		if(trim(securityname) ="") then
			securityname = "Kenya Electricity Generating Company"
		end if
		%>
		<input type = 'hidden' name ='action' id = 'action' value="Execute">&nbsp;
		<input type = 'hidden' name ='txtRenouncee' id = 'txtRenouncee' value='<%=Renouncee%>'>
		<input type = 'hidden' name ='txtDate' id = 'txtDate' value='<%=PDate%>'>
		<input type = 'hidden' name ='txtchkNo' id = 'txtchkNo' value='<%=ChkNo%>'>
		<input type = 'hidden' name ='txtChkAmt' id = 'txtChkAmt' value='<%=ChkAmt%>'>
		<input type = 'hidden' name ='cboBank' id = 'cboBank' value='<%=bank%>'>
		<input type = 'hidden' name ='txtNarrative' id = 'txtNarrative' value='<%=narrative%>'>			<input type = 'hidden' name ='cboPaymentTypes' id = 'cboPaymentTypes' value='<%=PaymentTypes%>'>
		<input type = 'hidden' name ='securityname' id = 'securityname' value='<%=securityname%>'>
      </td>
  </tr>
</table>
</form>


</body>

</html>
<%
end select
%>