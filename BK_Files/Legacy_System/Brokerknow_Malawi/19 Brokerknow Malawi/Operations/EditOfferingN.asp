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
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT>
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
		
			if(CCur(payable)>CCur(credit)) then
			ShowMessage "You have insuficient balance to complete this transaction"
			end if
			   
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

		price =drpOfferings[drpOfferings.selectedIndex].SearchPrice; 
		//alert(price);
		
		securityname = drpOfferings[drpOfferings.selectedIndex].text; 
		
		//alert(securityname);

		document.frmMain.elements("txtprice").value = price ;

		document.frmMain.elements("securityname").value = securityname ;
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
   dim RsEdit
   Dim guidStr 
   Dim guid 
   dim first
   dim securityname

   Set conn = GetActiveConnection("KBroker")
	
	action = Request.Form("action")
	first=Request.QueryString("first")		
		
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
			        dim idp
			        
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
					ID=Request.Form("txtid")	
					idp=Request.Form("idp")
					payable = Request.Form("txtpayable")
					PaymentTypes = Request.Form("cboPaymentTypes")
					securityname = Request.Form("securityname")

	if(trim(action)="" and first<>1) then
	ID = Request("ID")
	
	'Do not allow editing it has already been batched
    
	sqlstr = "SELECT  Offering_DPA_, Downloaded " & _
			 " FROM   dbo.Offerings " & _
			 " WHERE  (NOT (Batch_No IS NULL)) AND (Offering_DPA_ = " & ID & ")"
     
	 set rs = conn.execute(sqlstr)

	 if not (rs.bof or rs.eof) then
      %>
		<script language="javascript">
		alert('The record cannot be edited! It has already been batched.');
		</script>
		<%
		set conn = nothing
		Response.end
	 end if


	sqlStr = "SELECT * FROM OfferingsList WHERE Offering_DPA_=" & ID
    
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		window.self.ShowMessage "The selected <%=DataEntity%> cannot be retrieved for editing"
                		
                </script>
                <% response.end
        else
        palno = rs("PAL_No")
		clientName = rs("Client_DPA_")
		price = rs("Offering_Price")
		offering = rs("Offering")
		'narrative = rs("PaymentNarrative")
		AlRights = rs("Alloted_Rights")
		AcRights = rs("Accepted_Rights") 
		'bank = rs("BankAccount_DPA_")

		'Response.write(bank)
		'Response.end

		Renouncee = rs("Renouncee")
		'PDate = rs("PaymentPDate")					
		'idno=rs("ID_No")
		'ChkNo=rs("PaymentReference")
		'ChkAmt=rs("PaymentAmount")
		'idp=rs("Payment_DPA_")					
        End If          
	end if
	
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
		ID=Request.QueryString("OfferingId")
		idp=Request.QueryString("idp")
		payable = Request.Querystring("payable")
		PaymentTypes = Request.Querystring("PaymentTypes")
		securityname = Request.Querystring("securityname")
	end if
	
	if(action="Next") then
	if(Cint(first)<>1) then	
	WriteDialogRelocateScript "EditOfferingN.asp?action=" & action & "&first=1&palno=" & palno &"&clientName=" & clientName & _	
	                          "&price=" & price & "&offering=" & offering &"&narrative=" & narrative & "&AlRights=" & AlRights & _
	                          "&AcRights=" & AcRights & "&bank=" & bank & "&Renouncee=" & Renouncee & "&pdate=" & PDate & "&idno=" & idno & _
	                          "&ChkNo=" & ChkNo & "&ChkAmt=" & ChkAmt & "&OfferingId=" & ID & "&idp=" & idp & _
							  "&payable=" & payable & _
							  "&PaymentTypes=" & PaymentTypes & _
							  "&securityname=" & securityname
	end if
	end if
	
	if(action="Previous") then
	if(Cint(first)<>1) then
	WriteDialogRelocateScript "EditOfferingN.asp?first=1&palno=" & palno &"&clientName=" & clientName & _	
	                          "&price=" & price & "&offering=" & offering &"&narrative=" & narrative & "&AlRights=" & AlRights & _
	                          "&AcRights=" & AcRights & "&bank=" & bank & "&Renouncee=" & Renouncee & "&pdate=" & PDate & "&idno=" & idno & _
	                          "&ChkNo=" & ChkNo & "&ChkAmt=" & ChkAmt& "&OfferingId=" & ID & "&idp=" & idp & _
							  "&payable=" & payable & _
							  "&PaymentTypes=" & PaymentTypes & _
							  "&securityname=" & securityname
	end if
	end if
	
	UserId=Session("UserID")

	action = ucase(action)		
	
	if(trim(price)<>"" and trim(AlRights)<>"") then
	payable=price*AlRights
	else
	payable=0
	end if
	
	select case action
	case "EXECUTE" 
	Set conn = GetActiveConnection("KBroker")
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
					ID=Request.form("txtid")					 
					idp=Request.Form("idp")
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
					
					'Validate Price 					 
					If Trim(price) = "" Then%>
					    <script language = 'vbscript'>
					         	ShowMessage "Please Enter the Price "
					         	
					    </script>
					    <% response.end
					End If
					
					'ensure Amount is numeric
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
										
					'ensure Amount is numeric
					If (Alrights <> "") And (Not IsNumeric(Alrights)) Then%>
					    <script language = 'vbscript'>
							ShowMessage "Alloted Rights must be numeric"
							
					    </script>
					    <% response.end
					End If					
					
					payable = price*AlRights
					

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
					
					conn.BeginTrans										
					'save data
					sqlStr = "UPDATE Offerings SET PAL_No = '" & palno & "'" & _
									", Client_DPA_ = " & clientName & " " & _
									", Offering = " & offering & " " & _
									", ChangedBy = " & UserId & " " & _	
									", TimeChanged =GetDate()  " & _	
									", Offering_Price = " & ccur(price) & _
									", Alloted_Rights = " & AlRights & " " & _							
									", Renouncee = '" & Renouncee & "'" & _									
									" WHERE Offering_DPA_= " & ID
													
							sqlStr = SQLServerFormat(HandleQuote(sqlStr))						
												 
							conn.Execute sqlStr						
					
					conn.CommitTrans		
					
					conn.Close
					Set conn = Nothing
					WritefraEnabledDialogCloseScript
					Response.End
				
				Dim clientCode
        
				clientCode = "var validNavigate = true;" & chr(13)
				%>
				<script>
					<%=clientCode%>
				</script>
				<%
				response.End
    
	case "PRINT"
	%>
	<SCRIPT LANGUAGE="JAVASCRIPT">
		window.parent.parent.frames['maininfo'].location.reload();
	</SCRIPT>
	<%
	WriteDialogRelocateScript "ReceiptForm.asp?ID=" & idp	

   	case "NEXT"   

	
	if(PaymentTypes="") then
	PaymentTypes=4
	end if

   Set conn = GetActiveConnection("KBroker")   
   	%>
   	<SCRIPT language="JavaScript">
	var cal=new ctlSpiffyCalendarBox("cal", "frm<%=DataSource%>", "txtDate","cmdDate","<%= FormatDate(PDate) %>",1);
	</SCRIPT>   	
   	<form name = 'frmAddOffering' method = 'post' id="frmMain" action = "EditOfferingN.asp" >
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
                <td width="30%">Date Of Payment</td>
				<td width="70%"><SCRIPT language="JavaScript">cal.writeControl();</SCRIPT></td>   
  
              </tr>              
              <tr>
                <td width="30%">Amount</td>
				<td width="70%"><input type = 'text' name ='txtChkAmt' id = 'txtChkAmt' size="20" value='<%=ChkAmt%>' onblur='VerifyAmount(this)'></td>   
              </tr>                            

	 <tr>
	<td width="30%">Bank Account</td>
    <td width="70%"><select name = 'cboBank' id = 'cboBank' size="1" 
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
                        if Trim(rsEdit.Fields("Account_DPA_")) = Trim(bank) Then%>
                			<option selected SearchCode = "<%=rsEdit.Fields("AccountCode")%>" SearchText = "<%=rsEdit.Fields("AccountName")%>" value = '<%=rsEdit.Fields("Account_DPA_")%>'><%=rsEdit.Fields("AccountNameEx")%></option>
                		<%else%>
							<option SearchCode = "<%=rsEdit.Fields("AccountCode")%>" SearchText = "<%=rsEdit.Fields("AccountName")%>" value = '<%=rsEdit.Fields("Account_DPA_")%>'><%=rsEdit.Fields("AccountNameEx")%></option>
                     <%end if
						rsEdit.MoveNext
                Loop
        End If
%>

    </select></td>
    </tr>
	<tr>
    <td width="15%">Reference</td>
    <td width="54%"><input type = 'text' name ='txtchkNo' id = 'txtchkNo' size="20" value='<%=ChkNo%>'></td>    
  </tr>
  <%
  if(trim(narrative)="") then
   narrative = Alrights & " " & securityname & "@" & price 
  end if
  %>
 <tr>
    <td width="30%">Narrative</td>
    <td width="70%"><textarea name ='txtNarrative' id = 'txtNarrative'><%=narrative%></textarea></td>   	
  </tr>      
  <tr>
     <td width="100%" colspan="2" align=right>
		<BR>
		<BR>
		<BR>
		<input type = 'Submit' Class=Buttons name ='cmdPrevious' id = 'cmdCancel' value=" Previous " OnClick="JavaScript: UpdateAction('previous');">
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAddPrint' value="Print" OnClick="JavaScript: UpdateAction('print');">
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save ">		
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();">		
		&nbsp;&nbsp;
		
		<input type = 'hidden' name ='action' id = 'action' value="Execute">&nbsp;
		<input type = 'hidden' name ='txtpalno' id = 'txtpalno' value=<%=palno%>>		
		<input type = 'hidden' name ='txtid' id = 'txtid' value='<%=ID%>'>
		<input type = 'hidden' name ='IDP' id = 'IDP' value="<%=idp%>">
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
<form name = 'frmAddSecurity' method = 'post' id="frmMain" action = "EditOfferingN.asp" >
<table border="0" width="100%" cellpadding=2 cellspacing=2>
 <tr>
                <td width="30%">PAL NO</td>
                <td width="70%"><input type="text" name="txtpalno" id="txtName" size="25" value='<%=palno%>'></td>
              </tr>
              <tr>
                <td>Client</td>
    <td>
    &nbsp;<input type = 'text' name ='txtClientCode' id = 'txtClientCode' size="10" onBlur="txtval = this.value; selectItem(cboClient);UpdateCodes(true,cboClient,txtCdsNo);UpdateBalances();" value="<%=clientName%>">&nbsp;
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
							"  WHERE     (dbo.Client.Deleted = 0)  " & _
							"      ORDER BY LTRIM(RTRIM(dbo.Client.ClientName)) "

        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF

				thisClientName= trim(rs.Fields("ClientName"))
				
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
	 </tr>
              <tr>
                <td width="10%" nowrap>Offering Name</td>
    <td width="30%"><select name = 'cboOfferings' id = 'cboOfferings' size="1" onKeyup="UpdatePrice(this);" onChange="UpdatePrice(this);" onblur="UpdatePrice(this);">
    	
	<%
		Set conn = GetActiveConnection("KBroker")
        sqlStr = "SELECT * FROM [SecurityListOfferings] Order By SecurityName ASC"
        Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rsEdit.EOF Or rsEdit.BOF) Then
                rsEdit.MoveFirst
                Do Until rsEdit.EOF
				if trim(rsEdit.Fields("Security_DPA_")) = trim(offering) Then%>
                			<option selected SearchPrice = "<%=rsEdit.Fields("SecurityMktPrice")%>" value = '<%=rsEdit.Fields("Security_DPA_")%>'><%=rsEdit.Fields("SecurityName")%></option>
                		<%else%>                   						
			   				<option SearchPrice = "<%=rsEdit.Fields("SecurityMktPrice")%>" value = '<%=rsEdit.Fields("Security_DPA_")%>'><%=rsEdit.Fields("SecurityName")%></option>
             <%  end if         
                    rsEdit.MoveNext
                Loop
        End If
	%>

    </select></td>
    </td>
                          
     <tr>
	<td width="30%">Price</td>
	<td width="70%"><input type = 'text' name ='txtprice' id = 'txtprice' size="20" value='<%=price%>' onchange='UpdatePayable()' readonly class="readonlyex"></td>
	</td>   	
    </tr>          
 <tr>
    <td width="30%" nowrap>Quantity applied</td>
	<td width="70%"><input type = 'text' name ='txtAlloted' id = 'txtAlloted' size="20" value='<%=AlRights%>' onchange='UpdatePayable()'></td>
	</td>
  </tr>      
  <tr>
  <td width="30%">Payable</td>
  <td width="70%"><input type = 'text' name ='txtPayable' id = 'txtPayable' size="25" value='<%=formatnum(payable)%>' readonly class="readonlyex"></td>
  </tr>
 
  <tr>
     <td width="100%" colspan="2" align=right>
		<BR>
		<BR>
		<BR>
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save ">		
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();">
		&nbsp;&nbsp;
		
		<input type = 'hidden' name ='action' id = 'action' value="Execute">&nbsp;
		<input type = 'hidden' name ='txtRenouncee' id = 'txtRenouncee' value='<%=Renouncee%>'>
		<input type = 'hidden' name ='txtid' id = 'txtid' value='<%=ID%>'>
		<input type = 'hidden' name ='IDP' id = 'IDP' value="<%=idp%>">
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
<script language="javascript">
  UpdateBalances();
</script>
</body>

</html>
<%
end select
%>