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
<title>Edit Processed Request</title>
 
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
		 		 
		 dim RequestAmount
		 dim credit
		 dim credittxt
		 
		 price = 0
		 acrights = 0
		 payable = 0
		 credit = 0
		 credittxt = ""
		 
		 RequestAmount=document.frmMain.elements("txtAmount").value
		 
		 credittxt = document.frmMain.elements("txtAvailableCredit").value
		 credit = replace(credittxt,",","")	 	 		   
		   
		    if(RequestAmount<>"" and credit<>"") then
				if(CCur(RequestAmount)>CCur(credit)) then
				ShowMessage "You have insuficient balance to complete this transaction"
			
				document.frmMain.elements("cmdAdd").disabled=true
				else
				document.frmMain.elements("cmdAdd").disabled=false
				end if
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
			        
	
	action = ucase(action)			
	
	ID = Request("ID")
	Ids = Split(ID,"<->")
	
	ClientID = Ids(1)
	itemID =Ids(0)
	ProcessID = Ids(2)
	
	select case action
	case "EXECUTE" 
	   Dim palno
					entity=1					
			        reference = Request.Form("txtRef")
					PDate = Trim(Request.Form("txtDate1"))
					PDate = PDate & " " & Time
					narrative = Request.Form("txtPNarrative")
					ChequeStatus = Request.Form("ChequeCollection")
			        PaymentType = Request.Form("cboPaymentTypes")
			        bank = Request.Form("cboBank")
			        buttonAction = Trim(Ucase(Request.Form("cmdAdd"))) 
			           
					 If Trim(PaymentType) = "" Then%>
					         <script language = 'vbscript'>
					         		ShowMessage "Please specify the Payment Types."
					         		
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
					  					
					'PaymentTypes=4					
					
					Set conn = GetActiveConnection("KBroker")
					conn.BeginTrans		
							'Save payment
							
							if instr(1,buttonAction,"APPROVE") > 0 then																				
							sqlStr = "UPDATE PaymentRequests SET BankAccount_DPA_= " & "" & bank & "" & "," & _
									 "PaymentType_DPA_ = " & "" & PaymentType & "" & "," & _
									 "ModifiedBy = " & "" & UserId & "" & "," & _
									 "PaymentNarrative = " & "'" & narrative & "'" & "," & _
									 "PaymentReference = " & "'" & reference & "'" & "," & _
									 "PaidDate = " & "#" & PDate & "#" & "," & _
									 "TimeModified = GetDate(), " & _									 
									 "Approved = 1, " & _									 
									 "ApprovedBy = " & "" & UserId & "" & "," & _
									 "ApprovalDate = GetDate() " & _									 
									 " Where Processed_DPA_ = " & ProcessID & " and Approved=0"
							else
							sqlStr = "UPDATE PaymentRequests SET BankAccount_DPA_= " & "" & bank & "" & "," & _
									 "PaymentType_DPA_ = " & "" & PaymentType & "" & "," & _
									 "ModifiedBy = " & "" & UserId & "" & "," & _
									 "PaymentNarrative = " & "'" & narrative & "'" & "," & _
									 "PaymentReference = " & "'" & reference & "'" & "," & _
									 "PaidDate = " & "#" & PDate & "#" & "," & _
									 "TimeModified = GetDate(), " & _									 
									 "Approved = 2, " & _									 
									 "ApprovedBy = " & "" & UserId & "" & "," & _
									 "ApprovalDate = GetDate() " & _									 									 
									 " Where Processed_DPA_ = " & ProcessID & " and Approved=0"
							
							end if
									 
							sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))							
														
							'Response.Write sqlStr
							'Response.End
																		                                                     
							conn.Execute sqlStr												
														
							'conn.execute ("Exec ClientTotalProcedure " & ClientName)							
							'conn.execute ("Exec ClientBalanceProcedure " & ClientName)
											
					conn.CommitTrans
					conn.Close
					Set conn = Nothing
					WritefraEnabledDialogCloseScript
					Response.End	   	   	
   	Case ""
   	Set conn = GetActiveConnection("KBroker")
   	
   	sqlstr ="Select PaymentRequests.*,Client.ClientCDSNo from PaymentRequests inner join Client on PaymentRequests.client_DPA_=client.Client_DPA_ where request_DPA_=" & ID   	
   	   	
   	sqlstr ="SELECT     Client.ClientCDSNo,MAX(PaymentRequests.Request_DPA_) AS Request_DPA_, PaymentRequests.Client_DPA_, Client.ClientName,  " & _
			"                       SUM(PaymentRequests.PaymentAmount) AS PaymentAmount, MAX(PaymentRequests.RequestDate) AS RequestDate,  " & _
			"                       MAX(PaymentRequests.RequestPayDate) AS RequestPayDate, Users.Surname + ' ' + Users.OtherNames AS ModifiedBy,  " & _
			"                       MAX(PaymentRequests.TimeModified) AS TimeModified, Account.AccountName, CAST(FLOOR(CAST(PaymentRequests.PaidDate AS float)) AS Datetime)  " & _
			"                       AS PaidDate, PaymentTypes.Description, PaymentRequests.PaymentReference,MAX(PaymentRequests.RequestNarrative) AS RequestNarrative, PaymentRequests.Approved,  " & _
			"                       ISNULL(PaymentRequests.Processed_DPA_, 0) AS Processed_DPA_, PaymentRequests.BankAccount_DPA_,ltrim(rtrim(PaymentRequests.PaymentNarrative)) as PaymentNarrative" & _
			" FROM         PaymentRequests INNER JOIN " & _
			"                       Client ON PaymentRequests.Client_DPA_ = Client.Client_DPA_ INNER JOIN " & _
			"                       Users ON PaymentRequests.ModifiedBy = Users.UserID LEFT OUTER JOIN " & _
			"                       PaymentTypes ON PaymentRequests.PaymentType_DPA_ = PaymentTypes.PaymentTypes_DPA_ LEFT OUTER JOIN " & _
			"                       Account ON PaymentRequests.BankAccount_DPA_ = Account.Account_DPA_ " & _
			" WHERE     (PaymentRequests.Deleted = 0) AND (PaymentRequests.Approved = 0) and (PaymentRequests.Processed_DPA_ = " & ProcessID & ")" & _
			" GROUP BY Client.ClientCDSNo,PaymentRequests.Client_DPA_, Client.ClientName, Users.Surname, Users.OtherNames, Account.AccountName,  " & _
			"                       CAST(FLOOR(CAST(PaymentRequests.PaidDate AS float)) AS Datetime), PaymentTypes.Description, PaymentRequests.PaymentReference,  " & _
			"                       PaymentRequests.Approved, PaymentRequests.Processed_DPA_, PaymentRequests.BankAccount_DPA_,ltrim(rtrim(PaymentRequests.PaymentNarrative)) " & _
			" ORDER BY PaymentRequests.RequestDate DESC"
			
   	Set rsEdit = conn.Execute(sqlStr)   	
   
   'Response.Write rsEdit("")
   'Response.End
   	
   clientName = rsEdit("Client_DPA_") 	   	
%>
<body Class="Dialog">

<!--#include file="../libroutines.asp"-->
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>

<SCRIPT language="JavaScript">
	var cal=new ctlSpiffyCalendarBox("cal", "frm<%=DataSource%>", "txtDate1","cmdDate","<%= FormatDate(Date) %>",1);
</SCRIPT>

<form name = 'frmAddOffering' method = 'post' id="frmMain" action = "EditApproveRequest.asp" >
	<table border="0" width="100%" cellpadding=0 cellspacing=0>
		<tr><td colspan="2" height="1"><hr></hr></td></tr>
		<tr><td colspan="2" height="1" align="center"><b>Payment Request Details</b></td></tr>
		<tr><td colspan="2" height="1"><hr></hr></td></tr>
		<tr>
           <td>Client</td>
		   <td>
				<input type = 'text' name ='txtClientCode' id = 'txtClientCode' size="10" onBlur="txtval = this.value; selectItem(cboClient);UpdateCodes(true,cboClient,txtCdsNo);UpdateBalances();" value="<%=rsEdit("Client_DPA_")%>" readonly class="readonlyEx">&nbsp;
				<input type = 'text' name ='txtCdsNo' id = 'txtCdsNo' size="16" onBlur="txtval = this.value; selectItems(cboClient);UpdateCode(true,cboClient,txtClientCode);UpdateBalances();" value="<%=rsEdit("ClientCDSNo")%>" readonly class="readonlyEx">&nbsp;
					<select disabled name = 'cboClient' id = 'cbocboClient' size="1" 
					onKeypress="return (dodefaultaction()==''); " 
					onKeydown="return (dodefaultaction()==''); " 
					onKeyup="return (UpdateCode(change(cboClient,0),cboClient,txtClientCode));UpdateBalances();" 
					onChange="UpdateCode(true,cboClient,txtClientCode);UpdateCodes(true,cboClient,txtCdsNo);UpdateBalances();"
					onfocus="txtval = '';inputIsItemCode = 1;" 
					onblur="txtval = '';inputIsItemCode = 1;">
    	
					<%
		
					Dim thisClientName
					Dim NameClient
					
				    sqlStr = "SELECT LTRIM(RTRIM(ClientName)) as ClientName,Client_DPA_,ClientCDSNo FROM [client]  where Deleted=0 Order By ClientName ASC"
				    
				    sqlStr = " SELECT TOP 100 PERCENT ISNULL(dbo.ClientBalances.CurrentBal, 0) + ISNULL(dbo.Client.CreditLimit, 0) - ISNULL(dbo.ClientTotal.Total, 0) AS AvailableCredit,   " & _
							"         ISNULL(dbo.ClientBalances.CurrentBal, 0) AS CurrentBal," & _
							"         dbo.Client.Client_DPA_, dbo.Client.ClientName," & _
							"         dbo.Client.ClientCDSNo, dbo.ClientTotal.Total  " & _
							"  FROM dbo.Client LEFT OUTER JOIN  " & _
							"       dbo.ClientTotal ON dbo.Client.Client_DPA_ = dbo.ClientTotal.Client_DPA_ LEFT OUTER JOIN  " & _
							"       dbo.ClientBalances ON dbo.Client.Client_DPA_ = dbo.ClientBalances.client_DPA_ " & _							
							"  WHERE     (dbo.Client.Deleted = 0)  and Client.Client_DPA_ not in (Select Client_DPA_ from offerings where deleted=0)" & _
							"      ORDER BY LTRIM(RTRIM(dbo.Client.ClientName)) "

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
			<td nowrap>Request Amount</td>
			<td><input type = 'text' name ='txtAmount' id = 'txtAmount' size="25" onblur='UpdatePayable()' readonly class="readonlyEx" value='<%=FormatNum(rsEdit("PaymentAmount"))%>'></td>
		</tr>
		<tr>
		   <td width="20%" nowrap>Payment Date</td>
		   <td width="80%"><input type = 'text' name ='txtDate' id='txtDate' size ="25" value="<%= FormatDate(rsEdit("RequestPayDate")) %>" readonly class="readonlyEx"> </td>
		</tr>        	         
     <tr>
		<td nowrap>Narrative</td>
		<td><input type = 'text' name ='txtNarrative' id='txtNarrative' size ="45" value="<%=rsEdit("RequestNarrative")%>" readonly class="readonlyEx"> 
		</td>     
     </tr>
  
  <tr><td colspan="2" height="1"><hr></hr></td></tr>
		<tr><td colspan="2" height="1" align="center"><b>Process Payment</b></td></tr>
		<tr><td colspan="2" height="1"><hr></hr></td></tr>
		
  <tr>
    <td width="15%">Payment Type</td>
    <td width="54%"><select name = 'cboPaymentTypes' id = 'cboPaymentTypes' size="1" onchange='EnableCheqCollection(this.value)'>
    	
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
                Do Until rs.EOF
                		if rs.Fields("Account_DPA_") = rsEdit.Fields("BankAccount_DPA_") Then%>
                			<option selected SearchCode = "<%=rs.Fields("AccountCode")%>" SearchText = "<%=rs.Fields("AccountName")%>" value = '<%=rs.Fields("Account_DPA_")%>'><%=rs.Fields("AccountNameEx")%></option>
                		<%else%>
							<option SearchCode = "<%=rs.Fields("AccountCode")%>" SearchText = "<%=rs.Fields("AccountName")%>" value = '<%=rs.Fields("Account_DPA_")%>'><%=rs.Fields("AccountNameEx")%></option>
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
    <td width="15%">Reference</td>
    <td width="54%"><input type = 'text' name ='txtRef' id = 'txtRef' size="20" value='<%=rsEdit.Fields("PaymentReference")%>'></td>
    <td width="31%">

	</td>
  </tr>
  <tr>
    <td width="15%">Narrative</td>
    <td width="54%">
    <textarea name ='txtPNarrative' id = 'txtPNarrative' rows="1" cols="20" > <%=rsEdit.Fields("PaymentNarrative")%></textarea></td>
    <td width="31%">

	</td>
  </tr>
                          
  <tr>
     <td width="100%" colspan="2" align=right>		
		<BR>
		<!--<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Approve ">-->
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Reject ">
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();">
		&nbsp;&nbsp;
		<%
		if(trim(securityname) ="") then
			securityname = "Kenya Electricity Generating Company"
		end if
		%>
		<input type = 'hidden' name ='action' id = 'action' value="Execute">		
		<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
      </td>
  </tr>
</table>
</form>


</body>

</html>
<%
end select
%>