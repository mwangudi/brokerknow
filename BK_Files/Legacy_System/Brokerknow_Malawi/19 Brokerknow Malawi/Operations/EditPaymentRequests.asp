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
<title>Add Payment Request</title>
 
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
	
	select case action
	case "EXECUTE" 
	   Dim palno
					entity=1					
			        clientName = Request.Form ("cboClient")
					Amount = Replace(Request.Form("txtAmount"),",","")
					PayDate = Request.Form("txtDate")
					narrative = Request.Form("txtNarrative")	 							        
			         
					 If Trim(clientName) = "" Then%>
					         <script language = 'vbscript'>
					         		ShowMessage "Please Specify the client Name"
					         		
					         </script>
					         <% response.end
					 End If	 					 
					 					 
					'Check whether price has been entered
					If Trim(Amount) = "" Then%>
					    <script language = 'vbscript'>
					         	ShowMessage "Please Enter the Requested Amount "
					         	
					    </script>
					    <% response.end
					End If
					
					'ensure price is numeric
					If (Amount <> "") And (Not IsNumeric(price)) Then%>
					    <script language = 'vbscript'>
							ShowMessage "price Amount must be numeric"
							
					    </script>
					    <% response.end
					End If
										
					'PaymentTypes=4					
					
					Set conn = GetActiveConnection("KBroker")
					conn.BeginTrans		
							'Save payment																				
							sqlStr = "UPDATE PaymentRequests SET Client_DPA_= " & "" & ClientName & "" & "," & _
									 "PaymentAmount = " & "" & Amount & "" & "," & _
									 "ModifiedBy = " & "" & UserId & "" & "," & _
									 "RequestNarrative = " & "'" & narrative & "'" & "," & _
									 "RequestPayDate = " & "#" & PayDate & "#" & "," & _
									 "TimeModified = GetDate() " & _
									 " Where Request_DPA_=" & ID
									 
							sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))							
														
							'Response.Write sqlStr
							'Response.End
																		                                                     
							conn.Execute sqlStr												
														
							conn.execute ("Exec ClientTotalProcedure " & ClientName)							
							conn.execute ("Exec ClientBalanceProcedure " & ClientName)
											
					conn.CommitTrans
					conn.Close
					Set conn = Nothing
					WritefraEnabledDialogCloseScript
					Response.End	   	   	
   	Case ""
   	Set conn = GetActiveConnection("KBroker")
   	sqlstr ="Select PaymentRequests.*,Client.ClientCDSNo from PaymentRequests inner join Client on PaymentRequests.client_DPA_=client.Client_DPA_ where request_DPA_=" & ID
   	Set rsEdit = conn.Execute(sqlStr)   	
   	
   clientName = rsEdit("Client_DPA_") 	   	
%>
<body Class="Dialog">

<!--#include file="../libroutines.asp"-->
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>

<SCRIPT language="JavaScript">
	var cal=new ctlSpiffyCalendarBox("cal", "frm<%=DataSource%>", "txtDate","cmdDate","<%= FormatDate(rsEdit("RequestPayDate")) %>",1);
</SCRIPT>

<form name = 'frmAddOffering' method = 'post' id="frmMain" action = "EditPaymentRequest.asp" >
	<table border="0" width="100%" cellpadding=2 cellspacing=2>
		<tr>
           <td>Client</td>
		   <td>&nbsp;
				<input type = 'text' name ='txtClientCode' id = 'txtClientCode' size="10" onBlur="txtval = this.value; selectItem(cboClient);UpdateCodes(true,cboClient,txtCdsNo);UpdateBalances();" value="<%=rsEdit("Client_DPA_")%>">&nbsp;
				<input type = 'text' name ='txtCdsNo' id = 'txtCdsNo' size="16" onBlur="txtval = this.value; selectItems(cboClient);UpdateCode(true,cboClient,txtClientCode);UpdateBalances();" value="<%=rsEdit("ClientCDSNo")%>">&nbsp;
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
					
				    sqlStr = "SELECT LTRIM(RTRIM(ClientName)) as ClientName,Client_DPA_,ClientCDSNo FROM [client]  where Deleted=0 Order By ClientName ASC"
				    
				    sqlStr = " SELECT TOP 100 PERCENT ISNULL(dbo.ClientBalances.CurrentBal, 0) + ISNULL(dbo.Client.CreditLimit, 0) - ISNULL(dbo.ClientTotal.Total, 0) AS AvailableCredit,   " & _
							"         ISNULL(dbo.ClientBalances.CurrentBal, 0) AS CurrentBal," & _
							"         dbo.Client.Client_DPA_, dbo.Client.ClientName," & _
							"         dbo.Client.ClientCDSNo, dbo.ClientTotal.Total  " & _
							"  FROM dbo.Client LEFT OUTER JOIN  " & _
							"       dbo.ClientTotal ON dbo.Client.Client_DPA_ = dbo.ClientTotal.Client_DPA_ LEFT OUTER JOIN  " & _
							"       dbo.ClientBalances ON dbo.Client.Client_DPA_ = dbo.ClientBalances.client_DPA_ " & _							
							"  WHERE     (dbo.Client.Deleted = 0) " & _
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
			<td nowrap>Request Amount</td>
			<td><input type = 'text' name ='txtAmount' id = 'txtAmount' size="25" onblur='UpdatePayable()' value='<%=FormatNum(rsEdit("PaymentAmount"))%>'></td>
		</tr>
		<tr>
		   <td width="20%" nowrap>Payment Date</td>
		   <td width="80%"><SCRIPT language="JavaScript">cal.writeControl();</SCRIPT></td>
		</tr>        	         
     <tr>
		<td nowrap>Narrative</td>
		<td><textarea name ='txtNarrative' id = 'txtNarrative' rows="5" cols="30" ><%=rsEdit("RequestNarrative")%></textarea>
		</td>     
     </tr>                         
  <tr>
     <td width="100%" colspan="2" align=right>		
		<BR>
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save ">
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