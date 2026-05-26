<html>
<head>
<title>Edit IPO Forward</title>
 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css"> 
<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT>

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
	
	price=trim(document.frmMain.elements("txtprice").value)
	acrights=trim(document.frmMain.elements("txtAlloted").value)
	credittxt = trim(document.frmMain.elements("txtAvailableCredit").value)
	credit = replace(credittxt,",","")	 
		 		   
	if (price > 0 and acrights > 0) then		   
		payable = price*acrights
	end if
			
	document.frmMain.elements("txtPayable").value = Replace(FormatNumber(payable,0),",","")
	
	document.frmMain.optAcceptance(0).checked = false
	document.frmMain.optAcceptance(1).checked = false
	
	document.all.item("trPartial").style.display = "none"
	document.all.item("trFull").style.display = "none"
End Function

Function UpdatePartial()
	dim price
	dim acrights
	
	price = 0
	acrights = 0
	
	price=trim(document.frmMain.elements("txtprice").value)
	acrights=trim(document.frmMain.elements("txtPartial").value)
	
	if (price > 0 and acrights > 0) then		   
		payable = price*acrights
	end if
	
	document.frmMain.elements("txtPartialAmount").value = Replace(FormatNumber(payable,0),",","")  	
End Function	

Function UpdateFull()
	dim price
	dim acrights
	dim newrights
	
	price = 0
	acrights = 0
	newrights = 0 
	
	price = trim(document.frmMain.elements("txtprice").value)
	acrights = trim(document.frmMain.elements("txtAlloted").value)
	
	if (price > 0 and acrights > 0) then		   
		payable = price * acrights
	end if
	
	document.frmMain.elements("txtFull").value = acrights
	document.frmMain.elements("txtFullAmount").value = Replace(FormatNumber(payable,0),",","") 
	
	newrights = trim(document.frmMain.elements("txtNew").value)
	
	newpayable = price * newrights
	
	document.frmMain.elements("txtNewAmount").value = Replace(FormatNumber(newpayable,0),",","") 
	
	document.frmMain.elements("txtTotal").value = Replace(FormatNumber(cdbl(acrights)+cdbl(newrights),0),",","") 
	document.frmMain.elements("txtTotalAmount").value = Replace(FormatNumber((cdbl(acrights)+cdbl(newrights))*price,0),",","") 
End Function	
</script>

<script language="javascript">
function UpdateBalances()
	{
	client = document.frmMain.elements("cboclient")
		
	document.frmMain.elements("txtAvailableCredit").value = client[client.selectedIndex].Credit;
	document.frmMain.elements("txtCurrentBal").value = client[client.selectedIndex].CurrentBal;	
	
	ClientID = document.all.item("cboclient").value;
	SecurityID = document.frmMain.elements("cboOfferings")[document.frmMain.elements("cboOfferings").selectedIndex].ParentSecurity;
	
	//alert(SecurityID);
	
	//GetHoldings(document.all.item("cboclient").value,document.all.item("cboOfferings").value);
	var XMLHttpRequestObject = false;

	if (window.XMLHttpRequest)
	{
		XMLHttpRequestObject = new XMLHttpRequest();
	}
	else if (window.ActiveXObject)
	{
		XMLHttpRequestObject = new ActiveXObject("Microsoft.XMLHttp");
	}

	if (XMLHttpRequestObject)
		{	
		url = "GetHoldings.asp?cID="+ClientID+"&sID="+SecurityID;

		XMLHttpRequestObject.open("GET",url);
		
		XMLHttpRequestObject.onreadystatechange = function()
			{
			if (XMLHttpRequestObject.readyState == 4 && XMLHttpRequestObject.status == 200)
				{
				returnStr = XMLHttpRequestObject.responseText;
				
				var allot;
				allot = returnStr * document.all.item("txtRatio").value;
				allot = parseInt(allot,10)
				
				document.all.item("txtHoldings").value = returnStr;
				document.all.item("txtAlloted").value = allot;
				document.all.item("txtPayable").value = allot * document.all.item("txtPrice").value;
				}
			}
		}
	XMLHttpRequestObject.send(null);
	
	UpdatePayable();
	}
		
function UpdatePrice(drpOfferings)
	{
	var price;
	var ratio;
	//var securityname;

	price = drpOfferings[drpOfferings.selectedIndex].SearchPrice; 
	ratio = drpOfferings[drpOfferings.selectedIndex].Ratio; 
		
	//securityname = drpOfferings[drpOfferings.selectedIndex].text
	
	document.frmMain.elements("txtprice").value = price ;
	document.frmMain.elements("txtRatio").value = ratio ;
	//document.frmMain.elements("securityname").value = securityname ;
	
	//GetHoldings(document.all.item("cboclient").value,document.all.item("cboOfferings").value);
	}

function FullOrPartial(myOpt)
	{
	if (myOpt == 1)
		{
		document.frmMain.optAcceptance(0).checked = true;
		
		document.all.item("trPartial").style.display = 'none';
		document.all.item("trFull").style.display = '';
		
		UpdateFull();
		}
	else
		{
		document.frmMain.optAcceptance(1).checked = true;
		
		document.all.item("trPartial").style.display = '';
		document.all.item("trFull").style.display = 'none';
		}
	}

function FullOrPartial2(myOpt,oType)
	{
	if (oType == 2)
		{
		if (myOpt == 1)
			{
			document.frmMain.optAcceptance(0).checked = true;
			
			document.all.item("trPartial").style.display = 'none';
			document.all.item("trFull").style.display = '';
			
			UpdateFull();
			}
	
		if (myOpt == 2)
			{
			document.frmMain.optAcceptance(1).checked = true;
			
			document.all.item("trPartial").style.display = '';
			document.all.item("trFull").style.display = 'none';
			}
		}
	}
		
function showOfferType()
	{
	var offerType = document.all.item("cboOfferings").options[document.all.item("cboOfferings").selectedIndex].OfferType;
	
	if (offerType==2)
		{
		document.all.item("trHoldings").style.display = '';
		document.all.item("trAcceptance").style.display = '';
		}
	else
		{
		document.all.item("trHoldings").style.display = 'none';
		document.all.item("trAcceptance").style.display = 'none';
		}
	}
</script>
</head>

<body Class="Dialog">

<!--#include file="../libroutines.asp"-->
<%
Dim action
Dim conn 
Dim sqlStr
Dim rs
	
action = UCase(Request.Form("action"))
ID = Request("ID")
	
UserId = SESSION("UserID")

Set conn = GetActiveConnection("KBroker")

If action = "EXECUTE" Then
	ID = Request("ID")
	
	palno = Request.Form("txtPalNo")
	clientName = Request.Form("cboClient")
	CurrentBal = Request.Form("txtCurrentBal")
	AvailableCredit = Request.Form("txtAvailableCredit")
	offering = Request.Form("cboOfferings")
	price = Request.Form("txtPrice")
	AlRights = Request.Form("txtAlloted")
	payable = Request.Form("txtPayable")
	chkno = Request.Form("txtchkNo")
	narrative = Request.Form("txtNarrative")
	
	Accept = Request.Form("optAcceptance")
	Full = Request.Form("txtFull")''
	FullAmount = Request.Form("txtFullAmount")
	AddNew = Request.Form("txtNew")''
	NewAmount = Request.Form("txtNewAmount")
	Total = Request.Form("txtTotal")
	TotalAmount = Request.Form("txtTotalAmount")
	Partial = Request.Form("txtPartial")
	PartialAmount = Request.Form("txtPartialAmount")
	
	Select Case Accept
		Case 1
			''Full
			AlRights = Total
			payable = TotalAmount
		Case 2
			''Partial
			AlRights = Partial
			payable = PartialAmount
	End Select
	
	If Trim(palno) = "" Then%>
		<script language = 'vbscript'>
			ShowMessage "Please enter the PAL No"
		</script>
		<% response.end
	End If
					 
	If Trim(clientName) = "" Then%>
		<script language = 'vbscript'>
			ShowMessage "Please select the Client"
		</script>
		<% response.end
	End If	 
					 
	If Trim(offering) = "" Then%>
		<script language = 'vbscript'>
		ShowMessage "Please select the Offering"					         		
		</script>
		<% response.end
	End If
					 					 
	If Len(AlRights) = "" Then%>
		<script language = 'vbscript'>
			ShowMessage "Please enter the Alloted Rights"
		</script>
		<% response.end
	End If
										
	If (Alrights <> "") And (Not IsNumeric(Alrights)) Then%>
	    <script language = 'vbscript'>
			ShowMessage "Alloted Rights should be numeric"
		</script>
	    <% response.end
	End If					
					
IF ENABLED THEN	
	if(CCur(Alrights) < 500 ) then
		%>
		<script language="VBScript">
			ShowMessage "Applicable Quantity should be at least 500"		   
		</script>
		<%	Response.End 
	end if
		   
	if((Cdbl(Alrights) mod 100) <> 0 ) then
		%>
		<script language="VBScript">
			ShowMessage "The Applied Quantity should in multiples of hundreds"
		</script>
		<%	Response.End 
	end if
	
	if (CCur(payable) > CCur(AvailableCredit)) then
		%>
		<script language="VBScript">
			ShowMessage "This amount payable should be equal to the available credit which is <%=AvailableCredit%>"
		</script>
		<%	Response.end
	end if
END IF
	
	If Accept = "" Then Accept = 1
	
	conn.BeginTrans		

			sqlStr = "UPDATE Offerings SET" & _
				" Client_DPA_ = " & ClientName & "" & _
				" ,PAL_No = '" & palno & "'" & _
				" ,Offering = " & offering & "" & _
				" ,Offering_Price = " & Price & "" & _
				" ,Alloted_Rights = " & AlRights & "" & _
				" ,ChangedBy = " & UserId & "" & _
				" ,TimeChanged = GetDate()" & _
				" ,AcceptanceType = " & Accept & "" & _
				" ,Additional = " & AddNew & "" & _
				" ,OfferCheque = '" & chkno & "'" & _
				" ,OfferBank = '" & narrative & "'" & _
				" WHERE Offering_DPA_ = "& ID
			
			sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))							
			
			'Response.Write sqlStr
			'Response.End 
																				                                                     
			conn.Execute sqlStr																			
	conn.CommitTrans
	conn.Close

	Set conn = Nothing
	WritefraEnabledDialogCloseScript
	Response.End	
End If

sqlStr = "SELECT * FROM Offerings WHERE Offering_DPA_= " & ID
Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        
If rs.EOF Or rs.BOF Then
	if enabled then
	%>
	<script language = 'vbscript'>
		window.self.ShowMessage "The selected record cannot be retrieved for editing"
	</script>
	<%
	Response.End
ElseIf IsNumeric(rs("Batch_No")) Then
	WriteDialogRefuseOpenScript
	%>
	<script language = 'vbscript'>
		window.parent.dialogArguments.opener.alert "The selected record cannot be retrieved for editing" & Chr(13) & "as it has already been batched."
	</script>
	<%
	Response.End
end if
Else
	Offering_DPA_ = rs("Offering_DPA_")
	palno = rs("PAL_No")
	client = rs("Client_DPA_")
	price = rs("Offering_Price")
	offering = rs("Offering")
	
	payable = rs("Offering")*rs("Offering_Price")
	
	chkno = rs("OfferCheque")
	narrative = rs("OfferBank")
	
	AlRights = cdbl(rs("Alloted_Rights"))
	Accept = rs("AcceptanceType")
	
	Additional = rs("Additional")
End If          
%>
<form name = 'frmAddSecurity' method = 'post' id="frmMain" action = "EditForward.asp" >
	<table border="0" width="80%" cellpadding=2 cellspacing=2>
		<tr>
			<td width="20%" nowrap>PAL NO</td>
			<td width="80%" nowrap><input type="text" name="txtPalNo" id="txtPalNo" size="25" value="<%=palno%>"></td>
		</tr>
				
		<tr>
			<td width="20%" nowrap>Client</td>
			<td width="80%" nowrap>
				<input type = 'text' name ='txtClientCode' id = 'txtClientCode' value="<%=client%>" size="10" onBlur="txtval = this.value; selectItem(cboClient);UpdateCodes(true,cboClient,txtCdsNo);UpdateBalances();">&nbsp;
				<input type = 'text' name ='txtCdsNo' id = 'txtCdsNo' size="16" onBlur="txtval = this.value; selectItems(cboClient);UpdateCode(true,cboClient,txtClientCode);UpdateBalances();">&nbsp;
				<select name = 'cboClient' id = 'cboClient' size="1" 
				onKeypress="return (dodefaultaction()==''); " 
				onKeydown="return (dodefaultaction()==''); " 
				onKeyup="return (UpdateCode(change(cboClient,0),cboClient,txtClientCode));UpdateBalances();" 
				onChange="UpdateCode(true,cboClient,txtClientCode);UpdateCodes(true,cboClient,txtCdsNo);UpdateBalances();"
				onfocus="txtval = '';inputIsItemCode = 1;" 
				onblur="txtval = '';inputIsItemCode = 1;">
				<%
				sqlStr = " SELECT TOP 100 PERCENT ISNULL(dbo.ClientBalances.CurrentBal, 0) + ISNULL(dbo.Client.CreditLimit, 0) - ISNULL(dbo.ClientTotal.Total, 0) AS AvailableCredit,   " & _
					" ISNULL(dbo.ClientBalances.CurrentBal, 0) AS CurrentBal," & _
					" dbo.Client.Client_DPA_, dbo.Client.ClientName," & _
					" dbo.Client.ClientCDSNo  " & _
					" FROM dbo.Client LEFT OUTER JOIN  " & _
					" dbo.ClientTotal ON dbo.Client.Client_DPA_ = dbo.ClientTotal.Client_DPA_ LEFT OUTER JOIN  " & _
					" dbo.ClientBalances ON dbo.Client.Client_DPA_ = dbo.ClientBalances.client_DPA_ " & _							
					" WHERE (dbo.Client.Deleted = 0)" & _
					" ORDER BY LTRIM(RTRIM(dbo.Client.ClientName)) "
									
				Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
				
				intcountrs = rs.recordcount
				if intcountrs > 0 then
					clientdata = rs.getrows()
						     
					for intcount = 0 to intcountrs-1
						thisClientName = mid(trim(clientdata(3,intcount)),1,50)
						clientID = trim(clientdata(2,intcount))
						AvailableCredit = trim(clientdata(0,intcount))
						CreditBal = trim(clientdata(1,intcount))
						CDSNo = trim(clientdata(4,intcount))
						
						If (clng(trim(client)) = clng(trim(ClientID))) Then
							%>
							<option Selected Credit = "<%=FormatNumEx(AvailableCredit,2)%>" CurrentBal = "<%=FormatNumEx(CreditBal,2)%>" SearchCode = "<%=clientID%>" SearchText = "<%=thisClientName%>" SearchCds = "<%=CDSNo%>" value = '<%=clientID%>'><%=thisClientName%></option>
							<%
							AvailableCredit2 = trim(clientdata(0,intcount))
							CreditBal2 = trim(clientdata(1,intcount))
							CDSNo2 = trim(clientdata(4,intcount))
						Else
							%>
							<option Credit = "<%=FormatNumEx(AvailableCredit,2)%>" CurrentBal = "<%=FormatNumEx(CreditBal,2)%>" SearchCode = "<%=clientID%>" SearchText = "<%=thisClientName%>" SearchCds = "<%=CDSNo%>" value = '<%=clientID%>'><%=thisClientName%></option>
							<%
						End If
					next
				end if
				%>
				</select>
			</td>
		</tr>
		     
		<tr>
			<td width="20%" nowrap>&nbsp;</td>
			<td width="80%" nowrap>
				<table>
					<tr>
						<td>Client Balance</td>
						<td>Available Credit</td>
					</tr>     
					<tr>
						<td><input type = 'text' name ='txtCurrentBal' id = 'txtCurrentBal' readonly class="readonlyex" size="15" value="<%=FormatNumber(AvailableCredit2,2)%>"></td>
						<td><input type = 'text' name ='txtAvailableCredit' id = 'txtAvailableCredit' readonly class="readonlyex" size="15" value="<%=FormatNumber(CreditBal2,2)%>"></td>
					</tr>
				</table>
			</td>	         
		</tr>
			
		<tr>
			<td width="20%" nowrap>Offering Name</td>
			<td width="80%" nowrap>
				<select name = 'cboOfferings' id = 'cboOfferings' size="1" onChange="UpdatePrice(this);showOfferType();">
				<% 
				sqlStr = "SELECT * FROM [SecurityListOfferings] " & _
				" WHERE cast(floor(cast(ClosingDate as float)) as datetime) >= cast(floor(cast(GetDate() as float)) as datetime)" & _
				" Order By SecurityName ASC"
				Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
				
				If Not (rs.EOF Or rs.BOF) Then
					Do Until rs.EOF
							if rs("Security_DPA_") = offering Then
								%>
								<option selected ParentSecurity = "<%=Rs("ParentSecurity_DPA_")%>" Ratio = "<%=Rs("Ratio")%>" OfferType = "<%=Rs("OfferType_DPA_")%>" SearchPrice = "<%=rs.Fields("SecurityMktPrice")%>" value = '<%=rs.Fields("Security_DPA_")%>'><%=rs.Fields("SecurityName")%></option>
								<%
								price = rs("SecurityMktPrice")
								ratio = Rs("Ratio")
								ParentSecurity = rs("ParentSecurity_DPA_")
							else
								%>                   						
								<option ParentSecurity = "<%=Rs("ParentSecurity_DPA_")%>" Ratio = "<%=Rs("Ratio")%>" OfferType = "<%=Rs("OfferType_DPA_")%>" SearchPrice = "<%=rs.Fields("SecurityMktPrice")%>" value = '<%=rs.Fields("Security_DPA_")%>'><%=rs.Fields("SecurityName")%></option>
								<%
								'price = rs("SecurityMktPrice")
							end if
							
							OfferType = Rs("OfferType_DPA_")
						rs.MoveNext
					Loop
				End If
				%>
				</select>
			</td>     
		</tr>              
		                            
		<tr>
			<td width="20%" nowrap>Offering Price</td>
			<td width="80%" nowrap><input type = 'text' name ='txtPrice' id = 'txtPrice' size="20" value="<%=price%>" onchange='UpdatePayable()' readonly class="readonlyex"></td>
		</tr>          

		<tr id="trHoldings" name="trHoldings" style="display:none;">
			<td width="20%" nowrap>Holdings</td>
			<td width="80%" nowrap><input type = 'text' name ='txtHoldings' id = 'txtHoldings' size="20" value=0 readonly class="readonlyex"></td>
		</tr>     

		<tr>
			<td width="20%" nowrap>Quantity Applied</td>
			<td width="80%" nowrap><input type = 'text' name ='txtAlloted' id = 'txtAlloted' size="20" value="<%=AlRights%>" onchange='UpdatePayable()'></td>
		</tr>     
			 
		<tr>
			<td width="20%" nowrap>Amount Payable</td>
			<td width="80%" nowrap><input type = 'text' name ='txtPayable' id = 'txtPayable' size="25" value="<%=formatnum(payable)%>" readonly class="readonlyex"></td>
		</tr>
		
		<tr>
			<td width="20%">Reference</td>
			<td width="80%"><input type = 'text' name ='txtchkNo' id = 'txtchkNo' size="20" value='<%=chkno%>'></td>    
		</tr>

		<tr>
			<td width="20%">Bank</td>
			<td width="80%"><textarea name ='txtNarrative' id = 'txtNarrative' cols="50" rows="5"><%=narrative%></textarea></td>   	
		</tr>
		
		<tr id="trAcceptance" name="trAcceptance" style="display:none;">
			<td width="20%" nowrap>Acceptance</td>
			<td width="80%" nowrap>
			Full&nbsp;<input type = 'radio' name ='optAcceptance' id = 'optAcceptance' value=1 onclick="FullOrPartial(this.value);">
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			Partial&nbsp;<input type = 'radio' name ='optAcceptance' id = 'optAcceptance' value=2 onclick="FullOrPartial(this.value);">
			</td>
		</tr>
		
		<tr><td width="100%" nowrap colspan="2" align=right>&nbsp;</td></tr>
			
		<tr id="trFull" name="trFull" style="display:none;">
			<td width="20%" nowrap>&nbsp;</td>
			<td width="80%" nowrap bgcolor="gainsboro">
				<table border="0" width="100%" cellpadding=2 cellspacing=2>
					<tr>
						<td width="20%" nowrap>Full Acceptance</td>
						<td width="80%" nowrap><input type = 'text' name ='txtFull' id = 'txtFull' size="25" value=0 readonly class="readonlyex"></td>
					</tr>
					<tr>
						<td width="20%" nowrap>Amount for Full Acceptance</td>
						<td width="80%" nowrap><input type = 'text' name ='txtFullAmount' id = 'txtFullAmount' size="25" value=0 readonly class="readonlyex"></td>
					</tr>
					<tr>
						<td width="20%" nowrap>Additional New Shares</td>
						<td width="80%" nowrap><input type = 'text' name ='txtNew' id = 'txtNew' size="25" value=0 onblur="UpdateFull();"></td>
					</tr>
					<tr>
						<td width="20%" nowrap>Amount for Additional New Shares</td>
						<td width="80%" nowrap><input type = 'text' name ='txtNewAmount' id = 'txtNewAmount' size="25" value=0 readonly class="readonlyex"></td>
					</tr>
					<tr>
						<td width="20%" nowrap>Total</td>
						<td width="80%" nowrap><input type = 'text' name ='txtTotal' id = 'txtTotal' size="25" value=0 readonly class="readonlyex"></td>
					</tr>
					<tr>
						<td width="20%" nowrap>Total Amount</td>
						<td width="80%" nowrap><input type = 'text' name ='txtTotalAmount' id = 'txtTotalAmount' size="25" value=0 readonly class="readonlyex"></td>
					</tr>
				</table>
			</td>
		</tr>
		
		<tr id="trPartial" name="trPartial" style="display:none;">
			<td width="20%" nowrap>&nbsp;</td>
			<td width="80%" nowrap bgcolor="gainsboro">
				<table border="0" width="100%" cellpadding=2 cellspacing=2>
					<tr>
						<td width="20%" nowrap>Partial Acceptance</td>
						<td width="80%" nowrap><input type = 'text' name ='txtPartial' id = 'txtPartial' size="25" value="<%=Additional%>" onblur="UpdatePartial();"></td>
					</tr>
					<tr>
						<td width="20%" nowrap>Amount for Partial Acceptance</td>
						<td width="80%" nowrap><input type = 'text' name ='txtPartialAmount' id = 'txtPartialAmount' size="25" value="<%=Additional*price%>"></td>
					</tr>
				</table>
			</td>
		</tr>
		
		<tr>
			<td width="100%" nowrap colspan="2" align=right>
				<BR>
				<BR>
				<BR>
				<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save ">
				<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();">
				&nbsp;&nbsp;
				<input type = 'hidden' name ='action' id = 'action' value="Execute">&nbsp;
				
				<input type = 'hidden' name ='txtRatio' id = 'txtRatio' value="<%=ratio%>">
				<input type = 'hidden' name ='ID' id = 'ID' value='<%=Offering_DPA_%>'>
				<%payable2 = Price * AlRights%>
			</td>
		</tr>
	</table>
	
	<script language="javascript">
		showOfferType();
		FullOrPartial2(<%=Accept%>,<%=OfferType%>);
		document.all.item("txtCdsNo").value = '<%=CDSNo2%>';
		
		var XMLHttpRequestObject = false;

		if (window.XMLHttpRequest)
		{
			XMLHttpRequestObject = new XMLHttpRequest();
		}
		else if (window.ActiveXObject)
		{
			XMLHttpRequestObject = new ActiveXObject("Microsoft.XMLHttp");
		}

		if (XMLHttpRequestObject)
			{	
			url = "GetHoldings.asp?cID=<%=client%>&sID=<%=ParentSecurity%>";

			XMLHttpRequestObject.open("GET",url);
			
			XMLHttpRequestObject.onreadystatechange = function()
				{
				if (XMLHttpRequestObject.readyState == 4 && XMLHttpRequestObject.status == 200)
					{
					returnStr = XMLHttpRequestObject.responseText;
					
					//var allot;
					//allot = returnStr * document.all.item("txtRatio").value;
					//allot = parseInt(allot,10)
					
					document.all.item("txtHoldings").value = returnStr;
					//document.all.item("txtAlloted").value = allot;
					//document.all.item("txtPayable").value = allot * document.all.item("txtPrice").value;
					}
				}
			}
		XMLHttpRequestObject.send(null);
		
		document.all.item("txtPayable").value = '<%=payable2%>';
	</script>
</form>

</body>
</html>