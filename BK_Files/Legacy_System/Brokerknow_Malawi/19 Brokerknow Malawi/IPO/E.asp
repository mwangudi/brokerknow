<html>
<head>
<title>Edit Safaricom IPO Forward</title>
 
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
	cdsCharge = trim(document.frmMain.elements("txtCDSCharge").value)
	
		   
	if (price > 0 and acrights > 0) then		   
		payable = price * acrights
	end if
	if document.frmMain.elements("CDSPaid").checked = true then
		payable = (price * acrights) + cdsCharge
	end if 
	
			
	document.frmMain.elements("txtPayable").value = Replace(FormatNumber(payable,0),",","")
	'document.frmMain.elements("txtPayable").value = FormatNumber(payable,0)
	
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
function PLoadClient(guidStr)
	{
	var clientcode = document.frmMain.elements("txtClientCode").value
	var clientCDSNO = document.frmMain.elements("txtCdsNo").value
	var clientcobo = document.getElementById("cboclient");		 
		 
	xmlhttp = createXMLHTTPObj();
				
	url="GetList.asp?clientcode="+clientcode+"&cdsno="+clientCDSNO+"&action=PLoadClient&guidStr=" + guidStr;
				
	xmlhttp.open("GET",url,true);

	xmlhttp.onreadystatechange=function() 
	{
		if (xmlhttp.readyState==4) 
		{
			returnStr = xmlhttp.responseText;
			returnStr = getBodyHTML(returnStr);			
									
									
			myArray = returnStr.split(";");
									
			document.frmMain.elements("txtClientCode").value = myArray[2];
			document.frmMain.elements("txtCdsNo").value = myArray[4]; 								
									
									
			clientcobo[0].Credit = myArray[1];
			clientcobo[0].AvailCredit = myArray[0];
			clientcobo[0].SearchCode = myArray[2];
			clientcobo[0].SearchText = myArray[3];
			clientcobo[0].SearchCDS = myArray[4];
									
			clientcobo[0].text = myArray[3];
			clientcobo[0].value = myArray[2];

			document.frmMain.elements("txtAvailableCredit").value = myArray[0];
			document.frmMain.elements("txtCurrentBal").value = myArray[1];								
			//document.frmMain.elements("txtAgentID").value = myArray[7];
		}
	}
				 
	xmlhttp.setRequestHeader('Accept','message/x-jl-formresult');
	xmlhttp.send();
	}

	function ConfirmSerial(theList)
		    {
			var i = 0;
			var SerialNo = theList.value;
			var Offering = document.frmMain.elements("cboOfferings").value
			if (SerialNo == document.frmMain.elements("hiddenPal").value)
			{
				//alert("hapa");
				return;
			}
			var randNum = Math.random();
			
			//alert('tafuta shida');
			
			frm = document.frmMain;				
			xmlhttp = createXMLHTTPObj();
			
			url="GetList.asp?SerialNo="+SerialNo+"&Offering="+Offering+"&action=ConfirmSerial&guidd=" + randNum;
			
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
					if(myArray[0]==1)
						{
						alert('Serial No ' + SerialNo + ' has being used by Client ' + myArray[1]);
						theList.value=""
						theList.focus();
						} 
				    }
				}
			xmlhttp.setRequestHeader('Accept','message/x-jl-formresult');
			xmlhttp.send(); 
		
		    }
		    

function ClearFields(thefield)
	{
		if(thefield==1)
		{
		document.frmMain.elements("txtCdsNo").value = "";
		}
		else
		{
		document.frmMain.elements("txtClientCode").value = "";
		}
	}
	
function UpdateBalances()
	{
	client = document.frmMain.elements("cboclient")
		
	document.frmMain.elements("txtAvailableCredit").value = client[client.selectedIndex].Credit;
	document.frmMain.elements("txtCurrentBal").value = client[client.selectedIndex].CurrentBal;	
	
	ClientID = document.all.item("cboclient").value;
	SecurityID = document.frmMain.elements("cboOfferings")[document.frmMain.elements("cboOfferings").selectedIndex].ParentSecurity;
	
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
	var minQty;
	var stepQty;

	price = drpOfferings[drpOfferings.selectedIndex].SearchPrice; 
	ratio = drpOfferings[drpOfferings.selectedIndex].Ratio; 
	minQty = drpOfferings[drpOfferings.selectedIndex].minQty; 
	stepQty = drpOfferings[drpOfferings.selectedIndex].stepQty; 
	
	document.frmMain.elements("txtprice").value = price ;
	document.frmMain.elements("txtRatio").value = ratio ;
	document.frmMain.elements("txtMinQty").value = minQty ;
	document.frmMain.elements("txtStepQty").value = stepQty ;
	}

function FullOrPartial(myOpt)
	{
	for (i=0;i<document.frmMain.optAcceptance.length;i++)
		{
		if (document.frmMain.optAcceptance[i].checked)
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
		}
	}

function ShowEFTRefund(myOpt)
	{
	if (myOpt == 1)
		{
		if (document.frmMain.optDividend(1).checked == true) { return false; }
		document.all.item("trEFT").style.display = 'none';
		}
	else
		{
		document.all.item("trEFT").style.display = '';
		}
	}

function ShowEFTDividend(myOpt)
	{
	if (myOpt == 1)
		{
		if (document.frmMain.optRefund(1).checked == true) { return false; }
		document.all.item("trEFT").style.display = 'none';
		}
	else
		{
		document.all.item("trEFT").style.display = '';
		}
	}

function ShowPayment(myOpt)
	{
	if (myOpt == 1)
		{
		document.all.item("trPayment").style.display = 'none';
		}
		
	if (myOpt == 2)
		{
		document.all.item("trPayment").style.display = 'none';
		}
		
	if (myOpt == 3)
		{
		document.all.item("trPayment").style.display = '';
		
		//===================================================================
		document.all.item("trrefund").style.display = '';
		document.all.item("trtax").style.display = '';
		document.all.item("trdividend").style.display = '';
		//===================================================================
		
		document.all.item("txtPaymentRef").style.backgroundColor = '#b0c4de';
		document.all.item("txtPaymentBankRef").style.backgroundColor = '#b0c4de';
		document.all.item("txtPaymentBranchRef").style.backgroundColor = '#b0c4de';
		document.all.item("txtPaymentSortCode").style.backgroundColor = '#b0c4de';
		document.all.item("txtPaymentAccountNo").style.backgroundColor = '#b0c4de';
		
		
		document.frmMain.elements("txtPaymentBankRef").disabled = false;
		document.frmMain.elements("txtPaymentBranchRef").disabled = false;
		document.frmMain.elements("txtPaymentSortCode").disabled = false;
		document.frmMain.elements("txtPaymentAccountNo").disabled = false;
		}
	
	if (myOpt == 4)
		{
		document.all.item("trPayment").style.display = '';
		
		document.all.item("txtPaymentRef").style.backgroundColor = '#b0c4de';
		document.all.item("txtPaymentBankRef").style.backgroundColor = '#b0c4de';
		document.all.item("txtPaymentBranchRef").style.backgroundColor = '#b0c4de';
		document.all.item("txtPaymentSortCode").style.backgroundColor = '#b0c4de';
		document.all.item("txtPaymentAccountNo").style.backgroundColor = '#b0c4de';
		}
	
	if (myOpt == 5)
		{
		document.all.item("trPayment").style.display = 'none';//edited = '';
		
		//===================================================================
		document.all.item("trrefund").style.display = 'none';
		document.all.item("trtax").style.display = 'none';
		document.all.item("trdividend").style.display = 'none';
		//===================================================================
		
		document.all.item("txtPaymentRef").style.backgroundColor = '#b0c4de';
		document.all.item("txtPaymentBankRef").style.backgroundColor = '#808080';
		document.all.item("txtPaymentBranchRef").style.backgroundColor = '#808080';
		document.all.item("txtPaymentSortCode").style.backgroundColor = '#808080';
		document.all.item("txtPaymentAccountNo").style.backgroundColor = '#808080';
		
		document.frmMain.elements("txtPaymentBankRef").disabled = true;
		document.frmMain.elements("txtPaymentBranchRef").disabled = true;
		document.frmMain.elements("txtPaymentSortCode").disabled = true;
		document.frmMain.elements("txtPaymentAccountNo").disabled = true;
		
		//document.frmMain.elements("cboPaymentType").selectedIndex = 1;
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

function showSMS()
	{
		if (document.frmMain.useOwner.checked == true) 
		{ 
		document.all.item("ownerSelectRow").style.display = '';
		}
		else
		{ 
		document.all.item("ownerSelectRow").style.display = 'none';
		}
	}
function AddCDSCharge(obj)
{
	var availableCredit,amountPayable,tempAmountPayable,price,qty,cdsCharge;
	availableCredit = document.frmMain.elements("txtAvailableCredit").value;
	
	amountPayable=document.frmMain.elements("txtPayable").value;
	tempAmountPayable = amountPayable.toString();
	
	tempAmountPayable=tempAmountPayable.replace(",","") 
	
	amountPayable = parseFloat(tempAmountPayable);
	price = parseFloat(document.frmMain.elements("txtPrice").value);
	qty = parseFloat(document.frmMain.elements("txtAlloted").value);
	cdsCharge = parseFloat(document.frmMain.elements("txtcdsCharge").value);
	
	if (obj.checked==false)
	{	
		
		
		document.frmMain.elements("txtPayable").value = (price* qty) ;
	}
	else
	{
		
		document.frmMain.elements("txtPayable").value = (price* qty) + cdsCharge;
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
theBatchSize = Request.Form("txtbatchSize")
theoffering = Request.Form("cboOfferings")

Set conn = GetActiveConnection("KBroker")

Select Case action
	case "EXECUTE"
	
	Dim buttonAction
	buttonAction = Trim(Ucase(Request.Form("cmdAdd")))	
	
	If instr(1,buttonAction,"SAVE") > 0 then
		OfferingType = Request.Form("cboOfferingType")
	
		palno = Request.Form("txtPalNo")
		clientName = Request.Form("txtclientcode")
		CurrentBal = Request.Form("txtCurrentBal")
		AvailableCredit = Request.Form("txtAvailableCredit")
		offering = Request.Form("cboOfferings")
		price = Request.Form("txtPrice")
		AlRights = Request.Form("txtAlloted")
		payable = Request.Form("txtPayable")
		cdaCode = trim(Request.Form("txtCDAID"))
		Accept = Request.Form("optAcceptance")
		Full = Request.Form("txtFull")''
		FullAmount = Request.Form("txtFullAmount")
		AddNew = Request.Form("txtNew")''
		NewAmount = Request.Form("txtNewAmount")
		Total = Request.Form("txtTotal")
		TotalAmount = Request.Form("txtTotalAmount")
		Partial = Request.Form("txtPartial")
		PartialAmount = Request.Form("txtPartialAmount")
		
		MinQuantity = Request.Form("txtminQty")
		StepQuantity = Request.Form("txtstepQty")
		OfferType = Request.Form("txtOfferType")
		
		Category = Request.Form("txtCategory")
		PaymentType = Request.Form("cboPaymentType")
		PaymentRef = Request.Form("txtPaymentRef")
		PaymentBankRef = Request.Form("txtPaymentBankRef")
		PaymentBranchRef = Request.Form("txtPaymentBranchRef")
		PaymentSortCode = Request.Form("txtPaymentSortCode")
		PaymentAccountNo = Request.Form("txtPaymentAccountNo")
	
		RefundMethod = Request.Form("optRefund")
		TaxExempt = Request.Form("optTaxExempt")
		DividendMethod = Request.Form("optDividend")
		EFTBankRef = Request.Form("txtEFTBankRef")
		EFTBranchRef = Request.Form("txtEFTBranchRef")
		EFTSortCode = Request.Form("txtEFTSortCode")
		EFTAccountNo = Request.Form("txtEFTAccountNo")
		maintain = Request.Form("chkMaintain")
		maintainBatch = Request.Form("txtmaintain")
		mobileno = Request.Form("txtMobile")
		paymentradio = Request.Form("paymentradio")
		sms = Request.Form("useOwner")
		CDSCharge = Replace(Request.Form("txtCDSCharge"),",","")
		
		cdsPaid = trim(request.Form("CDSPaid"))
		
		agentcode = Request.Form("txtAgentCode")
		
						
		
		if maintain = "" then
			maintain = 0
		end if
		
						
		if (trim(sms) = "") then
			sms = 0
		else
			sms = 1
		end if
						
		if (trim(paymentradio) = "") then
			paymentradio = "NULL"
		end If
		
		Select Case Accept
			Case 1
				''Full
				AlRights = AlRights
				payable = TotalAmount + 30
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
	
		If (RefundMethod = 2) Or (DividendMethod = 2) Then
				If (EFTBankRef = "") Or (EFTAccountNo = "") Then%>
				<script language = 'vbscript'>
					ShowMessage "Please enter the EFT Bank Details"
				</script>
				<% response.end
		    End If
		End If	
	
		If (PaymentType = 3) Or (PaymentType = 4) Then
				If (PaymentRef = "") Or (PaymentAccountNo = "") Then%>
				<script language = 'vbscript'>
					ShowMessage "Please enter the Payment Bank Details"
				</script>
				<% response.end
		    End If
		End If	
					
		if (sms = 1 and mobileno = "") then
			%>
		    <script language = 'vbscript'>
		    	ShowMessage "Please enter the Mobile No."
		    </script>
		    <% response.end
		end if
						
		if (sms = 1 and mobileno <> "" and len(trim(mobileno)) < 5) then
			%>
		    <script language = 'vbscript'>
		    	ShowMessage "Invalid Mobile No. entered."
		    </script>
		    <% response.end
		end if
		
		'dim acode,bcode 
		'acode = Session("AgentCode")
		'if acode = "" then
		'agentCode = Session("BrokerCode")
		'else
		'-----------------------------------------------------------------------------------------
		'sqlAgent = "SELECT SubAgentCode FROM SubAgent WHERE SubAgentCode = " & Session("AgentCode")
		
		 'Set AgentRs = conn.Execute(sqlAgent)
		 'agentCode = Session("AgentCode")
		 
        'end if
		'GET BATCH NO
		'------------------------------------------------------------------------------------------------------------------
		sqlmax = "SELECT COUNT(ISNULL(Offerings.Offering_DPA_, 0)) AS Counter, Offerings.Batch_No, Security.BatchSize " & _
				 " FROM Offerings INNER JOIN " & _
				 " Security ON Offerings.Offering = Security.Security_DPA_ " & _
				 " WHERE (Offering = " & offering & ") AND (Offerings.Deleted = 0) AND (CreatedBy = " & Session("UserID") & ") AND (Batch_No = " & _
				 " (SELECT MAX(Batch_No) AS Maximum " & _
				 " FROM Offerings " & _
				 " WHERE (Offering = " & offering & ") AND (CreatedBy = " & Session("UserID") & ") AND (Offerings.Forward <> 1) and " & _
				 " (deleted <> 1) and " & _
				 " (Cast(floor(cast(Offerings_Date as float)) as DateTime)='" & FormatDate(Date) & "') )) AND (Offerings.AgentCode = " & agentCode & ") AND (Offerings.Forward = 1) AND " & _
				 " (ISNULL(Offerings.BatchClosed,0) <> 1) AND (Offerings.OfferingType = "& OfferingType &")" & _
				 " GROUP BY Offerings.Batch_No, Security.BatchSize"


		sqlmax="SELECT     COUNT(ISNULL(Offerings.Offering_DPA_, 0)) AS Counter, Offerings.Batch_No, Security.BatchSize " & _
				 " FROM         Offerings INNER JOIN " & _
				 "                       Security ON Offerings.Offering = Security.Security_DPA_ " & _
				 " WHERE     (Offerings.Offering = " & offering & ") AND (Offerings.Deleted = 0) AND (Offerings.CreatedBy = " & Session("UserID") & ") AND (Offerings.Batch_No = " & _
				 "                           (SELECT     MAX(Batch_No) AS Maximum " & _
				 "                             FROM          Offerings " & _
				 "                             WHERE      (Offerings.PaymentType =  "& PaymentType &") AND (Offering = " & offering & ") AND (CreatedBy = " & Session("UserID") & ") AND (Offerings.Forward =1) AND (deleted <> 1) AND  " & _
				 "                                                    (Cast(floor(cast(Offerings_Date AS float)) AS DateTime) = '" & FormatDate(Date) & "'))) AND (Offerings.Forward = 1) AND  " & _
				 "                       (ISNULL(Offerings.BatchClosed, 0) <> 1) AND (Offerings.OfferingType = "& OfferingType &") " & _
				" GROUP BY Offerings.Batch_No, Security.BatchSize"
		
		'Response.Write sqlmax
		'Response.End
		
		Set MaxRs = conn.Execute(sqlmax)

		intRecs = maxRs.RecordCount
		
		if intRecs >  0 then
			Count = MaxRs("Counter")
			Batch = MaxRs("Batch_No")
			BatchSize = MaxRs("BatchSize")
		else
			Count = 0
			Batch = 0
			BatchSize = 0
		end if

		if len(Trim(Count)) = 0 then
			Count = 0
		end if

		if len(Trim(Batch)) = 0 then
			Batch = 0
		end if

		if len(Trim(BatchSize)) = 0 then
			quicksql = "SELECT BatchSize " & _ 
					   " FROM Security " & _
					   " WHERE (Security_DPA_ = " & offering & ")"
			quickBatchSize = conn.Execute(quicksql)

			BatchSize = quickBatchSize
		end if

		if Count = BatchSize then
			NextCount = 1

			sqlmax="Select isnull(Max(Batch_No),0)+1 as NextBatch from offerings where offering = " & offering & " and deleted <> 1"

			Set MaxRs = conn.Execute(sqlmax)
			NextNo = MaxRs("NextBatch")

			NextBatch = NextNo
		else
			NextCount = Count+1
			NextBatch = Batch
		end if
		'------------------------------------------------------------------------------------------------------------------
		
		
		If Accept = "" Then Accept = 1
		set guid = server.createobject("NDUtils.CGUID")
		guidStr = guid.GenerateGUID
				
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
					" ,RefundMethod = " & RefundMethod & "" & _
					" ,TaxExempt = " & TaxExempt & "" & _
					" ,DividendMethod = " & DividendMethod & "" & _
					" ,EFTBankRef = '" & EFTBankRef & "'" & _
					" ,EFTBranchRef = '" & EFTBranchRef & "'" & _
					" ,EFTSortCode = '" & EFTSortCode & "'" & _
					" ,EFTAccountNo = '" & EFTAccountNo & "'" & _
					" ,PaymentType = " & PaymentType & "" & _
					" ,PaymentRef = '" & PaymentRef & "'" & _
					" ,PaymentBankRef = '" & PaymentBankRef & "'" & _
					" ,PaymentBranchRef = '" & PaymentBranchRef & "'" & _
					" ,PaymentSortCode = '" & PaymentSortCode & "'" & _
					" ,PaymentAccountNo = '" & PaymentAccountNo & "'" & _
					" ,OfferingType = '" & OfferingType & "'" & _
					" ,SMS = " & sms & "" & _
					" ,ClientCellTel = '" & mobileno & "'" & _
					" ,PaymentMode = " & paymentradio & "" & _
					" ,Application_Status = " & status & "" & _
					" WHERE Offering_DPA_ = "& ID

						'Delete Application
                   
                   'get voucher number
                    sqlStr =" SELECT Offerings.PAL_No, Offerings.Client_DPA_, Offerings.Offering_Price, Offerings.Alloted_Rights, " & _
                            " Offerings.Offerings_Date, Security.SecurityName, Offerings.Receipt,Offerings.CitiAccepted, Security.BankAccount_DPA_," & _
                            " Offerings.Batch_No,isnull(Certificate,0) as Certificate,isnull(Offerings.CDSCharge,0) as CDSCharge " & _
                            " ,isnull(Security.CDSBankAccount_DPA_,0) as CDSBankAccount_DPA_ FROM Offerings INNER JOIN Security " & _
                            " ON Offerings.Offering = Security.Security_DPA_ " & _
                            " WHERE Offerings.Offering_DPA_ = " & ID	
        					  
			        'Response.write sqlStr
			        'Response.End
        			

			        Set rs = conn.Execute(sqlStr)

			        If   (rs.BOF Or rs.EOF) Then%>
					        <script language = 'vbscript'>
							        ShowMessage "Serious error. The Application cannot be retrieved for deletion"
							        window.self.close
					        </script>
					        <%response.end
			        End If
        					
			        if isnull(rs.fields("Receipt")) then
					         Payment = 0
			        else
					        Payment = rs.fields("Receipt")
					        'voucherType = 3
			        end if
                           
                   
                   if (cdsPaid = "0") then
                   
					    'obtain values for journal

					    clientID = rs.fields("Client_DPA_")
					    oBank = rs.fields("BankAccount_DPA_")
					    oDate = FormatDate(rs.fields("Offerings_Date"))
					    oQty = trim(rs.fields("Alloted_Rights"))
					    oPrice = trim(rs.fields("Offering_Price"))
					    oSecurity = iif(Len(trim(rs.fields("SecurityName")))>20,Left(trim(rs.fields("SecurityName")),20) & "...",trim(rs.fields("SecurityName")))
					    oParticulars = FormatNumEx(oQty,0) & " " & oSecurity & " @" & FormatNumEx(oPrice,2) & " [ " & rs("Batch_No") & "]"
					    IssueCertificate = trim(rs.fields("Certificate"))
					    CDSCharge = trim(rs.fields("CDSCharge"))
					    CDSBank = trim(rs.fields("CDSBankAccount_DPA_"))

					    if cdbl(AvailableCredit) < cdbl(CDSCharge) then
						    %>
						    <script language = 'vbscript'>
								    ShowMessage "The Client has insufficient funds to cater for the CDS charge"
								    window.self.close
						    </script>
						    <%response.end
					    end if
					
					    if(IssueCertificate=true) then
						    TheCertificate = 1
					    else
						    TheCertificate = 0
					    end if

					    procStr = "@userID = " & Session("UserID") & ",@clientDPA = " & clientID & ",@offerBank = " & oBank & ",@jDate = '" & oDate & "',@offerQty = " & oQty & ",@offerPrice = " & oPrice & ",@JournalNarrative = '" & oParticulars & "',@Certificate = " & TheCertificate & ",@CDSCharge=" & CDSCharge & ",@CDSBank=" & CDSBank
    					
					    'response.write procStr :response.end
					    Conn.execute("DeleteForward " & procStr)
					    Conn.execute(sqlStr)
					    
			        end if
                   
                   
                   'checks whether downloaded....
	                '------------------------------------------------------------------------------------------------------
	                sqlchk  = " SELECT Downloaded " & _
	                          " FROM Offerings " & _
	                          " WHERE (Offering = "& offering &")"
	                            dim rsdownload,downloadvalue
				                set rsdownload = conn.execute(sqlchk)
	                            if NOT(rsdownload.EOF OR rsdownload.EOF)then
	                                downloadvalue = rsdownload("Downloaded")
				                end if
				                'response.Write downloadvalue: response.End 
                    '------------------------------------------------------------------------------------------------------
		            if downloadvalue <> "" then		
                        SQL="UPDATE Offerings SET  Deleted = 1, Application_Status = 2  WHERE (Offering_DPA_ = " & ID &")"
					else
					    SQL="UPDATE Offerings SET  Deleted = 1, Application_Status = 0  WHERE (Offering_DPA_ = " & ID &")"
                    end if
						conn.Execute(SQL)
						Certificate=0



				'-------------------------------------
		
				'Check PALNo status
					
					sqlpal=" SELECT  PAL_No " & _
						   " FROM Offerings " & _
						   " WHERE (RTRIM(LTRIM(PAL_No)) = N'"& palno&"') AND (Offering = "& offering &")"
						'response.write sqlpal:response.end
					dim rsPAL,status
					set rsPAL = conn.execute(sqlpal)

					if not(rsPAL.EOF or rsPAL.EOF)then

						status = 3 'Modified
					else
						status = 1 'New Application
					end if

					set rsPAL = nothing
								
				'-------------------------------------
				
				
				if maintain = 1 then
					
					sqlStr = "INSERT INTO Offerings (Offering_EIT_, PAL_No,Client_DPA_, " & _
					 " Offering,Offering_Price,Alloted_Rights,ChangedBy,AcceptanceType,Additional,Batch_No,BatchSeq," & _
					 " Forward,TimeCreated,CreatedBy,TimeChanged," & _
					 " RefundMethod,TaxExempt,EFTBankRef,EFTBranchRef,EFTSortCode,EFTAccountNo," & _
					 " Category,PaymentType,PaymentRef,PaymentBankRef,PaymentBranchRef,PaymentSortCode,PaymentAccountNo," & _
					 " OfferingType,SMS,ClientCellTel,PaymentMode,Branch_DPA_,ExtraQuantity,BatchPaymentMode,Certificate,CDSCharge,CDA_ID,AgentCode,Application_Status) " & _
					 " VALUES (#"& guidStr &"#,'" & palno & "'," & ClientName & "," & offering & "," & Price & _
					 "," & AlRights & "," & UserId & "," & Accept & ",0," & maintainBatch & ","& NextCount & _
					 ",1,GetDate()," & UserID & ",GetDate()"& _
					 "," & RefundMethod & "," & TaxExempt & ",'" & EFTBankRef & "','" & EFTBranchRef & "'" & _
					 ",'" & EFTSortCode & "','" & EFTAccountNo & "'" & _
					 "," & Category & "," & PaymentType & ",'" & PaymentRef & "','" & PaymentBankRef & "','" & PaymentBranchRef & "'" & _
					 ",'" & PaymentSortCode & "','" & PaymentAccountNo & "',"& OfferingType &","& sms &",'" & mobileno & "'" & _
					 ","& paymentradio &",1," & AddNew & ","& PaymentType &"," & Certificate & "," & CDSCharge & ",'" & cdacode & "','" & agentcode & "', " & status &")"
				else
					
					sqlStr = "INSERT INTO Offerings (Offering_EIT_, PAL_No,Client_DPA_, " & _
					 " Offering,Offering_Price,Alloted_Rights,ChangedBy,AcceptanceType,Additional,Batch_No,BatchSeq," & _
					 " Forward,TimeCreated,CreatedBy,TimeChanged," & _
					 " RefundMethod,TaxExempt,EFTBankRef,EFTBranchRef,EFTSortCode,EFTAccountNo," & _
					 " Category,PaymentType,PaymentRef,PaymentBankRef,PaymentBranchRef,PaymentSortCode,PaymentAccountNo," & _
					 " OfferingType,SMS,ClientCellTel,PaymentMode,Branch_DPA_,ExtraQuantity,BatchPaymentMode,Certificate,CDSCharge,CDA_ID,AgentCode,Application_Status) " & _
					 " VALUES (#"& guidStr &"#,'" & palno & "'," & ClientName & "," & offering & "," & Price & _
					 "," & AlRights & "," & UserId & "," & Accept & ",0," & NextBatch & ","& NextCount & _
					 ",1,GetDate()," & UserID &",GetDate()" & _
					 "," & RefundMethod & "," & TaxExempt & ",'" & EFTBankRef & "','" & EFTBranchRef & "'" & _
					 ",'" & EFTSortCode & "','" & EFTAccountNo & "'" & _
					 "," & Category & "," & PaymentType & ",'" & PaymentRef & "','" & PaymentBankRef & "','" & PaymentBranchRef & "'" & _
					 ",'" & PaymentSortCode & "','" & PaymentAccountNo & "',"& OfferingType &","& sms &",'" & mobileno & "'" & _
					 ","& paymentradio &",1," & AddNew & "," & PaymentType & "," & Certificate & "," & CDSCharge & ",'" & cdacode & "','" & agentcode & "'," & status &")"
				end if

		
			'Response.Write sqlStr
			'Response.End
			sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))							
			
			 
																				                                                     
			conn.Execute sqlStr																			
		conn.CommitTrans
		
		If instr(1,buttonAction,"SAVE & PRINT") > 0 Then
				sqlStr = "Select Offering_DPA_ From Offerings Where Offering_DPA_ = " & ID
				set rst = conn.execute(sqlStr)
				if not (rst.eof or rst.bof) then offeringID = rst("Offering_DPA_")
				%>
				<SCRIPT LANGUAGE="JAVASCRIPT">
					window.parent.parent.frames['maininfo'].location.reload();
				</SCRIPT>
				<%
				WriteDialogRelocateScript "OfferingForm.asp?frompage=1&ID=" & offeringID
			Else
				WritefraEnabledDialogCloseScript
			End If
			WritefraEnabledDialogCloseScript
		End If
		
		If instr(1,buttonAction,"PRINT BATCH") > 0 then
			Set conn = GetActiveConnection("KBroker")

			sqlStr = "SELECT MAX(Batch_No) AS MaxBatchNo " & _
					 " FROM (SELECT Offerings.Batch_No " & _
					 " FROM Offerings INNER JOIN  " & _
					 " Security ON Offerings.Offering = Security.Security_DPA_ " & _
					 " WHERE (Offerings.ChangedBy = "& UserId &") AND (Offerings.Offering = "& theoffering &") AND (Offerings.Deleted = 0) AND (Offerings.Forward = 1) " & _
					 " GROUP BY Offerings.Batch_No " & _
					 " HAVING (COUNT(Offerings.Offering_DPA_) = "& theBatchSize &") " & _
					 ") A"

			Set Rs = conn.execute(sqlStr)
			If Not(Rs.EOF Or Rs.BOF) Then
				If Not IsNumeric(Rs("MaxBatchNo")) Then
					BatchNumb = 0
					%>
					<script language = 'vbscript'>
					    ShowMessage "No FULL batch currently available for printing."
					</script>
					<% Response.end
				Else
					BatchNumb = Rs("MaxBatchNo")
				End If
			Else
				BatchNumb = 0
			End If
			
			BatchIDs = BatchNumb & "<->" & theoffering
			
			WriteDialogRelocateScript "../Reports/BatchedApplications.asp?From="&BatchIDs
		End If
		
		conn.Close
		Set conn = Nothing

		Response.End	
End Select

sqlStr = "SELECT * FROM Offerings WHERE Offering_DPA_= " & ID
sqlstr="SELECT     Client.ClientName, Client.ClientCDSNo, ISNULL(ClientBalances.CurrentBal, 0) + ISNULL(Client.CreditLimit, 0) - ISNULL(ClientTotal.Total, 0)  " & _
		 "                       AS AvailableCredit, ISNULL(ClientBalances.CurrentBal, 0) AS CurrentBal, Offerings.Offering_DPA_, Offerings.PAL_No, Offerings.Client_DPA_,  " & _
		 "                       Offerings.PaymentBankRef, Offerings.PaymentBranchRef, Offerings.PaymentSortCode, Offerings.Offering, Offerings.Offering_Price,  " & _
		 "                       Offerings.Alloted_Rights, Offerings.Batch_No, Offerings.Forward, Offerings.Offerings_Date, Offerings.Deleted, Offerings.Downloaded,  " & _
		 "                       Offerings.CreatedBy, Offerings.Branch_DPA_, Offerings.Agent_DPA_, Offerings.sms, Offerings.ClientCellTel, Offerings.AcceptanceType,  " & _
		 "                       Offerings.RefundMethod, Offerings.TaxExempt, Offerings.DividendMethod, Offerings.EFTBankRef, Offerings.EFTBranchRef, Offerings.EFTSortCode,  " & _
		 "                       Offerings.PaymentMode, Offerings.Additional, Offerings.EFTAccountNo, Offerings.Category, Offerings.PaymentType, Offerings.PaymentRef,  " & _
		 "                       Offerings.PaymentAccountNo, Offerings.OfferingType, Offerings.Certificate, Offerings.CDSCharge, ISNULL(Offerings.Hold, 0) AS Hold,  " & _
		 "                       Offerings.CDACode, Offerings.AgentCode, SubAgent.SubAgentName AS AgentName, Offerings.OfferCheque, Offerings.OfferBank  " & _
		 " FROM         ClientTotal RIGHT OUTER JOIN " & _
		 "                       ClientBalances RIGHT OUTER JOIN " & _
		 "                       SubAgent RIGHT OUTER JOIN " & _
		 "                       Offerings ON SubAgent.SubAgentCode = Offerings.AgentCode ON ClientBalances.client_DPA_ = Offerings.Client_DPA_ ON  " & _
		 "                       ClientTotal.Client_DPA_ = Offerings.Client_DPA_ LEFT OUTER JOIN " & _
		 "                       Client ON Offerings.Client_DPA_ = Client.Client_DPA_ " & _
		 " WHERE     (Offerings.Offering_DPA_ = " & ID & ")"

sqlstr="SELECT     Client.ClientName, Client.ClientCDSNo, ISNULL(ClientBalances.CurrentBal, 0) + ISNULL(Client.CreditLimit, 0) - ISNULL(ClientTotal.Total, 0)  " & _
		 "                       AS AvailableCredit, ISNULL(ClientBalances.CurrentBal, 0) AS CurrentBal, Offerings.Offering_DPA_, Offerings.PAL_No, Offerings.Client_DPA_,  " & _
		 "                       Offerings.PaymentBankRef, Offerings.PaymentBranchRef, Offerings.PaymentSortCode, Offerings.Offering, Offerings.Offering_Price,  " & _
		 "                       Offerings.Alloted_Rights, Offerings.Batch_No, Offerings.Forward, Offerings.Offerings_Date, Offerings.Deleted, Offerings.Downloaded,  " & _
		 "                       Offerings.CreatedBy, Offerings.Branch_DPA_, Offerings.Agent_DPA_, Offerings.sms, Offerings.ClientCellTel, Offerings.AcceptanceType,  " & _
		 "                       Offerings.RefundMethod, Offerings.TaxExempt, Offerings.DividendMethod, Offerings.EFTBankRef, Offerings.EFTBranchRef, Offerings.EFTSortCode,  " & _
		 "                       Offerings.PaymentMode, Offerings.Additional, Offerings.EFTAccountNo, Offerings.Category, Offerings.PaymentType, Offerings.PaymentRef,  " & _
		 "                       Offerings.PaymentAccountNo, Offerings.OfferingType, Offerings.Certificate, Offerings.CDSCharge," & _
		 "                       Offerings.CDAcode, Offerings.AgentCode, Offerings.OfferCheque, Offerings.OfferBank  " & _
		 " FROM         ClientTotal RIGHT OUTER JOIN " & _
		 "                       ClientBalances RIGHT OUTER JOIN " & _
		 "                       Offerings  ON ClientBalances.client_DPA_ = Offerings.Client_DPA_ ON  " & _
		 "                       ClientTotal.Client_DPA_ = Offerings.Client_DPA_ LEFT OUTER JOIN " & _
		 "                       Client ON Offerings.Client_DPA_ = Client.Client_DPA_ " & _
		 " WHERE     (Offerings.Offering_DPA_ = " & ID & ") "
'response.write " Hapa    " & offering & "hapa" :response.end
Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        
If rs.EOF Or rs.BOF Then
	%>
	<script language = 'vbscript'>
		window.self.ShowMessage "The selected record cannot be retrieved for editing"
	</script>
	<%
	Response.End
'ElseIf IsNumeric(rs("Batch_No")) Then
'ElseIf rs("BatchClosed") = 1 Then
	'WriteDialogRefuseOpenScript
	%>
	<script language = 'vbscript'>
		'window.parent.dialogArguments.opener.alert "The selected record cannot be retrieved for editing" & Chr(13) & "as 'it has already been batched."
	</script>
	<%
	'Response.End
Else
	OfferingType = RS("OfferingType")
	Offering_DPA_ = rs("Offering_DPA_")
	palno = rs("PAL_No")
	client = rs("Client_DPA_")
	price = rs("Offering_Price")
	offering = Trim(rs("Offering"))
	maintainBatch = rs("Batch_No")
	
	payable = rs("Offering")*rs("Offering_Price")
	
	chkno = rs("OfferCheque")
	narrative = rs("OfferBank")
	
	AlRights = cdbl(rs("Alloted_Rights"))
	Accept = rs("AcceptanceType")
	
	Additional = rs("Additional")
	
	RefundMethod = rs("RefundMethod")
	TaxExempt = rs("TaxExempt")
	DividendMethod = rs("DividendMethod")
	EFTBankRef = rs("EFTBankRef")
	EFTBranchRef = rs("EFTBranchRef")
	EFTSortCode = rs("EFTSortCode")
	EFTAccountNo = rs("EFTAccountNo")
	cdaID= rs("CDAcode")
	Category = rs("Category")
	PaymentType = rs("PaymentType")
	PaymentRef = rs("PaymentRef")
	PaymentBankRef = rs("PaymentBankRef")
	PaymentBranchRef = rs("PaymentBranchRef")
	PaymentSortCode = rs("PaymentSortCode")
	PaymentAccountNo = rs("PaymentAccountNo")
	
	SMS = rs("SMS")
	ClientCellTel = rs("ClientCellTel")
	PaymentMode = rs("PaymentMode")
	CDSNo = rs("ClientCDSNo")
	ClientName = rs("ClientName")
	
	AgentCode = rs("AgentCode")
	'AgentName = rs("AgentName")

End If 
'response.write " Hapa    " & offering & "hapa" :response.end
%>

<form name = 'frmMain' method = 'post' id="frmMain" action = "EditSafForward.asp" >
	<table border="0" width="80%" cellpadding=2 cellspacing=2>
		<tr>
			<td width="20%" nowrap>Type</td>
			<td width="80%" nowrap>
				<select name = 'cboOfferingType' id = 'cboOfferingType' size="1" onchange="EditCategory(this.value);">
				         <% Select case offeringtype %>
				         <%case 1%>
						<option value = "1" <%if offeringtype= 1 then Response.Write "selected"%>>Standard Application</option>
						 <%case 3 %>
						<option value = "3" <%if offeringtype= 3 then Response.Write "selected"%>>QII's</option>
						<%  End Select%>
				</select>
			</td>
		</tr>		
		<tr>
			<td width="20%" nowrap>PAL NO</td>
			<INPUT TYPE="hidden" NAME="hiddenPal" ID="hiddenPal" value="<%=palno%>">
			<td width="80%" nowrap><input type="text" name="txtPalNo" id="txtPalNo" size="25" onblur="ConfirmSerial(this);" value="<%=palno%>"></td>
		</tr>
				
		<tr>
			<td width="20%" nowrap>Client</td>
			<td width="80%" nowrap>
				<!--<input class="Readonly" Readonly="true" name ='txtClientCode' id = 'txtClientCode' value="<%=client%>" size="10" onBlur="" onChange="">&nbsp;
				<input  class="Readonly" Readonly="true" name ='txtCdsNo' id = 'txtCdsNo' size="16" onBlur="" onChange=";">&nbsp;

				<select name = 'cboAccount' id = "cboAccount" size="1">
				<option SearchCode = "" SearchText = "" value = ''>Load Account</option>
				</select>-->
				 <table>
					<tr>
						<td>Client Code</td>
						<td>CDS NO</td>
						<td>Client Name</td>
						<td>Names</td>
					</tr>
					
					<tr>
						<td>
							<input  name ='txtClientCode' class="Readonly" Readonly="true" id = 'txtClientCode' size="10"  value="<%=client%>">
						</td>
						<td>
							<input  name ='txtCdsNo' id = 'txtCdsNo' class="Readonly" Readonly= size="16"  value="<%=CDSNo%>">
						</td>
						<td>
							<input type = 'text' name ='txtClientName' id = 'txtClientName' class="Readonly" Readonly="true" size="15" onBlur="" value="<%=ClientName%>">
						</td>
						<td id="comboclient" name="comboclient">
							<select name = 'cboClient' id = 'cboClient' size="1" readonly>					                  
								<option AgentReturnable= "" OrderContact = "" Iscustodian = "" AgentID = "" Agent = "" OwnerID = "" Owner = "" Credit="" CurrentBal="" SearchCode = "" SearchText = "" SearchCds = "" value = ''>Load Client</option>				
							</select>
						</td>
					</tr>
				</table>
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
						<td><input  name ='txtCurrentBal' id = 'txtCurrentBal' readonly class="readonly" size="15" value="<%=FormatNumber(AvailableCredit2,2)%>"></td>
						<td><input  name ='txtAvailableCredit' id = 'txtAvailableCredit' readonly class="readonly" size="15" value="<%=FormatNumber(CreditBal2,2)%>"></td>
					</tr>
				</table>
			</td>	         
		</tr>		
		<tr id="ownerSelectRow" align="right">	    
			<td width="20%" nowrap>&nbsp;</td>
			<td width="80%" nowrap>
				<table width="100%" border=0 cellspacing=2 cellpadding=2>
					<tr>
						<td width="20%" nowrap>Mobile No.</td>
						<td width="80%" nowrap><input type='text' name='txtMobile' id = 'txtMobile' size="15" value="<%=ClientCellTel%>"></td>
					</tr>					
				</table>
			</td>		
		</tr>	
		<tr >
			<td width="20%" nowrap>Agent Code</td>
			<td width="30%" height="22">
			   <input type="text" id="txtAgentCode" name="txtAgentCode" size="5" value = "<%=AgentCode%>" onBlur="LoadAgent();">&nbsp;&nbsp;Agent Name
			   <input type="text" readonly id="txtAgentName" name="txtAgentName" value = "<%=AgentName%>">
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
			<td width="20%" nowrap>Client Category</td>
			<td width="80%" nowrap>
			<input type = 'radio' disabled name ='optCategory' id = 'optCategory' value=1 >&nbsp;Individual
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<input type = 'radio' disabled name ='optCategory' id = 'optCategory' value=2>&nbsp;Joint
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<input type = 'radio' disabled name ='optCategory' id = 'optCategory' value=3>&nbsp;Corporate
			</td>
		</tr>
		
		<tr>
			<td width="20%" nowrap>CDS Paid</td>
			<td width="80%" nowrap>

				<input type = 'checkbox'<% =response.write(iif(cdsPaid=1,"Checked","") )%> name ='CDSPaid' id = 'CDSPaid' value=1 onCLick="AddCDSCharge(this);">			
			</td>
		</tr>
		<tr>
			<td width="20%" nowrap>Payment Type</td>
			<td width="80%" nowrap>
				<select name = 'cboPaymentType' id = 'cboPaymentType' size="1" onChange="ShowPayment(this.value);">
					<!--  <option value = "1">Credit</option>
					<option value = "2">Cash</option>-->
					<option selected value = "3">Cheque</option>
					<!-- <option value = "4">Financed</option>
					<option value = "5">Institutional Investor</option>-->
					<option value = "1">Guarantee</option>
				</select>
			</td>
		</tr>
				
		<tr id="trPayment" name="trPayment" style="display:none;">
			<td width="20%" nowrap>&nbsp;</td>
			<td width="80%" nowrap>
				<table border="0" width="100%" cellpadding=2 cellspacing=2>
					<tr>
						<td width="10%" nowrap>Cheque Number</td>
						<td width="90%" nowrap><input type = 'text' name ='txtPaymentRef' id = 'txtPaymentRef' size="20" value="<%=PaymentRef%>"></td>
					</tr>
					<tr>
						<td width="10%" nowrap>Bank Ref</td>
						<td width="90%" nowrap><input type = 'text' name ='txtPaymentBankRef' id = 'txtPaymentBankRef' size="20" value="<%=PaymentBankRef%>"></td>
					</tr>
					<tr>
						<td width="10%" nowrap>Branch Ref</td>
						<td width="90%" nowrap><input type = 'text' name ='txtPaymentBranchRef' id = 'txtPaymentBranchRef' size="20" value="<%=PaymentBranchRef%>"></td>
					</tr>
					<tr>
						<td width="10%" nowrap>Sort Code</td>
						<td width="90%" nowrap><input type = 'text' name ='txtPaymentSortCode' id = 'txtPaymentSortCode' size="20" value="<%=PaymentSortCode%>"></td>
					</tr>
					<tr>
						<td width="10%" nowrap>Account No</td>
						<td width="90%" nowrap><input type = 'text' name ='txtPaymentAccountNo' id = 'txtPaymentAccountNo' size="20" value="<%=PaymentAccountNo%>"></td>
					</tr>
				</table>
			</td>
		</tr>		
		<tr id ="trrefund"name = "trrefund">
			<td width="20%" nowrap>Refund Method</td>
			<td width="80%" nowrap>
			<input type = 'radio' name ='optRefund' id = 'optRefund' value=1 onclick="ShowEFTRefund(this.value);">&nbsp;Cheque
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<input type = 'radio' name ='optRefund' id = 'optRefund' value=2 onclick="ShowEFTRefund(this.value);">&nbsp;EFT
			</td>
		</tr>
		
		<tr id = "trtax"name = "trtax">
			<td width="20%" nowrap>Tax Exempt</td>
			<td width="80%" nowrap>
			<input type = 'radio' name ='optTaxExempt' id = 'optTaxExempt' value=1>&nbsp;No
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<input type = 'radio' name ='optTaxExempt' id = 'optTaxExempt' value=2>&nbsp;Yes
			</td>
		</tr>
		
		<tr id = "trdividend"name = "trdividend">
			<td width="20%" nowrap>Dividend Method</td>
			<td width="80%" nowrap>
			<input type = 'radio' name ='optDividend' id = 'optDividend' value=1 onclick="ShowEFTDividend(this.value);">&nbsp;Cheque
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<input type = 'radio' name ='optDividend' id = 'optDividend' value=2 onclick="ShowEFTDividend(this.value);">&nbsp;EFT
			</td>
		</tr>
		
		<tr id="trEFT" name="trEFT" style="display:none;">
			<td width="20%" nowrap>&nbsp;</td>
			<td width="80%" nowrap>
				<table border="0" width="100%" cellpadding=2 cellspacing=2>
					<tr>
						<td width="10%" nowrap>Bank Ref</td>
						<td width="90%" nowrap><input type = 'text' name ='txtEFTBankRef' id = 'txtEFTBankRef' size="20" value="<%=EFTBankRef%>"></td>
					</tr>
					<tr>
						<td width="10%" nowrap>Branch Ref</td>
						<td width="90%" nowrap><input type = 'text' name ='txtEFTBranchRef' id = 'txtEFTBranchRef' size="20" value="<%=EFTBranchRef%>"></td>
					</tr>
					<tr>
						<td width="10%" nowrap>Sort Code</td>
						<td width="90%" nowrap><input type = 'text' name ='txtEFTSortCode' id = 'txtEFTSortCode' size="20" value="<%=EFTSortCode%>"></td>
					</tr>
					<tr>
						<td width="10%" nowrap>Account No</td>
						<td width="90%" nowrap><input type = 'text' name ='txtEFTAccountNo' id = 'txtEFTAccountNo' size="20" value="<%=EFTAccountNo%>"></td>
					</tr>
				</table>
			</td>
		</tr>
		
		<tr id="trAcceptance" name="trAcceptance" style="display:none;">
			<td width="20%" nowrap>Acceptance</td>
			<td width="80%" nowrap>
			Full&nbsp;<input type = 'radio' name ='optAcceptance' id = 'optAcceptance' value=1 onclick="FullOrPartial(this.value);">
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			Partial&nbsp;<input type = 'radio' name ='optAcceptance' id = 'optAcceptance' value=2 onclick="FullOrPartial(this.value);">
			</td>
		</tr>		
		<tr>
			<td width="100%" nowrap colspan="2" align=right>&nbsp;</td>
		</tr>			
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
				<font size="2"><b>Maintain Application Batch</font>&nbsp;&nbsp;&nbsp;&nbsp;<input type = 'checkbox' Class=checkbox name ='chkMaintain' id = 'chkMaintain' value="1" >
				<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdds' value=" Save " OnClick='frmMain.elements("cmdAdds").style.display="none"'>
				<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAddPrint' value=" Save & Print " OnClick='frmMain.elements("cmdAddPrint").style.display="none"'>
				<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAddPrintBatch' value=" Print Batch ">
				<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();">
				&nbsp;&nbsp;
				<input type = 'hidden' name ='action' id = 'action' value="Execute">&nbsp;
				
				<input type = 'hidden' name ='txtRatio' id = 'txtRatio' value="<%=ratio%>">
				<input type = 'hidden' name ='txtRatio' id = 'txtRatio' value="<%=ratio%>">
				<input type = 'hidden' name ='txtMinQty' id = 'txtMinQty' value='<%=minQty%>'>
				<input type = 'hidden' name ='txtStepQty' id = 'txtStepQty' value='<%=Quantity%>'>
				<input type = 'hidden' name ='txtBatchSize' id = 'txtBatchSize' value='<%=batchSize%>'>
				<input type = 'hidden' name ='txtOfferType' id = 'txtOfferType' value='<%=offerType%>'>
				<input type = 'hidden' name ='txtmaintain' id = 'txtmaintain' value='<%=maintainBatch%>'>
				<input type = 'hidden' name ='txtCDAID' id = 'txtCDAID' value='<%=cdaID%>'>
				<input type = 'hidden' name ='txtCDSCharge' id = 'txtCDSCharge' value='<%=TheCDSCharge%>'>	
				<input type = 'hidden' name ='txtCategory' id = 'txtCategory' value='<%=Category%>'>
				<input type = 'hidden' name ='ID' id = 'ID' value='<%=Offering_DPA_%>'>
				<%payable2 = Price * AlRights%>
				<input type = 'hidden' name ='txtCDSPaidStatus' id = 'txtCDSPaidStatus' value='<%=CDSPaid%>'>
			</td>
		</tr>
	</table>
	
	<script language="javascript">
		showOfferType();
		FullOrPartial(<%=Accept%>);
		document.all.item("txtCdsNo").value = '<%=CDSNo%>';
		
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
	
	<script language="javascript">
		var payList = document.frmMain.elements("cboPaymentType");
		for (i=0; i < payList.options.length; i++)
			{
			if (payList.options(i).value == <%=PaymentType%>)
				{
				payList.options(i).selected = true;
				}
			}
	</script>
	
	<%
	If Category = 1 Then
		%><script language="javascript">document.frmMain.optCategory(0).checked = true</script><%
	ElseIf Category = 2 Then
		%><script language="javascript">document.frmMain.optCategory(1).checked = true</script><%
	ElseIf Category = 3 Then
		%><script language="javascript">document.frmMain.optCategory(2).checked = true</script><%
	End If
	
	If PaymentType = 1 Then
		%><script language="javascript">ShowPayment(1);</script><%
	ElseIf PaymentType = 2 Then
		%><script language="javascript">ShowPayment(2);</script><%
	ElseIf PaymentType = 3 Then
		%><script language="javascript">ShowPayment(3);</script><%
	ElseIf PaymentType = 4 Then
		%><script language="javascript">ShowPayment(4);</script><%
	ElseIf PaymentType = 5 Then
		%><script language="javascript">ShowPayment(5);</script><%
	End If
	
	If RefundMethod = 1 Then
		%><script language="javascript">document.frmMain.optRefund(0).checked = true; ShowEFTRefund(1);</script><%
	ElseIf RefundMethod = 2 Then
		%><script language="javascript">document.frmMain.optRefund(1).checked = true; ShowEFTRefund(2);</script><%
	End If
	
	If TaxExempt = 1 Then
		%><script language="javascript">document.frmMain.optTaxExempt(0).checked = true</script><%
	ElseIf TaxExempt = 2 Then
		%><script language="javascript">document.frmMain.optTaxExempt(1).checked = true</script><%
	End If
	
	If DividendMethod = 1 Then
		%><script language="javascript">document.frmMain.optDividend(0).checked = true; ShowEFTDividend(1);</script><%
	ElseIf DividendMethod = 2 Then
		%><script language="javascript">document.frmMain.optDividend(1).checked = true; ShowEFTDividend(2);</script><%
	End If
	
	%>
</form>
</body>
</html>

