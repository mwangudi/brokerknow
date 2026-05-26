<html>
<head>
<title>Add Order</title>

<!--#include file="../libroutines.asp"-->

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>

<%theDate = Date + 30 %>

<SCRIPT language="JavaScript">
	var cal=new ctlSpiffyCalendarBox("cal", "frmAddOrder", "txtDate","cmdDate","<%=FormatDate(Date)%>",1);
	var calValidity=new ctlSpiffyCalendarBox("calValidity", "frmAddOrder", "txtValidity","cmdValDate","<%=FormatDate(theDate)%>",1);
	var calReleaseDate=new ctlSpiffyCalendarBox("calReleaseDate", "frmAddOrder", "txtReleaseDate","cmdReleaseDate","<%=FormatDate(Date)%>",1);
</SCRIPT>
<!--END CALENDAR -->

<SCRIPT>
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
		
		function UpdateTitles()
		{
			var securitytype = document.getElementById("cboOrderType").value ;
			var price = document.getElementById("cboprice").value ;
		         
			if (securitytype == 1 && price ==  "BF")
				{
					document.getElementById("Quantity").innerText = "Limit" ;
					//document.getElementById("QuantityItem").style.display = "none" ;
					//document.getElementById("LimitItem").style.display = "" ;
					//document.getElementById("LimitItem").value = ""
				}
			else
				{
					document.getElementById("Quantity").innerText = "Quantity" ;
					//document.getElementById("QuantityItem").style.display = "" ;
					//document.getElementById("LimitItem").style.display = "none" ;
				}
			 	
			if (securitytype==2)
				{
					document.all.item("cboPayOptions").style.display = "";
					document.all.item("lbSale").style.display = "";
				}
			else
				{
					document.all.item("cboPayOptions").style.display = "none";
					document.all.item("lbSale").style.display = "none";
					document.all.item("txtPayAmt").style.display = "none";
					document.all.item("lbPartial").style.display = "none";
				} 
		}

		function ShowSecurityHoldings()
		{
			var security = document.getElementById("cboOrderSecType").value
					 
			if (security==2)
				{
					//document.frmMain.elements("txtCurrentHoldings").value = "";
					//document.frmMain.elements("txtAvailableHoldings").value = "";
				}
			else
				{
					//document.frmMain.elements("txtCurrentHoldings").value = 'N/A';
					//document.frmMain.elements("txtAvailableHoldings").value = 'N/A';
				}
		}
		
		function UpdateClientHoldings()
		{
			var client = document.frmMain.elements("cboclient").value
			var security = document.frmMain.elements("cboSecurity").value
			var securityType = document.getElementById("cboOrderSecType").value
			     
			if (securityType == 2)
				{
					//document.getElementById("txtCurrentHoldings").innerText = "";
					//document.getElementById("txtAvailableHoldings").innerText = "";

					xmlhttp = createXMLHTTPObj();
									
					url="GetList.asp?client="+client+"&security="+security+"&action=GetClientHoldings";
									 
					xmlhttp.open("GET",url,true);

					xmlhttp.onreadystatechange=function() 
					{
						if (xmlhttp.readyState==4) 
							{
								returnStr = xmlhttp.responseText;
								returnStr = getBodyHTML(returnStr);
																					
								myArray = returnStr.split(";");

								//document.getElementById("txtCurrentHoldings").innerText = myArray[0]
								//document.getElementById("txtAvailableHoldings").innerText = myArray[1]
							}
					}
				xmlhttp.setRequestHeader('Accept','message/x-jl-formresult');
				xmlhttp.send();
				}
			else
				{
					//document.getElementById("txtCurrentHoldings").innerText = 'N/A';
					//document.getElementById("txtAvailableHoldings").innerText = 'N/A';
				}
		}
		
		function updatefields()
		{
			var clientcobo = document.getElementById("cboClient");
			var clientcode = clientcobo[clientcobo.selectedIndex].value;
			var clientcds = '';
			var clientname = '';
			var clientcobo = '';
			var x_clientname;
	     
			document.frmMain.elements("txtClientCode").value = clientcode;
			LoadMyClient();
		}


		
		function ClearFields(element)
		{
		
			document.frmMain.elements("txtAvailableCredit").value = '';
			document.frmMain.elements("txtCurrentBal").value = '';
			
		   if (element == 'txtClientCode')
		   {
			document.frmMain.elements("txtClientCode").value = ''
			document.frmMain.elements("txtCdsNo").value = 'CSD No.'
			document.frmMain.elements("txtclientname").value = 'Client Name'
			return;
		   }
		   if (element == 'txtCdsNo')
		   {
			document.frmMain.elements("txtClientCode").value = 'Code'
			document.frmMain.elements("txtCdsNo").value = ''
			document.frmMain.elements("txtclientname").value = 'Client Name'
			return;
		   }
		   if (element == 'txtclientname')
		   {
		    document.frmMain.elements("txtclientname").value = ''
			document.frmMain.elements("txtClientCode").value = 'Code'
			document.frmMain.elements("txtCdsNo").value = 'CSD No.'
			return;
		   }		
		   
		}
		
		function LoadMyClient()
		{
			var clientcode = document.frmMain.elements("txtClientCode").value
			var clientcds = document.frmMain.elements("txtCdsNo").value
			var clientcobo = document.getElementById("cboclient");		 
			var guidstr = Math.random();
			
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
									
					document.frmMain.elements("txtClientCode").value = myArray[5];
					document.frmMain.elements("txtCdsNo").value = myArray[9]; 
					document.frmMain.elements("txtAvailableCredit").value = myArray[0];
					document.frmMain.elements("txtCurrentBal").value = myArray[1];
					document.frmMain.elements("txtContact").value = myArray[6];
					document.frmMain.elements("txtAgent").value = myArray[2];
					document.frmMain.elements("AgentID").value = myArray[4];
					document.frmMain.elements("txtAccManager").value = myArray[3];
					document.frmMain.elements("AccManagerID").value = myArray[8];
					document.frmMain.elements("txtClientName").value = myArray[7];
         

				}
		   }
				 
		xmlhttp.setRequestHeader('Accept','message/x-jl-formresult');
		xmlhttp.send();
		}
		
		function LoadClient(accountno, element, guidstr)
		{
		 var clientcode = document.frmMain.elements("txtClientCode").value
		 var clientcds = document.frmMain.elements("txtCdsNo").value
		 var clientname = document.frmMain.elements("txtclientname").value
		 var clientcobo = document.getElementById("cboClient");		 
		 
	     var x_clientname;
			
		 if (element == 'txtClientCode')
		 {
			clientcds = ''
			clientname = ''
			
			if (clientcode == '')
			{
			document.frmMain.elements("txtClientCode").value = 'Code'
			clientcobo.length = 1;
			clientcobo[0].text = 'Load Account';
			clientcobo[0].value = '';
			return;
			}
			
		 }
		 else if (element == 'txtCdsNo')
		 {
			clientcode = ''
			clientname = ''

			if (clientcds == '')
			{
			document.frmMain.elements("txtCdsNo").value = 'CSD No.'
			clientcobo.length = 1;
			clientcobo[0].text = 'Load Account';
			clientcobo[0].value = '';
			return;
			}
						
		 }
		 else if (element == 'txtclientname')
		 {
			clientcode = ''
			clientcds = ''

			if (clientname == '')
			{
			document.frmMain.elements("txtclientname").value = 'Client Name';
			clientcobo.length = 1;
			clientcobo[0].text = 'Load Account';
			clientcobo[0].value = '';
			return;
			}
			
		 }

				xmlhttp = createXMLHTTPObj();
				
				url="GetList.asp?clientcode="+clientcode+"&cdsno="+clientcds+"&clientname="+clientname+"&action=SLoadClient&guidstr="+guidstr;
				
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
								
								//document.getElementById("cboClient").options.length = 0;
								clientcobo.length = 1;
								if (element != 'txtclientname')
								{
									document.frmMain.elements("txtClientCode").value = myArray[5];
									document.frmMain.elements("txtCdsNo").value = myArray[9]; 
		
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
									document.frmMain.elements("txtClientname").value = myArray[7]; 
									document.frmMain.elements("txtAvailableCredit").value = myArray[0];
									document.frmMain.elements("txtCurrentBal").value = myArray[1];
									document.frmMain.elements("txtContact").value = myArray[6];
									document.frmMain.elements("txtAgent").value = myArray[2];
									document.frmMain.elements("AgentID").value = myArray[4];
									document.frmMain.elements("txtAccManager").value = myArray[3];
									document.frmMain.elements("AccManagerID").value = myArray[8];
									
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
										
										document.frmMain.elements("txtClientCode").value = '';
										document.frmMain.elements("txtCdsNo").value = '';
										document.frmMain.elements("txtclientname").value = '';

										document.frmMain.elements("txtClientCode").value = myArray[5];
										document.frmMain.elements("txtCdsNo").value = myArray[9]; 
										document.frmMain.elements("txtClientname").value = myArray[7]; 
										//document.getElementById("cboClient").options[i] = new Option(myArrayz[6],myArrayz[10],myArrayz[4],myArrayz[8],myArrayz[3],myArrayz[3],myArrayz[0],myArrayz[1],myArrayz[5],myArrayz[7],myArrayz[9]);
										document.getElementById("cboClient").options[i] = new Option(myArrayz[7],myArrayz[5],myArrayz[6],myArrayz[10],myArrayz[4],myArrayz[8],myArrayz[3],myArrayz[3],myArrayz[0],myArrayz[1],myArrayz[5],myArrayz[7],myArrayz[9]);
										
																			
										//document.getElementById("cboClient").
										//document.getElementById("cboClient").
										
										
																										
									}
									
								}
							}
					}
				 
				 xmlhttp.setRequestHeader('Accept','message/x-jl-formresult');
				 xmlhttp.send();
		}
		
		
		function UpdateBalances()
		{
			client = document.frmMain.elements("cboclient")
					
			document.frmMain.elements("txtAvailableCredit").value = client[client.selectedIndex].Credit;
			document.frmMain.elements("txtCurrentBal").value = client[client.selectedIndex].CurrentBal;
			document.frmMain.elements("txtContact").value = client[client.selectedIndex].OrderContact;

			if (client[client.selectedIndex].Iscustodian=='True') 
				{
					document.frmMain.elements("txtCDA").value= 1;
					document.frmMain.elements("chkCDA").checked= true;
				}
			else
				{
					document.frmMain.elements("txtCDA").value= 0;
					document.frmMain.elements("chkCDA").checked= false;
				}
			         
			if (client[client.selectedIndex].AgentReturnable==1) 
				{
					document.frmMain.elements("txtReturnable").value= 1;
				}
			else
				{
					document.frmMain.elements("txtReturnable").value= 0;
				}
				
			document.frmMain.elements("txtAgent").value = client[client.selectedIndex].Agent;
			document.frmMain.elements("AgentID").value = client[client.selectedIndex].AgentID;
			document.frmMain.elements("txtAccManager").value = client[client.selectedIndex].Owner;
			document.frmMain.elements("AccManagerID").value = client[client.selectedIndex].OwnerID;
		}

		function UpdateCertField()
		{
			var i = 0;
			var theList = document.getElementById("cboOrderType");
								
			if(theList.value == '1' && document.frmMain.elements("cboprice").value=='BF') 
				{
					document.frmMain.elements("txtAmount").disabled = false;
					document.frmMain.elements("txtLimit").disabled = false;
					document.frmMain.elements("txtQty").disabled=true;	
					document.frmMain.elements("txtLimit").focus();
				}
			else
				{
					document.frmMain.elements("txtAmount").value="";
					document.frmMain.elements("txtAmount").disabled=true;
					document.frmMain.elements("txtQty").disabled=false;
					document.frmMain.elements("txtLimit").disabled = true;
					document.frmMain.elements("txtQty").focus();
				}
				
			document.frmMain.elements("txtAmount").value="";
					
			//Diasble To Pay field in case of Purchase Orders
			if(theList.value == '1')
				{
					document.frmMain.elements("chkPay").checked = false;
					document.frmMain.elements("chkPay").disabled = true;
				}
			else 
				{
					document.frmMain.elements("chkPay").checked = true;
					document.frmMain.elements("chkPay").disabled = false;
				}
		}
		
		function UpdateSecurityListing(theList)
		{
			//swap lists
			if(theList.currentSecType == "S")
				{
					document.frmMain.elements("cboFixed").style.display = "block";
					document.frmMain.elements("cboFixed").name = "cboSecurity";
									
					document.frmMain.elements("cboSecurity").style.display = "none";
					document.frmMain.elements("cboSecurity").name = "cboSecurityHidden";
									
					theList.currentSecType = "F"
					document.frmMain.elements("cboIssue").disabled = false;
					document.frmMain.elements("cboPrice").selectedIndex = 0;
					document.frmMain.elements("txtPrice").value = "";
					document.frmMain.elements("cboPrice").disabled = true;
				}
			else
				{
					document.frmMain.elements("cboFixed").style.display = "none";
					document.frmMain.elements("cboFixed").name = "cboSecurityHidden";
									
					document.frmMain.elements("cboSecurity").style.display = "block";
					document.frmMain.elements("cboSecurity").name = "cboSecurity";
									
					theList.currentSecType = "S"
					document.frmMain.elements("cboIssue").disabled = true;
					document.frmMain.elements("cboPrice").disabled = false;
				}
		}
		
		function ChangeCalendar(thechk)
		{
			if(thechk.checked)
				{
					document.frmMain.elements("txtValidity").disabled=false
					document.frmMain.elements("txtcalendar").value="1"
				}
			else
				{
					document.frmMain.elements("txtValidity").disabled=true
					document.frmMain.elements("txtcalendar").value="0"
				}
		}

		function ChangePrice(thecbo)
		{
			if(thecbo.value == "P")
				{		
					document.frmMain.elements("txtPrice").disabled=false
					document.frmMain.elements("txtPrice").value=""
					document.frmMain.elements("txtQty").disabled=false;
					document.frmMain.elements("txtQty").focus();
					document.frmMain.elements("txtAmount").disabled=true;
					document.frmMain.elements("txtbest").value="0";
				}
			else
				{
					document.frmMain.elements("txtbest").value="1";
					document.frmMain.elements("txtPrice").value='Best'
					document.frmMain.elements("txtPrice").disabled=true			
			
					if(document.frmMain.elements("cboOrderType").value=='1') 
						{
							document.frmMain.elements("txtAmount").disabled=false;
							document.frmMain.elements("txtLimit").disabled=false;
							//document.frmMain.elements("txtLimit").focus();
							document.frmMain.elements("txtQty").disabled=true;
						}
					else
						{
							document.frmMain.elements("txtQty").focus();
						}
						
					document.frmMain.elements("txtPrice").value="Best"
				}
		}

		function UpdateIssue(theid)
		{
			var DataCombo = document.all.item("cboIssueList");		

			for (var i=0;i<DataCombo.length;i++)
				{
				ComboID = DataCombo.options[i].id;		
				}
		}

		function FetchAccounts(theList)
				{
					var i = 0;
					var entity = theList.value;
					var toList = document.frmMain.cboIssue;
											
					frm = document.frmMain;				
					xmlhttp = createXMLHTTPObj();
					
					url="GetList.asp?ID="+entity+"&action=GetBondList";
					xmlhttp.open("GET",url,true);
					xmlhttp.onreadystatechange=function() {
						if (xmlhttp.readyState==4) {
						returnStr = xmlhttp.responseText;
						returnStr = getBodyHTML(returnStr);
					
						var secList = "<select name = '" + toList.name + "' id = '" + toList.name + "' size='1' ";
						secList += "OnClick='event.cancelBubble=true;' " ;
						secList += "onChange='event.cancelBubble=true;' " ;
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
				
				
				}
		function formatnumber(theTxt)
			{	
			var theprice = theTxt.value;
			var thesectype=document.frmMain.elements("cboOrderSecType").value
			
			if(thesectype==1)
				{
				theprice =format_number(theprice,4);
				}
				else
				{
				theprice =format_number(theprice,2);
				}
				theTxt.value=theprice;
			}
			
		function format_number(p,d) 
			{
		  	var r;
		  	if(p<0)
		  		{
		  		p=-p;
		  		r=format_number2(p,d);
		  		r="-"+r;
		  		}
		  	else
		  		{
		  		r=format_number2(p,d);
		  		}
		  return r;
			}

		function format_number2(pnumber,decimals) 
			{
		  	var strNumber = new String(pnumber);
		  	var arrParts = strNumber.split('.');
		  	var intWholePart = parseInt(arrParts[0],10);
		  	var strResult = '';
		  	if (isNaN(intWholePart))
		    intWholePart = '0';
		  	if(arrParts.length > 1)
		  		{
		    	var decDecimalPart = new String(arrParts[1]);
		    	var i = 0;
		    	var intZeroCount = 0;
		     	while ( i < String(arrParts[1]).length )
		     		{
		       		if( parseInt(String(arrParts[1]).charAt(i),10) == 0 )
		       			{
		         		intZeroCount += 1;
		         		i += 1;
		       			}
		       		else
		         	break;
		    		}
		    	decDecimalPart = parseInt(decDecimalPart,10)/Math.pow(10,parseInt(decDecimalPart.length-decimals-1)); 
		    	Math.round(decDecimalPart); 
		    	decDecimalPart = parseInt(decDecimalPart)/10; 
		    	decDecimalPart = Math.round(decDecimalPart); 

		    	//If the number was rounded up from 9 to 10, and it was for 1 'decimal' 
		    	//then we need to add 1 to the 'intWholePart' and set the decDecimalPart to 0. 

		    	if(decDecimalPart==Math.pow(10, parseInt(decimals)))
		    		{ 
		      		intWholePart+=1; 
		      		decDecimalPart="0"; 
		    		} 
		    	var stringOfZeros = new String('');
		    	i=0;
		    	if( decDecimalPart > 0 )
		    		{
		      		while( i < intZeroCount)
		      			{
		        		stringOfZeros += '0';
		        		i += 1;
		      			}
		    		}
		    	decDecimalPart = String(intWholePart) + "." + stringOfZeros + String(decDecimalPart); 
		    	var dot = decDecimalPart.indexOf('.');
		    	if(dot == -1)
		    		{
		      		decDecimalPart += '.'; 
		      		dot = decDecimalPart.indexOf('.'); 
		    		} 
		    	var l=parseInt(dot)+parseInt(decimals); 
		    	while(decDecimalPart.length <= l) 
		    		{
		      		decDecimalPart += '0'; 
		    		}
		    	strResult = decDecimalPart;
		  		}
		  	else
		  		{
		    	var dot; 
		    	var decDecimalPart = new String(intWholePart); 

		    	decDecimalPart += '.'; 
		    	dot = decDecimalPart.indexOf('.'); 
		    	var l=parseInt(dot)+parseInt(decimals); 
		    	while(decDecimalPart.length <= l) 
		    		{
		      		decDecimalPart += '0'; 
		    		}
		    	strResult = decDecimalPart;
		  		}
		  	return strResult;
			}


			
</SCRIPT>

</head>
<%
	
	Dim UserId
	Dim action
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim guidStr 
	Dim guid 
	Dim buttonAction
	
	action = ucase(Request.Form("action"))
	UserId=Session("UserID")	
	
	if action = "EXECUTE" then					
		Dim reloadRequired
		
		reloadRequired = false
		buttonAction = Trim(Trim(Ucase(Request.Form("cmdAdd"))))
		
		if instr(1,buttonAction,"SAVE") > 0 Or instr(1,buttonAction,"ADD") > 0 then
				'Save Details
				Dim branch
				Dim client
				Dim orderType
				Dim orderType2
				Dim hold
				Dim orderDate
				Dim ref
				Dim security
				Dim qty
				Dim price
				Dim cert
				Dim validity
				Dim secType
				Dim holdType
				Dim releaseDate
				Dim compound
				Dim bond
				Dim savedate
				Dim amount
			 	Dim Best
			 	Dim chkprice
			 	Dim interbank
				Dim Limit
				Dim IsCustodian
				Dim ToPay
                Dim Agent
				Dim AccManager
				Dim OrderContact
				Dim AgentReturnable
			 	Dim Fundmanager
			 	Dim Comment
			 	Dim orderTime
			 	Dim LAction
				Dim PT
				Dim PayOptions
				Dim PartialAmount
				Dim AvailableCredit
				
				LAction = "ADD"
				
				PayOptions = Request.Form("cboPayOptions")
				PartialAmount = Replace(Request.Form("txtPayAmt"),",","")
				
				if(Cint(PayOptions)=3) then
					if(trim(PartialAmount)="") then
						%>
						<script language="javascript">
						alert('Please specify the Partial amount')
						</script>
						<%
						Response.End 
					end if
					
					'Ensure Partial Amount is numeric
					If  (Not IsNumeric(PartialAmount)) Then%>
						<script language = 'vbscript'>
							ShowMessage "PartialAmount must be numeric."
							
						</script>
						<% response.end
					End If
                 
				else
					PartialAmount=0
				end if
			
			
				bond = Request.Form("cboIssue")
				compound = cint(Request.Form("chkCompound"))
				interbank= cint(Request.Form("chkInterbank"))
				holdType = Request.Form("cboOrderHoldType")		'This hold Type is actually order hold options ID
				releaseDate = Request.Form("txtReleaseDate")
				branch = Request.Form("cboBranch")
				client = Request.Form("cboClient")
				orderType = Request.Form("cboOrderType")
				secType = Request.Form("cboOrderSecType")
				hold = Request.Form("cboHold")
				orderDate = Request.Form("txtDate")
				ref = Request.Form("txtRef")
				security = Request.Form("cboSecurity")
				qty = Request.Form("txtQty")
				price = Request.Form("txtPrice")
				cert = Request.Form("txtCert")
				savedate=Request.Form("txtCalendar")
				amount=Replace(Request.Form("txtamount"),",","")
				validity = Request.Form("txtValidity")
				best=Request.Form("txtbest")
				chkprice=Request.Form("cboprice")	 
				Limit = Request.Form("txtLimit")
				IsCustodian = Request.Form("txtCDA")
				ToPay = Request.Form("chkPay")
				Agent = Request.Form("AgentID")
				AccManager = Request.Form("AccManagerID")
				OrderContact = Request.Form("txtContact")
				AgentReturnable = Request.Form("txtReturnable")
				Fundmanager = Request.Form("cboFundManager")
				Comment = Request.Form("txtComment")
				CreateTime = Now()
				PT = Cint(Request.Form("chkPT"))
				AvailableCredit = trim(Request.Form("txtAvailableCredit")) 
				
				'Assign Default values
				if(bond="") then
					bond = 0
				end if
				
				if(trim(amount)="") then
					amount = 0
				end if
					
				if(trim(price)="") then
					price = 0
				end if

				if(Ucase(trim(price))="BEST") then
					best = 1
				end if	
				
				if (trim(IsCustodian) = "") then
					IsCustodian = 0
				end if

				if (trim(ToPay) = "") then
					ToPay = 0
				end if

				if (trim(AgentReturnable) = "") then
					AgentReturnable = 0
				end if

				if (trim(compound) = "") then
					compound = 0
				end if

				if (trim(interbank) = "") then
					interbank = 0
				end if
                
				if (trim(Agent) = "") then
					Agent = "NULL"
				end if

				if (trim(AccManager) = "") then
					AccManager = "NULL"
				end if
				
				if (trim(Fundmanager) = "") then
					Fundmanager = "NULL"
				end if
				
				'this is to check for both purchase and sale duplicate orders 
				'i.e client should not have p and s orders for same security
				if orderType = 1 then
					orderType2 = 2
				else
					orderType2 = 1
				end if		
				
				Limit = trim(replace(Limit,",",""))
				AvailableCredit = ccur(AvailableCredit)
				
				'--- Order Header Info Validation ----

				'validate Branch
				 If Trim(branch) = "" Then%>
				         <script language = 'vbscript'>
				         		ShowMessage "Please specify the Branch"
				         </script>
				         <% response.end
				 End If
                 
				 'check Order's date validity
				if cdate(orderDate) < Date() then%>
				         <script language = 'vbscript'>
							ShowMessage "The Order Date cannot be in the past"
				         </script>
				         <% response.end
				 End If
                  
				  'validate Order Type
				 If Trim(orderType) = "" Or Trim(orderType) = "0" Then%>
				         <script language = 'vbscript'>
				         		ShowMessage "Please specify the Order Type"
				         </script>
				         <% response.end
				 End If
                 
				 'validate security type
				 If Trim(secType) = "" Then%>
				         <script language = 'vbscript'>
				         		ShowMessage "Please specify the security type"
				         </script>
				         <% response.end
				 End If
                 
				 'validate Order Hold Option 
				 If Trim(holdType) = "none" Then%>
				         <script language = 'vbscript'>
				         		ShowMessage "Please specify the Order Hold Option"
				         </script>
				         <% response.end
				 End If
                 
				 'validate Hold
				 If Trim(Hold) = "" Then%>
				         <script language = 'vbscript'>
				         		ShowMessage "Please specify the Hold status"
				         </script>
				         <% response.end
				 End If
                 
				 'validate Client
				 If Trim(Client) = "" Then%>
				         <script language = 'vbscript'>
				         		ShowMessage "Please specify the Client"
				         </script>
				         <% response.end
				 End If

                 'validate size of Reference No.
				 If Len(Ref) > 100 Then%>
				         <script language = 'vbscript'>
					         ShowMessage "Reference No. can only be 100 characters in length"
				         </script>
				         <% response.end
				 End If
				 
				 'validate size of remarks/comments
				 If Len(Comment) > 50 Then%>
						<script language = 'vbscript'>
							ShowMessage "Remarks can only be 50 characters in lenght"
						</script>
						<% response.end
				 End If
                
				'--- Order Grid Info Validation ----

                'validate Security 
				 If Trim(Security) = "" Then%>
						<script language = 'vbscript'>
								ShowMessage "Please specify the Security "
						</script>
						<% response.end
				 End If	
                 
				 'validate Estimated Quantity
				 If Trim(qty)= "" and Best = 0 Then%>
						<script language = 'vbscript'>
								ShowMessage "Please specify the Quantity "
						</script>
						<% response.end
				 End If
                        
                 'Ensure Order Detail Estimated Quantity is numeric
				 If (Qty <> "") And (Not IsNumeric(Qty)) Then%>
					<script language = 'vbscript'>
						ShowMessage "Order Detail Quantity must be numeric."
						
					</script>
					<% response.end
				 End If
                 
				 'validate Order Limit for Purchase Best
				 'If Trim(Limit)= "" and Best = 1 and orderType = 1 Then%>
						<script language = 'vbscript'>
								'ShowMessage "Please specify the Limit."
						</script>
						<% 'response.end
				'End If

				 'ensure Order Limit is numeric
				'If (Limit <> "") And (Not IsNumeric(Limit)) Then%>
					<script language = 'vbscript'>
						'ShowMessage "Order Detail Limit must be numeric."
					</script>
					<% 'response.end
				'End If
               
				 'ensure Order Detail Price is numeric
				If (Price <> "") And (Not IsNumeric(Price)) and (chkprice="P") Then%>
					<script language = 'vbscript'>
						ShowMessage "Order Detail Price must be numeric."
					</script>
					<% response.end
				End If

				'Check for Amount to be more than 0
				if(Cint(best)=1 and Cint(orderType)=1 and Cdbl(amount)<=0) then%>
					<script language = 'vbscript'>
				         ShowMessage "Order Detail Amount Must be more than Zero"
				    </script>
				    <% response.end
				end if
				
				'Ensure vaildity canot be extended beyond one month from the order date
				if(CDate(validity) > CDate(theDate)) then%>
					<script language = 'vbscript'>
						ShowMessage "The validity of this order cannot exceed one month from the order date"
					</script>
					<% response.end
				end if
				
					'Validate available balance for purchase orders
					if Cint(orderType)=1 then
						'AvailableCredit 
						if Best = 0 then
							if AvailableCredit < 0 AND Qty > 0 then
							  %>
								<script language = 'vbscript'>
									ShowMessage "Order cannot be registered! Insufficient funds available for account [<%=client%>]."
								</script>
								<% response.end
							end if
						end if
				
						if IsNumeric(Price) AND IsNumeric(Qty) then
						 amt = cdbl(Price*Qty)
						 
							if AvailableCredit < amt AND amt > 0 then
							  %>
								<script language = 'vbscript'>
									ShowMessage "Order cannot be registered! Insufficient funds available for account [<%=client%>]."
								</script>
								<% response.end
							end if
						end if
					end if
				
				'save header
				if holdType = 4 then 'hold until given date
				   	releaseDate = "#" & releaseDate & "#"
				else
				   	releaseDate = "NULL"
				end if
				
				validity = "NULL"
				if savedate = 1 then 'save only the date check box checked
				   	validity = "#" & validity & "#"
				else
				   	validity = "NULL"
				end if
				 
				if(trim(Limit)="") then
					Limit = 0
				end if
                
				Set conn = GetActiveConnection("KBroker")
				
					'Do not allow Duplicate Orders
					 sqlstr = "SELECT     dbo.tbOrder.Order_DPA_ " & _
							" FROM         dbo.tbOrder INNER JOIN " & _
							"                       dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN " & _
							"                       dbo.OrdDetail ON dbo.tbOrder.Order_DPA_ = dbo.OrdDetail.Order_DPA_ INNER JOIN " & _
							"                           (SELECT     SUM(LotQty) AS ContractQty, OrdDetail_DPA_ " & _
							"                             FROM          dbo.Lot " & _
							"                             WHERE      (Deleted = 0) " & _
							"                             GROUP BY OrdDetail_DPA_) contracted ON dbo.OrdDetail.OrdDetail_DPA_ = contracted.OrdDetail_DPA_ " & _
							" WHERE     (dbo.tbOrder.Deleted = 0) AND (dbo.OrdDetail.Deleted = 0) AND (dbo.tbOrder.OrderCanceled <> 1) AND (dbo.tbOrder.Client_DPA_ = " & client & ") AND  " & _
							"                       (dbo.OrdDetail.Security_DPA_ = " & security & ") AND (dbo.OrderType.OrderType_DPA_ = " & orderType & ") AND (dbo.tbOrder.OrderSecType_DPA_ = " & secType & ") AND  " & _
							"                       (dbo.OrdDetail.OrdDetailQty - contracted.ContractQty > 0) "
							
					Set rs = conn.Execute(HandleQuote(sqlStr))
				
					If  Not (rs.EOF Or rs.BOF) Then
					    OrderDPA = trim(rs.fields("Order_DPA_"))
						%>
			 			<script language = 'vbscript'>
			 					ShowMessage "This is a duplicate order. Please follow up the original Order (<%=OrderDPA%>)."
											
			 			</script>
			 			<% response.end
					End If
				
				'Save Order
				set guid = server.createobject("NDUtils.CGUID")
				guidStr = guid.GenerateGUID
				 
				sqlStr = "INSERT INTO [tbOrder] (OrderDate,OrderHold,OrderRef,Order_DPA_,Order_EIT_,Branch_DPA_," & _
					" OrderSecType_DPA_,Client_DPA_,OrderType_DPA_,OrderAutoReleaseDate,OrderHoldType_DPA_," & _
					" OrderCompounded,InterBank,ChangedBy,IsCustodian,ToPay,PT,PayOption,PartialAmount," & _
					" Remarks, Agent_DPA_, AgentReturnable,TimeChanged) SELECT " & "#" & FormatDate(orderDate) & "#" & " as OrderDate," & " " & hold & " " & " as OrderHold" & _
					"," & "'" & ref & "'" & " as OrderRef," & " " & "iif(isnull(max([Order_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'tbOrder'),max([Order_DPA_]) + 1)" & " " & " as Order_DPA_" & _
					"," & "'" & guidStr & "'" & " as Order_EIT_," & " " & branch & " " & " as Branch_DPA_," & " " & secType & " " & " as OrderSecType_DPA_" & _
					"," & " " & client & " " & " as Client_DPA_" & _
					"," & " " & orderType & " " & " as OrderType_DPA_" & _
					"," & " " & releaseDate & " " & " as OrderAutoReleaseDate" & _
					"," & " " & holdType & " " & " as OrderHoldType_DPA_" & _
					"," & " " & compound & " " & " as OrderCompounded" & _
					"," & " " & interbank & " " & " as InterBank" & _
				   	"," & " " & UserId & " " & " as ChangedBy " & _
				   	"," & " " & IsCustodian & " " & " as IsCustodian" & _
				   	"," & " " & ToPay & " " & " as ToPay" & _
				   	"," & " " & PT & " " & " as PT" & _
				   	"," & " " & PayOptions & " " & " as PayOption" & _
				   	"," & " " & PartialAmount & " " & " as PartialAmount" & _
				   	"," & "'" & Comment & "'" & " as Remarks, "& Agent &" AS Agent, "& AgentReturnable &" AS AgentReturnable" & _
				   	",getdate() as TimeChanged FROM [tbOrder]"       
				' response.write sqlStr
				' response.end
				 conn.BeginTrans
						
						sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
						
						conn.Execute sqlStr
				     
						'obtain header key value
						sqlStr = "SELECT tbOrder.Order_DPA_ FROM tbOrder WHERE tbOrder.Order_EIT_ = " & "'" & guidStr & "'"
						Set rs = conn.Execute(HandleQuote(sqlStr))
						
						If (rs.EOF Or rs.BOF) Then%>
				         		<script language = 'vbscript'>
				         				ShowMessage "A serious error has been encountered while saving the data. Try saving again"
				         		</script>
				         		<% response.end
						End If

						'save detail data
						sqlStr = "INSERT INTO [OrdDetail] (OrdDetailCertNo,OrdDetailPrice,Amount,Best,Bond_DPA_,OrdDetailQty,OrdDetailValidity" & _
							",OrdDetail_DPA_,Order_DPA_, Limit,Security_DPA_) SELECT " & "'" & cert & "'" & " as OrdDetailCertNo" & _
							"," & "'" & price & "'" & " as OrdDetailPrice" & _
							"," & " " & amount & " " & " as Amount" & _
							"," & " " & Best & " " & " as Best" & _
							"," & " " & bond & " " & " as Bond_DPA_" & _
							"," & " " & CDbl(qty) & " " & " as OrdDetailQty" & _
							"," & " " & validity & " " & " as OrdDetailValidity" & _
							"," & " " & "iif(isnull(max([OrdDetail_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'OrdDetail'),max([OrdDetail_DPA_]) + 1)" & " " & " as OrdDetail_DPA_" & _
							"," & " " & rs.Fields("Order_DPA_") & " " & " as Order_DPA_" & _
							"," & " " & CDbl(Limit) & " " & " as Limit " & _
							"," & " " & security & " " & " as Security_DPA_ " & _
							" FROM [OrdDetail]"										
						
						sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
						conn.Execute sqlStr
				
				conn.CommitTrans
				
				'retrieve the item ID
				OrderID = rs.Fields("Order_DPA_")
				sqlStr = "SELECT OrdDetail.OrdDetail_DPA_, Security_DPA_  FROM OrdDetail WHERE OrdDetail.Order_DPA_=" & rs.Fields("Order_DPA_")
				Set rs = conn.Execute(HandleQuote(sqlStr))
				If (rs.EOF Or rs.BOF) Then%>
					<script language = 'vbscript'>
						ShowMessage "An error has been encountered while saving the order. Try editing the Order if you wish to add more entries"
					</script>
					<% response.end
				End If
				
				conn.execute ("ClientTotalProcedure " & client)
				'conn.execute ("ClientTotalsDelete")
				'conn.execute ("ClientTotalsProcedure")
				'conn.execute ("ClientBalancesDelete")
				'conn.execute ("ClientBalancesProcedure")


				if instr(1,buttonAction,"LINE/S") > 0 then
					%>
					<SCRIPT LANGUAGE="JAVASCRIPT">
						window.parent.parent.frames['maininfo'].location.reload();
					</SCRIPT>
					<%

					WriteDialogRelocateScript "EditOrder.asp?ID=" & rs.Fields("OrdDetail_DPA_")
				elseif instr(1,buttonAction,"PRINT") > 0 then
					%>
					<SCRIPT LANGUAGE="JAVASCRIPT">
						window.parent.parent.frames['maininfo'].location.reload();
					</SCRIPT>
					<%
					WriteDialogRelocateScript "OrderForm.asp?order_id=" & OrderID & "&returnPath=EditOrder.asp&OrdDetailID=" & rs.Fields("OrdDetail_DPA_")

				elseif instr(1,buttonAction,"ADD NEW") > 0 then%>
					<SCRIPT LANGUAGE="JAVASCRIPT">
						window.parent.parent.frames['maininfo'].location.reload();
					</SCRIPT>
					<%
					WriteDialogRelocateScript "AddOrder.asp"
				Else
					WritefraEnabledDialogCloseScript
					%>
					<script>
						//window.parent.parent.frames["maininfo"].location.reload();
						//try
						//{
						//	window.parent.parent.frames["maininfo"].location.href="/Operations/OrderList.asp";
						//}
						//catch(err)
						//{
							//alert(err.description);
						//}
						//window.parent.closeDocOpener();
					</script>
					<%
					'WritefraEnabledDialogCloseScript
				end if
				
				Response.End
		end if
   	end If
%>
		
<body Class="Dialog">

<div id="spiffycalendar" class="text" STYLE="z-Index: 1000; width:8; height:130"></div>

<form name = 'frmAddOrder' method = 'post' action = 'AddOrder.asp' id = 'frmMain'>

<table border="0" width="830" height="357" cellspacing=2 cellpadding=2>
  <tr>
    <td width="98" height="22">Branch</td>
    <td width="200" height="22"><select name = 'cboBranch' id = 'cboBranch' size="1">
    	<option selected value = ''></option>
		<%
        Set conn = GetActiveConnection("KBroker")
        
        sqlStr = "SELECT * FROM [BranchList]"
        Set rs = conn.Execute(sqlStr)
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF
                        if cbool(rs.Fields("DefaultSelection")) then%>
								<option selected value = '<%=rs.Fields("Branch_DPA_")%>'><%=rs.Fields("BranchName")%></option>
                        <%else%>
								<option value = '<%=rs.Fields("Branch_DPA_")%>'><%=rs.Fields("BranchName")%></option>
                        <%end if
                        rs.MoveNext
                Loop
        End If
		%>
    </select></td>
    <td width="141" height="22">Client</td>
    <td width="676" height="22">
    &nbsp;<input type = 'text' value='Code' name ='txtClientCode' id = 'txtClientCode' size="10" onBlur="txtval = this.value; LoadClient(cboClient, this.name, '<%=guidStr%>');" onChange="selectItem(cboClient); UpdateCodes(true,cboClient,txtCdsNo);UpdateBalances();UpdateClientHoldings();" onClick  = "ClearFields(this.name);">&nbsp;
	<input type = 'text' tabindex=-1 name ='txtCdsNo' value = 'CSD No.' id = 'txtCdsNo' size="16" onBlur="txtval = this.value; LoadClient(cboClient, this.name, '<%=guidStr%>');" onClick  = "ClearFields(this.name);">&nbsp;
    <input type = 'text' tabindex=-1  name ='txtclientname' id = 'txtclientname' size="15" value = 'Client Name' onBlur="txtval = this.value;LoadClient(cboClient, this.name, '<%=guidStr%>')" onClick  = "ClearFields(this.name)">&nbsp;&nbsp;
    &nbsp;&nbsp;<select name = 'cboClient' id = 'cboClient' size="1" readonly onChange=updatefields();>
		<%
		dim ClientName
		dim NameClient      
				
				%>                    
				<option OrderContact = "" Iscustodian = "" AgentID = "" Agent = "" OwnerID = "" Owner = "" Credit="" CurrentBal="" SearchCode = "" SearchText = "" SearchCds = "" value = ''>Load Client</option>
				<%
			'Next
		'End If
		%>
    </select></td>
  </tr>
  <tr>
    <td width="98" height="18">Date</td>
    <td width="200" height="18"><SCRIPT language="JavaScript">cal.writeControl();</SCRIPT></td>
    <td width="141" height="18">&nbsp;</td>
    <td width="676" height="18" align=left>
    &nbsp;
    Current Balance
    &nbsp;
    <input type = 'text' name ='txtCurrentBal' id = 'txtCurrentBal' readonly class="readonlyex" size="15">
    &nbsp;
    Available Credit
    &nbsp;
    <input type = 'text' name ='txtAvailableCredit' id = 'txtAvailableCredit' readonly class="readonlyex" size="15">
    </td>
  </tr>
  <tr>
    <td width="98" height="28">Order Type</td>
    <td width="200" height="28"><select name = 'cboOrderType' id = 'cboOrderType' size="1" onchange='UpdateTitles();UpdateCertField(); ' >
    	<option selected value = ''></option>
		<%
        sqlStr = "SELECT * FROM [OrderTypeList]"
        Set rs = conn.Execute(sqlStr)
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
		%>
			<option selected value = '0'>Select Type ...</option>
		<%
                Do Until rs.EOF
                        if cbool(rs.Fields("DefaultSelection")) then%>
								<option value = '<%=rs.Fields("OrderType_DPA_")%>'><%=rs.Fields("OrderTypeName")%></option>
                        <%else%>
								<option value = '<%=rs.Fields("OrderType_DPA_")%>'><%=rs.Fields("OrderTypeName")%></option>
                        <%end if
                        rs.MoveNext
                Loop
        End If
		%>
    </select></td>
    <td width="141" height="28">Account Manager &nbsp;&nbsp;&nbsp;</td>
	<td width="676" height="28">&nbsp;<input type = 'text' name ='txtAccManager' id = 'txtAccManager'  readonly class="readonlyex" size="30">
	<input type = 'hidden' name ='AccManagerID' id = 'AccManagerID'></td>
  </tr>
   <tr>
    <td width="98" height="20">Security Type</td>
    <td width="200" height="20">
<b>
<select name = 'cboOrderSecType' id = 'cboOrderSecType' size="1" currentSecType = "S" onchange='UpdateSecurityListing(this);UpdateClientHoldings();'>
    	 <%sqlStr = "SELECT * FROM [OrderSecTypeList]"
        Set rs = conn.Execute(sqlStr)
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF
                        if cbool(rs.Fields("DefaultSelection")) then%>
								<option selected value = '<%=rs.Fields("OrderSecType_DPA_")%>'><%=rs.Fields("OrderSecTypeDisplayName")%></option>
                        <%else%>
								<option value = '<%=rs.Fields("OrderSecType_DPA_")%>'><%=rs.Fields("OrderSecTypeDisplayName")%></option>
                        <%end if
                        rs.MoveNext
                Loop
        End If%>
    </select></b></td>
    <td width="141" height="20">Agent</td>
<td width="676" height="20">
&nbsp;<input type = 'text' name ='txtAgent' id = 'txtAgent'  readonly class="readonlyex" size="30">
</td>
<input type = 'hidden' name ='AgentID' id = 'AgentID' >

  </tr>
  <tr>
    <td width="98" height="22">Hold</td>
    <td width="200" height="22">
    <select name = 'cboOrderHoldType' id = 'cboOrderHoldType' size="1" onChange="javascript: toggleDate(this)">
    	<%sqlStr = "SELECT * FROM [OrderHoldOptions]"
        Set rs = conn.Execute(sqlStr)
        If Not (rs.EOF Or rs.BOF) Then
            rs.MoveFirst
            Do Until rs.EOF
                if UCase(Trim(rs.Fields("Description").Value)) = "AWAITING MANUAL RELEASE" then%>
						<option id="<%=rs.Fields("RequiresDate")%>" selected value='<%=rs.Fields("OrderHoldOptionID")%>'><%=rs.Fields("Description")%></option>
                <%else%>
						<option id="<%=rs.Fields("RequiresDate")%>" value='<%=rs.Fields("OrderHoldOptionID")%>'><%=rs.Fields("Description")%></option>
                <%end if
                rs.MoveNext
            Loop
        End If%>
    </select>&nbsp;
    <select name = 'cboHold' id = 'cboHold' size="1" style="display:none">
    	<option selected value = '1'>Yes</option>
    	<option value = '0'>No</option>
    </select></td>
    <td width="141" height="22">Contact Name</td>
    <td width="676" height="22">
    &nbsp;<input type = 'text' name ='txtContact' id = 'txtContact' readonly size="40"></td>

  </tr>
  <tr id="HeldDate">
	<td width="98" height="20">Held Date</td>
	<td width="200" height="20"><SCRIPT language="javascript">calReleaseDate.writeControl();</SCRIPT>&nbsp;</td>
	<td width="141" height="20">Ref No.</td>
	<td width="676" height="20">&nbsp;<input type = 'text' name ='txtRef' id = 'txtRef' size="20"></td>

  </tr>
  
  <tr>
    <td width="98" height="20">CDA</td>
    <td width="200" height="20">&nbsp;<input type=checkbox   value='1' name='chkCDA'></td>
    <td width="141" height="20" style="display:none">PT</td>
    <td width="676" height="20" style="display:none">&nbsp;<input type=checkbox   value='1' name='chkPT' id="chkPT"></td>
  </tr>
  
  <tr>
    <td width="98" height="20">Compound</td>
    <td width="200" height="20">&nbsp;<input type=checkbox   value='1' name='chkCompound' id="chkCompound"></td>
    <td id="td1" name="td1" style="display:none" width="141" height="20">To Pay</td>
    <td id="td2" name="td2" style="display:none" width="676" height="20">&nbsp;<input type=checkbox id="chkPay"  value='1' name='chkPay'></td>
  </tr>
  
  <tr>
    <td width="98" height="20"><!--Inter Bank--></td>
    <td width="200" height="20">&nbsp;<input type=hidden   value='1' name='chkInterbank' id='chkInterbank'></td>
    <td width="141" height="20">Remarks/Comments</td>
	<td width="676" height="20">&nbsp;<TEXTAREA name ='txtComment' id = 'txtComment' rows='2' cols="30"></TEXTAREA></td>

  </tr>
  <tr>
	<td><label style="display:none" name="lbSale" id="lbSale">Pay&nbsp;Options</label></td>
	<td>
		<select name ='cboPayOptions' id='cboPayOptions' size=1 onchange='UpdatePartial(this);' style="display:none">
			<option value="1">Reinvest Funds</option>
			<option value="2">Full Payment</option>
			<option value="3">Partial Payment</option>
		</select>
	</td>
	<td><label style="display:none" name="lbPartial" id="lbPartial">Partial Amount</label></td>
	<td>&nbsp;<input type = 'text' name ='txtPayAmt' id = 'txtPayAmt' size="20" style="display:none" value='' OnBlur="JavaScript: format2Number(this)"></td>
  </tr>    
  
  <tr><td colspan = '4'>
	<table border="0">
		<tr>
		  <td style="border-left-width: 1; border-right-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" bordercolor="#000000"><b><font color="#000080">Security</font></b></td>
		  <td style="border-left-width: 1; border-right-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" bordercolor="#000000"><b><font color="#000080">Bond</font></b></td>
		  <td style="border-left-width: 1; border-right-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" bordercolor="#000000"><b><font color="#000080"><span id="Quantity">Quantity</span></font></b></td>
		 <td  style="border-left-width: 1; border-right-width: 1; Display:none; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" bordercolor="#000000"><b><font color="#000080">Limit</font></b></td>
		  <td colspan="2" style="border-left-width: 1;  border-right-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" bordercolor="#000000"><b><font color="#000080"><span id="Price">Price</span></font></b></td>
		  <td style="border-left-width: 1; border-right-width: 1;Display:none; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" bordercolor="#000000"><b><font color="#000080"><span id="Amount">Amount</span></font></b></td>
		  <td style="border-left-width: 1; border-right-width: 1;  border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" bordercolor="#000000"><b><font color="#000080"><span id="Amount">Certificate</span></font></b></td>
		  <td colspan="2" style="border-left-width: 1; border-right-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" bordercolor="#000000"><b><font color="#000080">Validity</font></b></td>
		</tr>
	  
	  <tr>
      <td nowrap>
     <!-- <div style="position: relative; width: 50px; z-Index: -1">-->
      <select name = 'cboSecurity' id = 'cboSecurity' size="1" style="FONT-FAMILY:  Trebuchet MS, ARIAL, TAHOMA; FONT-SIZE: 8PT; WIDTH: 200PX" 
			onKeypress="return (dodefaultaction()==''); " 
			onKeydown="return (dodefaultaction()==''); " 
			onKeyup="return (change(cboSecurity));" 
			onfocus="txtval = '';inputIsItemCode = 1;" 
			onblur="txtval = '';inputIsItemCode = 1;">
    	<option selected SearchCode = "0" SearchText = ""  value = ''></option>
<%
        Set conn = GetActiveConnection("KBroker")
        response.Write GetSecurityList(2)
%>

    </select>
    
   <!-- </div>-->
    <select name = 'cboSecurityHidden' id = "cboFixed" size="1" style="display:none"  onchange='FetchAccounts(this);' style="FONT-FAMILY:  Courier, Trebuchet MS, ARIAL, TAHOMA; FONT-SIZE: 8PT; WIDTH: 200PX">
    	<option selected value = ''></option>
<%
        Set conn = GetActiveConnection("KBroker")
        response.Write GetSecurityList(1)
%>

    </select> </td>
    <td><select name = 'cboIssue' id = 'cboIssue' size="1" onchange='' >    			
    </select></td>
	<td>	
      <input type = 'text' name ='txtQty' id = 'txtQty' size="9" OnBlur="JavaScript: format2NumberCommasOnly(this)" onChange="JavaScript:UpdateTotalAmount(this.name)">
      </td>
	  <td style="Display:none;"><input type = 'text' name ='txtLimit' id = 'txtLimit' size="9" value="0" disabled OnBlur="JavaScript: formatnumber(this)"></td>
      <td style="Display:none;">&nbsp;<select name="cboprice" onchange='ChangePrice(this);'>			
			<option selected SearchCode = "0" SearchText = "P" value = 'P'>P</option>			
			<option value='BF'>BF</option>			
		</select>
      </td>
	  <td ><input type = 'text' name ='txtPrice' id = 'txtPrice' size="9" OnKeyUp="JavaScript: updateBestPrice(this)" OnBlur="JavaScript: formatnumber(this)" onChange="JavaScript:UpdateTotalAmount(this.name)"></td>
      <td>
	  <input  name ='txtAmount' id = 'txtAmount' size="5" readonly="true" class="readonly" OnChange="JavaScript: format2Number(this)"></td>
      <td><input type = 'text' name ='txtCert' id = 'txtCert' size="15"></td>
      <td>
      <input type ='checkbox' name ='chkdate' id='chkdate' size="1" class='BorderLess' onClick = 'ChangeCalendar(this);' value="ON"></td>
      <td><SCRIPT language="JavaScript">
      calValidity.writeControl();
      document.frmMain.elements("txtValidity").disabled=true; //Disable Validity date by default
      </SCRIPT></td>
    </tr>
	</table>
  </td></tr>
 
</table>

<table border="0"  width="830">
	  <tr>
    <td width="20%"  align=right>
		
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAddMore' value=" Add Line/s " onclick = "AllowedNavigation()">&nbsp;&nbsp;&nbsp; <input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAddMoreNew' value=" Add New Order " onclick = "AllowedNavigation()">
        &nbsp;&nbsp; <input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAddPrint' value=" Save & Print " onclick = "AllowedNavigation()">&nbsp;&nbsp;
        <input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save & Close " onclick = "AllowedNavigation()">
        &nbsp;&nbsp; <input type = 'button' Class=Buttons name ='cmdClose' id = "cmdClose" value=" Close " onclick = "window.self.close();">
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		<input type = 'hidden' name ='txtcalendar' id = 'txtcalendar' value='0'>
		<input type = 'hidden' name ='txtCDA' id = 'txtCDA' value='0'>
		<input type = 'hidden' name ='txtReturnable' id = 'txtReturnable' value='0'>
		<input type = 'hidden' name ='txtbest' id = 'txtbest' value='0'>
		<select name = 'cboIssueList' size="1" style="display:none">    	
		<%
        sqlStr = "SELECT DISTINCT SecurityCode + ' - ' + BondIssue AS Issue,Security_DPA_,Bond_DPA_ FROM IssueList"
        Set rs = conn.Execute(sqlStr)
		%>
    </select>&nbsp;&nbsp;&nbsp;
	</td>
  </tr>
</table>
</form>
</body>
  <script language="javascript">
  function toggleDate(theCombo)
  {
	if (theCombo.options[theCombo.options.selectedIndex].id==1)
	{
		document.all.item("HeldDate").style.display = "";
	}
	else
	{
		document.all.item("HeldDate").style.display = "none";
	}
  }
  
  
  function UpdatePartial(theCombo)
  {
	var comboid=theCombo.value;
	
	if(comboid==3)
	{
	document.all.item("txtPayAmt").style.display = "";
	document.all.item("lbPartial").style.display = "";
	//document.all.item("td1").style.display = "";
	//document.all.item("td2").style.display = "";
	}
	else
	{
	document.all.item("txtPayAmt").style.display = "none";
	document.all.item("lbPartial").style.display = "none";
	//document.all.item("td1").style.display = "none";
	//document.all.item("td2").style.display = "none";
	}
	
	if(comboid==1) 
	{	
	//document.all.item("td1").style.display = "";
	//document.all.item("td2").style.display = "";
	document.all.item("chkPay").checked=false
	//document.all.item("td1").style.display = "none";
	//document.all.item("td2").style.display = "none";
	}
	else
	{
	//document.all.item("td1").style.display = "";
	//document.all.item("td2").style.display = "";
	document.all.item("chkPay").checked=true;
	}
  }

	function UpdateTotalAmount(element)
	{
		var Price = document.frmMain.elements("txtPrice").value;
		Price=Price.replace(',','')
		var qty = document.frmMain.elements("txtQty").value;
		qty=qty.replace(',','')
		var Amt ;
		//Validate against charcters
		if ( Price.length>0  && qty.length>0 )
		{
			Amt=parseFloat(Price)*parseInt(qty);
			if (isNaN(Amt)==true)
			{
				document.getElementById("txtAmount").value=0;
			}
			else
			{
				document.getElementById("txtAmount").value=Amt;
			}
		}
		else
		{
			document.getElementById("txtAmount").value=0;
		}
		
	}
  </script>

</html>
<%
	function GetSecurityList(secType)
			Dim optionList
			'2nd where removes serena coz its suspended from trading
			sqlStr = "SELECT * FROM [SecurityList] WHERE OrderSecType_DPA_ = " & secType & " ORDER BY SecurityName"
			Set rs = conn.Execute(sqlStr)
			If Not (rs.EOF Or rs.BOF) Then
					rs.MoveFirst
					Do Until rs.EOF
							optionList = optionList & "<option TITLETEXT=""" & rs.Fields("SecurityCode") & """ SearchCode = """ & rs.Fields("SecurityCode") & """ SearchText = """ & rs.Fields("SecurityCode") & """  value = """ & rs.Fields("Security_DPA_") & """>" & rs.Fields("SecurityCode") & "</option>" & chr(13)
							rs.MoveNext
					Loop
			End If
			GetSecurityList = optionList
	end function
%>