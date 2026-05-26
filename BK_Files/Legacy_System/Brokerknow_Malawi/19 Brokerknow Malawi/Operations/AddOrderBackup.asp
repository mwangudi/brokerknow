<html>
<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add Order</title>
<!--#include file="../libroutines.asp"-->

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
<SCRIPT language="JavaScript">
	var cal=new ctlSpiffyCalendarBox("cal", "frmAddOrder", "txtDate","cmdDate","<%=FormatDate(Date)%>",1);
	var calValidity=new ctlSpiffyCalendarBox("calValidity", "frmAddOrder", "txtValidity","cmdValDate","<%=FormatDate(Date)%>",1);
	var calReleaseDate=new ctlSpiffyCalendarBox("calReleaseDate", "frmAddOrder", "txtReleaseDate","cmdReleaseDate","<%=FormatDate(Date)%>",1);
</SCRIPT>
<!--END CALENDAR -->

<script >
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
		
		function UpdateCertField(theList)
		{
			var i = 0;
			//Change the amount
			if(theList.value=='1' && document.frmMain.elements("cboprice").value=='BF') 
			{
			document.frmMain.elements("txtAmount").disabled=false;
			document.frmMain.elements("txtAmount").focus;
			document.frmMain.elements("txtQty").disabled=true;			
			}
			else
			{
			document.frmMain.elements("txtAmount").value='0';
			document.frmMain.elements("txtAmount").disabled=true;
			document.frmMain.elements("txtQty").disabled=false;
			}
			//handle the certificate
			
			for (i=0; i < theList.options.length; i++) {
				if((theList.options(i).selected))
				{
					if(theList.options(i).RequireCertificate == "True")
					{
						document.frmMain.elements("txtCert").disabled = false;
					}
					else
					{
						document.frmMain.elements("txtCert").disabled = true;
						document.frmMain.elements("txtCert").value = "";
					}
				}
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
		
		function  UpdateCompoundStatus(theChk)
		{
			var holdVal = "0"; //order to be compounded
			if (theChk.checked)
			{
				holdVal = "1";//order not to be compounded
			}
				
			document.frmMain.elements("CompoundStatus").value = holdVal;
		}
		
		function  UpdateInterBankStatus(theChk)
		{
			var holdVal = "0"; //order to be compounded
			if (theChk.checked)
			{
				holdVal = "1";//order not to be compounded
			}
				
			document.frmMain.elements("InterBankStatus").value = holdVal;
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
		document.frmMain.elements("txtPrice").focus();
		document.frmMain.elements("txtAmount").disabled=true;
		document.frmMain.elements("txtQty").disabled=false;
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
			document.frmMain.elements("txtAmount").focus();
			document.frmMain.elements("txtQty").disabled=true;
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
		//document.all.item("cboIssue").add
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

</script>
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
	
	'Response.End
	if action = "EXECUTE" then					
		Dim reloadRequired
		
		reloadRequired = false
		buttonAction = Trim(Trim(Ucase(Request.Form("cmdAdd"))))
		
		if instr(1,buttonAction,"SAVE") > 0 Or instr(1,buttonAction,"ADD") > 0 then
				Dim branch
				Dim client
				Dim orderType
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
			 	
			 	
				bond = Request.Form("cboIssue")
				compound = cint(Request.Form("CompoundStatus"))
				interbank= cint(Request.Form("InterBankStatus"))
				holdType = Request.Form("cboOrderHoldType")		'This holdType is actually order hold options ID
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
				RequiresDate = 	Request.Form("RequiresDate")			
		
		     
				if(bond="") then
				bond=0
				end if
				
				if(trim(amount)="") then
				 amount=0
				end if
				
								
				if(trim(price)="") then
				price=0
				end if
				
				if(Ucase(trim(price))="BEST") then
				best=1
				end if							
				         		
				'Check for Amount to be more than 0
				if(Cint(best)=1 and Cint(orderType) = 1 and Cdbl(amount)<=0) then
				%>
				<script language = 'vbscript'>
				         		ShowMessage "Order Detail Amount Must be more than Zero"
				         </script>
				         <% response.end
				end if
				
				 'validate Branch
				 If Trim(branch) = "" Then%>
				         <script language = 'vbscript'>
				         		ShowMessage "Please specify the Branch"
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
				 'validate Client
				 If Trim(Client) = "" Then%>
				         <script language = 'vbscript'>
				         		ShowMessage "Please specify the Client"
				         </script>
				         <% response.end
				 End If
				 'validate Order Type
				 If Trim(orderType) = "" Then%>
				         <script language = 'vbscript'>
				         		ShowMessage "Please specify the Order Type"
				         		
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
				 'validate security type
				 If Trim(secType) = "" Then%>
				         <script language = 'vbscript'>
				         		ShowMessage "Please specify the security type"
				         		
				         </script>
				         <% response.end
				 End If
				 
				 'validate Reference No.
				 'If Trim(ref) = "" Then%>
				         <script language = 'vbscript'>
				         		'ShowMessage "Please specify the Reference No."
				     	
				         </script>
				         <% 'response.end
				 'End If
				 'validate size of Reference No.
				 If Len(Ref) > 100 Then%>
				         <script language = 'vbscript'>
				         ShowMessage "Reference No. can only be 100 characters in length"
				         
				         </script>
				         <% response.end
				 End If
				 
				'check date validity
				if cdate(orderDate) < Date() then%>
				         <script language = 'vbscript'>
				         ShowMessage "The Order Date cannot be in the past"
				         
				         </script>
				         <% response.end
				 End If
				 
					'validate detail info
				         'validate Security 
				         If Trim(Security) = "" Then%>
				         		<script language = 'vbscript'>
				         				ShowMessage "Please specify the Security "
				         				
				         		</script>
				         		<% response.end
				         End If		
				         		         
				         'validate Estimated Quantity
				         If Trim(qty)= "" and Best=0 Then%>
				         		<script language = 'vbscript'>
				         				ShowMessage "Please specify the Quantity "
				         				
				         		</script>
				         		<% response.end
				         End If
				         'validate size of Order Detail Certificate Number
							If Len(Cert) > 100 Then%>
				         		<script language = 'vbscript'>
									ShowMessage "Order Detail Certificate Number can only be 100 characters in length"
									
				         		</script>
				         		<% response.end
							End If
							
							'ensure Order Detail Estimated Quantity is numeric
							If (Qty <> "") And (Not IsNumeric(Qty)) Then%>
				         		<script language = 'vbscript'>
									ShowMessage "Order Detail Estimated Quantity must be numeric"
									
				         		</script>
				         		<% response.end
							End If
				         
				         'ensure Order Detail Price is numeric
							If (Price <> "") And (Not IsNumeric(Price)) and (chkprice="P") Then%>
				         		<script language = 'vbscript'>
									ShowMessage "Order Detail Price must be numeric"
									
				         		</script>
				         		<% response.end
							End If

				 'Save only relevant dates 
				 
				 if RequiresDate = 1 then 
						releaseDate = "#" & releaseDate & "#"
				 else
						releaseDate = "NULL"
				 end if       		
				 
				 if savedate = 1 then 'save only the date check box checked
						validity = "#" & validity & "#"
				 else
						validity = "NULL"
				 end if
				 
				 set guid = server.createobject("NDUtils.CGUID")
				 guidStr = guid.GenerateGUID
				 
				 sqlStr = "INSERT INTO [tbOrder] (OrderDate,OrderHold,OrderRef,Order_DPA_,Order_EIT_,Branch_DPA_,OrderSecType_DPA_,Client_DPA_,OrderType_DPA_,OrderAutoReleaseDate,OrderHoldType_DPA_,OrderCompounded,InterBank,ChangedBy) SELECT " & "#" & FormatDate(orderDate) & "#" & " as OrderDate," & " " & hold & " " & " as OrderHold" & _
				         "," & "'" & ref & "'" & " as OrderRef," & " " & "iif(isnull(max([Order_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'tbOrder'),max([Order_DPA_]) + 1)" & " " & " as Order_DPA_" & _
				         "," & "'" & guidStr & "'" & " as Order_EIT_," & " " & branch & " " & " as Branch_DPA_," & " " & secType & " " & " as OrderSecType_DPA_" & _
				         "," & " " & client & " " & " as Client_DPA_" & _
				         "," & " " & orderType & " " & " as OrderType_DPA_" & _
				         "," & " " & FormatDate(releaseDate) & " " & " as OrderAutoReleaseDate" & _
				         "," & " " & holdType & " " & " as OrderHoldType_DPA_" & _
				         "," & " " & compound & " " & " as OrderCompounded" & _
				         "," & " " & interbank & " " & " as InterBank" & _
				         "," & " " & Userid & " " & " as ChangedBy " & _
				         " FROM [tbOrder]"
				         
				 Set conn = GetActiveConnection("KBroker")
				 conn.BeginTrans
						sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
										
						conn.Execute sqlStr
				     
						'obtain header key value
						sqlStr = "SELECT [tbOrder.Order_DPA_] FROM [tbOrder] WHERE [tbOrder.Order_EIT_] = " & "'" & guidStr & "'"
				     
						Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
						If (rs.EOF Or rs.BOF) Then%>
				         			<script language = 'vbscript'>
				         					ShowMessage "A serious error has been encountered while saving the data. Try saving again"
				         					
				         			</script>
				         			<% response.end
						End If
				     					

						'save detail data
						sqlStr = "INSERT INTO [OrdDetail] (OrdDetailCertNo,OrdDetailPrice,Amount,Best,Bond_DPA_,OrdDetailQty,OrdDetailValidity" & _
								",OrdDetail_DPA_,Order_DPA_,Security_DPA_) SELECT " & "'" & cert & "'" & " as OrdDetailCertNo" & _
								"," & "'" & price & "'" & " as OrdDetailPrice" & _
								"," & " " & amount & " " & " as Amount" & _
								"," & " " & Best & " " & " as Best" & _
								"," & " " & bond & " " & " as Bond_DPA_" & _
								"," & " " & CDbl(qty) & " " & " as OrdDetailQty" & _
								"," & "" & FormatDate(validity) & "" & " as OrdDetailValidity" & _
								"," & " " & "iif(isnull(max([OrdDetail_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'OrdDetail'),max([OrdDetail_DPA_]) + 1)" & " " & " as OrdDetail_DPA_" & _
								"," & " " & rs.Fields("Order_DPA_") & " " & " as Order_DPA_" & _
								"," & " " & security & " " & " as Security_DPA_" & _
								" FROM [OrdDetail]"										
                          
						
						sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
						
						conn.Execute = sqlStr
						
				conn.CommitTrans
				
				 'retrieve the item ID
				OrderID = rs.Fields("Order_DPA_")
				sqlStr = "SELECT OrdDetail.OrdDetail_DPA_  FROM OrdDetail WHERE OrdDetail.Order_DPA_=" & rs.Fields("Order_DPA_")
				Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
				If (rs.EOF Or rs.BOF) Then%>
				         	<script language = 'vbscript'>
				         			ShowMessage "An error has been encountered while saving the order. Try editing the Order if you wish to add more entries"
				         			
				         	</script>
				         	<% response.end
				End If
				
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
						WriteDialogRelocateScript "OrderForm.asp?order_id=" & OrderID
				elseif instr(1,buttonAction,"ADD NEW") > 0 then%>
						<SCRIPT LANGUAGE="JAVASCRIPT">
							window.parent.parent.frames['maininfo'].location.reload();
						</SCRIPT>
						<%
						WriteDialogRelocateScript "AddOrder.asp"
				else
						WritefraEnabledDialogCloseScript
				end if
				
				Response.End
		end if
   	end If
%>
		

<body Class="Dialog">

<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>


<form name = 'frmAddOrder' method = 'post' action = 'AddOrder.asp' id = 'frmMain'>


<table border="0">
  <tr>
    <td>Branch</td>
    <td><select name = 'cboBranch' id = 'cboBranch' size="1">
    	<option selected value = ''></option>
		<%
        Set conn = GetActiveConnection("KBroker")
        
        sqlStr = "SELECT * FROM [BranchList]"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
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
  </tr>
  <tr>
    <td>Date</td>
    <td><SCRIPT language="JavaScript">cal.writeControl();</SCRIPT></td>
  </tr>
  <tr>
    <td>Order Type</td>
    <td><select name = 'cboOrderType' id = 'cboOrderType' size="1" onchange='UpdateCertField(this)' >
    	<option selected value = ''></option>
		<%
        sqlStr = "SELECT * FROM [OrderTypeList]"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF
                        if cbool(rs.Fields("DefaultSelection")) then%>
								<option RequireCertificate='<%=rs.Fields("RequireCertificate")%>' selected value = '<%=rs.Fields("OrderType_DPA_")%>'><%=rs.Fields("OrderTypeName")%></option>
                        <%else%>
								<option RequireCertificate='<%=rs.Fields("RequireCertificate")%>' value = '<%=rs.Fields("OrderType_DPA_")%>'><%=rs.Fields("OrderTypeName")%></option>
                        <%end if
                        rs.MoveNext
                Loop
        End If
		%>
    </select></td>
  </tr>
  
   <tr>
    <td>Security Type</td>
    <td>
<b>

<select name = 'cboOrderSecType' id = 'cboOrderSecType' size="1" currentSecType = "S" onchange='UpdateSecurityListing(this)'>
    	 <%sqlStr = "SELECT * FROM [OrderSecTypeList]"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
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
  </tr>
  <tr>
    <td>Hold</td>
    <td>
    <select name = 'cboOrderHoldType' id = 'cboOrderHoldType' size="1" onChange="javascript: toggleDate(this)" onload"javascript: toggleDate(this)">
    <option id="0"  value="none" > None </option>
    	<%
        Dim styleHeldDate
        
        styleHeldDate = "none"
        
    	sqlStr = "SELECT * FROM [OrderHoldOptions]"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
            rs.MoveFirst
            Do Until rs.EOF
            
                if rs.Fields("DefaultSelection") then
                
                  if Cint(rs.Fields("RequiresDate")) then
                     styleHeldDate = ""
                  end if 
                %>
						<option id="<%=rs.Fields("RequiresDate")%>" selected value='<%=rs.Fields("OrderHoldOptionID")%>'><%=rs.Fields("Description")%></option>
                <%else%>
						<option id="<%=rs.Fields("RequiresDate")%>" value='<%=rs.Fields("OrderHoldOptionID")%>'><%=rs.Fields("Description")%></option>
                <%end if
                rs.MoveNext
            Loop
        End If%>
    </select>
    
    <input type="hidden" name="RequiresDate" id ="RequiresDate" value="0">
    
    <select name = 'cboHold' id = 'cboHold' size="1" style="display:none">
    	<option selected value = '1'>Yes</option>
    	<option value = '0'>No</option>
    </select></td>
  </tr>
  
  <tr id="HeldDate" style="display:<%=styleHeldDate%>">
	<td>Held Date</td>
	<td><SCRIPT language="javascript">calReleaseDate.writeControl();</SCRIPT></td>
  </tr>
  
  <tr>
    <td>Client</td>
    <td>
    <input type = 'text' name ='txtClientCode' id = 'txtClientCode' size="10" onBlur="txtval = this.value; selectItem(cboClient);">
    <select name = 'cboClient' id = 'cboClient' size="1" 
			onKeypress="return (dodefaultaction()==''); " 
			onKeydown="return (dodefaultaction()==''); " 
			onKeyup="return (UpdateCode(change(cboClient,0),cboClient,txtClientCode));" 
			onChange="UpdateCode(true,cboClient,txtClientCode);"
			onfocus="txtval = '';inputIsItemCode = 1;" 
			onblur="txtval = '';inputIsItemCode = 1;" readonly>
    	<option selected SearchCode = "" SearchText = "" value = ''></option>
		<%
		dim ClientName
		dim NameClient

        sqlStr = "SELECT * FROM [ClientList]"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF					                
                
                ClientName = rs.Fields("ClientName")
			    NameClient=rs.Fields("Client_DPA_") & " " & Mid(ClientName,1,20)
					    %>                    
                        <option SearchCode = "<%=rs.Fields("Client_DPA_")%>" SearchText = "<%=rs.Fields("ClientName")%>" value = '<%=rs.Fields("Client_DPA_")%>'><%=ClientName%></option>
                        <%rs.MoveNext
                Loop
        End If
		%>
    </select></td>
  </tr>
  <tr>
    <td>Ref No.</td>
    <td><input type = 'text' name ='txtRef' id = 'txtRef' size="20"></td>
  </tr>
  <tr>
    <td>Compound</td>
    <td><input type=checkbox   value='False' name='chkCompound' onClick = 'UpdateCompoundStatus(this);'> 
      </td>
  </tr>
  <tr>
    <td>Inter Bank</td>
    <td><input type=checkbox   value='False' name='chkInterBank' onClick = 'UpdateInterBankStatus(this);'> 
      </td>
  </tr>    
  <tr>
  <td colspan = '2'>
  
  <table border="0">
    <tr>
      <td><b><font color="#000080">Security</font></b></td>
      <td><b><font color="#000080">Bond</font></b></td>
      <td><b><font color="#000080">Quantity</font></b></td>
      <td  colspan="2"><b><font color="#000080">Price</font></b></td>
      <td><b><font color="#000080">Amount</font></b></td>
      <td><b><font color="#000080">Certificate</font></b> <b><font color="#000080">No</font></b>.</td>
      <td colspan="2"><b><font color="#000080">Validity</font></b></td>
    </tr>
    <tr>
      <td nowrap>
      <div style="position: relative; width: 50px; z-Index: -1">
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
    
    </div>
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
      <input type = 'text' name ='txtQty' id = 'txtQty' size="9" OnBlur="JavaScript: format2NumberCommasOnly(this)">
      </td>
      <td><select name="cboprice" onchange='ChangePrice(this);'>			
			<option selected SearchCode = "0" SearchText = "P" value = 'P'>P</option>			
			<option value='BF'>BF</option>			
		</select>
      </td>
      <td><input type = 'text' name ='txtPrice' id = 'txtPrice' size="9" OnKeyUp="JavaScript: updateBestPrice(this)" OnBlur="JavaScript: formatnumber(this)"></td>
      <td><input type = 'text' name ='txtAmount' id = 'txtAmount' size="9" disabled OnBlur="JavaScript: format2Number(this)"></td>
      <td><input type = 'text' name ='txtCert' id = 'txtCert' size="18"></td>
      <td>
      <input type ='checkbox' name ='chkdate' id='chkdate' size="1" class='BorderLess' onClick = 'ChangeCalendar(this);' value="ON"></td>
      <td><SCRIPT language="JavaScript">
      calValidity.writeControl();
      document.frmMain.elements("txtValidity").disabled=true; //Disable Validity date by default
      </SCRIPT></td>
    </tr>
  </table>
  
  </td>
  </tr>
 
</table>

<table>
	  <tr>
    <td width="20%" colspan=4 align=right>
		
        &nbsp;&nbsp; <input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAddMore' value=" Add Line/s " onclick = "AllowedNavigation()">
        &nbsp;&nbsp; <input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAddMoreNew' value=" Add New Order " onclick = "AllowedNavigation()">
       &nbsp;&nbsp; <input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAddPrint' value=" Save & Print " onclick = "AllowedNavigation()">
        <input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save & Close " onclick = "AllowedNavigation()">
        &nbsp;&nbsp; <input type = 'button' Class=Buttons name ='cmdClose' id = "cmdClose" value=" Close " onclick = "window.self.close();">
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		<input type = 'hidden' name ='CompoundStatus' id = 'CompoundStatus' value='0'>
		<input type = 'hidden' name ='InterBankStatus' id = 'InterBankStatus' value='0'>
		<input type = 'hidden' name ='txtcalendar' id = 'txtcalendar' value='0'>
		<input type = 'hidden' name ='txtbest' id = 'txtbest' value='0'>
		<select name = 'cboIssueList' size="1" style="display:none">    	
		<%
        sqlStr = "SELECT DISTINCT SecurityCode + ' - ' + BondIssue AS Issue,Security_DPA_,Bond_DPA_ FROM IssueList"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))        
		%>
    </select>
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
		document.all.item("RequiresDate").value = 1;
		
	}
	else
	{
		document.all.item("HeldDate").style.display = "none";
		document.all.item("RequiresDate").value = 0;
	}
  }
  </script>

</html>
<%
	function GetSecurityList(secType)
			Dim optionList
			
			sqlStr = "SELECT * FROM [SecurityList] WHERE OrderSecType_DPA_ = " & secType & " ORDER BY SecurityName"
			Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
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