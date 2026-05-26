<html>

<head>
<title>Manage Offerings</title>
 
<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css"> 
<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT>

<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<script language="JavaScript" src="CALENDAR/calendar.js"></script>
<!--END CALENDAR -->

<script language="javascript">
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

function  UpdateExtra(theChk)
	{
		if (theChk.checked)
		{
			document.frmMain.elements("txtExtra").value = "1";
		}
		else
		{
			document.frmMain.elements("txtExtra").value = "0";
		}
	}

function  UpdateHoldings(theChk)
	{
		if (theChk.checked)
		{
			document.frmMain.elements("txtHoldings").value = "1";
		}
		else
		{
			document.frmMain.elements("txtHoldings").value = "0";
		}
	}

function  UpdateDefault(theChk)
	{
		if (theChk.checked)
		{
			document.frmMain.elements("txtDefault").value = "1";
		}
		else
		{
			document.frmMain.elements("txtDefault").value = "0";
		}
	}
		
function showParentsecurity()
	{
	var offerType = document.getElementById("cboOfferType").value;
				 
	if (offerType==2)
		{
		document.getElementById("ParentTitle").style.display = '';
		document.getElementById("MinQty").style.display = '';
		document.getElementById("StepQty").style.display = '';
		document.getElementById("Extra").style.display = '';
		document.getElementById("Holdings").style.display = '';
		document.getElementById("Ratio").style.display = '';
		}
	else
		{
		document.getElementById("ParentTitle").style.display = 'none';
		document.getElementById("MinQty").style.display = 'none';
		document.getElementById("StepQty").style.display = 'none';
		document.getElementById("Extra").style.display = 'none';
		document.getElementById("Holdings").style.display = 'none';
		document.getElementById("Ratio").style.display = 'none';
		}
	}

function ClearFields(element)
		{
		
			document.frmMain.elements("txtAvailableCredit").value = '';
			document.frmMain.elements("txtCurrentBal").value = '';
			
		   if (element == 'txtClientCode')
		   {
			document.frmMain.elements("txtClientCode").value = '';
			document.frmMain.elements("txtCdsNo").value = 'CSD No.';
			document.frmMain.elements("txtclientname").value = 'Client Name';
			return;
		   }
		   if (element == 'txtCdsNo')
		   {
			document.frmMain.elements("txtClientCode").value = 'Code';
			document.frmMain.elements("txtCdsNo").value = '';
			document.frmMain.elements("txtclientname").value = 'Client Name';
			return;
		   }
		   if (element == 'txtclientname')
		   {
		    document.frmMain.elements("txtclientname").value = '';
			document.frmMain.elements("txtClientCode").value = 'Code';
			document.frmMain.elements("txtCdsNo").value = 'CSD No.';

			return;
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
		 var clientcode = document.frmMain.elements("txtClientCode").value;
		 var clientcds = document.frmMain.elements("txtCdsNo").value;
		 var clientname = document.frmMain.elements("txtclientname").value;
		 //alert();
		 var clientcobo = document.getElementById("cboClient");		 
		 
		 var guid = Math.random();     
			
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
				
				url="GetList.asp?clientcode="+clientcode+"&cdsno="+clientcds+"&clientname="+clientname+"&action=SLoadClient&guidstr="+guid;
				
				//alert(url);
				
				var x_clientname;
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
									//document.frmMain.elements("txtContact").value = myArray[6];
									//document.frmMain.elements("txtAgent").value = myArray[2];
									//document.frmMain.elements("AgentID").value = myArray[4];
									//document.frmMain.elements("txtAccManager").value = myArray[3];
									//document.frmMain.elements("AccManagerID").value = myArray[8];
									
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
										document.frmMain.elements("txtAvailableCredit").value = myArray[0];
									document.frmMain.elements("txtCurrentBal").value = myArray[1];
										/*document.frmMain.elements("txtAvailableCredit").value = myArray[0];
										document.frmMain.elements("txtCurrentBal").value = myArray[1];
										document.frmMain.elements("txtAgent").value = myArray[2];
										document.frmMain.elements("AgentID").value = myArray[4];
										document.frmMain.elements("txtAccManager").value = myArray[3];
										document.frmMain.elements("AccManagerID").value = myArray[8];*/
																			
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
		
	action = ucase(Request.Form("action"))
	
	UserID = Session("UserID")
	
	if action = "EXECUTE" then
		immob = trim(Request.Form("txtImmobilised"))
		code = trim(Request.Form("txtCode"))
		name = trim(Request.Form("txtName"))
		addr = trim(Request.Form("txtAddr"))
		mktPrice = trim(Request.Form("txtMktPrice"))
		adate = trim(Request.Form("txtADate"))
		fee = trim(Request.Form("txtFee"))
		Sector= trim(Request.Form("cboSector"))
		importCode= trim(Request.Form("txtimportCode"))
		ClosingDate = trim(Request.Form("txtADate1"))
		cantrade = trim(Request.Form("txtCanTrade"))
		account = trim(Request.Form("cbobank"))
		BatchSize =  trim(Request.Form("txtBatchSize")) 
		OfferType = trim(Request.Form("cboOfferType"))
		ParentSecurity = trim(Request.Form("cboParentSecurity"))
        
		MinQty = trim(Request.Form("txtMinQty"))
		StepQty = trim(Request.Form("txtStepQty"))
		
		Extra = trim(Request.Form("txtExtra"))
		Holdings = trim(Request.Form("txtHoldings"))
		isDefault = trim(Request.Form("txtDefault"))
		
		numerator = trim(Request.Form("txtR1"))
		denominator = trim(Request.Form("txtR2"))

        secType =  2
      
        'validate Name
        If Trim(Name) = "" Then%>
				<script language="vbscript">
                						ShowMessage "Please specify the Name"
								</script>
				<% response.end
        End If
        
        'validate Market Price
        If Trim(mktPrice) = "" Then%>
				<script language="vbscript">
                						ShowMessage "Please specify the Market Price"
								</script>
				<% response.end
        End If
        'validate size of Address
        If Len(Addr) > 100 Then%>
				<script language="vbscript">
										ShowMessage "Address can only be 100 characters in length"
								</script>
				<% response.end
        End If
        'validate size of Code
        If Len(Code) > 25 Then%>
				<script language="vbscript">
										ShowMessage "Code can only be 25 characters in length"
								</script>
				<% response.end
		End If
        'ensure Market Price is numeric
        If (MktPrice <> "") And (Not IsNumeric(MktPrice)) Then%>
			<script language="vbscript">
										ShowMessage "Market Price  must be numeric"
							</script>
			<% response.end
        End If
        'validate size of Name
        If Len(Name) > 100 Then%>
			<script language="vbscript">
										ShowMessage "Name can only be 100 characters in length"
							</script>
			<% response.end
        End If
        'validate Offer Type
        If Trim(OfferType) = "" Then%>
				<script language="vbscript">
                						ShowMessage "Please specify the Offer Type."
        						</script>
				<% response.end
        End If
        
     	'validate Security Type
        If Trim(secType) = "" Then%>
				<script language="vbscript">
                						ShowMessage "Please specify the Security Type"
        						</script>
				<% response.end
        End If
        'validate Transfer Fee
        If Trim(Fee) = "" Then%>
        		<script language = 'vbscript'>
								        ShowMessage "Please specify the Fee"
        		</script>
        		<% response.end
        End If
        'ensure Transfer Fee is numeric
        If (Fee <> "") And (Not IsNumeric(Fee)) Then%>
        		<script language = 'vbscript'>
						                ShowMessage "Fee must be numeric"
        		</script>
        		<% response.end
        End If
        'validate Batch Size
        If Trim(BatchSize) = "" Then%>
        		<script language = 'vbscript'>
								        ShowMessage "Please specify the batch size."
        		</script>
        		<% response.end
        End If
        'ensure Batch Size is numeric
        If (BatchSize <> "") And (Not IsNumeric(BatchSize)) Then%>
        		<script language = 'vbscript'>
						                ShowMessage "Batch size must be numeric"
        		</script>
        		<% response.end
        End If
        'Validate Parent security
        if OfferType = 2 then 'Rights Issue
			If Trim(ParentSecurity) = "" Then%>
        			<script language = 'vbscript'>
										ShowMessage "Please specify the parent security."
        			</script>
        			<% response.end
			 End If
			 
			If Trim(MinQty) = "" Then
				%>
				<script language = 'vbscript'>
					ShowMessage "Please specify the Minimum Quantity"
				</script>
				<% response.end
			End If

			If (MinQty <> "") And (Not IsNumeric(MinQty)) Then
				%>
				<script language = 'vbscript'>
					ShowMessage "Minimum Quantity must be numeric"
				</script>
				<% response.end
			End If

			If Trim(StepQty) = "" Then
				%>
				<script language = 'vbscript'>
					ShowMessage "Please specify the Step Quantity"
				</script>
				<% response.end
			End If

			If (StepQty <> "") And (Not IsNumeric(StepQty)) Then
				%>
				<script language = 'vbscript'>
					ShowMessage "Step Quantity must be numeric"
				</script>
				<% response.end
			End If     

			If Trim(numerator) = "" Then
				%>
				<script language = 'vbscript'>
					ShowMessage "Please make sure that the ratio is correct"
				</script>
				<% response.end
			End If

			If (numerator <> "") And (Not IsNumeric(numerator)) Then
				%>
				<script language = 'vbscript'>
					ShowMessage "Ratio values must be numeric"
				</script>
				<% response.end
			End If

			If Trim(denominator) = "" Then
				%>
				<script language = 'vbscript'>
					ShowMessage "Please make sure that the ratio is correct"
				</script>
				<% response.end
			End If

			If (denominator <> "") And (Not IsNumeric(denominator)) Then
				%>
				<script language = 'vbscript'>
					ShowMessage "Ratio values must be numeric"
				</script>
				<% response.end
			End If
			
			if (numerator > 0) and (denominator > 0) then
				theRatio = Numerator/Denominator
			else
				theRatio = 1
			end if
        else 'IPO
			ParentSecurity = "NULL"
			MinQty = 0
			StepQty = 0
			theRatio = 0
        end if
                
        set guid = server.createobject("NDUtils.CGUID")
        guidStr = guid.GenerateGUID
        
        Set conn = GetActiveConnection("KBroker")
        
		'***************************** Changes made by Peter Muchiri ***********************
        ' The IIF function is not standard ANSI SQL. Standard ANSI SQL, and T-SQL, support the CASE expression which can do
        'whatever the IIF function can do in Access.
		'***********************************************************************************

       sqlStr = "INSERT INTO [Security] (SecurityAddr,SecurityCode,SecurityMktPrice,SecurityName" & _
                "       ,Security_DPA_,OrderSecType_DPA_,Sector_DPA_,Immobilised,Offerings,CanTrade,BankAccount_DPA_" & _
                ",ImportCode,Security_EIT_,ClosingDate,BatchSize,OfferType_DPA_,TimeModified,ModifiedBy,ParentSecurity_DPA_,RequiresExtra,DefaultSelection,Ratio,MinimumQty,StepQty,RequiresHoldings) " & _
                " SELECT " & "'" & addr & "'" & " as SecurityAddr," & "'" & Code & "'" & " as SecurityCode" & _
                "       ," & " " & MktPrice & " " & " as SecurityMktPrice" & _
                "       ," & "'" & Name & "'" & " as SecurityName," & _
                "		" & "case isnull(max([Security_DPA_]),0)  when 0 then (SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Security') " & _
                "		else max([Security_DPA_]) + 1 End as Security_DPA_ " & _
                "		," & " " & secType & " " & " as OrderSecType_DPA_" & _
                "		," & " " & Sector & " " & " as Sector_DPA_" & _
                "		," & " " & immob & " " & " as Immobilised" & _
				"		," & " " & 1 & " " & " as Offerings" & _
				"		," & " " & cantrade & " " & " as CanTrade" & _
				"		," & " " & account & " " & " as BankAccount_DPA_" & _				
                "		," & "'" & ImportCode & "'" & " as ImportCode" & _
                "       ," & "'" & guidStr & "'" & " as Security_EIT_ " & _
				"		," & "#" & FormatDate(ClosingDate) & "#" & " as ClosingDate" & _
				"		," & " " & BatchSize & " " & " as BatchSize" & _
				"		," & " " & OfferType & " " & " as OfferType_DPA_" & _
				"		," & "#" & FormatDate(Date()) & "#" & " as TimeModified" & _
				"		," & " " & UserID & " " & " as ModifiedBy" & _
				"		," & " " & ParentSecurity & " " & " as ParentSecurity_DPA_" & _
				"		," & " " & Extra & " " & " as RequiresExtra" & _
				"		," & " " & isDefault & " " & " as DefaultSelection" & _
				"		," & " " & theRatio & " " & " as Ratio" & _
				"		," & " " & MinQty & " " & " as MinimumQty" & _
				"		," & " " & StepQty & " " & " as StepQty" & _
				"		," & " " & Holdings & " " & " as RequiresHoldings" & _
				"       FROM [Security]"

'Response.Write sqlstr
'Response.End 
                   
        conn.BeginTrans			
			conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
	        			
			sqlStr = "SELECT [Security.Security_DPA_] FROM [Security] WHERE [Security.Security_EIT_] = " & "'" & guidStr & "'"
	        
			Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
			If (rs.EOF Or rs.BOF) Then%>
					<script language = 'vbscript'>
                			ShowMessage "A serious error has been encountered while saving the data. Try saving again"
            		</script>
					<% response.end
			End If
	        
			sqlStr = "INSERT INTO [SecTransFee] (SecTransFeeADate,SecTransFeeFee,SecTransFee_DPA_" & _
                "                       ,Security_DPA_) SELECT " & "#" & adate & " " & Time & "#" & " as SecTransFeeADate" & _
                "       ," & " " & fee & " " & " as SecTransFeeFee" & _
                "       ," & " " & "case isnull(max([SecTransFee_DPA_]),0)  when 0 then (SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'SecTransFee') " & _
                "  else max([SecTransFee_DPA_]) + 1 End as SecTransFee_DPA_ " & _
                "       ," & " " & rs.Fields("Security_DPA_") & " " & " as Security_DPA_" & _
                "        FROM [SecTransFee]"
	
			conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
			
			If isDefault = 1 Then
				SqlStr = "UPDATE Security SET DefaultSelection = 0"
				conn.Execute(sqlStr)
				
				SqlStr = "UPDATE Security SET DefaultSelection = 1 WHERE Security_DPA_ = " & rs.Fields("Security_DPA_")
				conn.Execute(sqlStr)
			End If
		conn.CommitTrans
		
        Set Conn = Nothing
        
        WritefraEnabledDialogCloseScript
        Response.End 
   	end If
%>
<form name = 'frmAddSecurity' method = 'post' id="frmMain" action = "AddOffering.asp" >
<table border="0" width="100%" cellpadding=2 cellspacing=2>
	<tr>
		<td width="20%">Name</td>
		<td width="80%"><input type="text" name="txtName" id="txtName" size="25"></td>
	</tr>
	
	<tr>
		<td width="20%">Code</td>
		<td width="80%"><input type="text" name="txtCode" id="txtCode" size="25"></td>
	</tr>
	
	<tr>
		<td width="20%">Address</td>
		<td width="80%"><textarea rows=3 name ='txtAddr' id = "txtAddr"></textarea></td>
	</tr>
	
	<tr>
		<td width="20%">Offer Type</td>
		<td width="20%">
			<select name = 'cboOfferType' id = 'cboOfferType' size="1" onChange="javascript: showParentsecurity();">
			<%
			Set conn = GetActiveConnection("KBroker")
						
			sqlStr = "SELECT OfferType_DPA_,Description FROM [OfferTypeList] Order By Description ASC"
			Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
								
			intrscount = rs.recordcount
								
			if intrscount > 0 then
										
					rs.movefirst
										
					rsdata = rs.getrows()
										
					for intcount = 0 to intrscount-1
					%>
					<option value = '<%=trim(rsdata(0,intcount))%>'><%=trim(rsdata(1,intcount))%></option>
					<% 
					next
			end if
			%>
			</select>
		</td>		 
	</tr> 
	              
	<tr>
		<td width="20%">Offering Price</td>
		<td width="80%"><input type="text" name="txtMktPrice" id="txtMktPrice" size="25"></td>
	</tr>
	              
	<tr>
		<td width="20%">Activation Date</td>
		<td width="80%">
			<SCRIPT language="JavaScript">			
			var cal=new ctlSpiffyCalendarBox("cal", "frmAddSecurity", "txtADate","cmdDate","<%= FormatDate(Date) %>",1);
			cal.writeControl();
			</SCRIPT>
		</td>
	</tr>          
	
	<tr>
		<td width="20%">Closing Date</td>
		<td width="80%">
			<SCRIPT language="JavaScript">			
			var cal1=new ctlSpiffyCalendarBox("cal1", "frmAddSecurity", "txtADate1","cmdDate1","<%= FormatDate(Date) %>",1);
			cal1.writeControl();
			</SCRIPT>
		</td>
	</tr>
	              
	<tr>
		<td width="20%">Market Sector</td>
		<td width="20%">
		<select name = 'cboSector' id = 'cboSector' size="1">
		<%
		sqlStr = "SELECT Sector_DPA_,ShortDescription FROM [MarketSectorList] Order By ShortDescription ASC"
		Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
							
		intrscount = rs.recordcount
							
		if intrscount > 0 then
			rs.movefirst
								
			rsdata = rs.getrows()
								
			for intcount = 0 to intrscount-1
			%>
			<option value = '<%=trim(rsdata(0,intcount))%>'><%=trim(rsdata(1,intcount))%></option>
			<% 
			next
		end if
		%>
		</select>
		</td>		 
	</tr>

	<tr id=ParentTitle name=ParentTitle style="display:none">
		<td width="18%" >Parent Security</td>
		<td width="82%">
		<select name = 'cboParentSecurity' id = 'cboParentSecurity' size="1">
		<option value = ''></option>
		<% 
		sqlStr = " SELECT security_DPA_,securityname FROM [SecurityList] " & _
		" WHERE     (OrderSecType_DPA_ = 2) AND (OfferType IS NULL) " & _
		" order by securityname "
													 
		Set rs = conn.Execute(sqlStr)
											
		intrscount = rs.recordcount
											
		if intrscount > 0 then
			rs.movefirst
			rsdata =  rs.getrows()
											  
			for intcount = 0 to intrscount-1
			%>                   						
			<option value = '<%=trim(rsdata(0,intcount))%>'><%=trim(rsdata(1,intcount))%></option>
			<% 
			next
		end if
		%>
		</select>
		</td>
	</tr>
	
	<tr id=MinQty name=MinQty style="display:none">
		<td width="20%">Minimum Quantity</td>
		<td width="80%"><input type="text" name="txtMinQty" id="txtMinQty" size="25"></td>
	</tr>
		
	<tr id=StepQty name=StepQty style="display:none">
		<td width="20%">Step Quantity</td>
		<td width="80%"><input type="text" name="txtStepQty" id="txtStepQty" size="25"></td>
	</tr>
	
	<tr id=Extra name=Extra style="display:none">
		<td width="20%">Requires Extra</td>
		<td width="80%"><input type=checkbox Class="BorderLess"   value='False' name='chkRequiresExtra' onClick = 'UpdateExtra(this);'></td>
	</tr>
	
	<tr id=Holdings name=Holdings style="display:none">
		<td width="20%">Requires Holdings</td>
		<td width="80%"><input type=checkbox Class="BorderLess"   value='False' name='chkRequiresHoldings' onClick = 'UpdateHoldings(this);'></td>
	</tr>
	
	<tr id=Ratio name=Ratio style="display:none">
		<td width="20%">Ratio</td>
		<td width="80%"><input type="text" name="txtR1" id="txtR1" size="5">&nbsp;for every&nbsp;<input type="text" name="txtR2" id="txtR2" size="5">&nbsp;held</td>
	</tr>
		
	<tr>
		<td width="20%">Account</td>
		<td width="20%">
		<select name = 'cbobank' id = 'cbobank' size="1">
		<%
		sqlStr = "SELECT Account_DPA_,AccountName FROM [AccountList] Order By AccountName ASC"
		Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
						    
		intrscount = rs.recordcount
							
		if intrscount > 0 then
								
			rs.movefirst
								
			rsdata = rs.getrows()
								
			for intcount = 0 to intrscount-1
			%>
			<option value = '<%=trim(rsdata(0,intcount))%>'><%=trim(rsdata(1,intcount))%></option>
			<% 
			next
		end if
		%>
		</select>
		</td> 
	</tr>
	              
	<tr>
		<td width="20%">Can Trade</td>
		<td width="80%"><input type=checkbox Class="BorderLess"   value='False' name='chkCanTrade' onClick = 'UpdateCanTrade(this);'></td>
	</tr>  

	<tr>
		<td width="20%">Immobilised</td>
		<td width="80%"><input type=checkbox Class="BorderLess"   value='False' name='chkImmobilised' onClick = 'UpdateImmobilised(this);'></td>
	</tr>

	<tr>
		<td width="20%">Default selection</td>
		<td width="80%"><input type=checkbox Class="BorderLess"   value='False' name='chkDefault' onClick = 'UpdateDefault(this);'></td>
	</tr>
	
	<tr>
		<td width="20%">Transfer Fee</td>
		<td width="80%"><input type = 'text' name ='txtFee' id = 'txtFee' size="25"></td>
	</tr>

	<tr>
		<td width="20%">Import Code</td>
		<td width="80%"><input type = 'text' name ='txtImportCode' id = 'txtImportCode' size="25"></td>
	</tr>

	<tr>
		<td width="20%">Batch Size</td>
		<td width="80%"><input type = 'text' name ='txtBatchSize' id = 'txtBatchSize' size="10"></td>
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
			<input type = 'hidden' name ='txtImmobilised' id = 'txtImmobilised' value='0'>
			<input type = 'hidden' name ='txtCanTrade' id = 'txtCanTrade' value='0'>
			<input type = 'hidden' name ='txtExtra' id = 'txtExtra' value='0'>
			<input type = 'hidden' name ='txtHoldings' id = 'txtHoldings' value='0'>
			<input type = 'hidden' name ='txtDefault' id = 'txtDefault' value='0'>
		</td>
	</tr>
</table>

</form>

</body>

</html>
