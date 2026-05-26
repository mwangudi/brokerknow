<html>

<head>
<title>Edit Manage Offerings</title>
 
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
		ID = Request.Form("ID")
		
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
        
		sqlStr = "UPDATE Security SET" & _
			" SecurityAddr = '" & addr & "'"& _
			" ,SecurityCode = '" & Code & "'"& _
			" ,SecurityMktPrice = " & MktPrice & ""& _
			" ,SecurityName = '" & Name & "'"& _
			" ,OrderSecType_DPA_ = " & secType & ""& _
			" ,Sector_DPA_ = " & Sector & ""& _
			" ,Immobilised = " & immob & ""& _
			" ,Offerings = 1"& _
			" ,CanTrade = " & cantrade & ""& _
			" ,BankAccount_DPA_ = " & account & ""& _
			" ,ImportCode = '" & ImportCode & "'"& _
			" ,ClosingDate = #" & FormatDate(ClosingDate) & "#"& _
			" ,BatchSize = " & BatchSize & ""& _
			" ,OfferType_DPA_ = " & OfferType & ""& _
			" ,TimeModified = GetDate()"& _
			" ,ModifiedBy = " & UserID & ""& _
			" ,ParentSecurity_DPA_ = " & ParentSecurity & ""& _
			" ,RequiresExtra = " & Extra & ""& _
			" ,DefaultSelection = " & isDefault & ""& _
			" ,Ratio = " & theRatio & ""& _
			" ,MinimumQty = " & MinQty & ""& _
			" ,StepQty = " & StepQty & ""& _
			" ,RequiresHoldings =  "& Holdings & ""& _
			" WHERE Security_DPA_ = " & ID

'Response.Write sqlstr
'Response.End 
                   
        conn.BeginTrans			
			conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
	        			
			If isDefault = 1 Then
				SqlStr = "UPDATE Security SET DefaultSelection = 0"
				conn.Execute(sqlStr)
				
				SqlStr = "UPDATE Security SET DefaultSelection = 1 WHERE Security_DPA_ = " & ID
				conn.Execute(sqlStr)
			End If
		conn.CommitTrans
		
        Set Conn = Nothing
        
        WritefraEnabledDialogCloseScript
        Response.End 
   	end If
   	
	Set conn = GetActiveConnection("KBroker")
    
    ID = Request("ID")
    
	sqlStr = "SELECT * FROM [Security] WHERE Security_DPA_  = " & ID
        
    sqlStr = SQLServerFormat(HandleQuote(sqlStr))
    
    Set rsEdit = conn.Execute(sqlStr)
    
    If rsEdit.EOF Or rsEdit.BOF Then
		%>
		<script language = 'vbscript'>
			ShowMessage "The selected Security cannot be retrieved for editing"
		</script>
		<%
		response.end
    End If
    %>
<form name = 'frmAddSecurity' method = 'post' id="frmMain" action = "EditOffering.asp" >
<table border="0" width="100%" cellpadding=2 cellspacing=2>
	<tr>
		<td width="20%">Name</td>
		<td width="80%"><input type="text" name="txtName" id="txtName" size="35" value='<%=rsEdit("SecurityName")%>'></td>
	</tr>
	
	<tr>
		<td width="20%">Code</td>
		<td width="80%"><input type="text" name="txtCode" id="txtCode" size="25" value='<%=rsEdit("SecurityCode")%>'></td>
	</tr>
	
	<tr>
		<td width="20%">Address</td>
		<td width="80%"><textarea rows=3 name ='txtAddr' id = "txtAddr"><%=rsEdit("SecurityAddr")%></textarea></td>
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
					
					OfferType = rsEdit("OfferType_DPA_")
					
					if trim(rsdata(0,intcount)) = trim(rsEdit("OfferType_DPA_")) then
					%>
					<option selected value = '<%=trim(rsdata(0,intcount))%>'><%=trim(rsdata(1,intcount))%></option>
					<% 
					else
					%>
					<option value = '<%=trim(rsdata(0,intcount))%>'><%=trim(rsdata(1,intcount))%></option>
					<% 
					end if
					next
			end if
			%>
			</select>
		</td>		 
	</tr> 
	              
	<tr>
		<td width="20%">Offering Price</td>
		<td width="80%"><input type="text" name="txtMktPrice" id="txtMktPrice" size="25" value='<%=rsEdit("SecurityMktPrice")%>'></td>
	</tr>
	              
	<%
	if(isnull(rsEdit("ClosingDate"))) then
		ClosingDate=Date
	else
		ClosingDate = CDate(rsEdit("ClosingDate"))
	end if
	%>
	<tr>
		<td width="20%">Closing Date</td>
		<td width="80%">
			<SCRIPT language="JavaScript">			
			var cal1=new ctlSpiffyCalendarBox("cal1", "frmAddSecurity", "txtADate1","cmdDate1","<%= FormatDate(ClosingDate) %>",1);
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
			
			if trim(rsdata(0,intcount)) = trim(rsEdit("Sector_DPA_")) then
				%>
				<option selected value = '<%=trim(rsdata(0,intcount))%>'><%=trim(rsdata(1,intcount))%></option>
				<% 
			else
				%>
				<option value = '<%=trim(rsdata(0,intcount))%>'><%=trim(rsdata(1,intcount))%></option>
				<% 
			end if
			
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
			if trim(rsEdit("ParentSecurity_DPA_")) = trim(rsdata(0,intcount)) then
				%>                   						
				<option selected value = '<%=trim(rsdata(0,intcount))%>'><%=trim(rsdata(1,intcount))%></option>
				<% 
			else
				%>                   						
				<option value = '<%=trim(rsdata(0,intcount))%>'><%=trim(rsdata(1,intcount))%></option>
				<% 
			end if
			next
		end if
		%>
		</select>
		</td>
	</tr>
	
	<tr id=MinQty name=MinQty style="display:none">
		<td width="20%">Minimum Quantity</td>
		<td width="80%"><input type="text" name="txtMinQty" id="txtMinQty" size="25" value='<%=rsEdit("MinimumQty")%>'></td>
	</tr>
		
	<tr id=StepQty name=StepQty style="display:none">
		<td width="20%">Step Quantity</td>
		<td width="80%"><input type="text" name="txtStepQty" id="txtStepQty" size="25" value='<%=rsEdit("StepQty")%>'></td>
	</tr>
	
	<%
	if rsEdit("RequiresExtra") then
		chkState = "checked"
		Extra = "1"
	else
		chkState = ""
		Extra = "0"
	end if
	%>
	<tr id=Extra name=Extra style="display:none">
		<td width="20%">Requires Extra</td>
		<td width="80%"><input type=checkbox <%=chkState%> Class="BorderLess"   value='False' name='chkRequiresExtra' onClick = 'UpdateExtra(this);'></td>
	</tr>
	
	<%
	if rsEdit("RequiresHoldings") then
		chkState = "checked"
		Holdings = "1"
	else
		chkState = ""
		Holdings = "0"
	end if
	%>
	<tr id=Holdings name=Holdings style="display:none">
		<td width="20%">Requires Holdings</td>
		<td width="80%"><input type=checkbox <%=chkState%> Class="BorderLess"   value='False' name='chkRequiresHoldings' onClick = 'UpdateHoldings(this);'></td>
	</tr>
	
	<%
	if rsEdit("Ratio")>0 then
		R1 = 1
		R2 = cdbl(1/rsEdit("Ratio"))
	else
		R1 = 1
		R2 = 0
	end if
	%>
	<tr id=Ratio name=Ratio style="display:none">
		<td width="20%">Ratio</td>
		<td width="80%"><input type="text" name="txtR1" id="txtR1" size="5" value="<%=R1%>">&nbsp;for every&nbsp;<input type="text" name="txtR2" id="txtR2" size="5" value="<%=R2%>">&nbsp;held</td>
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
	
	<%
	if rsEdit("CanTrade") then
		chkState = "checked"
		CanTrade = "1"
	else
		chkState = ""
		CanTrade = "0"
	end if
	%>           
	<tr>
		<td width="20%">Can Trade</td>
		<td width="80%"><input type=checkbox <%=chkState%> Class="BorderLess" value='False' name='chkCanTrade' onClick = 'UpdateCanTrade(this);'></td>
	</tr>  
	
	<%
	if rsEdit("Immobilised") then
		chkState2 = "checked"
		Immobilised = "1"
	else
		chkState2 = ""
		Immobilised = "0"
	end if
	%> 
	<tr>
		<td width="20%">Immobilised</td>
		<td width="80%"><input type=checkbox <%=chkState2%> Class="BorderLess"   value='False' name='chkImmobilised' onClick = 'UpdateImmobilised(this);'></td>
	</tr>
	
	<%
	if rsEdit("DefaultSelection") then
		chkState3 = "checked"
		DefaultSelection = "1"
	else
		chkState3 = ""
		DefaultSelection = "0"
	end if
	%> 
	<tr>
		<td width="20%">Default selection</td>
		<td width="80%"><input type=checkbox <%=chkState3%> Class="BorderLess"   value='False' name='chkDefault' onClick = 'UpdateDefault(this);'></td>
	</tr>
	
	<tr>
		<td width="20%">Import Code</td>
		<td width="80%"><input type = 'text' name ='txtImportCode' id = 'txtImportCode' size="25" value='<%=rsEdit("ImportCode")%>'></td>
	</tr>

	<tr>
		<td width="20%">Batch Size</td>
		<td width="80%"><input type = 'text' name ='txtBatchSize' id = 'txtBatchSize' size="10" value='<%=rsEdit("BatchSize")%>'></td>
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
			<input type = 'hidden' name ='txtImmobilised' id = 'txtImmobilised' value='<%=Immobilised%>'>
			<input type = 'hidden' name ='txtCanTrade' id = 'txtCanTrade' value='<%=CanTrade%>'>
			<input type = 'hidden' name ='txtExtra' id = 'txtExtra' value='<%=Extra%>'>
			<input type = 'hidden' name ='txtHoldings' id = 'txtHoldings' value='<%=Holdings%>'>
			<input type = 'hidden' name ='txtDefault' id = 'txtDefault' value='<%=DefaultSelection%>'>
			
			<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
		</td>
	</tr>
</table>

<script language="javascript">
var offerType = <%=OfferType%>;
				 
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
</script>
</form>

</body>

</html>
