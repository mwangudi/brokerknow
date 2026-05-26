<!--#include virtual="libroutines.asp"-->
<%
	const UDLName = "KBroker"
	const DataSource = "ChequeCollection"
	const DataEntity = "ChequeCollection"
	const DataEntityPlural = "Cheque Collection"
	const ActionFolder = "Operations"
	
	Dim action
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim ID
	dim currentEntityType
	
	action = ucase(Request.QueryString("action"))
	ID = Request.Form("ID")
	currentEntityType = 1
	
	Set conn = GetActiveConnection("KBroker")
	
	select case action
		case "EXECUTE"
			
			Dim buttonAction
			Dim reloadRequired
		
			'reloadRequired = true
			       
			Dim clientVoucher
			Dim brokerVoucher
			querystr = Request.QueryString
			CollectedCheques = Request.QueryString("chqs")
			CollectionDate = Request.QueryString("Date")
			narrative =Request.QueryString("narrative")
			
			CollectedChequesArray = split(CollectedCheques,",")
			
			userID = Session("UserID")
			
			conn.BeginTrans
				for i = 0 to UBound(CollectedChequesArray)
					if Len(CollectedChequesArray(i)) <> 0 then
						sqlStr = "UPDATE Payment SET ChequeCollection = N'COLLECTED', ChequeCollectionDate = '" & FormatDate(CollectionDate) & "'," & _
						         "ChequeCollectionNarrative='" & narrative & "',ChequeCollectionUser= " & userID & " WHERE (Payment_DPA_ = " & CollectedChequesArray(i) & ")"
						
						conn.Execute sqlStr
					end if
				next
			conn.CommitTrans
			conn.Close
			Set conn = Nothing
			WriteDialogCloseScript
			%>
<Script Language="JavaScript">
				//try
				//{
					//window.opener.location.reload();
					//window.parent.frames["maininfo"].location.reload();
					//window.parent.closeDocOpener();
					//window.parent.frames["maininfo"].location.replace("ChequeCollectionList.asp");
					//alert(window.name) 
					window.self.close();
				//}
				//catch(e){}//window.self.close();}	
			</Script>
<%
			Response.End
		'	Dim clientCode
		'	
		'	clientCode = "var validNavigate = true;" & chr(13)
			%>
<script>
				<%'=clientCode%>
			</script>
			<%
		'	response.End
   	end select
   	
%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add <%=DataEntity%></title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></script>
<SCRIPT language="Javascript" src="../scripts/fhsupport.js"></SCRIPT>
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
	<script language="JavaScript" src="CALENDAR/calendar.js"></script>
	<!--END CALENDAR -->
	<%
'Client verification procedure
dim VerifyAction
'stop
VerifyAction = trim(Request.QueryString("v_action"))
tempAction = Request.Form("action")
FilteredClient = Request.QueryString("client")

if VerifyAction = "" and UCase(Trim(Request.Form("action"))) = "" then
	dim myconn, myRs
	
	Set myconn = GetActiveConnection("KBroker")
	set myRs = server.CreateObject("ADODB.Recordset")
	
	set myRs = myconn.execute("SELECT ClientIDPass, ClientCDSNo, Client_DPA_, ClientName FROM Client")
	Response.Write "<select name='AllClients' size='5' style='width: 150px;display: none'>" & vbCrLf
	if not (myRs.BOF or myRs.EOF) then
		do until myRs.EOF
			DataStr = myRs("ClientIDPass") & "|" & myRs("ClientCDSNo") & "|" & myRs("ClientName") & "|" & myRs("Client_DPA_")
			Response.Write "<option id='" & DataStr & "' value='" & myRs("Client_DPA_") & "'>" & myRs("ClientName") & "</option>" & vbCrLf
			myRs.MoveNext
		loop
	end if
	Response.Write "</select>" & vbCrLf
	%>
	<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="80%">
		<tr>
			<td width="100%" colspan="4">
				<b>Client Verification</b>
			</td>
		</tr>
		<tr>
			<td width="100%" colspan="4">
				&nbsp;
			</td>
		</tr>
		<tr>
			<td width="25%" id="ClientCodeTitle" nowrap>
			</td>
			<td width="75%" id="ClientCode" colspan="3">
			</td>
		</tr>
		<tr>
			<td width="25%" id="IDPassTitle" nowrap>
			</td>
			<td width="75%" id="IDPass" colspan="3">
			</td>
		</tr>
		<tr>
			<td width="25%" id="UserNameTitle" nowrap>
			</td>
			<td width="75%" id="UserName" colspan="3">
			</td>
		</tr>
		<tr>
			<td width="25%">
				<input type="button" name="Previous" value=" Prev " style="display: none" onClick="javascript: goPrevious()">
			</td>
			<td width="30%" id="LabelField" nowrap>
				Enter Client Code
			</td>
			<td width="30%">
				<input type="text" name="txtQueryItem" id="txtQueryItem" value="" size="20">
			</td>
			<td width="15%">
				<input type="button" name="Next" value=" Next " onClick="javascript: goNext()">
			</td>
		</tr>
		<tr>
			<td width="100%" colspan="4">
				&nbsp;
			</td>
		</tr>
		<input type="hidden" name="step_no" value="1">
	</table>
	<script language="javascript">
		//Global array for storing data entered in successive steps
		var StepData = new Array()
		StepData.length = 15;
		for (var k=0;k<15;k++)
		{
			StepData[k] = 0;
		}
		
		if(document.all.item("txtQueryItem").value!=null)
		{
			document.all.item("txtQueryItem").focus();
		}
		
		function goNext()
		{
			var QueryItem = document.all.item("txtQueryItem").value;
			var DataCombo = document.all.item("AllClients");
			var ComboID, ComboData, increment = false, tempVal;
			var step = document.all.item("step_no").value;
			
			for (var i=0;i<DataCombo.length;i++)
			{
				ComboID = DataCombo.options[i].id;
				ComboData = ComboID.split("|");
				
				switch (step)
				{
					case '1':
						if (ComboData[3]==QueryItem)
						{
							StepData[0] = QueryItem;
							document.all.item("LabelField").innerHTML = "Enter ID No/Passport No/Cert No&nbsp;&nbsp;";
							document.all.item("txtQueryItem").value = "";
							document.all.item("ClientCodeTitle").innerHTML = "Client Code&nbsp;&nbsp;";
							document.all.item("ClientCode").innerHTML = ComboData[3];
							document.all.item("Previous").style.display = "";
							increment = true;
							i = DataCombo.length;
							document.all.item("txtQueryItem").focus();
						}
						break;
					case '2':
						if (ComboData[0]==QueryItem)
						{
							StepData[1] = QueryItem;
							document.all.item("LabelField").innerHTML = "Enter CDS No";
							document.all.item("txtQueryItem").value = "";
							document.all.item("IDPassTitle").innerHTML = "ID No/Passport No/Cert No&nbsp;&nbsp;";
							document.all.item("IDPass").innerHTML = ComboData[0];
							document.all.item("Previous").style.display = "";
							increment = true;
							i = DataCombo.length;
							document.all.item("txtQueryItem").focus();
							document.all.item("UserName").innerHTML = ComboData[2];
							document.all.item("UserNameTitle").innerHTML = "Client Name&nbsp;&nbsp;";
						}
						break;
					case '3':
						if (ComboData[1]==QueryItem)
						{
							window.location.href = "AddChequeCollection.asp?v_action=go_on&client=" + ComboData[3];
							increment = true;
							i = DataCombo.length;
						}
						break;
				}
			}
			
			if (increment==true)
			{
				//Indicate one step forward moved
				document.all.item("step_no").value = parseFloat(document.all.item("step_no").value) + 1;
			}
			else
			{
				alert("Wrong detail entered. Please try again");
				document.all.item("txtQueryItem").focus();
			}
		}
	
		function goPrevious()
		{
			var QueryItem = document.all.item("txtQueryItem").value;
			var decrement = false;
			var step = document.all.item("step_no").value;
			
			switch (step)
			{
				case '2':
					document.all.item("txtQueryItem").value = StepData[0];
					document.all.item("LabelField").innerHTML = "Enter Client Code";
					document.all.item("txtQueryItem").focus();
					document.all.item("Previous").style.display = "none";
					break;
				case '3':
					document.all.item("txtQueryItem").value = StepData[1];
					document.all.item("LabelField").innerHTML = "Enter ID No/Passport No/Cert No";
					document.all.item("txtQueryItem").focus();
					break;
			}
			
			//Indicate one step backward moved
			document.all.item("step_no").value = parseFloat(document.all.item("step_no").value) - 1;
		}
	</script>
	<%
	Response.End
end if
%>
	<script language='vbscript'>
	'function EntitySelected(itemID)
 	'		frm<%=DataSource%>.elements("ID").value = itemID
 	'		frm<%=DataSource%>.elements("action").value = "Fetch_Accounts"
 	'		frm<%=DataSource%>.submit
	'end function
	</script>
	<script language='javascript'>
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
		
		var currentEntityType = <%=currentEntityType%>
		function setFilter(obj,entityType)
		{
			var IDName = "";
			var frameName = "";
			var framePageName = "";
			var voucherParamName = "";
			var saleType = "";
			var selValue = "";
			switch (entityType)
			{
			case 1:
				IDName = "Client_DPA_";
				frameName = "fraInnerSelects";
				framePageName = "inner_select_chequecollect";
				voucherParamName = "ContractClientVouchered";
				saleType = "1";
				selValue = obj.value;
				break;
			case 3:
				IDName = "BrokerCode";
				frameName = "fraInnerBrokerSelects";
				framePageName = "inner_select_chequecollect";
				voucherParamName = "ContractVouchered";
				saleType = "0";
				selValue = obj.options[obj.selectedIndex].SearchCode;
				break;
			default:
				return;
			}
			
			var fra = document.getElementById(frameName);
			
			var pageTo = framePageName + '.asp?client=' + document.all.item("cboEntity").value + '&ent_type=' + document.all.item("cboAccount").value; 
			
			if(document.all.item("cboEntity").value!=0 && document.all.item("cboAccount").value!=0)
			{
				pageTo = framePageName + '.asp';
			}
			
			fra.src = pageTo;
		}
		
		function FetchAccounts(theList)
		{
			var i = 0;
			var entity = theList.value;
			var toList = document.frmMain.cboAccount;
			
			currentEntityType = entity;
			frm = document.frmMain;				
			xmlhttp = createXMLHTTPObj();
			
			url="GetList.asp?ID="+entity+"&action=GetAccountList";
			xmlhttp.open("GET",url,true);
			xmlhttp.onreadystatechange=function() {
				if (xmlhttp.readyState==4) {
				returnStr = xmlhttp.responseText;
				returnStr = getBodyHTML(returnStr);
				
				var secList = "<select name = '" + toList.name + "' id = '" + toList.id + "' size='1' ";
				secList += "onChange='setFilter(this," + currentEntityType + ");' " ;
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
			
			var fra = document.getElementById('fraInnerSelects');		
			var pageTo = 'inner_select_vouchers.asp';  
			fra.src = pageTo;
			fra = document.getElementById('fraInnerBrokerSelects');		
			pageTo = 'inner_select_voucherBroker.asp';
			fra.src = pageTo;
			
			totalContractAmt = 0;
			document.all.item("txtTotal").value = totalContractAmt;
			document.all.item("ContractsSel").value = "";	
			
			if (entity==3){
				document.getElementById('brokerVoucherRow').style.display = '';
				document.getElementById('clientVoucherRow').style.display = 'none';
				document.getElementById('txtVoucherType').value = currentEntityType;
			} 
			
			else{
				 if (entity==1){
					document.getElementById('brokerVoucherRow').style.display = 'none';
					document.getElementById('clientVoucherRow').style.display = '';
					document.getElementById('txtVoucherType').value = currentEntityType;
				}			
				else {
					document.getElementById('brokerVoucherRow').style.display = 'none';
					document.getElementById('clientVoucherRow').style.display = 'none';
					document.getElementById('txtVoucherType').value = '0';
				}
			}	
		}
		
		var totalContractAmt = 0;
		var CheckedPayIDs = new Array()
		CheckedPayIDs.length = 20;
		var l = 0;
		
		function UpdateVoucherAmount(PayID, theAction)
		{
			if (theAction=="add")
			{
				CheckedPayIDs[l] = PayID;
			}
			else
			{
				for(var j=0;i<CheckedPayIDs.length;i++)
				{
					if(CheckedPayIDs[j]==PayID)
					{
						CheckedPayIDs.pop(j);
					}
				}
			}
			l += 1;
			//alert(CheckedPayIDs[0])
		}
		
		function SubmitForm()
		{
			var form_action = "Add<%=trim(DataSource)%>.asp?action=execute&chqs=" + CheckedPayIDs + '&narrative=' + document.all.item('txtNarrative').value + '&date=' + document.all.item('txtDate').value;
			window.location.href = form_action;
			//document.all.item("frmMain").action = form_action;
			//a/lert(form_action)
			//document.getElementById("frmMain").action=""
			//document.all.item("frmMain").action = "<%=trim(DataSource)%>.asp?chqs=" + CheckedPayIDs;
			//document.all.item("frm<%=DataSource%>").submit();
		}
	</script>
	</head>
	<body Class="Dialog">
		<div id="spiffycalendar" class="text" STYLE="z-Index: 1000">
		</div>
		<SCRIPT language="JavaScript">
	var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "txtDate","cmdDate","<%= FormatDate(Date) %>",1);
		</SCRIPT>
		<form name='frmMain' method='post' id="frmMain">
			<table border="0" width="100%" cellspacing="1" cellpadding="1">
				<tr>
					<td width="15%">
						Date
					</td>
					<td width="54%">
						<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>
					</td>
				</tr>
				<tr>
					<td>
						Client
					</td>
					<%
					 sqlStr = "SELECT * FROM [ClientList] WHERE (Client_DPA_ = " & FilteredClient & ")"
                      Set rs = conn.Execute(sqlstr)
                      If Not (rs.EOF Or rs.BOF) Then
                      ClientID = rs.fields("Client_DPA_")
                      end if
					%>
					<td>
						<input type='text' name='txtClientCode' value="<%=ClientID%>" id='txtClientCode' size="10" onBlur="txtval = this.value; selectItem(cboClient);" readonly>
						<select name='cboClient' id='cboClient' size="1" onKeypress="return (dodefaultaction()==''); " onKeydown="return (dodefaultaction()==''); " onKeyup="return (UpdateCode(change(cboClient,0),cboClient,txtClientCode));" onChange="UpdateCode(true,cboClient,txtClientCode);" onfocus="txtval = '';inputIsItemCode = 1;" onblur="txtval = '';inputIsItemCode = 1;" readonly>
							<!--option selected SearchCode = "" SearchText = "" value = ''></option-->
							<%
        
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF%>
							<option SearchCode="<%=rs.Fields("Client_DPA_")%>" SearchText="<%=rs.Fields("ClientName")%>" value='<%=rs.Fields("Client_DPA_")%>'>
								<%=rs.Fields("ClientName")%>
							</option>
							<%rs.MoveNext
                Loop
        End If
		%>
						</select>
					</td>
				</tr>
				<tr id="brokerVoucherRow" style="display: none">
					<td width="15%" valign="top">
						Cheques
					</td>
					<td width="54%">
						<iframe id="fraInnerBrokerSelects" name="fraInnerBrokerSelects" width="400px" height="200px" src="inner_select_chequecollect.asp?ent_type=<%=FilteredClient%>"></iframe>
						<BR>
					</td>
					<td width="31%">
						<input type='hidden' name='txtVoucherType' id='txtVoucherType' value="1">
					</td>
				</tr>
				<tr id="clientVoucherRow">
					<td width="15%" valign="top">
						Cheques
					</td>
					<td width="54%">
						<iframe id="fraInnerSelects" name="fraInnerSelects" width="400px" height="200px" src="inner_select_chequecollect.asp?ent_type=<%=FilteredClient%>"></iframe>
						<BR>
					</td>
					<td width="31%">
					</td>
				</tr>
				<tr>
					<td width="15%">
						Narrative
					</td>
					<td width="54%">
						<textarea name='txtNarrative' id='txtNarrative'></textarea>
					</td>
					<td width="31%">
					</td>
				</tr>
				<tr>
					<td width="100%" colspan="3" align="right" valign="absBottom">
						<BR>
						<BR>
						<input type='button' Class="Buttons" name='cmdAdd' id='cmdAdd' value="Save" onclick="javascript: SubmitForm()">
						<input type='button' Class="Buttons" name='cmdCancel' id="cmdCancel" value="Cancel" onclick="JavaScript: window.self.close()">
						<input type='hidden' name='action' id='action' value="Execute"> <input type='hidden' name='ID' id='ID'>
					</td>
				</tr>
			</table>
		</form>
	</body>
</html>
