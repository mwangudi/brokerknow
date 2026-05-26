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
	Dim rsClient
	Dim ID
	dim currentEntityType
	Dim ClientID
	Dim EditNarrative
	action = ucase(Request.Form("action"))
	ID = Request("ID")
	
	Set conn = GetActiveConnection("KBroker")
	
	Set rsClient = Server.CreateObject("ADODB.Recordset")
		
	rsClient.CursorLocation = adUseClient
		
	sqlStr="Select Entity_DPA_,ChequeCollectionNarrative From Payment Where(Payment_DPA_=" & ID & ")"
	
	set rsClient=conn.execute(sqlStr)
	if not(rsClient.eof and rsClient.bof) then
	ClientID=rsClient("Entity_DPA_")
	EditNarrative=rsClient("ChequeCollectionNarrative")
	end if
	
	currentEntityType = 1
	
	
	
	
	select case action
		case "EXECUTE"
			
			Dim buttonAction
			Dim reloadRequired
		
			'reloadRequired = true		
			       
			Dim clientVoucher
			Dim brokerVoucher
			Dim narrative
			Dim PaymentID
			
			narrative=Request.Form("txtNarrative")
						
			userID = Session("UserID")
			PaymentID=Request.Form("ID")
			
			conn.BeginTrans
						sqlStr = "UPDATE Payment SET ChequeCollectionNarrative='" & narrative & "',ChequeCollectionModifyUser="& userID & _
						         ",ChequeCollectionModifyDate=GetDate() Where(Payment_DPA_=" & PaymentID & ")"				
				
						conn.Execute sqlStr
				
			conn.CommitTrans
			conn.Close
			Set conn = Nothing
			WritefraEnabledDialogCloseScript
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
			//
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
						//j = CheckedPayIDs.length;
					}
				}
			}
			l += 1;
			//alert(CheckedPayIDs[0])
		}
		
		function SubmitForm()
		{
			var form_action = "Edit<%=trim(DataSource)%>.asp?action=execute&chqs=" + CheckedPayIDs + '&date=' + document.all.item('txtDate').value;
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
		<form name='frmMain' method='post' id="frmMain" action = "EditChequeCollection.asp">
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
					<td>						
						<select name='cboClient' id='cboClient' size="1" readonly=true>
							<!--option selected SearchCode = "" SearchText = "" value = ''></option-->
							<%
        sqlStr = "SELECT * FROM [ClientList] where(Client_DPA_=" & ClientID & ")"
        Set rs = conn.Execute(sqlstr)
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
						<iframe id="fraInnerBrokerSelects" name="fraInnerBrokerSelects" width="400px" height="200px" src="inner_select_chequecollect.asp?action=edit&ent_type=<%=ClientID%>"></iframe>
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
						<iframe id="fraInnerSelects" name="fraInnerSelects" width="400px" height="200px" src="inner_select_chequecollect.asp?action=edit&ent_type=<%=ClientID%>"></iframe>
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
						<textarea name='txtNarrative' id='txtNarrative'><%=EditNarrative%></textarea>
					</td>
					<td width="31%">
					</td>
				</tr>
				<tr>
					<td width="100%" colspan="3" align="right" valign="absBottom">
						<BR>
						<BR>
						<!--onclick="javascript: SubmitForm()"-->
						<input type='Submit' Class="Buttons" name='cmdAdd' id='cmdAdd' value="Save">
						<input type='button' Class="Buttons" name='cmdCancel' id="cmdCancel" value="Cancel" onclick="JavaScript: window.self.close()">
						<input type='hidden' name='action' id='action' value="Execute"> 
						<input type='hidden' name='ID' id='ID'value='<%=ID%>'>
					</td>
				</tr>
			</table>
		</form>
	</body>
</html>
