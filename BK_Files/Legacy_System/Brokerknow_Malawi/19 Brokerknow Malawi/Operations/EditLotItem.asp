<!--#include file="../libroutines.asp"-->

	<%
	const UDLName = "KBroker"
	const DataSource = "EditLot"
	const DataEntity = "Lot"
	const DataEntityPlural = "Lots"
	const ActionFolder = "Operations"
	
	const LinkedIndependent = 1
    const LinkedDependent = 2
	
	Dim action
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim guidStr 
	Dim guid
	Dim ID
	Dim ItemID
	Dim rsEdit
	Dim IDHolder
	Dim IDArray
	
	UserId=Session("UserID")
	
	action = ucase(Request.Form("action"))
	IDHolder = Request("ID")
	
	'action = "EXECUTE_DETAIL"
	'IDHolder = "3320<->7171"
	
	'IDHolder = "1<->2"
	
	If (Trim(IDHolder) = "") or (IDHolder = "0") Then%>
			<script language = 'vbscript'>
                	ShowMessage "Please select an Order item for Lot allocation"
                	window.self.close
			</script>
			<%response.end
	End If
	
	IDArray = split(IDHolder,"<->")
	ID = IDArray(lbound(IDArray))
	
	'ItemID = IDArray(ubound(IDArray))
	ItemID = Request.Form("ItemID")
	
	'ItemID = -1
	
	if action <> "" then
			If Trim(ItemID) = "" Then%>
					<script language = 'vbscript'>
            				ShowMessage "No item specified"                				
					</script>
					<%response.end
			End If
	end if
	
	select case action 
		case "EXECUTE_DETAIL"

				
			Dim slip
			Dim broker
			Dim tDate
			Dim qty
			Dim varBalanceQty
			Dim price
			Dim orderType
			Dim orderIsSaleType
			Dim securityID
			Dim commission	
			Dim orderSecType
			Dim agentCommission
			Dim staffCommission
			Dim volComm
			Dim volBound
			Dim minComm
			Dim cma
			Dim imobRate
			Dim secImob
			Dim regularComm			
			Dim interbank
			Dim ContractDPA
			Dim custodian
			Dim client
			Dim clientEntity
						
			Dim pDate
			Dim payRS
			Dim ContractsSel
			Dim receiptVoucher
			Dim amount
			Dim bank
			Dim entity
			Dim account
			Dim custOrder
			Dim clientVoucher
			Dim ClientClass  
			Dim SettlementDate

			if itemID = "-1" then
				broker = Request.Form("cboBroker")
				tDate = Trim(Request.Form("txtTDate"))
				tDate = tDate & " " & Time
				slip = Request.Form("txtSlip")
				qty = Request.Form("txtQty")
				varBalanceQty = Request.Form("BalanceQty")
				price = Request.Form("txtPrice")
				SettlementDate = Trim(Request.Form("txtSDate"))
				SettlementDate = SettlementDate & " " & Time
			else
				broker = Request.Form("cboBrokerInPlace")
				tDate = Trim(Request.Form("Date"))
				tDate = tDate & " " & Time
				slip = Request.Form("Slip")
				qty = Request.Form("Quantity")
				varBalanceQty = Request.Form("BalanceQty")
				price = Request.Form("Price")
				SettlementDate = Request.Form("SettlementDate")
				SettlementDate = SettlementDate & " " & Time
			end if
						
			sDate = FormatDate(SettlementDate)
						
			orderType = Request.Form("txtOrderType")
			orderIsSaleType = cbool(Request.Form("txtOrderIsSaleType"))
			securityID = Request.Form("txtSecurityID")
			regularComm = Request.Form("txtCommission")
			orderSecType = Request.Form("txtInstrument")
			agentCommission = Request.Form("txtAgentCommission")
			staffCommission = Request.Form("txtStaffCommission")
			volComm = Request.Form("txtVolumeCommission")
			volBound = ccur(Request.Form("txtVolumeBoundary"))
			minComm = ccur(Request.Form("txtMinimumCommission"))
			cma = Request.Form("txtCMA")
			imobRate = Request.Form("txtPostImmobilisedRate")
			secImob = Request.Form("txtSecurityImmobilised")
			interbank=Cint(Request.Form("txtinterbank"))       
			custodian=Cint(Request.Form("txtcustodian"))       
			client =Request.Form("txtClientDPA")       
			cliententity =Cint(Request.Form("txtEntityDPA"))       
			clientClass =Cint(Request.Form("txtClass"))
						
			Set conn = GetActiveConnection("KBroker")
			          
			'validate Broker
			If Trim(Broker) = "" Then%>
				<script language = 'vbscript'>
					ShowMessage "Please specify the Broker"
				</script>
				<SCRIPT LANGUAGE="JAVASCRIPT">					
					window.parent.frames("detail").location="EditLotItem.asp?ID=" +<%= ID%>;
				</SCRIPT>
				<% 
				
				response.end
			End If
			'validate Slip
			If Trim(Slip) = "" Then%>
				<script language = 'vbscript'>
					ShowMessage "Please specify the Ref No."
				</script>
				<SCRIPT LANGUAGE="JAVASCRIPT">					
					window.parent.frames("detail").location="EditLotItem.asp?ID=" +<%= ID%>;
				</SCRIPT>
				<% response.end
			End If

			'validate Estimated Price
			If Trim(Price) = "" Then%>
				<script language = 'vbscript'>
					ShowMessage "Please specify the Price "
				</script>
				<SCRIPT LANGUAGE="JAVASCRIPT">					
					window.parent.frames("detail").location="EditLotItem.asp?ID=" +<%= ID%>;
				</SCRIPT>
			<% response.end
			End If

			'validate Estimated Quantity
			If Trim(qty)= "" Then%>
				<script language = 'vbscript'>
					ShowMessage "Please specify the Quantity "
				</script>
				<SCRIPT LANGUAGE="JAVASCRIPT">					
					window.parent.frames("detail").location="EditLotItem.asp?ID=" +<%= ID%>;
				</SCRIPT>
				<% response.end
			End If
			'ensure Order Detail Estimated Price is numeric
			If (Price <> "") And (Not IsNumeric(Price)) Then%>
				<script language = 'vbscript'>
					ShowMessage "Price must be a number"
				</script>
				<SCRIPT LANGUAGE="JAVASCRIPT">					
					window.parent.frames("detail").location="EditLotItem.asp?ID=" +<%= ID%>;
				</SCRIPT>
				<% response.end
			End If
			'ensure Order Detail Estimated Quantity is numeric
			If (Qty <> "") And (Not IsNumeric(Qty)) Then%>
				<script language = 'vbscript'>
					ShowMessage "Quantity must be a number"
				</script>
				<SCRIPT LANGUAGE="JAVASCRIPT">					
					window.parent.frames("detail").location="EditLotItem.asp?ID=" +<%= ID%>;
				</SCRIPT>
				<% response.end
			End If
			'ensure date is valid format
			If Not IsDate(tDate) Then%>
				<SCRIPT LANGUAGE="JAVASCRIPT">					
					alert ("Please specify a valid date.");
				</SCRIPT>
				<SCRIPT LANGUAGE="JAVASCRIPT">					
					window.parent.frames("detail").location="EditLotItem.asp?ID=" +<%= ID%>;
				</SCRIPT>
				<% 
				response.end
			End If

			'ensure balance qty on order is not negative
			'response.write varBalanceQty

			If cdbl(varBalanceQty) <= 0 Then%>
				<SCRIPT LANGUAGE="JAVASCRIPT">					
					alert ("Order already filled. Please place a new order.");
					window.parent.frames("detail").location="EditLotItem.asp?ID=" +<%= ID%>;
				</SCRIPT>
				<% 
				response.end
			End If

			'ensure balance qty is not exceeded
			If cdbl(varBalanceQty) < cdbl(qty) Then%>
				<SCRIPT LANGUAGE="JAVASCRIPT">					
					alert ("Specified quantity exceeds balance quantity on order.");
					window.parent.frames("detail").location="EditLotItem.asp?ID=" +<%= ID%>;
				</SCRIPT>
				<% 
				response.end
			End If

				sqlStr = "execute cont_CreateContract " & ID & ", " & broker & ", " & price & ", " & qty & ", '" & slip & "', '" & FormatDate(tDate) & "', '" & FormatDate(sDate) & "', " & UserId & ""

				conn.BeginTrans
					conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
				conn.CommitTrans

			conn.Close
			Set conn = Nothing
			%>
			<SCRIPT LANGUAGE="JAVASCRIPT">					
				this.location="EditLotItem.asp?ID=" +<%= ID%>;
			</SCRIPT>
			<%
			'WriteDialogRelocateScript "EditLotItem.asp?ID=" & IDHolder
			Response.End
			
   	case else
   			sqlStr = "SELECT * FROM OrdDetailList WHERE OrdDetailList.OrdDetail_DPA_= " & ID
		   	
   			Set conn = GetActiveConnection("KBroker")
   			set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
   			If rs.EOF Or rs.BOF Then%>
					<script language = 'vbscript'>
                			window.self.ShowMessage "The selected Order item cannot be retrieved for lot allocation"
		                	
					</script>
					<% response.end
			End If
   	end select

%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit <%=DataEntity%> Item</title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
 <script language='vbscript'>
			function ItemSelected(itemID)
					
 					frm<%=DataSource%>Item.elements("ItemID").value = itemID
			end function
			
			
			function SaveInPlaceEdit()
				    Dim myOwnerFrame				
					'UpdateID
					'Set window.parent.dialogArguments.opener.parent.frames("footer").editDocOpener = window.self
					'frm<%=DataSource%>Item.target = "deleteFrame" 					
					frm<%=DataSource%>Item.submit
			end function
		</script>
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
		
		<!-- ActiveUI stylesheet and scripts -->
		<link href="../runtime/classic/activeui.css" rel="stylesheet" type="text/css">
		<script src="../runtime/activeui.js"></script>
		<!-- Include patches here -->
		<script src="../runtime/paging1.js"></script>
		<!-- grid format -->
		<style> 
			.active-controls-grid {height: 100%; font: menu;}
			.active-row-highlight .active-row-cell {background-color: skyblue}
		    
		    
		     	
			.active-column-0 {width: 50px;}
			.active-column-1 {width: 70px;}
			.active-column-2 {width: 120px;}
			.active-column-3 {width: 80px;}
			.active-column-4 {width: 80px;}
			.active-column-5 {width: 200px;}
			.active-column-6 {width: 100px;}
			.active-column-7 {width: 100px;}
			
			
			.active-grid-row,
			.active-grid-row.active-list-item,
			.active-scroll-left .active-list-item {height: 22px;}
			
			
			.active-selection-true, .active-selection-true .active-row-cell {
				color: blue!important;
				background-color: bisque!important;
				}
		</style>
		

</head>

<body Class="Dialog" SCROLL="No">	
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<form name = 'frm<%=DataSource%>Item' id='frmMain' method = 'post' action = '<%=DataSource%>Item.asp' OnSubmit="UpdateID();">

<SCRIPT language="JavaScript">
	var calTDate;
	function changeDateInterface(selCol){
		try{
			calTDate = new ctlSpiffyCalendarBox('calTDate', 'frm<%=DataSource%>Item', 'txtTDate', 'cmdTDate','<%= FormatDate(Date) %>', 1); 
			calTDate.readonly = false;
			calTDate.returnOutStringOnWrite(); 
			//var parentDiv = document.all.item("txtTDate").parentNode;
			//parentDiv.innerHTML = calTDate.writeControl();
			//if (selCol==null || selCol == "undefined"){
				document.all.item("txtTDate").outerHTML = calTDate.writeControl();
		//	}
		//	else{
		//		document.all.item(selCol).outerHTML = calTDate.writeControl();
				
		//	}	
			//parentDiv.style.zIndex = 10;
			//parentDiv.childNodes(1).style.zIndex = 10	;
			
			//Contract Settlement Date
			calSDate = new ctlSpiffyCalendarBox('calSDate', 'frm<%=DataSource%>Item', 'txtSDate', 'cmdSDate','<%= FormatDate(Date) %>', 1); 
			calSDate.readonly = false;
			calSDate.returnOutStringOnWrite(); 

			document.all.item("txtSDate").outerHTML = calSDate.writeControl();
		
		}
		
		catch(e){}	
		
		document.all.item('txtSlip').focus();
	}
	
	document.body.onload = changeDateInterface;
	
	//function ShowGrid()
	//{
	///	ShowMessage(document.all.item("GridCell").innerText);
	//}
</SCRIPT>
 
  
 <%
 Set conn = GetActiveConnection("KBroker")
 
	Dim rowCount
	Dim brokerList
	Dim quote 
	
	quote =  chr(34)
	brokerList = GetBrokerList("cboBroker")
	       
    sqlStr = "SELECT * FROM LotList WHERE LotList.OrdDetail_DPA_= " & ID
    sqlStr1 = "SELECT * FROM OrdDetailList WHERE OrdDetailList.OrdDetail_DPA_= " & ID
       
    Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr1)))
    
    'Response.Write(sqlStr1)
    'Response.End
    
    If (rs.EOF Or rs.BOF) Then
    %>
					<script language = 'vbscript'>
                			window.self.close		                	
					</script>
					<% 
				  'WriteDialogRelocateScript "AddLot.asp?ID=" & DetailID					
				  Response.redirect "AddLot.asp?ID=" & ID					
				  response.end
			
    end if
    
    If Not(rs.EOF Or rs.BOF) Then
    clientdpa=rs.Fields("Client_DPA_")
   
    
		'store the Order Type%>
        <input type = 'hidden' name ='txtOrderType' id = 'txtOrderType' size='9' value = '<%=rs.Fields("OrdDetailType")%>'>
        <input type = 'hidden' name ='txtInstrument' id = "txtInstrument" size='9' value = '<%=rs.Fields("OrdDetailSecType")%>'>
        <input type = 'hidden' name ='txtOrderIsSaleType' id = "txtOrderIsSaleType" size="20" value = '<%=rs.Fields("OrderTypeSale")%>'>
        <input type = 'hidden' name ='txtSecurityID' id = "txtSecurityID" size="20" value = '<%=rs.Fields("Security_DPA_")%>'>
        <input type = 'hidden' name ='txtAgentCommission' id = "txtAgentCommission" size="20" value = '<%=rs.Fields("AgentCommission")%>'>
        <input type = 'hidden' name ='txtStaffCommission' id = "txtStaffCommission" size="20" value = '<%=rs.Fields("StaffCommission")%>'>
        <input type = 'hidden' name ='txtCommission' id = "txtCommission" size="20" value = '<%=rs.Fields("CommissionRate")%>'>
        <input type = 'hidden' name ='txtVolumeCommission' id = "txtVolumeCommission" size="20" value = '<%=rs.Fields("VolumeRate")%>'></td>
	  <input type = 'hidden' name ='txtVolumeBoundary' id = "txtVolumeBoundary" size="20" value = '<%=rs.Fields("VolumeBoundary")%>'></td>
	  <input type = 'hidden' name ='txtMinimumCommission' id = "txtMinimumCommission" size="20" value = '<%=rs.Fields("MinimumCommission")%>'></td>
	  <input type = 'hidden' name ='txtCMA' id = "txtCMA" size="20" value = '<%=rs.Fields("CMARegulated")%>'></td>
		<input type = 'hidden' name ='txtPostImmobilisedRate' id = "txtPostImmobilisedRate" size="20" value = '<%=rs.Fields("PostImmobilisedRate")%>'></td>
		<input type = 'hidden' name ='txtSecurityImmobilised' id = "txtSecurityImmobilised" size="20" value = '<%=rs.Fields("SecurityImmobilised")%>'></td>
		<input type = 'hidden' name ='txtClientDPA' id = "txtClientDPA" size="20" value = '<%=clientdpa%>'></td>
		<input type = 'hidden' name ='txtEntityDPA' id = "txtEntityDPA" size="20" value = '<%=rs.Fields("EntityType_DPA_")%>'></td>
            <input type = 'hidden' name ='txtClass' id = "txtClass" size="20" value = '<%=rs.Fields("Class")%>'>
            
            <input type = 'hidden' name ='txtSettlementDate' id = "txtSettlementDate" size="20" value = ''>
            </td>

		<% 
		if rs.Fields("InterBank")=true then
		%>
		<input type = 'hidden' name ='txtinterbank' id = "txtinterbank" size="20" value = '1'></td>
		<%
		else
		%>
		<input type = 'hidden' name ='txtinterbank' id = "txtinterbank" size="20" value = '0'></td>
		<%
		end if
		%>
	 
		<% 
		if rs.Fields("IsCustodian")=true then
		%>
		<input type = 'hidden' name ='txtcustodian' id = "txtcustodian" size="20" value = '1'></td>
		<%
		else
		%>
		<input type = 'hidden' name ='txtcustodian' id = "txtcustodian" size="20" value = '0'></td>
		<%
		end if
		%>

		<!-- grid data -->
		<% 'row data
	end if
	
	Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
	If Not(rs.EOF Or rs.BOF) Then
		
		rowCount = 0 
		rs.MoveFirst
		Do Until rs.EOF 
		
'======================= Begin_Alter_Across_Entities =================================
			'row ID 
			
			rowData = rowData & quote & rs.Fields("Lot_DPA_") & quote & " : " 
			
			'row data 
			rowData = rowData & "[" 
			rowData = rowData & quote & rs.Fields("Lot_DPA_") & quote & "," 
			rowData = rowData & quote & rs.Fields("LotSlipNo") & quote & ","
			rowData = rowData & quote & FormatDate(rs.Fields("LotTDate")) & quote & ","
			rowData = rowData & quote & rs.Fields("LotQty") & quote & ","
			rowData = rowData & quote & rs.Fields("LotPrice") & quote & ","
			rowData = rowData & quote & rs.Fields("BrokerCode") & " : " & rs.Fields("BrokerName") & quote & ","
			rowData = rowData & quote & FormatDate(rs.Fields("ContractSettlementDate")) & quote & ","
			rowData = rowData & quote & " " & quote  
			rowData = rowData & "]" 
			
			rowIDs = rowIDs & quote & rs.Fields("Lot_DPA_") & quote 
			rowCount = rowCount + 1
		
		
			rs.MoveNext 
			
			
				'build the row IDs array
				rowIDs = rowIDs & "," 
				rowData = rowData & ","	
			
'======================= End_Alter_Across_Entities =================================

			
		Loop

		
		rs.MoveFirst
		%><script language="javascript">
			//update the quantity balance
			try{
				window.parent.frames("header").document.frmMain.elements("txtBalance").value = '<%= FormatNum(rs.Fields("BalanceQty")) %>';
			}
			catch(e){}	
		 </script><%
	End if
	
		'row ID 	
		rowData = rowData & quote & -1 & quote & " : " 
				
		'row data 
		rowData = rowData & "[" 
		rowData = rowData & quote & "New Line" & quote & "," 
		'rowData = rowData & quote & "<input type = 'text' name ='txtSlip' id = 'txtSlip' size='9' onChange = 'AddRowInProgress();event.cancelBubble=true;' OnClick='event.cancelBubble=true;'>" & quote & ","
		'rowData = rowData & quote & "<input type='text' name='txtTDate' size=40 value='" & FormatDate(Date) & "' onChange = 'AddRowInProgress();event.cancelBubble=true;' OnClick='event.cancelBubble=true;'>" & quote & ","
		'rowData = rowData & quote & "<input type = 'text' name ='txtQty' id = 'txtQty' size='9' onChange = 'AddRowInProgress();event.cancelBubble=true;' OnClick='event.cancelBubble=true;'>" & quote & ","
		'rowData = rowData & quote & "<input type = 'text' name ='txtPrice' id = 'txtPrice' size='9' onChange = 'AddRowInProgress();event.cancelBubble=true;' OnClick='event.cancelBubble=true;'>" & quote & ","
		rowData = rowData & quote & "<input type = 'text' name ='txtSlip' id = 'txtSlip' size='7' OnClick='event.cancelBubble=true;'>" & quote & ","
		rowData = rowData & quote & "<input type='text' name='txtTDate' size=40 value='" & FormatDate(Date) & "' OnClick='event.cancelBubble=true;'>" & quote & ","
		rowData = rowData & quote & "<input type = 'text' name ='txtQty' id = 'txtQty' size='9' OnClick='event.cancelBubble=true;'>" & quote & ","
		rowData = rowData & quote & "<input type = 'text' name ='txtPrice' id = 'txtPrice' size='9' OnClick='event.cancelBubble=true;'>" & quote & ","
		rowData = rowData & quote & brokerList & quote  & ","
		rowData = rowData & quote & "<input type='text' name='txtSDate' size=40 value='" & FormatDate(conn.execute("select   getdate() - 0 + 7 AS SettlementDate")(0)) & "' OnClick='event.cancelBubble=true;'>" & quote  & ","
		rowData = rowData & quote & "<input type=button value='Add' Class=Buttons OnClick='JavaScript: AddRowInProgress();'>&nbsp;&nbsp;<input type='reset' value='Cancel' Class=Buttons>" & quote 
		rowData = rowData & "]" 
	

		'build the row IDs array 
		rowIDs = rowIDs & quote & -1 & quote 
		rowCount = rowCount + 1
		
'======================= Begin_Alter_Across_Entities =================================%> 
		<script language="javascript">
			//column titles 
			var colCount = 8;
			var colNames = ["", "Ref",  
					 "Date", "Quantity","Price","Broker","SettlementDate", ""];
			
			var myColumns = ["Lot No", "Ref",  
					 "Date", "Quantity","Price","Broker","SettlementDate" ,""];
		</script>
<%'======================= End_Alter_Across_Entities =================================%>			
		<script language="javascript">
		
			
			//data
			var myData = {<%=rowData%>}; 
			var myRowIDs = [<%=rowIDs%>]; 
			
			
			//editing
			var inPlaceEdit = false;
			var addInProgress = false;
			var clickedRowID = -1; 
			var dataChanged = false;
			var prevRow = -1;//the row currently under in-place edit
			
			function EditInPlaceDataChanged()
			{
				dataChanged = true;
			}
			
			function AddRowInProgress()
			{
				addInProgress = true;
			}
			
			var currentBrokerName = "";
			var RowEditFn = function(src)
			{
				var rowIndex = src.getProperty("row/index");
				var i;
				if (rowIndex==0) return;
				for(i = 0; i < colCount; i++)
				{
					if(colNames[i] != "")
					{
						if(prevRow >= 0)
						{
							if(colNames[i]=="Broker")
							{
								myData[prevRow][i] = currentBrokerName;
							}
							else
							{
								myData[prevRow][i] = document.frmMain.elements(colNames[i]).value;
							}
						}
						if(colNames[i]=="Broker")
						{	
							currentBrokerName =  myData[rowIndex][i];
							myData[rowIndex][i] = inPlaceList;
							
						}
						else
						{
							if(colNames[i]=="Date" || colNames[i]=="SettlementDate")
							{
								currentDate =  myData[rowIndex][i];								
								myData[rowIndex][i] = "<INPUT TYPE='text' NAME='" + colNames[i] + "' ID='" + colNames[i] + "' VALUE='" + currentDate + "' onChange = 'EditInPlaceDataChanged();event.cancelBubble=true;' OnClick='event.cancelBubble=true;'>";
								//show the calendar
								changeDateInterface(colNames[i]);
				
								
							}
							else {
								myData[rowIndex][i] = "<INPUT TYPE='text' NAME='" + colNames[i] + "' ID='" + colNames[i] + "' VALUE='" + myData[rowIndex][i] + "' onChange = 'EditInPlaceDataChanged();event.cancelBubble=true;' OnClick='event.cancelBubble=true;'>";
							}	
						}
					}
				}
				
				myData[rowIndex][colCount - 1] = "<INPUT TYPE='button' class='Buttons' VALUE='Save' onClick = 'SaveInPlaceEdit();event.cancelBubble=true;'>&nbsp;<INPUT TYPE='button' class='Buttons' VALUE='Cancel' onClick = 'cancelEditRow();event.cancelBubble=true;'>";
				inPlaceEdit = true;
				prevRow = rowIndex;
				grid.refresh();
				//select the appropriate item
				var secList = document.frmMain.elements("cboBrokerInPlace");
				for (i=0; i < secList.options.length; i++) {
					if(secList.options(i).text == currentBrokerName)
					{
							secList.options(i).selected = true;
					}
				}
				
				
			}
			
			function cancelEditRow(){
					var i;
							for(i = 0; i < colCount; i++)
							{
								if(colNames[i] != "")
								{
									if(colNames[i]=="Broker")
									{
										myData[prevRow][i] = currentBrokerName;
									}
									else
									{
										myData[prevRow][i] = document.frmMain.elements(colNames[i]).value;
									}
								}
							}
							myData[prevRow][colCount - 1] = "";						
							inPlaceEdit = false;
							prevRow = -1;
							grid.refresh();
			}
			
			var RowChangeFn = function(src)
			{
				if(inPlaceEdit || addInProgress)
				{
					if(dataChanged || addInProgress)
					{
						ItemSelected(prevRow);
						SaveInPlaceEdit();
					}
					else
					{
						if(prevRow != clickedRowID)
						{
							cancelEditRow();
						}
					}
				}
			}
			
			var HandleClick = function(src)
			{
				clickedRowID = src.getProperty("row/index");
				ItemSelected(clickedRowID);
			}
			
			var headerID;
			
			function UpdateID(){
				headerID = window.parent.frames["header"].document.all.item("ID").value;
				//headerID = window.document.all.item("ID").value;
				document.all.item("ID").value = headerID;
				
				theSDate = window.parent.frames["header"].document.all.item("txtSDate").value;
				//theSDate = window.document.all.item("txtSDate").value;
				document.all.item("txtSettlementDate").value = theSDate;
			}
			
			function restoreID(){
				document.all.item("ID").value = headerID;
			}
			
			function HandleDeleteAction()
			{
					document.frmMain.elements("action").value = "Execute_Delete"
					SaveInPlaceEdit();
					document.frmMain.elements("action").value = "Execute_Detail"
			}
			
			function HandleSaveAction()
			{
					document.frmMain.elements("action").value = "Execute_Save"
					SaveInPlaceEdit();
					document.frmMain.elements("action").value = "Execute_Detail"
			}
			//get ready for in-place edit
			var inPlaceList = "<%=GetBrokerList("cboBrokerInPlace")%>"
		</script> 
		
		<script language="javascript"> 

			// create ActiveUI Grid javascript object 
			var grid = new Active.Controls.Grid; 
			
			
			// set rows ids 
			grid.setRowValues(myRowIDs); 
			

			// set number of columns 
			grid.setColumnCount(colCount); 

			// provide cells and headers text 
			grid.setDataText(function(i, j){return myData[i][j]}); 
			grid.setColumnText(function(i){return myColumns[i]}); 

			// set click action handler 
			grid.setAction("click", HandleClick); 
			//grid.setAction("dblclick", RowEditFn); 
			grid.setAction("selectionChanged", RowChangeFn);
			
			//stripes 
			var alternate = function(){ return this.getProperty("row/order") % 2 ? "gainsboro" : "white";} 
			var row = new Active.Templates.Row; row.setStyle("background", alternate); 
			row.setEvent("onmouseover", "mouseover(this, 'active-row-highlight')"); 
			row.setEvent("onmouseout", "mouseout(this, 'active-row-highlight')"); 
			grid.setTemplate("row", row); 
			var column = new Active.Templates.Text; 
			column.setStyle("border-right", "1px solid white");  
			grid.setTemplate("column", column);  grid.setRowHeaderWidth("0px"); 
			
			//disable sort
			grid.getTemplate("top/item").setEvent("onmousedown", null);
			
			// write grid html to the page 
			document.write(grid); 
			
			//let grid be aware of composite layout
			grid.getLayoutTemplate().action("adjustSize");
		</script> 
        <%
       
  function GetBrokerList(listName)
		Dim secList
		
		secList = "<select name = '" & listName & "' id = '" & listName & "' size='1' "
		secList = secList & "OnClick='event.cancelBubble=true;' "  
		secList = secList & "onChange='event.cancelBubble=true;' " 
		secList = secList & "onKeypress='return (dodefaultaction()==\""\""); ' "  
		secList = secList & "onKeydown='return (dodefaultaction()==\""\"");event.cancelBubble=true;' "  
		secList = secList & "onKeyup='return (change(" & listName & "));' "  
		secList = secList & "onfocus='txtval = \""\"";inputIsItemCode = 1;' "  
		secList = secList & "onblur='txtval = \""\"";inputIsItemCode = 1;'>"
		
		secList = secList & "<option selected SearchCode = '0' SearchText = ''  value = ''></option>"
		
        sqlStr = "SELECT * FROM [BrokerList] Order By BrokerCode"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF
                        secList = secList & "<option SearchCode = '" & rs.Fields("BrokerCode") & "' SearchText = '" & rs.Fields("BrokerName") & "'  value = '" & rs.Fields("Broker_DPA_") & "'>" & rs.Fields("BrokerNameEx") & "</option>"
                        rs.MoveNext
                Loop
        End If
	    secList = secList & "</select>"
	    GetBrokerList = secList
  end function
  
  function DeleteItem(EntityName,KeyField,DelItemID)
		dim delRS
			'find out whether any child records exist
			sqlStr = "SELECT Child,DeletionMessage,ParentKey FROM [_Parent_Child_Links_] WHERE (Parent = '" & EntityName & "') AND (ChildType = " & LinkedIndependent & ")"
			Set delRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
			If Not (delRS.BOF Or delRS.EOF) Then
					Dim childRS
					Dim tableName
					
					delRS.MoveFirst
					Do Until delRS.EOF
                			tableName = delRS.Fields("Child")
							sqlStr = "SELECT TOP 1 * FROM [" & tableName & "] WHERE " & KeyField & " = " & DelItemID & " and Deleted <>1"							
					
							Set childRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
							If Not (childRS.BOF Or childRS.EOF) Then%>
                					<script language = 'vbscript'>
                						ShowMessage "<%=delRS.Fields("DeletionMessage")%>"
	                					
                					</script>
                					<%response.end
							End If
							delRS.MoveNext
					Loop
			End If
			
			'delete from database
			if(ucase(Trim(EntityName))="LOT" OR ucase(Trim(EntityName))="PAYMENT") then
			sqlStr = "Update  [" & EntityName & "] Set Deleted = 1,ChangedBy=" & UserId & ",TimeChanged=GetDate() WHERE " & KeyField & " = " & DelItemID
			else
			sqlStr = "Update  [" & EntityName & "] Set Deleted = 1 WHERE " & KeyField & " = " & DelItemID
			end if
			
			conn.Execute SQLServerFormat(HandleQuote(sqlStr))
  end function
  

 %>
 <table border="0" width="100%" ID="Table1">
<tr><td>
<input type = 'hidden' name ='ItemID' id = 'ItemID'>
<input type = 'hidden' name ='ID' id = 'ID' value="<%= IDHolder %>">
<input type = 'hidden' name ='action' id = 'action' value="Execute_Detail">
<input type = 'hidden' name ='BalanceQty' id = 'BalanceQty'>
</td>
</tr>
</table>

	<script language="javascript">
		try{
			window.parent.frames("detail").document.frmMain.elements("BalanceQty").value = window.parent.frames("header").document.frmMain.elements("txtBalance").value;
			//alert(window.parent.frames("header").document.frmMain.elements("txtBalance").value);
			//alert(window.parent.frames("detail").document.frmMain.elements("BalanceQty").value);
		}
		catch(e){}	
	</script>

</form>
</body>

</html>