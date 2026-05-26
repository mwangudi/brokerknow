<!--#include file="../libroutines.asp"-->
<%
'======================= Begin_Alter_Across_Entities =================================
		
		const UDLName = "KBroker"
		const DataSource = "WebtbOrderList"
		const ActionPage = "OnlineOrderList"
		const DataEntity = "OnlineOrder"
		const DataEntityPlural = "OnlineOrders"
		const ActionFolder = "Operations"
'======================= End_Alter_Across_Entities =================================		
		
		Dim conn 
		Dim sqlStr
		Dim rs
		Dim bcolor
		Dim rowIDs
		Dim rowData
		Dim quote
		Dim sortQryStr
		
		sortQryStr = Request.Form("SelectedSortArgs")
		filterStr = Request.Form("SelectedFilterArgs")
		sqlStrOrig = "SELECT * FROM [" & DataSource & "]"
		
        Set conn = GetActiveConnection(UDLName)
        Set Rs = Server.CreateObject("ADODB.Recordset")
		Rs.CursorLocation = adUseClient 
		
		'=====================================================================
		'Save Approval Stuff
		action = Request.QueryString("action")
		
		if trim(lcase(action)) = "save_approval" then
			OrderID = Request.QueryString("orderID")
			Action = Trim(Request.QueryString("appaction"))
			Reason = Replace(Request.QueryString("reason"),"'","''")
			
			set Rs = conn.execute("SELECT OtherNames + ' ' + Surname AS UserName FROM Users WHERE (UserID = " & Session("UserID") & ")")
			
			if not (rs.BOF or rs.EOF) then
				UserName = Rs("UserName")
			else
				UserName = ""
			end if
			
			sqlstr = "UPDATE WebtbOrder SET ApprovalAction = N'" & Action & "', Reason = N'" & Reason & "', UserName = '" & UserName & "', ActionDate = '" & FormatDate(Date()) & "', ActionTime = '" & Time() & "' WHERE (Order_DPA_ = " & OrderID & ")"
			'stop
			conn.execute(sqlstr)
			'rs.Open sqlStr, Conn.ConnectionString, adOpenKeySet, adLockOptimistic
			
			sqlstr = "SELECT Client.ClientEmail, (SELECT TOP 1 SendDisplayName FROM MailConfiguration) AS FromAddress" & _
				"	FROM         WebtbOrder INNER JOIN" & _
				"	                      Client ON WebtbOrder.Client_DPA_ = Client.Client_DPA_" & _
				"	WHERE     (WebtbOrder.Order_DPA_ = " & OrderID & ")"
			
			set Rs = conn.execute(sqlstr)	'rs.Open SQLServerFormat(HandleQuote(sqlStr)), Conn.ConnectionString, adOpenKeySet, adLockOptimistic
			
			if not (rs.BOF or rs.EOF) then
				ClientEmail = Rs("ClientEmail")
				FromAddress = rs("FromAddress")
			else
				'Addresses needed...
				ClientEmail = "client@email.com"
				FromAddress = "admin@knowing.com"
			end if
			
			select case UCase(Trim(Action))
				case "ACCEPT"
					AcceptOrder(OrderID)
					Subject = "Online Order Approved"
					Body = "Your Online Order has been approved" & vbCrLf
					Body = Body & "Date : " & Date() & vbCrLf
					Body = Body & "Online Order No : " & OrderID & vbCrLf
					Body = Body & "Approved Order No : " & vbCrLf
				case "REVIEW"
					Subject = "Online Order Reviewed"
					Body = "Your Online Order is under review." & vbCrLf
					Body = Body & "Date : " & Date() & vbCrLf
					Body = Body & "Online Order No : " & OrderID & vbCrLf
					Body = Body & "Reason : " & Reason & vbCrLf
					Body = Body & "NB: Please do not reply to this email address." & vbCrLf
				case else
					Action = UCase(Trim(Request.QueryString("appaction")))
					if Action = "REJECT-INSUFFICIENT FUNDS" then
						Reason = "Insufficient funds"
					elseif Action = "REJECT-PRICE HAS CHANGED" then
						Reason = "Price has changed"
					else
						Reason = "User request"
					end if
					
					Subject = "Online Order Rejected"
					Body = "Your Online Order has been rejected." & vbCrLf
					Body = Body & "Date : " & Date() & vbCrLf
					Body = Body & "Online Order No : " & OrderID & vbCrLf
					Body = Body & "Reason : " & Reason & vbCrLf
					Body = Body & "NB: Please do not reply to this email address." & vbCrLf
			end select
			
			SendMail ClientEmail, Subject, Body
			
			Response.Redirect "OnlineOrderList.asp"
		end if
		'=====================================================================
		
		sqlStr = "SELECT * FROM [" & DataSource & "]"
		
		If filterStr <> "" Then
			sqlStr = "SELECT * FROM [" & DataSource & "] WHERE " & filterStr
		End If
		
		If sortQryStr <> "" Then		
			sqlStr = "SELECT * FROM [" & DataSource & "] ORDER BY " & sortQryStr	
		End If 
		
		Rs.Open  SQLServerFormat(HandleQuote(sqlStr)), Conn.ConnectionString, adOpenKeySet, adLockOptimistic
        
        If (rs.EOF Or rs.BOF) Then
        
		 Else
		 	intPage = IntToNull(Request("Page"))
		 	If intPage < 0 Then intPage = 1
			If IsNull(intPage) Then intPage = 1
			Rs.PageSize = intPageSize
			If intPage > Rs.PageCount Then intPage = Rs.PageCount
			Rs.AbsolutePage = intPage
        End If
%>
<html>
	<head>
		<title><%=DataEntityPlural%></title>
		<meta http-equiv="Content-Language" content="en-us">
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
		<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
		<meta name="ProgId" content="FrontPage.Editor.Document">
		<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css">
			<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
				<SCRIPT language="Javascript" src="../scripts/common.js"></SCRIPT>
				<script language='vbscript'>
					function ItemSelected(itemID)
 							frm<%=ActionPage%>.elements("ID").value = itemID
					end function
					
					function SaveInPlaceEdit()
							frm<%=ActionPage%>.submit
					end function
				</script>
				<!-- ActiveUI stylesheet and scripts -->
				<link href="../runtime/styles/xp/grid.css" rel="stylesheet" type="text/css">
				<script src="../runtime/activeui.js"></script>
				<!-- Include patches here -->
				<script src="../runtime/paging1.js"></script>
				<!-- grid format -->
				<style> 
					.active-controls-grid {height: 100%; font: menu;}
					.active-row-highlight .active-row-cell {background-color: skyblue}
					.active-selection-true, .active-selection-true .active-row-cell {
						color: blue!important;
						background-color: bisque!important;
						}
					
					.active-column-0 {width: 90px;}
					.active-column-1 {width: 70px;}
					.active-column-2 {width: 70px; text-align: right;}
					.active-column-3 {width: 30px; text-align: left;}
					.active-column-4 {width: 100px;}
					.active-column-5 {width: 70px;}
					.active-column-6 {width: 70px; text-align: right;}
					.active-column-7 {width: 70px; text-align: right;}
					.active-column-8 {width: 90px; }					
				</style>
										
</head>
	<body leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
	<form name='frm<%=ActionPage%>' method='post' id='frmMain' action='AcceptOrder.asp'>
			<!-- grid data -->
		<% 'row data
		Dim rowCount
		Dim orderHeld
		Dim AcceptSelected
		Dim ReviewSelected
		Dim RejectFundsSelected
		Dim RejectPriceSelected
		Dim RejectUserSelected
		quote = chr(34)
		intRecord = 1
		Do Until rs.EOF 

'======================= Begin_Alter_Across_Entities =================================
			
			select case Trim(rs.Fields("ApprovalAction"))
				case "ACCEPT"
					AcceptSelected = " selected"
				case "REVIEW"
					ReviewSelected = " selected"
				case "REJECT-INSUFFICIENT FUNDS"
					RejectFundsSelected = " selected"
				case "REJECT-PRICE HAS CHANGED"
					RejectPriceSelected = " selected"
				case "REJECT-USER REQUEST"
					RejectUserSelected = " selected"
			end select
			'orderAccept = "<input type=checkbox class='BorderLess' style=BorderLess  name='chkAccept' onClick = 'AcceptOrder(this, " & rs.Fields("Order_DPA_") & ");'>"
			orderAccept = "<select name='cboAction" & Trim(rs.Fields("Order_DPA_")) & "' size='1' style='Borderless' class='Borderless'"
			orderAccept = orderAccept & " onClick='event.cancelBubble=true;' "  
			orderAccept = orderAccept & " onChange='event.cancelBubble=true;' " 
			orderAccept = orderAccept & " onKeypress='return (dodefaultaction()==\""\""); ' "  
			orderAccept = orderAccept & " onKeydown='return (dodefaultaction()==\""\"");event.cancelBubble=true;' "  
			orderAccept = orderAccept & " onKeyup='return (change(" & listName & "));' "  
			orderAccept = orderAccept & " onfocus='txtval = \""\"";inputIsItemCode = 1;' "  
			orderAccept = orderAccept & " onblur='txtval = \""\"";inputIsItemCode = 1;'>"
			orderAccept = orderAccept & "<option value='ACCEPT' " & AcceptSelected & ">Accept</option>"
			orderAccept = orderAccept & "<option value='REVIEW' " & ReviewSelected & ">Review</option>"
			orderAccept = orderAccept & "<option value='REJECT-INSUFFICIENT FUNDS' " & RejectFundsSelected & ">Reject: Insufficient Funds</option>"
			orderAccept = orderAccept & "<option value='REJECT-PRICE HAS CHANGED' " & RejectPriceSelected & ">Reject: Price Has Changed</option>"
			orderAccept = orderAccept & "<option value='REJECT-USER REQUEST' " & RejectUserSelected & ">Reject: User Request</option>"
			orderAccept = orderAccept & "</select>"
			
			ReasonField = "<input type='text' name='Reason" & rs.Fields("Order_DPA_") & "' value='" & rs.Fields("Reason") & "' onClick='event.cancelBubble=true;' "  
			ReasonField = ReasonField & " onChange='event.cancelBubble=true;' " 
			ReasonField = ReasonField & " onKeypress='return (dodefaultaction()==\""\"");event.cancelBubble=true; ' "  
			ReasonField = ReasonField & " onKeydown='return (dodefaultaction()==\""\"");event.cancelBubble=true;' "  
			ReasonField = ReasonField & " onKeyup='return (dodefaultaction()==\""\"");event.cancelBubble=true;' "  
			ReasonField = ReasonField & " onfocus='txtval = \""\"";inputIsItemCode = 1;' "  
			ReasonField = ReasonField & " onblur='txtval = \""\"";inputIsItemCode = 1;'>"
			
			AcceptSelected = ""
			ReviewSelected = ""
			RejectFundsSelected = ""
			RejectPriceSelected = ""
			RejectUserSelected = ""
			
			'row ID 
			rowData = rowData & quote & rs.Fields("Order_DPA_") & quote & " : " 
			
			'row data 
			rowData = rowData & "[" 
			rowData = rowData & quote & FormatDate(rs.Fields("OrderDate")) & quote & ","
			rowData = rowData & quote & rs.Fields("OrderTypeName") & quote & ","
			rowData = rowData & quote & rs.Fields("Order_DPA_") & quote & ","  
			rowData = rowData & quote & rs.Fields("Client_DPA_") & quote & ","  
			rowData = rowData & quote & rs.Fields("ClientName") & quote & ","
			rowData = rowData & quote & rs.Fields("SecurityCode") & quote & ","
			rowData = rowData & quote & FormatNumCommasOnly(rs.Fields("OrdDetailQty")) & quote & ","
			rowData = rowData & quote & FormatNum(rs.Fields("OrdDetailPrice")) & quote & ","
			rowData = rowData & quote & FormatDate(rs.Fields("OrdDetailValidity")) & quote & ","
			rowData = rowData & "]"
			
			'build the row IDs array 
			rowIDs = rowIDs & quote & rs.Fields("Order_DPA_") & quote 
			rowCount = rowCount + 1
			
'======================= End_Alter_Across_Entities =================================

			rs.MoveNext 
			if not(rs.eof) then 
				If  intRecord >= intPageSize Then	Exit Do				
				intRecord = intRecord + 1	
				rowIDs = rowIDs & "," 
				rowData = rowData & "," 
			end if 
		Loop
		
'======================= Begin_Alter_Across_Entities =================================%> 
		<script>
			//column titles 
			var colCount = 9;
			var colNames = ["",  
			"", "", "","", "", "","","","",""];
			
			var myColumns = ["Date",  "Type","Order No","Code",   
			"Client", "Security", "Qty", "Price","Validity"];
		</script>
<%'======================= End_Alter_Across_Entities =================================%>			
		<script>
			//data
			var myData = {<%=rowData%>}; 
			var myRowIDs = [<%=rowIDs%>]; 
			
			//editing
			var inPlaceEdit = false;
			var clickedRowID = 0; 
			var dataChanged = false;
			var prevRow = -1;//the row currently under in-place edit
			
			//Wakaria
			function gotoApprove(theOrderID)
			{
				var action = document.all.item("cboAction" + theOrderID).value;
				var reason = document.all.item("Reason" + theOrderID).value;
				window.location.href = "OnlineOrderList.asp?action=save_approval&orderID=" + theOrderID + "&appaction=" + action + "&reason=" + reason;
			}
			
			function EditInPlaceDataChanged()
			{
				dataChanged = true;
			}
			
			var RowEditFn = function(src)
			{
				var rowIndex = src.getProperty("row/index");
				var i;
				for(i = 0; i < colCount; i++)
				{
					if(colNames[i] != "")
					{
						if(prevRow >= 0)
						{
							myData[prevRow][i] = document.frmMain.elements(colNames[i]).value;
						}
						myData[rowIndex][i] = "<INPUT TYPE='text' NAME='" + colNames[i] + "' ID='" + colNames[i] + "' VALUE='" + myData[rowIndex][i] + "' onChange = 'EditInPlaceDataChanged();'>";
						
					}
				}
				inPlaceEdit = true;
				prevRow = rowIndex;
				grid.refresh();
			}
			
			var first = 1;
			var RowChangeFn = function(src)
			{
				if (first!=1)
				{
					var RowID = document.all.item("ID").value;
					document.all.item("cboAction" + RowID).focus();
				}
				first = 0;
			}
			
			var HandleClick = function(src)
			{
				clickedRowID = src.getProperty("row/index");
				ItemSelected(clickedRowID);
			}
			
			function  AcceptOrder(theCbo, theItem)
			{
				var theVal = theCbo.options[theCbo.options.selectedIndex].value;
				var reason = document.all.item('Reason' + theItem).value;
				
				if (theVal=="REJECT" || theVal=="REVIEW")
				{
					alert('Please enter a reason for the action');
					return false;
				}
				
				return false
				document.frmMain.elements("Accept").value = "1";
				document.frmMain.elements("delAction").value = "Execute";
				ItemSelected(theItem);
				SaveInPlaceEdit();
			}
			
			window.onload = function()
			{
				//select first item
				grid.setSelectionValues([myRowIDs[0]]);
				ItemSelected(myRowIDs[0]);
			}
		</script> 
		<script> 

			// create ActiveUI Grid javascript object 
			var grid = new Active.Controls.Grid; 
			//grid.setModel("row", new Active.Rows.Page);
			
			// set rows ids 
			//grid.setProperty("row/count", <%=rowCount%>);
			//grid.setProperty("row/values", myRowIDs);
			grid.setRowValues(myRowIDs); 
			

			// set number of columns 
			//grid.setProperty("column/count", colCount);
			grid.setColumnCount(colCount); 

			// provide cells and headers text 
			//grid.setProperty("data/text", function(i, j){return myData[i][j]});
			grid.setDataText(function(i, j){return myData[i][j]}); 
			//grid.setProperty("column/texts", myColumns);
			grid.setColumnText(function(i){return myColumns[i]}); 

			//introduce paging
			//grid.setProperty("row/pageSize", 5);
			
			// set click action handler 
			grid.setAction("click", HandleClick); 
			//grid.setAction("dblclick", RowEditFn); 
			//grid.setAction("selectionChanged", RowChangeFn); 

			//stripes 
			var alternate = function(){ return this.getProperty("row/order") % 2 ? "gainsboro" : "white";} 
			var row = new Active.Templates.Row; row.setStyle("background", alternate); 
			row.setEvent("onmouseover", "mouseover(this, 'active-row-highlight')"); 
			row.setEvent("onmouseout", "mouseout(this, 'active-row-highlight')"); 
			grid.setTemplate("row", row); 
			var column = new Active.Templates.Text; 
			column.setStyle("border-right", "1px solid white");  
			grid.setTemplate("column", column);  grid.setRowHeaderWidth("0px"); 

			// write grid html to the page 
			document.write(grid); 
			
			//let grid be aware of composite layout
			grid.getLayoutTemplate().action("adjustSize");
		</script> 
			
				<input type = 'hidden' name ='ID' id = 'ID'>
				<input type = 'hidden' name ='Accept' id = 'Accept'>
				
				<!--The message in the input below is meant for the Footer page. It is replaced with a different string if deletion proceeds-->
				<input type = 'hidden' name ='delAction' id = 'delAction' value="">
				<!-- ----------------------------------------------------------------------------------------------------------------------- -->
				
				<input type = 'hidden' name ='EditPage' id = "EditPage" value =  "<%=ActionFolder%>/Edit<%=DataEntity%>.asp">
				<input type = 'hidden' name ='AddPage' id = "AddPage" value = ""> 
				<input type = 'hidden' name ='DeletePage' id = "DeletePage" value = "<%=ActionFolder%>/Delete<%=DataEntity%>.asp">
				<input type = 'hidden' name ='ActionPage' id = "ActionPage" value = "<%=ActionPage%>.asp">
				<input type = 'hidden' name ='SQLStatement' id = "SQLStatement" value = "<%= sqlStrOrig %>">

<%'======================= Begin_Alter_Across_Entities =================================%>
				
				<input type = 'hidden' name ='FilterArgs' id="FilterArgs" value="OrderDate:Date*3;OrderTypeName:Type*0;Order_DPA_:Order No*0;OrderRef:Reference*0;ClientName:Client*0">
				<input type = 'hidden' name ='SortArgs' id="SortArgs" value="OrderDate:Date;OrderTypeName:Type;Order_DPA_:Order No;OrderRef:Reference;ClientName:Client">
				<input type = 'hidden' name ='SearchArgs' id="SearchArgs" value="OrderDate:Date*3;OrderTypeName:Type*0;Order_DPA_:Order No*0;OrderRef:Reference*0;ClientName:Client*0">
				<input type = 'hidden' name ='dialogLayout' id="dialogLayout" value="height:25em;width:45em">
				
<%'======================= End_Alter_Across_Entities =================================%>
				
				<input type = 'hidden' name ='SelectedSortArgs' id="SelectedSortArgs" value="<%= sortQryStr %>">
				<input type = 'hidden' name ='SelectedFilterArgs' id="SelectedFilterArgs" value="<%= filterStr %>">
				<input type = 'hidden' name ='SelectedSearchArgs' id="SelectedSearchArgs" value="<%= searchStr %>">
				
				<table STYLE="position:absolute; margin-left: 0;" ID="BottomDiv"  Class="footerHighlightNav">
						<%If Rs.PageCount > 1 Then%>
							<tr>
								<td align="center"><%=Paging (intPage, Rs.PageCount, Rs.RecordCount)%></td>
							</tr>
						<%End If%>
				</table>
			<%			
			Set Rs = Nothing
			Set Conn = Nothing
			%>	
				
			</form> 
			
			
</body> 
</html>

<%
sub AcceptOrder(ID)
sqlStr = "SELECT * FROM WebtbOrder WHERE (Order_DPA_ = " & ID & ")"
Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
If (rs.EOF Or rs.BOF) Then%>
 			<script language = 'vbscript'>
 					ShowMessage "The specified order cannot be found."
	         		window.history.back			
 			</script>
 			<% response.end
End If
			
set guid = server.createobject("NDUtils.CGUID")
 guidStr = guid.GenerateGUID
				 
 sqlStr = "INSERT INTO [tbOrder] (OrderDate,OrderHold,OrderRef,Order_DPA_,Order_EIT_,Branch_DPA_,OrderSecType_DPA_,Client_DPA_,OrderType_DPA_,OrderAutoReleaseDate,OrderHoldType_DPA_,OrderCompounded) SELECT " & _
		 "#" & rs.fields("OrderDate") & "#" & " as OrderDate" & _
		 "," & " " & 1 & " " & " as OrderHold" & _
         "," & "'" & rs.fields("OrderRef") & "'" & " as OrderRef" & _ 
         "," & " " & "iif(isnull(max([Order_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'tbOrder'),max([Order_DPA_]) + 1)" & " " & " as Order_DPA_" & _
         "," & "'" & guidStr & "'" & " as Order_EIT_" & _ 
         "," & " " & Session("Branch_DPA_") & " " & " as Branch_DPA_" & _
         "," & " " & rs.fields("OrderSecType_DPA_") & " " & " as OrderSecType_DPA_" & _
         "," & " " & rs.fields("Client_DPA_") & " " & " as Client_DPA_" & _
         "," & " " & rs.fields("OrderType_DPA_") & " " & " as OrderType_DPA_" & _
         "," & " NULL " & " as OrderAutoReleaseDate" & _
         "," & " " & 2 & " " & " as OrderHoldType_DPA_" & _
         "," & " " & abs(cint(rs.fields("OrderCompounded"))) & " " & " as OrderCompounded" & _
         " FROM [tbOrder]"
 Set conn = GetActiveConnection("KBroker")
 conn.BeginTrans
		sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
		conn.Execute sqlStr
				     
		'obtain header key value
		sqlStr = "SELECT [tbOrder.Order_DPA_] FROM [tbOrder] WHERE [tbOrder.Order_EIT_] = " & "'" & guidStr & "'"
				     
	     Dim orderRS
		Set orderRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		If (orderRS.EOF Or orderRS.BOF) Then%>
         			<script language = 'vbscript'>
         					ShowMessage "A serious error has been encountered while saving the data. Try saving again"
				         					
         			</script>
         			<% response.end
		End If
				     
		'save detail data
		sqlStr = "INSERT INTO [OrdDetail] (OrdDetailCertNo,OrdDetailPrice,OrdDetailQty,OrdDetailValidity" & _
				",OrdDetail_DPA_,Order_DPA_,Security_DPA_) SELECT " & _
				"''" & " as OrdDetailCertNo" & _
				"," & "'" & rs.fields("OrdDetailPrice") & "'" & " as OrdDetailPrice" & _
				"," & " " & CDbl(rs.fields("OrdDetailQty")) & " " & " as OrdDetailQty" & _
				"," & "#" & rs.fields("OrdDetailValidity") & "#" & " as OrdDetailValidity" & _
				"," & " " & "iif(isnull(max([OrdDetail_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'OrdDetail'),max([OrdDetail_DPA_]) + 1)" & " " & " as OrdDetail_DPA_" & _
				"," & " " & orderRS.Fields("Order_DPA_") & " " & " as Order_DPA_" & _
				"," & " " & rs.fields("Security_DPA_") & " " & " as Security_DPA_" & _
				" FROM [OrdDetail]"
				     
		sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
		conn.Execute = sqlStr
					
		sqlStr = "UPDATE WebtbOrder SET Accepted = 1 WHERE(Order_DPA_ = " & RS.Fields("Order_DPA_") & ")"
		conn.Execute = sqlStr
conn.CommitTrans
conn.Close
			
end sub

sub SendMail(toRecipient, subject, bodyText)
	cc = ""
	bcc = ""
	
	Set Conn = GetActiveConnection("KBroker")
	SQL = "SELECT * FROM MAilConfigList"
	Set Rs = Conn.Execute(SQL)
	
	If (Rs.EOF Or Rs.BOF) Then%>
		<Script Language="JavaScript">
			alert("The mail configurations have not been set.");
		</Script>
		<%Response.End 
	End If
				
	Const cdoSchema = "http://schemas.microsoft.com/cdo/configuration/" 
	Set objMsg = CreateObject("CDO.Message") 
	
	With objMsg
	
		.Configuration.Fields.Item(cdoSchema & "sendusing") = Rs.Fields("SendUsingMethod").Value 
		.Configuration.Fields.Item(cdoSchema & "smtpserver") = Rs.Fields("SMTPServer").Value 
		.Configuration.Fields.Item(cdoSchema & "smtpserverport") = Rs.Fields("SMTPServerPort").Value 
		.Configuration.Fields.Update 		 
		.Subject = subject
		.Sender = Rs.Fields("SendDisplayName").Value 
		.To = toRecipient
		If cc <> "" Then
			.CC = cc
		End If
		
		If bcc <> "" Then
			.BCC = bcc
		End If
		
		If bodyText <> "" Then
			.TextBody = bodyText
		End If
		
		.HTMLBody = bodyText	' "file://" & pathTo
		
		On Error Resume Next
		.Send 
		
		If Err.number > 0 Then%>
			<Script Language="JavaScript">
				alert("The page could not be sent. An unexpected error occured: <%= Err.description %>")
			</Script>
			<%Response.End 
		End If
	End With
	
	Set Rs = Nothing
	Set Conn = Nothing
	Set objMsg = Nothing
	Set Fso = Nothing%>
		<Script Language="JavaScript">
			alert("The page was sent successfully.")		
		</Script>
	<%	
end sub
%>
