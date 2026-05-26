<!--#include file="../libroutines.asp"-->
<%
'======================= Begin_Alter_Across_Entities =================================
		
		const UDLName = "KBroker"		
		const ActionPage = "SaleRequests"		
		const DataSource = "SaleRequestList"
		const DataEntity = "SaleRequest"
		const DataEntityPlural = "SaleRequests"
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
		searchStr = Request.Form("SelectedSearchArgs")
		
		sqlStrOrig = "SELECT * FROM [" & DataSource & "]"
		
		sqlStr = "SELECT     Client.Client_DPA_, Client.ClientCDSNo, Client.ClientName, tbOrder.Order_DPA_, lot.Contract_DPA_, Security.SecurityCode, ClientBalances.CurrentBal,  " & _
				"                       Lot.ContractNumber, Lot.LotGrossAmount - a.Levy AS NetAmount,  " & _
				"                       CASE WHEN ClientBalances.CurrentBal < CASE WHEN tbOrder.PayOption = 3 THEN tbOrder.PartialAmount ELSE Lot.LotGrossAmount - a.Levy END THEN " & _
				"                        ClientBalances.CurrentBal ELSE CASE WHEN tbOrder.PayOption = 3 THEN tbOrder.PartialAmount ELSE Lot.LotGrossAmount - a.Levy END END AS Amount, " & _
				"                        cast(floor(cast(Contract.ContractSettlementDate AS float)) AS DateTime) AS ContractSettlementDate, Lot.LotTDate AS TradeDate, b.Balance " & _
				" FROM         OrdDetail INNER JOIN " & _
				"                       tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN " & _
				"                       Lot ON OrdDetail.OrdDetail_DPA_ = Lot.OrdDetail_DPA_ INNER JOIN " & _
				"                       Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ INNER JOIN " & _
				"                       ClientBalances ON Client.Client_DPA_ = ClientBalances.client_DPA_ INNER JOIN " & _
				"                       Security ON OrdDetail.Security_DPA_ = Security.Security_DPA_ LEFT OUTER JOIN " & _
				"                       PaymentRequests ON Lot.Contract_DPA_ = PaymentRequests.Contract_DPA_ INNER JOIN " & _
				"                           (SELECT     SUM(LevyAmount) AS Levy, Contract_DPA_ " & _
				"                             FROM          Levycontract " & _
				"                             WHERE      deleted = 0 " & _
				"                             GROUP BY Contract_DPA_) a ON Lot.Contract_DPA_ = a.Contract_DPA_ INNER JOIN " & _
				"                       Contract ON Lot.Contract_DPA_ = Contract.Contract_DPA_ INNER JOIN " & _
				"                           (SELECT     CASE WHEN tbOrder.PayOption = 3 THEN tbOrder.PartialAmount ELSE SUM(lot.LotGrossAmount) - SUM(d .LevyAmount)  " & _
				"                                                    END - SUM(isnull(PaymentRequests.PaymentAmount, 0)) AS Balance, tbOrder.Order_DPA_ " & _
				"                             FROM          OrdDetail INNER JOIN " & _
				"                                                    Lot ON OrdDetail.OrdDetail_DPA_ = Lot.OrdDetail_DPA_ INNER JOIN " & _
				"                                                    tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN " & _
				"                                                        (SELECT     SUM(LevyAmount) AS LevyAmount, contract_DPA_ " & _
				"                                                          FROM          LevyContract " & _
				"                                                          WHERE      deleted = 0 " & _
				"                                                          GROUP BY Contract_DPA_) d ON Lot.Contract_DPA_ = d .Contract_DPA_ LEFT OUTER JOIN " & _
				"                                                    PaymentRequests ON Lot.Contract_DPA_ = PaymentRequests.Contract_DPA_ " & _
				"                             WHERE      (tbOrder.Deleted = 0) AND (OrdDetail.Deleted = 0) AND (Lot.Deleted = 0) " & _
				"                             GROUP BY tbOrder.PartialAmount, tbOrder.Order_DPA_, tbOrder.PayOption) b ON tbOrder.Order_DPA_ = b.Order_DPA_ " & _
				" WHERE     (OrdDetail.Deleted = 0) AND (Lot.Deleted = 0) AND (tbOrder.Deleted = 0) AND (Client.Deleted = 0) AND (tbOrder.OrderType_DPA_ = 2) AND  " & _
				"                       (tbOrder.PayOption <> 1) AND (PaymentRequests.Request_DPA_ IS NULL) and b.balance<>0 " & _
				" ORDER BY lot.contract_DPA_"
		
		'Response.Write sqlStr
		'Response.End
		
		If filterStr <> "" Then	
			sqlStr = "SELECT * FROM [" & DataSource & "] WHERE " & filterStr
			
			sqlStr = "SELECT     Client.Client_DPA_, Client.ClientCDSNo, Client.ClientName, tbOrder.Order_DPA_, lot.Contract_DPA_, Security.SecurityCode, ClientBalances.CurrentBal,  " & _
					"                       Lot.ContractNumber, Lot.LotGrossAmount - a.Levy AS NetAmount,  " & _
					"                       CASE WHEN ClientBalances.CurrentBal < CASE WHEN tbOrder.PayOption = 3 THEN tbOrder.PartialAmount ELSE Lot.LotGrossAmount - a.Levy END THEN " & _
					"                        ClientBalances.CurrentBal ELSE CASE WHEN tbOrder.PayOption = 3 THEN tbOrder.PartialAmount ELSE Lot.LotGrossAmount - a.Levy END END AS Amount, " & _
					"                        cast(floor(cast(Contract.ContractSettlementDate AS float)) AS DateTime) AS ContractSettlementDate, Lot.LotTDate AS TradeDate, b.Balance " & _
					" FROM         OrdDetail INNER JOIN " & _
					"                       tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN " & _
					"                       Lot ON OrdDetail.OrdDetail_DPA_ = Lot.OrdDetail_DPA_ INNER JOIN " & _
					"                       Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ INNER JOIN " & _
					"                       ClientBalances ON Client.Client_DPA_ = ClientBalances.client_DPA_ INNER JOIN " & _
					"                       Security ON OrdDetail.Security_DPA_ = Security.Security_DPA_ LEFT OUTER JOIN " & _
					"                       PaymentRequests ON Lot.Contract_DPA_ = PaymentRequests.Contract_DPA_ INNER JOIN " & _
					"                           (SELECT     SUM(LevyAmount) AS Levy, Contract_DPA_ " & _
					"                             FROM          Levycontract " & _
					"                             WHERE      deleted = 0 " & _
					"                             GROUP BY Contract_DPA_) a ON Lot.Contract_DPA_ = a.Contract_DPA_ INNER JOIN " & _
					"                       Contract ON Lot.Contract_DPA_ = Contract.Contract_DPA_ INNER JOIN " & _
					"                           (SELECT     CASE WHEN tbOrder.PayOption = 3 THEN tbOrder.PartialAmount ELSE SUM(lot.LotGrossAmount) - SUM(d .LevyAmount)  " & _
					"                                                    END - SUM(isnull(PaymentRequests.PaymentAmount, 0)) AS Balance, tbOrder.Order_DPA_ " & _
					"                             FROM          OrdDetail INNER JOIN " & _
					"                                                    Lot ON OrdDetail.OrdDetail_DPA_ = Lot.OrdDetail_DPA_ INNER JOIN " & _
					"                                                    tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN " & _
					"                                                        (SELECT     SUM(LevyAmount) AS LevyAmount, contract_DPA_ " & _
					"                                                          FROM          LevyContract " & _
					"                                                          WHERE      deleted = 0 " & _
					"                                                          GROUP BY Contract_DPA_) d ON Lot.Contract_DPA_ = d .Contract_DPA_ LEFT OUTER JOIN " & _
					"                                                    PaymentRequests ON Lot.Contract_DPA_ = PaymentRequests.Contract_DPA_ " & _
					"                             WHERE      (tbOrder.Deleted = 0) AND (OrdDetail.Deleted = 0) AND (Lot.Deleted = 0) " & _
					"                             GROUP BY tbOrder.PartialAmount, tbOrder.Order_DPA_, tbOrder.PayOption) b ON tbOrder.Order_DPA_ = b.Order_DPA_ " & _
					" WHERE     (OrdDetail.Deleted = 0) AND (Lot.Deleted = 0) AND (tbOrder.Deleted = 0) AND (Client.Deleted = 0) AND (tbOrder.OrderType_DPA_ = 2) AND  " & _
					"                       (tbOrder.PayOption <> 1) AND (PaymentRequests.Request_DPA_ IS NULL) " & _
					" and b.balance <> 0 and " & filterStr & " "
		
		End If
		
		If searchStr <> "" Then
			If InStr(1, sqlStr, " WHERE ") > 0 Then
				sqlStr = sqlStr & " AND " & searchStr				
			Else
				sqlStr = "SELECT * FROM [" & DataSource & "] WHERE " & searchStr
				
				sqlStr = "SELECT     Client.Client_DPA_, Client.ClientCDSNo, Client.ClientName, tbOrder.Order_DPA_, lot.Contract_DPA_, Security.SecurityCode, ClientBalances.CurrentBal,  " & _
						"                       Lot.ContractNumber, Lot.LotGrossAmount - a.Levy AS NetAmount,  " & _
						"                       CASE WHEN ClientBalances.CurrentBal < CASE WHEN tbOrder.PayOption = 3 THEN tbOrder.PartialAmount ELSE Lot.LotGrossAmount - a.Levy END THEN " & _
						"                        ClientBalances.CurrentBal ELSE CASE WHEN tbOrder.PayOption = 3 THEN tbOrder.PartialAmount ELSE Lot.LotGrossAmount - a.Levy END END AS Amount, " & _
						"                        cast(floor(cast(Contract.ContractSettlementDate AS float)) AS DateTime) AS ContractSettlementDate, Lot.LotTDate AS TradeDate, b.Balance " & _
						" FROM         OrdDetail INNER JOIN " & _
						"                       tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN " & _
						"                       Lot ON OrdDetail.OrdDetail_DPA_ = Lot.OrdDetail_DPA_ INNER JOIN " & _
						"                       Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ INNER JOIN " & _
						"                       ClientBalances ON Client.Client_DPA_ = ClientBalances.client_DPA_ INNER JOIN " & _
						"                       Security ON OrdDetail.Security_DPA_ = Security.Security_DPA_ LEFT OUTER JOIN " & _
						"                       PaymentRequests ON Lot.Contract_DPA_ = PaymentRequests.Contract_DPA_ INNER JOIN " & _
						"                           (SELECT     SUM(LevyAmount) AS Levy, Contract_DPA_ " & _
						"                             FROM          Levycontract " & _
						"                             WHERE      deleted = 0 " & _
						"                             GROUP BY Contract_DPA_) a ON Lot.Contract_DPA_ = a.Contract_DPA_ INNER JOIN " & _
						"                       Contract ON Lot.Contract_DPA_ = Contract.Contract_DPA_ INNER JOIN " & _
						"                           (SELECT     CASE WHEN tbOrder.PayOption = 3 THEN tbOrder.PartialAmount ELSE SUM(lot.LotGrossAmount) - SUM(d .LevyAmount)  " & _
						"                                                    END - SUM(isnull(PaymentRequests.PaymentAmount, 0)) AS Balance, tbOrder.Order_DPA_ " & _
						"                             FROM          OrdDetail INNER JOIN " & _
						"                                                    Lot ON OrdDetail.OrdDetail_DPA_ = Lot.OrdDetail_DPA_ INNER JOIN " & _
						"                                                    tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN " & _
						"                                                        (SELECT     SUM(LevyAmount) AS LevyAmount, contract_DPA_ " & _
						"                                                          FROM          LevyContract " & _
						"                                                          WHERE      deleted = 0 " & _
						"                                                          GROUP BY Contract_DPA_) d ON Lot.Contract_DPA_ = d .Contract_DPA_ LEFT OUTER JOIN " & _
						"                                                    PaymentRequests ON Lot.Contract_DPA_ = PaymentRequests.Contract_DPA_ " & _
						"                             WHERE      (tbOrder.Deleted = 0) AND (OrdDetail.Deleted = 0) AND (Lot.Deleted = 0) " & _
						"                             GROUP BY tbOrder.PartialAmount, tbOrder.Order_DPA_, tbOrder.PayOption) b ON tbOrder.Order_DPA_ = b.Order_DPA_ " & _
						" WHERE     (OrdDetail.Deleted = 0) AND (Lot.Deleted = 0) AND (tbOrder.Deleted = 0) AND (Client.Deleted = 0) AND (tbOrder.OrderType_DPA_ = 2) AND  " & _
						"                       (tbOrder.PayOption <> 1) AND (PaymentRequests.Request_DPA_ IS NULL) " & _
						" and b.balance <> 0 and " & searchStr & " "
		
			End If
		End If
		
		If sortQryStr <> "" Then
			If InStr(1, sqlStr, " ORDER BY ") > 0 Then
				sqlStr = sqlStr & " , " & sortQryStr 'hoping that the last clause in the sql will always be an order by
			Else
				sqlStr = sqlStr & " ORDER BY " & sortQryStr	
			End	If
		End If 
			
	    
	    'Response.Write sqlStr
	    'Response.End
	    
        Set conn = GetActiveConnection(UDLName)
		 
		Set Rs = Server.CreateObject("ADODB.Recordset")
		Rs.CursorLocation = adUseClient 
		Rs.Open  SQLServerFormat(HandleQuote(sqlStr)), Conn.ConnectionString, adOpenKeySet, adLockOptimistic
        
        If rs.EOF Or rs.BOF Then
        
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
					.active-column-0 {width: 50px; text-align: left;}	
					.active-column-1 {width: 50px; text-align: left;}
					.active-column-2 {width: 100px;}					
					.active-column-3 {width: 150px; text-align: left;}
					.active-column-4 {width: 50px;}
					.active-column-5 {width: 80px;}
					.active-column-6 {width: 75px;}					
					.active-column-7 {width: 80px; text-align: right;}
					.active-column-8 {width: 80px; text-align: right;}
					.active-column-9 {width: 80px; text-align: right;}										
					.active-column-10 {width: 80px; text-align: right;}										
					
				</style>
										
</head>
	<body leftMargin=0 topMargin=0 marginheight="0" marginwidth="0"> 
	<form name='frm<%=ActionPage%>' method='post' id='frmMain' action='ConfirmRequest.asp''>
			<!-- grid data -->
		<% 'row data
		Dim rowCount
		quote = chr(34) 
		intRecord = 1
		
		Topay=0
		
		Do Until rs.EOF 

'======================= Begin_Alter_Across_Entities =================================
			'row ID 			
			'ContractID = rs.Fields("Contract_DPA_") & "<->" & rs("Client_DPA_")
			
			if(CCur(rs.Fields("CurrentBal"))=0) then
			ConfirmSale=""			
			else
			ConfirmSale =  "<input type=checkbox class='BorderLess' style=BorderLess  value='False' name='chkRelease' onClick = 'UpdateReleaseStatus(this, " & rs.Fields("Contract_DPA_") & "," & rs("Amount") & "," & rs("Client_DPA_") & "," & rs("ContractSettlementDate") & ");'>"
			end if
			rowData = rowData & quote & rs.Fields("Contract_DPA_") & "<->" & rs("Client_DPA_") & quote & " : " 
			'rowData = rowData & quote & rs.Fields("Request_DPA_") & "<->" & rs.Fields("Client_DPA_") & "<->" & rs("Processed_DPA_") & quote & " : "
			'row data 
			rowData = rowData & "["
			rowData = rowData & quote & ConfirmSale & quote	& ","		
			rowData = rowData & quote & rs.Fields("Client_DPA_") & quote & ","
			rowData = rowData & quote & rs.Fields("ClientCDSNo") & quote & "," 
			rowData = rowData & quote & Mid(rs.Fields("ClientName"),1,30) & quote & ","						
			rowData = rowData & quote & rs.Fields("Contract_DPA_") & quote & "," 
			rowData = rowData & quote & FormatDate(rs.Fields("TradeDate")) & quote & ","
			rowData = rowData & quote & rs.Fields("SecurityCode") & quote & ","			
			
			if(Ccur(rs.Fields("NetAmount"))<CCur(rs.Fields("Balance"))) then
			Topay = rs.Fields("NetAmount")
			else
			Topay = rs.Fields("Balance")
			end if
			
			if(CCur(Topay) > CCur(rs.Fields("CurrentBal"))) then			
			Topay = rs.Fields("CurrentBal")
			end if
			
			rowData = rowData & quote & FormatNum(rs.Fields("NetAmount")) & quote & ","
			rowData = rowData & quote & FormatNum(rs.Fields("Balance")) & quote & "," 			 			
			rowData = rowData & quote & FormatNum(rs.Fields("CurrentBal")) & quote & ","
			rowData = rowData & quote & FormatNum(Topay) & quote & "," 			
			
			rowData = rowData & "]" 
			'build the row IDs array 
			rowIDs = rowIDs & quote & rs.Fields("Contract_DPA_") & "<->" & rs("Client_DPA_") & quote 
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
			var colCount = 11;
			var colNames = ["", "", "", "", "","","","","","",""];
			
			var myColumns = ["Confirm","Code","CDS NO","Client","Contract","Trade Date","Security", "Amount", "Balance","Cash","To Pay"];
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
/////==========================  Begin_CHanged  05 Apr 2004 ============================================================///////////////////////////////			
			var RowChangeFn = function(src)
			{
				if(inPlaceEdit)
				{
					if(dataChanged)
					{
						ItemSelected(prevRow);
						SaveInPlaceEdit();
					}
					else
					{
						if(prevRow != clickedRowID)
						{
							var i;
							for(i = 0; i < colCount; i++)
							{
								if(colNames[i] != "")
								{
									myData[prevRow][i] = document.frmMain.elements(colNames[i]).value;
								}
							}
							inPlaceEdit = false;
							prevRow = -1;
							grid.refresh();
						}
					}
				}
			}
			
			var HandleClick = function(src)
			{
				clickedRowID = src.getProperty("row/index");
				ItemSelected(clickedRowID);
			}
			
			function  UpdateReleaseStatus(theChk, theItem,RequestAmount,ClientID,PDate)
			{
				var relVal = "0";
				if (theChk.checked)
				{
					relVal = "1";
				}
				
				//alert(ClientID);
				//document.frmMain.elements("Release").value = relVal;
				//document.frmMain.elements("ReleaseDate").value = theChk.ReleaseDate;
				document.frmMain.elements("delAction").value = "Execute";
				document.frmMain.elements("ClientDPA").value = ClientID;
				//document.frmMain.elements("ClientDPA") = ClientID;
				//document.frmMain.elements("PayDate") = PDate;
				
				ItemSelected(theItem);
				SaveInPlaceEdit();
			}
			
			window.onload = function()
			{
				//select first item
				grid.setSelectionValues([myRowIDs[0]]);
				ItemSelected(myRowIDs[0]);
			}
							
/////==========================  End_CHanged   05 Apr 2004 =============================================================///////////////////////////////
		</script> 
		<script> 

			// create ActiveUI Grid javascript object 
			var grid = new Active.Controls.Grid; 
			//grid.setModel("row", new Active.Controls.Page);
			
			// set rows ids 
			//grid.setProperty("row/count", <%=rowCount%>);
			//grid.setProperty("row/values", myRowIDs);//
			grid.setRowValues(myRowIDs); 
			

			// set number of columns 
			//grid.setProperty("column/count", colCount);//
			grid.setColumnCount(colCount); 

			// provide cells and headers text 
			//grid.setProperty("data/text", function(i, j){return myData[i][j]});//
			grid.setDataText(function(i, j){return myData[i][j]}); 
			//grid.setProperty("column/texts", myColumns);//
			grid.setColumnText(function(i){return myColumns[i]}); 

			//introduce paging
			//grid.setProperty("row/pageSize", 5);
			
			// set click action handler 
			grid.setAction("click", HandleClick); 
			grid.setAction("dblclick", RowEditFn); 
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

			// write grid html to the page 
			document.write(grid);
			
			//let grid be aware of composite layout
			grid.getLayoutTemplate().action("adjustSize");
		 
		</script> 
			
				<input type = 'hidden' name ='ID' id = 'ID'>
				<input type = 'hidden' name ='Release' id = 'Release'>
				<input type = 'hidden' name ='ReleaseDate' id = 'ReleaseDate'>
				<input type = 'hidden' name ='ClientDPA' id = 'ClientDPA'>
				<input type = 'hidden' name ='Amount' id = 'Amount'>
				<input type = 'hidden' name ='PayDate' id = 'PayDate'>
				
				<input type = 'hidden' name ='delAction' id = 'delAction' value="">
				<input type = 'hidden' name ='EditPage' id = "EditPage" value = "<%=ActionFolder%>/Edit<%=DataEntity%>.asp">
				<input type = 'hidden' name ='AddPage' id = "AddPage" value = "<%=ActionFolder%>/Add<%=DataEntity%>.asp"> 
				<input type = 'hidden' name ='DeletePage' id = "DeletePage" value = "Delete<%=DataEntity%>.asp">
				<input type = 'hidden' name ='ActionPage' id = "ActionPage" value = "<%=ActionPage%>.asp">

<%'======================= Begin_Alter_Across_Entities =================================%>
				
				<input type = 'hidden' name ='FilterArgs' id="FilterArgs" value="Client_DPA_:Code*0;ClientName:Client Name*0;Amount:Amount*0;ClientCDSNo:CDS No*0;CurrentBal:Credit*2;SecurityCode:Security*0">
				<input type = 'hidden' name ='SortArgs' id="SortArgs" value="Client_DPA_:Code;ClientName:Client Name;Amount:Amount;ClientCDSNo:CDS No;CurrentBal:Credit;SecurityCode:Security">
				<input type = 'hidden' name ='SearchArgs' id="SearchArgs" value="Client_DPA_:Code*0;ClientName:Client Name*0;Amount:Amount*0;ClientCDSNo:CDS No*0;CurrentBal:Credit*2;SecurityCode:Security*0">
				<input type = 'hidden' name ='dialogLayout' id="dialogLayout" value="height:28em;width:38em">
				
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