<!--#include file="../libroutines.asp"-->
<%
'======================= Begin_Alter_Across_Entities =================================
		
		const UDLName = "KBroker"
		'const DataSource = "DB_OrderList"
		const DataSource = "[dbo].[DB_OrderList]"
		
		const ActionPage = "OrderList"
		const DataEntity = "Order"
		const DataEntityPlural = "Orders"
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
		Dim filtered
		
		'filtered=""
		
		sortQryStr = Request.Form("SelectedSortArgs")
		filterStr = Request.Form("SelectedFilterArgs")
		
		'filtered=Request.QueryString("filtered")
		
		sqlStrOrig = "SELECT TOP 100 * FROM [" & DataSource & "]"
		
		if(filterStr="") then		
		'sqlStr = "SELECT TOP 100 * FROM [" & DataSource & "] "

		sqlStr = "SELECT DISTINCT  " & _
				 "                       TOP 100 CONVERT(DATETIME, tbOrder.OrderDate, 108) AS OrderDate, tbOrder.OrderRef, OrdDetail.OrdDetailPrice, OrdDetail.OrdDetailQty,  " & _
				 "                        tbOrder.OrderHold, tbOrder.Order_DPA_, tbOrder.OrderAutoReleaseDate, tbOrder.Client_DPA_,   " & _
				 "                       tbOrder.OrderDateReleased, OrdDetail.Amount, OrdDetail.Best,  " & _
				 "                       tbOrder.TimeChanged, tbOrder.InterBank, Client.ClientName, Client.ClientCDSNo, OrderType.OrderTypeDescription AS OrderTypeName,  " & _
				 "                       OrderType.OrderTypeSale, Security.Security_DPA_, Security.SecurityCode, OrdDetail.OrdDetail_DPA_,OrdDetail.Limit,  " & _
				 "                       Users.OtherNames + ' ' + Users.Surname AS ChangedBy,dbo.OrdDetail.OrdDetailQty - ISNULL(OrdDetailContractedQtyList.ContractQty, 0) AS BalanceQty " & _
				 " FROM         Security INNER JOIN " & _
				 "                       OrdDetail ON Security.Security_DPA_ = OrdDetail.Security_DPA_ INNER JOIN " & _
				 "                       tbOrder INNER JOIN " & _
				 "                       Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ INNER JOIN " & _
				 "                       OrderType ON tbOrder.OrderType_DPA_ = OrderType.OrderType_DPA_ ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ LEFT OUTER JOIN " & _
				 "                           (SELECT     SUM(cast(isnull(dbo.OrdDetail.OrdDetailQty, 0) AS float)) AS TotalHoldings, dbo.OrdDetail.Security_DPA_, dbo.tbOrder.Client_DPA_ " & _
				 "                             FROM          dbo.OrdDetail INNER JOIN " & _
				 "                                                    dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ " & _
				 "                             GROUP BY dbo.OrdDetail.Security_DPA_, dbo.tbOrder.Client_DPA_) a ON Client.Client_DPA_ = a.Client_DPA_ AND  " & _
				 "                       Security.Security_DPA_ = a.Security_DPA_ full outer JOIN " & _
				 "                       Users ON tbOrder.ChangedBy = Users.UserID LEFT OUTER JOIN " & _
				 "                           (SELECT     dbo.OrdDetail.OrdDetail_DPA_, SUM(dbo.Lot.LotQty) AS ContractQty " & _
				 "                             FROM          dbo.OrdDetail INNER JOIN " & _
				 "                                                    dbo.Lot ON dbo.OrdDetail.OrdDetail_DPA_ = dbo.Lot.OrdDetail_DPA_ " & _
				 "                             WHERE      (dbo.Lot.Deleted = 0) AND (dbo.OrdDetail.Deleted = 0) " & _
				 "                             GROUP BY dbo.OrdDetail.OrdDetail_DPA_) OrdDetailContractedQtyList ON  " & _
				 "                       OrdDetail.OrdDetail_DPA_ = OrdDetailContractedQtyList.OrdDetail_DPA_  " & _
				 " WHERE     (tbOrder.OrderCanceled = 0) AND (Client.Deleted = 0) AND (OrdDetail.Deleted = 0) AND (tbOrder.Deleted = 0) " & _
				 " ORDER BY tbOrder.Order_DPA_ DESC"
		end if
		
		If filterStr <> "" Then
			'sqlStr = "SELECT TOP 10 * FROM [" & DataSource & "] WHERE " & filterStr

			sqlStr = "SELECT DISTINCT  " & _
				 "                       TOP 100 CONVERT(DATETIME, tbOrder.OrderDate, 108) AS OrderDate, tbOrder.OrderRef, OrdDetail.OrdDetailPrice, OrdDetail.OrdDetailQty,  " & _
				 "                        tbOrder.OrderHold, tbOrder.Order_DPA_, tbOrder.OrderAutoReleaseDate, tbOrder.Client_DPA_,   " & _
				 "                       tbOrder.OrderDateReleased, OrdDetail.Amount, OrdDetail.Best,  " & _
				 "                       tbOrder.TimeChanged, tbOrder.InterBank, Client.ClientName, Client.ClientCDSNo, OrderType.OrderTypeDescription AS OrderTypeName,  " & _
				 "                       OrderType.OrderTypeSale, Security.Security_DPA_, Security.SecurityCode, OrdDetail.OrdDetail_DPA_,OrdDetail.Limit,  " & _
				 "                       Users.OtherNames + ' ' + Users.Surname AS ChangedBy,dbo.OrdDetail.OrdDetailQty - ISNULL(OrdDetailContractedQtyList.ContractQty, 0) AS BalanceQty " & _
				 " FROM         Security INNER JOIN " & _
				 "                       OrdDetail ON Security.Security_DPA_ = OrdDetail.Security_DPA_ INNER JOIN " & _
				 "                       tbOrder INNER JOIN " & _
				 "                       Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ INNER JOIN " & _
				 "                       OrderType ON tbOrder.OrderType_DPA_ = OrderType.OrderType_DPA_ ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ LEFT OUTER JOIN " & _
				 "                           (SELECT     SUM(cast(isnull(dbo.OrdDetail.OrdDetailQty, 0) AS float)) AS TotalHoldings, dbo.OrdDetail.Security_DPA_, dbo.tbOrder.Client_DPA_ " & _
				 "                             FROM          dbo.OrdDetail INNER JOIN " & _
				 "                                                    dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ " & _
				 "                             GROUP BY dbo.OrdDetail.Security_DPA_, dbo.tbOrder.Client_DPA_) a ON Client.Client_DPA_ = a.Client_DPA_ AND  " & _
				 "                       Security.Security_DPA_ = a.Security_DPA_ full outer JOIN " & _
				 "                       Users ON tbOrder.ChangedBy = Users.UserID LEFT OUTER JOIN " & _
				 "                           (SELECT     dbo.OrdDetail.OrdDetail_DPA_, SUM(dbo.Lot.LotQty) AS ContractQty " & _
				 "                             FROM          dbo.OrdDetail INNER JOIN " & _
				 "                                                    dbo.Lot ON dbo.OrdDetail.OrdDetail_DPA_ = dbo.Lot.OrdDetail_DPA_ " & _
				 "                             WHERE      (dbo.Lot.Deleted = 0) AND (dbo.OrdDetail.Deleted = 0) " & _
				 "                             GROUP BY dbo.OrdDetail.OrdDetail_DPA_) OrdDetailContractedQtyList ON  " & _
				 "                       OrdDetail.OrdDetail_DPA_ = OrdDetailContractedQtyList.OrdDetail_DPA_  " & _
				 " WHERE     (tbOrder.OrderCanceled = 0) AND (Client.Deleted = 0) AND (OrdDetail.Deleted = 0) AND (tbOrder.Deleted = 0) and " & filterStr & _
				 " ORDER BY tbOrder.Order_DPA_ DESC"
		End If
		
		If sortQryStr <> "" Then		
			sqlStr = sqlStr & "," & sortQryStr			
		End If
		intPage = IntToNull(Request("Page"))
		'Response.write(intPage)
		

		'if((filtered="" and IsNull(intPage)) or (filtered="0" and IsNull(intPage))) then
		'if(filterStr="") then
		'SqlStr="SELECT TOP 0 * FROM DB_OrderList"
		'end if

        Set conn = GetActiveConnection(UDLName)
        
		Conn.CommandTimeout = 0

        Conn.Execute("BestBalanceQtys")
        'conn.execute ("ClientTotalsDelete")
		'conn.execute ("ClientTotalsProcedure")
		'conn.execute ("ClientBalancesDelete")
		'conn.execute ("ClientBalancesProcedure")

        Set Rs = Server.CreateObject("ADODB.Recordset")
		Rs.CursorLocation = adUseClient 
		
		'response.write(HandleQuote(sqlStr))
		'Response.end

'		Rs.Open  SQLServerFormat(HandleQuote(sqlStr)), Conn.ConnectionString, adOpenKeySet, adLockOptimistic
		'Rs.Open  SQLServerFormat(HandleQuote(sqlStr)), Conn.ConnectionString, 0, 1
        
		Set Rs=Conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))

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
		<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
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
					
					.active-column-0 {width:70px;}
					.active-column-1 {width: 30px; text-align: center;}
					.active-column-2 {width: 50px; text-align: left;}
					.active-column-3 {width: 60px;}
					.active-column-4 {width: 60px;}
					.active-column-5 {width: 60px; text-align: right;}
					.active-column-6 {width: 70px; text-align: right;}
					.active-column-7 {width: 60px; text-align: right;}	
					.active-column-8 {width: 60px; text-align: center;}	
					.active-column-9 {width:200px; }	
					.active-column-10 {width: 50px;}
					.active-column-11 {width: 70px;}	
					.active-column-12 {width: 100px;}
					.active-column-13 {width: 70px;}	
							
										
				</style>
										
</head>
	<body leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">	 
	<form name='frm<%=ActionPage%>' method='post' id='frmMain' action='Edit<%=DataEntity%>InPlace.asp'>	
			<!-- grid data -->
		<% 'row data
		Dim rowCount
		quote = chr(34) 
		intRecord = 1
		Do Until rs.EOF 

'======================= Begin_Alter_Across_Entities =================================
			'row ID 
			rowData = rowData & quote & rs.Fields("OrdDetail_DPA_") & quote & " : " 
			
			'row data 
			rowData = rowData & "[" 
			%><script language='javascript'>//alert('<%=rs.Fields("OrderTypeSale")%>');</script><%
			rowData = rowData & quote & FormatDate(rs.Fields("OrderDate")) & quote & ","
			rowData = rowData & quote & IIf(CBool(rs.Fields("OrderTypeSale")) = True, "S", "P") & quote & "," 
			rowData = rowData & quote & rs.Fields("Order_DPA_") & quote & "," 			
			rowData = rowData & quote & rs.Fields("SecurityCode") & quote & ","	
			'rowData = rowData & quote & FormatNumCommasOnly(rs.Fields("Balance")) & quote & ","	
			'rowData = rowData & quote & FormatNumEx(rs.Fields("AvailableBalance"),0) & quote & ","
			Best="BEST"	
			
			if(rs.Fields("Best")=true) then
				if(rs.Fields("OrderTypeSale")=false) then
				Amount=rs("Amount")
				else
				Amount=0
				end if
			  rowData = rowData & quote & Best & quote & ","
            'Limit = rs.Fields("Limit")
			   'if (trim(rs.Fields("Limit")) = 0 or isnull(rs.Fields("Limit"))) then
                ' rowData = rowData & quote & "none" & quote & ","
			   'else 
               '  rowData = rowData & quote & FormatNumCommasOnly(rs.Fields("Limit")) & quote & ","
			   'end if
			  
              'rowData = rowData & quote & FormatNumCommasOnly(rs.Fields("Limit")) & quote & ","
			  rowData = rowData & quote & FormatNumCommasOnly(rs.Fields("OrdDetailQty")) & quote & ","			 
			else
			 rowData = rowData & quote & (FormatNumEx(rs.Fields("OrdDetailPrice"),4)) & quote & ","
			 'rowData = rowData & quote & "none" & quote & ","
			 rowData = rowData & quote & FormatNumCommasOnly(rs.Fields("OrdDetailQty")) & quote & ","			 
					
			 Amount=0
			end if
			rowData = rowData & quote & FormatNum(rs.Fields("BalanceQty")) & quote & ","
			'rowData = rowData & quote & FormatNum(Amount) & quote & ","			
			rowData = rowData & quote & rs.Fields("OrderRef") & quote & ","
			rowData = rowData & quote & rs.Fields("Client_DPA_") & quote & ","  
			'rowData = rowData & quote & rs.Fields("ClientCDSNo") & quote & "," 
			rowData = rowData & quote & rs.Fields("ClientName") & quote & "," 	
			'rowData = rowData & quote & FormatNumEx(rs.Fields("CurrentBal"),2) & quote & ","	
			'rowData = rowData & quote & FormatNumEx(rs.Fields("AvailableCredit"),2) & quote & ","	
			rowData = rowData & quote & IIf(rs.Fields("OrderHold") = True, "Yes", "No") & quote & "," 
			'rowData = rowData & quote & IIf(rs.Fields("InterBank") = True, "Yes", "No") & quote & "," 
					
			rowData = rowData & quote & FormatDate(rs.Fields("OrderDateReleased")) & quote & ","
			rowData = rowData & quote & FormatDate(rs.Fields("ChangedBy")) & quote  & ","
			'rowData = rowData & quote & FormatDate(rs.Fields("TimeChanged"),vbshorttime) & "  " & FormatDate(rs.Fields("TimeChanged")) & quote  
			rowData = rowData & quote & rs.Fields("TimeChanged") & quote  
			rowData = rowData & "]" 

			'build the row IDs array 
			rowIDs = rowIDs & quote & rs.Fields("OrdDetail_DPA_") & quote 
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
			var colCount = 14;
			var colNames = ["", "", "", "", "", "", 
			"", "", "","",""];
			
			var myColumns = ["Date",  "Type","Order","Security", "Price",  "Quantity","Balance",  
			"Reference","Code", "Client","Hold","Date Released","Last Modified By","Time Modified"];
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
				
				<!--The message in the input below is meant for the Footer page. It is replaced with a different string if deletion proceeds-->
				<input type = 'hidden' name ='delAction' id = 'delAction' value="This action will delete the Order that contains the selected item. If you want to delete the whole Order, click Yes. Otherwise, click No and then click Edit to delete the item from the Order.">
				<!-- ----------------------------------------------------------------------------------------------------------------------- -->
				
				<input type = 'hidden' name ='EditPage' id = "EditPage" value = "<%=ActionFolder%>/Edit<%=DataEntity%>.asp">
				<input type = 'hidden' name ='AddPage' id = "AddPage" value = "<%=ActionFolder%>/Add<%=DataEntity%>.asp"> 
				<input type = 'hidden' name ='DeletePage' id = "DeletePage" value = "Delete<%=DataEntity%>.asp">
				<input type = 'hidden' name ='ActionPage' id = "ActionPage" value = "<%=ActionPage%>.asp">
				<input type = 'hidden' name ='SQLStatement' id = "SQLStatement" value = "<%= sqlStrOrig %>">

<%'======================= Begin_Alter_Across_Entities =================================%>
				
				<input type = 'hidden' name ='FilterArgs' id="FilterArgs" value="OrderDate:Date*3;OrderTypeName:Type*0;tbOrder.Order_DPA_:Order No*0;SecurityCode:Security*0;OrdDetailPrice:Price*0;Limit:   Limit*2;OrdDetailQty:Quantity*2;OrderRef:Reference*0;Client.Client_DPA_:Client Code*0;ClientName:Client*0;OrderHold:Hold*5;OrderDateReleased:Date Released*5;ClientCDSNo:CDS No*0;(dbo.OrdDetail.OrdDetailQty - ISNULL(OrdDetailContractedQtyList.ContractQty, 0)):Balance*2">
				<input type = 'hidden' name ='SortArgs' id="SortArgs" value="OrderDate:Date;OrderTypeName:Type;tbOrder.Order_DPA_:Order No;SecurityCode:Security;OrdDetailPrice:Price;OrdDetailQty:Limit:  Limit;Quantity;OrderRef:Reference;Client.Client_DPA_:Client Code;ClientName:Client;OrderHold:Hold;(dbo.OrdDetail.OrdDetailQty - ISNULL(OrdDetailContractedQtyList.ContractQty, 0)):Balance">
				<input type = 'hidden' name ='SearchArgs' id="SearchArgs" value="OrderDate:Date*3;OrderTypeName:Type*0;tbOrder.Order_DPA_:Order No*0;SecurityCode:Security*0;OrdDetailPrice:Price*0;Limit: Limit*2;OrdDetailQty:Quantity*2;OrderRef:Reference*0;Client.Client_DPA_:Client Code*0;ClientName:Client*0;OrderHold:Hold*5;OrderDateReleased:Date Released*5;ClientCDSNo:CDS No*0;(dbo.OrdDetail.OrdDetailQty - ISNULL(OrdDetailContractedQtyList.ContractQty, 0)):Balance*2">
				<input type = 'hidden' name ='dialogLayout' id="dialogLayout" value="height:35em;width:55em">
				
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