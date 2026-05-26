<!--#include file="../libroutines.asp"-->
<%
'======================= Begin_Alter_Across_Entities =================================
		
		const UDLName = "KBroker"
		const DataSource = "FullOrderList"
		const ActionPage = "OrderListFO"
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
		
		sortQryStr = Request.Form("SelectedSortArgs")
		filterStr = Request.Form("SelectedFilterArgs")
		sqlStrOrig = "SELECT * FROM [" & DataSource & "]"
		
		sqlStr = "SELECT * FROM [" & DataSource & "]"
		
		If filterStr <> "" Then
			sqlStr = "SELECT * FROM [" & DataSource & "] WHERE " & filterStr
		End If
		
		If sortQryStr <> "" Then		
			sqlStr = "SELECT * FROM [" & DataSource & "] ORDER BY " & sortQryStr	
		End If 
		
		
        Set conn = GetActiveConnection(UDLName)
        
        Conn.Execute("BestBalanceQtys")

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
					.active-column-1 {width: 30px;}
					.active-column-2 {width: 50px; text-align: left;}
					.active-column-3 {width: 30px;}
					.active-column-4 {width: 70px;}
					.active-column-5 {width: 70px; text-align: right;}
					.active-column-6 {width: 70px; text-align: right;}	
					.active-column-7 {width: 70px; text-align: right;}	
					.active-column-8 {width: 60px;}	
					.active-column-9 {width: 30px; text-align: right;}		
					.active-column-10 {width: 200px;}
					.active-column-11 {width: 30px;}
					.active-column-12 {width: 70px; text-align: right;}	
					.active-column-13 {width: 60px;}			
					.active-column-14 {width: 100px; }			
					.active-column-15 {width: 100px;}
					.active-column-16 {width: 100px;}
					.active-column-17 {width: 100px;}
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
			if isdate(rs.Fields("OrderDate")) then
			  OrderDate = FormatDate(rs.Fields("OrderDate"))
			else
			  OrderDate = ""
			end if
			
			'row data 
			rowData = rowData & "[" 
			rowData = rowData & quote & OrderDate & quote & ","
			rowData = rowData & quote & IIf(CBool(rs.Fields("OrderTypeSale")) = True, "S", "P") & quote & "," 
			rowData = rowData & quote & rs.Fields("Order_DPA_") & quote & "," 
			rowData = rowData & quote & rs.Fields("OrdDetailPos") & quote & ","
			
			if(trim(rs.Fields("ordDetailSecType"))="Fixed") then
			 rowData = rowData & quote & rs.Fields("BondDescription") & quote & ","
			else
			 rowData = rowData & quote & rs.Fields("SecurityCode") & quote & ","
			end if			
			
			Best="BEST"		
				
			if(rs.Fields("Best")=true) then
				if(rs.Fields("OrderTypeSale")=false) then
				Amount=rs("Amount")
				else
				Amount=0
				end if
				
			rowData = rowData & quote & Best & quote & ","
			
				if(rs("Amount")<>0) then
					if(rs("Price")<>0) then
					rowData = rowData & quote & FormatNumCommasOnly(rs.Fields("Amount")/(rs("Price")*1.020825)) & quote & ","			 
					else
					rowData = rowData & quote & FormatNumCommasOnly(rs.Fields("Amount")/(1.020825)) & quote & ","			 
					end if
				else
				rowData = rowData & quote & FormatNumCommasOnly(rs.Fields("OrdDetailQty")) & quote & ","			 
				end if
			else
			
					if lcase(rs.Fields("OrdDetailSecType")) = "fixed" Then
							rowData = rowData & quote & (FormatNumEx(rs.Fields("OrdDetailPrice"),4)) & quote & ","
					else
							rowData = rowData & quote & (FormatNum(rs.Fields("OrdDetailPrice"))) & quote & ","
					end if
					rowData = rowData & quote & FormatNumCommasOnly(rs.Fields("OrdDetailQty")) & quote & ","			 
					'rowData = rowData & quote & FormatNum(rs.Fields("OrdDetailPrice")) & quote & ","
			Amount=0
			end if
			
			if isdate(rs.Fields("TimeChanged")) then
			 TimeChanged = FormatDateTime(rs.Fields("TimeChanged"),vbshorttime) & "  " & FormatDate(rs.Fields("TimeChanged"))
			else
			 TimeChanged = ""
			end if
			
			rowData = rowData & quote & FormatNum(Amount) & quote & ","			
			rowData = rowData & quote & rs.Fields("OrderRef") & quote & ","
			rowData = rowData & quote & rs.Fields("Client_DPA_") & quote & ","  
			rowData = rowData & quote & rs.Fields("ClientName") & quote & "," 			
			rowData = rowData & quote & IIf(rs.Fields("OrderHold") = True, "Yes", "No") & quote & "," 
			rowData = rowData & quote & FormatNum(rs.Fields("BalanceQty")) & quote & ","
			rowData = rowData & quote & IIf(rs.Fields("InterBank") = True, "Yes", "No") & quote & "," 
			rowData = rowData & quote & rs.Fields("OrderReleasedBy") & quote & "," 			
			rowData = rowData & quote & FormatDate(rs.Fields("OrderDateReleased")) & quote & ","
			rowData = rowData & quote & rs.Fields("ChangedBy") & quote  & ","
			rowData = rowData & quote & TimeChanged & quote  
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
			var colCount = 18;
			var colNames = ["", "", "", "", "", "", 
			"", "", "","","","","","","","","",""];
			
			var myColumns = ["Date",  "Type","Order No", "Line", "Security", "Price", "Quantity","Amount",  
			"Reference","Code","Client","Hold","Balance","Inter Bank","Released By","Date Released","Last Modified By","Time Modified"];
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
				<input type = 'hidden' name ='AddPage' id = "AddPage" value = "<%=ActionFolder%>/Add<%=DataEntity%>FO.asp"> 
				<input type = 'hidden' name ='DeletePage' id = "DeletePage" value = "Delete<%=DataEntity%>.asp">
				<input type = 'hidden' name ='ActionPage' id = "ActionPage" value = "<%=ActionPage%>.asp">
				<input type = 'hidden' name ='SQLStatement' id = "SQLStatement" value = "<%= sqlStrOrig %>">

<%'======================= Begin_Alter_Across_Entities =================================%>
				
				<input type = 'hidden' name ='FilterArgs' id="FilterArgs" value="OrderDate: Date*1;Order_DPA_: Order No*2;SecurityCode: Security*0;OrdDetailPos: Line*2;OrderRef: Reference*0;Client_DPA_: Code*2;ClientName: Client*0;OrderHold: Hold*6;BalanceQty: Balance*2;InterBank: Inter Bank*6;OrderReleasedBy: Released By*0;OrderDateReleased: Date Released*1;ChangedBy: Last Modified By*0;TimeChanged: Time Modified*1">
				<input type = 'hidden' name ='SortArgs' id="SortArgs" value="OrderDate: Date;Order_DPA_: Order No;SecurityCode: Security;OrdDetailPos: Line;OrderRef: Reference;Client_DPA_: Code;ClientName: Client;OrderHold: Hold;BalanceQty: Balance;InterBank: Inter Bank;OrderReleasedBy: Released By;OrderDateReleased: Date Released;ChangedBy: Last Modified By;TimeChanged: Time Modified">
				<input type = 'hidden' name ='SearchArgs' id="SearchArgs" value="OrderDate: Date*1;Order_DPA_: Order No*2;SecurityCode: Security*0;OrdDetailPos: Line*2;OrderRef: Reference*0;Client_DPA_: Code*2;ClientName: Client*0;OrderHold: Hold*6;BalanceQty: Balance*2;InterBank: Inter Bank*6;OrderReleasedBy: Released By*0;OrderDateReleased: Date Released*1;ChangedBy: Last Modified By*0;TimeChanged: Time Modified*1">
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