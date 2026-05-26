<!--#include file="../libroutines.asp"-->
<%
'======================= Begin_Alter_Across_Entities =================================
		
		const UDLName = "KBroker"
		const DataSource = "CDSMatchedTradesList"
		const DataEntity = "CDSTrade"
		const DataEntityPlural = "CDSTrades"
		const ActionFolder = "Imports"
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
		
		sqlStr = "SELECT * FROM [" & DataSource & "]"
		
		If filterStr <> "" Then
			sqlStr = "SELECT * FROM [" & DataSource & "] WHERE " & filterStr
		End If
		
		If searchStr <> "" Then
			If InStr(1, sqlStr, " WHERE ") > 0 Then
				sqlStr = sqlStr & " AND " & searchStr
			Else
				sqlStr = "SELECT * FROM [" & DataSource & "] WHERE " & searchStr
			End If
		End If
		
		If sortQryStr <> "" Then
			If InStr(1, sqlStr, " ORDER BY ", vbTextCompare) > 0 Then					
				sqlStr = sqlStr & " , " & sortQryStr	
				'hoping that the last clause in the sql will always be an order by
			Else
				sqlStr = sqlStr & " ORDER BY " & sortQryStr	
			End	If
		End If 
			
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
		<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
		<meta name="ProgId" content="FrontPage.Editor.Document">
		<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css">
			<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
				<SCRIPT language="Javascript" src="../scripts/common.js"></SCRIPT>
				<script language='vbscript'>
					function ItemSelected(itemID)
 							frm<%=DataSource%>.elements("ID").value = itemID
					end function
					
					function SaveInPlaceEdit()
							frm<%=DataSource%>.submit
					end function
				</script>
				<!-- ActiveUI stylesheet and scripts -->
				<link href="../runtime/styles/xp/grid.css" rel="stylesheet" type="text/css">
				<script src="../runtime/activeui.js"></script>
				<!-- Include patches here -->
				<script src="../runtime/paging1.js"></script>
				<!-- grid format -->
				<%	Dim colIndex
					Dim styleStr
					
					for colIndex = 13 to 17
						styleStr = styleStr & ".active-column-" & colIndex & " {"
						styleStr = styleStr & "width: 0px;}" & chr(13)
					next
				%>
				
				
				<style> 					
					.active-controls-grid {height: 100%; font: menu;}
					.active-row-highlight .active-row-cell {background-color: skyblue}
					.active-selection-true, .active-selection-true .active-row-cell {
						color: blue!important;
						background-color: bisque!important;
						}
						
						<%=styleStr%>
						
					.active-column-0 {width: 50px;}
					.active-column-1 {width: 50px;}
					.active-column-2 {width: 70px; text-align: right;}
					.active-column-3 {width: 60px;}
					.active-column-4 {width: 70px;}
					.active-column-5 {width: 40px; text-align: center;}
					.active-column-6 {width: 200px;}
					.active-column-7 {width: 50px;}
					.active-column-8 {width: 80px;text-align: right;}	
					.active-column-9 {width: 80px;}	
					.active-column-10 {width: 70px;}	
					.active-column-11 {width: 50px; text-align: center;}	
					.active-column-12 {width: 80px; text-align: right;}
				//	.active-column-12 {width: 50px;}
							
					
					.active-grid-row,
					.active-grid-row.active-list-item,
					.active-scroll-left .active-list-item {height: 18px;}
				</style>
				<!--CALENDAR -->
				<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
				<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
										
</head>
	<body leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" Scroll="No"> 
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form name='frm<%=DataSource%>' method='post' id='frmMain' action='CommitCDSTrade.asp'>
			<!-- grid data -->
		<% 'row data
		Dim rowCount
		Dim entryID
		Dim displayColor
		Dim commitTrade
		
		quote = chr(34) 
		intRecord = 1
		
		first=1

		Do Until rs.EOF 

'======================= Begin_Alter_Across_Entities =================================
			'if rs.Fields("BalanceQty") > 0 then
					'displayColor = "DarkBlue"
			'else
					displayColor = "Black"
			'end if
			
			entryID = trim(rs.Fields("CDSImport_DPA_")) & "<->" & trim(rs.Fields("OrdDetail_DPA_"))
			
			''Response.write(entryID)
			''Response.end
			
			commitTrade = "<input type=checkbox class='BorderLess' name='chkCommit' onClick = 'CommitCDSTrade(this, " & rs.Fields("CDSImport_DPA_") & ", " & rs.Fields("OrdDetail_DPA_") & ");'>"
            
			if Cint(first) =1 then
			commitTradeAll = "<input type=checkbox class='BorderLess' name='chkCommitAll' onClick = 'CommitAll(this);'>"
			else
			commitTradeAll=""
			end if
			
			first=0

            'row ID
            rowData = rowData & quote & entryID & quote & " : "

            'row data
            rowData = rowData & "["
			rowData = rowData & quote & commitTradeAll & quote & ","
            rowData = rowData & quote & commitTrade & quote & ","
			rowData = rowData & quote & rs.Fields("CommissionRate") & quote & ","            
            rowData = rowData & quote & ApplyDisplayColor(rs.Fields("Order_DPA_")) & quote & ","
            rowData = rowData & quote & ApplyDisplayColor(FormatDate(rs.Fields("OrderDate"))) & quote & ","
            rowData = rowData & quote & ApplyDisplayColor(IIf(CBool(rs.Fields("OrderTypeSale")) = True, "S", "P")) & quote & ","
            rowData = rowData & quote & ApplyDisplayColor(rs.Fields("OrdDetailClient")) & quote & ","
            rowData = rowData & quote & ApplyDisplayColor(rs.Fields("SecurityCode")) & quote & ","
            rowData = rowData & quote & ApplyDisplayColor(FormatNum(rs.Fields("OrdDetailQty"))) & quote & ","
            rowData = rowData & quote & ApplyDisplayColor(FormatNum(rs.Fields("BalanceQty"))) & quote & ","
            rowData = rowData & quote & ApplyDisplayColor(FormatDate(rs.Fields("TradeDate"))) & quote & ","
            rowData = rowData & quote & ApplyDisplayColor(rs.Fields("CDSRef")) & quote & ","
            rowData = rowData & quote & ApplyDisplayColor(FormatNumCommasOnly(rs.Fields("Quantity"))) & quote & ","
            rowData = rowData & quote & ApplyDisplayColor(FormatNum(rs.Fields("Price"))) & quote & ","
            rowData = rowData & quote & ApplyDisplayColor(rs.Fields("BrokerCode")) & quote & ","
            rowData = rowData & quote & ApplyDisplayColor(rs.Fields("OrdDetailType")) & quote & ","
            rowData = rowData & quote & ApplyDisplayColor(rs.Fields("OrdDetailSecType")) & quote & ","
            rowData = rowData & quote & ApplyDisplayColor(rs.Fields("OrderTypeSale")) & quote & ","
            rowData = rowData & quote & ApplyDisplayColor(rs.Fields("Security_DPA_")) & quote & ","
           ' rowData = rowData & quote & ApplyDisplayColor(rs.Fields("Commission")) & quote
            rowData = rowData & "]"
            'build the row IDs array
            rowIDs = rowIDs & quote & entryID & quote
            RowCount = RowCount + 1


			
'======================= End_Alter_Across_Entities =================================

			rs.MoveNext 
			If Not (rs.EOF) Then
                If intRecord >= intPageSize Then Exit Do
                intRecord = intRecord + 1
                rowIDs = rowIDs & ","
                rowData = rowData & ","
            End If


		Loop

'======================= Begin_Alter_Across_Entities =================================%> 
		<script>
			//column titles 
			var colCount = 19;
			var colNames = ["","", "txtCommission", "", "", "","", 
			"", "","", "", "", "","","","","","",""];
			
			var myColumns = ["Commit","Check","Comm Rate","Order No","Date", "Type", "Client", "Security", 
			 "Limit","Balance", "Trade Date", "CDS Ref", "Quantity", "Price","Broker","OrderType","OrderSecType","OrderIsSaleType","SecurityID"];
			
		
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
						myData[rowIndex][i] = "<INPUT TYPE='text' NAME='" + colNames[i] + "' ID='" + colNames[i] + "' VALUE='" + myData[rowIndex][i] + "'onChange = 'EditInPlaceDataChanged();'>";
						
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
						document.frmMain.elements("delAction").value = "Update";			
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
									if(colNames[i]=="Broker")
									{
										myData[prevRow][i] = currentBrokerName;
									}
									else
									{
										if(colNames[i]=="txtDDate")
										{
											myData[prevRow][i] = currentDate;
										}
										else
										{
											myData[prevRow][i] = document.frmMain.elements(colNames[i]).value;
										}
									}
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
			
			function CommitAll(theChk)
			{
			 if (theChk.checked)
			 {
			 SaveInPlaceEdit();
			 }
			}

			function  CommitCDSTrade(theChk,theItem,OrderDetail_DPA_)
			{
			var thestring;
			var laststring;

				document.frmMain.elements("delAction").value = "Execute";
				document.frmMain.elements("OrderDetail_DPA_").value = OrderDetail_DPA_;
				
				//alert (theChk.checked);

				if (theChk.checked)
				{
					if (document.frmMain.elements("CommitParams").value =="")
					{
					document.frmMain.elements("CommitParams").value = theItem + '<->' + OrderDetail_DPA_;
					}
					else
						{
						document.frmMain.elements("CommitParams").value = document.frmMain.elements("CommitParams").value + ',' + theItem + '<->' + OrderDetail_DPA_;
						}
				}
				else
				{
				//alert(theChk.checked);
				thestring=',' + theItem + '<->' + OrderDetail_DPA_;
				laststring=theItem + '<->' + OrderDetail_DPA_;

					if(document.frmMain.elements("CommitParams").value==laststring)
					{
					document.frmMain.elements("CommitParams").value="";
					}
					else
					{
					document.frmMain.elements("CommitParams").value = replaceSubstring(document.frmMain.elements("CommitParams").value,thestring,'') ;
					}
				}
				
				//alert(document.frmMain.elements("CommitParams").value);

				ItemSelected(theItem);
				//SaveInPlaceEdit();
			}			
			
			function replaceSubstring(inputString, fromString, toString) {
				   // Goes through the inputString and replaces every occurrence of fromString with toString
				   var temp = inputString;
				   if (fromString == "") {
					  return inputString;
				   }
				   if (toString.indexOf(fromString) == -1) { // If the string being replaced is not a part of the replacement string (normal situation)
					  while (temp.indexOf(fromString) != -1) {
						 var toTheLeft = temp.substring(0, temp.indexOf(fromString));
						 var toTheRight = temp.substring(temp.indexOf(fromString)+fromString.length, temp.length);
						 temp = toTheLeft + toString + toTheRight;
					  }
				   } else { // String being replaced is part of replacement string (like "+" being replaced with "++") - prevent an infinite loop
					  var midStrings = new Array("~", "`", "_", "^", "#");
					  var midStringLen = 1;
					  var midString = "";
					  // Find a string that doesn't exist in the inputString to be used
					  // as an "inbetween" string
					  while (midString == "") {
						 for (var i=0; i < midStrings.length; i++) {
							var tempMidString = "";
							for (var j=0; j < midStringLen; j++) { tempMidString += midStrings[i]; }
							if (fromString.indexOf(tempMidString) == -1) {
							   midString = tempMidString;
							   i = midStrings.length + 1;
							}
						 }
					  } // Keep on going until we build an "inbetween" string that doesn't exist
					  // Now go through and do two replaces - first, replace the "fromString" with the "inbetween" string
					  while (temp.indexOf(fromString) != -1) {
						 var toTheLeft = temp.substring(0, temp.indexOf(fromString));
						 var toTheRight = temp.substring(temp.indexOf(fromString)+fromString.length, temp.length);
						 temp = toTheLeft + midString + toTheRight;
					  }
					  // Next, replace the "inbetween" string with the "toString"
					  while (temp.indexOf(midString) != -1) {
						 var toTheLeft = temp.substring(0, temp.indexOf(midString));
						 var toTheRight = temp.substring(temp.indexOf(midString)+midString.length, temp.length);
						 temp = toTheLeft + toString + toTheRight;
					  }
				   } // Ends the check to see if the string being replaced is part of the replacement string or not
				   return temp; // Send the updated string back to the user
				} // Ends the "replaceSubstring" function

		    </script> 
		<script> 

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
				<input type = 'hidden' name ='delAction' id = 'delAction' value="">
				<input type = 'hidden' name ='CommitParams' id = 'CommitParams' value="">
                <input type = 'hidden' name ='OrderDetail_DPA_' id = 'OrderDetail_DPA_' value="">
				<input type = 'hidden' name ='EditPage' id = "EditPage" value = "<%=ActionFolder%>/Edit<%=DataEntity%>.asp">
				<input type = 'hidden' name ='AddPage' id = "AddPage" value = "<%=ActionFolder%>/Add<%=DataEntity%>.asp"> 
				<input type = 'hidden' name ='DeletePage' id = "DeletePage" value = "Delete<%=DataEntity%>.asp">
				<input type = 'hidden' name ='ActionPage' id = "ActionPage" value = "<%=DataSource%>.asp">

<%'======================= Begin_Alter_Across_Entities =================================%>
				
				<input type = 'hidden' name ='FilterArgs' id="FilterArgs" value="Order_DPA_:Order No*2;OrderDate:Date*1;OrderTypeSale:Type*0;OrdDetailClient:Client*0;OrdDetailSecurity:Security*0;OrdDetailQty:Limit*2;BalanceQty:Balance*2;CDSRef:CDS Ref*0;Quantity:Quantity*2;Price:Price*2;BrokerCode:Broker*0;TradeDate:Trade Date*1">
				<input type = 'hidden' name ='SortArgs' id="SortArgs" value="Order_DPA_:Order No;OrderDate:Date;OrderTypeSale:Type;OrdDetailClient:Client;OrdDetailSecurity:Security;OrdDetailQty:Limit;BalanceQty:Balance;CDSRef:CDS Ref;Quantity:Quantity;Price:Price;BrokerCode:Broker;TradeDate:Trade Date">
				<input type = 'hidden' name ='SearchArgs' id="SearchArgs" value="Order_DPA_:Order No*2;OrderDate:Date*1;OrderTypeSale:Type*0;OrdDetailClient:Client*0;OrdDetailSecurity:Security*0;OrdDetailQty:Limit*2;BalanceQty:Balance*2;CDSRef:CDS Ref*0;Quantity:Quantity*2;Price:Price*2;BrokerCode:Broker*0;TradeDate:Trade Date*1">
				<input type = 'hidden' name ='dialogLayout' id="dialogLayout" value="height:10em;width:15em">
				
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
			
				''UPDATE COMMISSIONS FOR COMPOUNDED CONTRACTS	
					 
				''RUN STORED PROCEDURE
				Conn.execute("UpdateCompoundedContractCommissions")	

				Set Rs = Nothing
				Set Conn = Nothing
			%>	
			</form> 
			
	
</body> 
</html>