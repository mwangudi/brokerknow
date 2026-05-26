<!--#include file="../libroutines.asp"-->
<%
'======================= Begin_Alter_Across_Entities =================================
		
		const UDLName = "KBroker"
		const DataSource = "BrokerList"
		const DataEntity = "Broker"
		const DataEntityPlural = "Brokers"
		const ActionFolder = "Data"
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
			If InStr(1, sqlStr, " ORDER BY ") > 0 Then
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
		<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
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
					<!-- grid format -->
					<style> 
						
						.active-controls-grid {height: 100%; font: menu;}
						.active-row-highlight .active-row-cell {background-color: skyblue}
						.active-selection-true, .active-selection-true .active-row-cell {
							color: blue!important;
							background-color: bisque!important;
							}
							
						.active-column-0 {width: 200px;}
						.active-column-1 {width: 70px;}		
						.active-column-2 {width: 130px;}
						.active-column-3 {width: 100px;}		
						.active-column-4 {width: 250px;}
						
					</style>
										
</head>
	<body> 
	<form name='frm<%=DataSource%>' method='post' id='frmMain' action='Edit<%=DataEntity%>InPlace.asp'>
			<!-- grid data -->
		<% 'row data
		quote = chr(34) 
		intRecord = 1 
		Do Until rs.EOF 

'======================= Begin_Alter_Across_Entities =================================
			'row ID 
			rowData = rowData & quote & rs.Fields("Broker_DPA_") & quote & " : " 
			
			'row data 
			rowData = rowData & "[" 
			rowData = rowData & quote & rs.Fields("BrokerName") & quote & "," 
			rowData = rowData & quote & rs.Fields("BrokerCode") & quote & ","
			rowData = rowData & quote & rs.Fields("BrokerOfficeTel") & quote & "," 
			rowData = rowData & quote & rs.Fields("BrokerFax") & quote & "," 
			rowData = rowData & quote & rs.Fields("BrokerAddr") & quote 
			rowData = rowData & "]" 
			'build the row IDs array 
			rowIDs = rowIDs & quote & rs.Fields("Broker_DPA_") & quote 
			
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
			var colCount = 5;
			var colNames = ["", "", "", 
				 "", ""];
			
			var myColumns =  ["Name", "Code", "Office Phone", 
				 "Fax", "Address"];
		</script>
<%'======================= End_Alter_Across_Entities =================================%>			
		<script>
			//data
			var myData = {<%=rowData%>}; 
			var myRowIDs = [<%=rowIDs%>]; 
			
			//editing
			var inPlaceEdit = false;
			var dataChanged = false;
			var prevRow = -1;
			
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
				}
			}
			
			var HandleClick = function(src)
			{
				ItemSelected(src.getProperty("row/index"));
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

			// write grid html to the page 
			document.write(grid);
			
			//let grid be aware of composite layout
			grid.getLayoutTemplate().action("adjustSize"); 
		</script> 
			
				<input type = 'hidden' name ='ID' id = 'ID'>
				<input type = 'hidden' name ='delAction' id = 'delAction' value="">
				<input type = 'hidden' name ='EditPage' id = "EditPage" value = "<%=ActionFolder%>/Edit<%=DataEntity%>.asp">
				<input type = 'hidden' name ='AddPage' id = "AddPage" value = "<%=ActionFolder%>/Add<%=DataEntity%>.asp"> 
				<input type = 'hidden' name ='DeletePage' id = "DeletePage" value = "Delete<%=DataEntity%>.asp">
				<input type = 'hidden' name ='ActionPage' id = "ActionPage" value = "<%=DataSource%>.asp">

<%'======================= Begin_Alter_Across_Entities =================================%>
				
				<input type = 'hidden' name ='FilterArgs' id="FilterArgs" value="BrokerName:Name*0;BrokerOfficeTel:Office Phone*0;BrokerFax:Fax*0;BrokerAddr:Address*0">
				<input type = 'hidden' name ='SortArgs' id="SortArgs" value="BrokerName:Name;BrokerOfficeTel:Office Phone;BrokerFax:Fax;BrokerAddr:Address">
				<input type = 'hidden' name ='SearchArgs' id="SearchArgs" value="BrokerName:Name*0;BrokerOfficeTel:Office Phone*0;BrokerFax:Fax*0;BrokerAddr:Address*0">
				<input type = 'hidden' name ='dialogLayout' id="dialogLayout" value="height:17em;width:17em">
				
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