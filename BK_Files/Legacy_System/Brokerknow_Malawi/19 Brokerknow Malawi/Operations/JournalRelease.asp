<!--#include file="../libroutines.asp"-->
<%
'======================= Begin_Alter_Across_Entities =================================
		
		const UDLName = "KBroker"
		const DataSource = "JournalRelease"
		const ActionPage = "JournalRelease"
		const DataEntity = "Journal"
		const DataEntityPlural = "Journals"
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
					.active-column-0 {width: 50px;}
					.active-column-1 {width: 80px;}
					.active-column-2 {width: 80px;}
					.active-column-3 {width: 80px;}
					.active-column-4 {width: 80px;}
					.active-column-5 {width: 80px; text-align: right;}	
					.active-column-6 {width: 80px; text-align: right;}	
					.active-column-7 {width: 80px;}		
					.active-column-8 {width: 80px;}		
				</style>
										
</head>
	<body leftMargin=0 topMargin=0 marginheight="0" marginwidth="0"> 
	<form name='frm<%=ActionPage%>' method='post' id='frmMain' action='ReleaseJournal.asp'>
			<!-- grid data -->
		<% 'row data
		Dim rowCount
		Dim Release
		quote = chr(34) 
		intRecord = 1
		Do Until rs.EOF 

'======================= Begin_Alter_Across_Entities =================================
			Release = "<input type=checkbox class='BorderLess' style=BorderLess  value='False' name='chkRelease' onClick = 'UpdateReleaseStatus(this, " & rs.Fields("Journal_DPA_") & ");'>"

			'row ID 
			rowData = rowData & quote & rs.Fields("Journal_DPA_") & quote & " : " 
			
			'row data 
			rowData = rowData & "[" 
			rowData = rowData & quote & Release & quote	& ","		
			rowData = rowData & quote & rs.Fields("Journal_DPA_") & quote & ","
			rowData = rowData & quote & FormatDate(rs.Fields("JournalDate")) & quote & ","
			rowData = rowData & quote & rs.Fields("JournalEntryUser") & quote & "," 
			rowData = rowData & quote & rs.Fields("DebitAmount") & quote & ","			 
			rowData = rowData & quote & rs.Fields("CreditAmount") & quote & "," 			
			rowData = rowData & "]" 
			'build the row IDs array 
			rowIDs = rowIDs & quote & rs.Fields("Journal_DPA_") & quote 
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
			var colCount = 6;
			var colNames = ["", "", "", "", "", ""];
			
			var myColumns = ["Release", "No",  "Date", "User", "Debit", "Credit"];
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
			
			function  UpdateReleaseStatus(theChk, theItem)
			{
				var relVal = "0";
				if (theChk.checked)
				{
					relVal = "1";
				}
				
				document.frmMain.elements("Release").value = relVal;
				document.frmMain.elements("ReleaseDate").value = theChk.ReleaseDate;
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
				
				<!--The message in the input below is meant for the Footer page. It is replaced with a different string if deletion proceeds-->
				<input type = 'hidden' name ='delAction' id = 'delAction' value="This action will delete the <%=DataEntity%> that contains the selected item. If you want to delete the whole <%=DataEntity%>, click Yes. Otherwise, click No and then click Edit to delete the item from the <%=DataEntity%>.">
				<!-- ----------------------------------------------------------------------------------------------------------------------- -->
				
				<input type = 'hidden' name ='EditPage' id = "EditPage" value = "<%=ActionFolder%>/Edit<%=DataEntity%>.asp">
				<input type = 'hidden' name ='AddPage' id = "AddPage" value = "<%=ActionFolder%>/Add<%=DataEntity%>.asp"> 
				<input type = 'hidden' name ='DeletePage' id = "DeletePage" value = "Delete<%=DataEntity%>.asp">
				<input type = 'hidden' name ='ActionPage' id = "ActionPage" value = "<%=ActionPage%>.asp">
				<input type = 'hidden' name ='SQLStatement' id = "SQLStatement" value = "<%= sqlStrOrig %>">

<%'======================= Begin_Alter_Across_Entities =================================%>
				
				<input type = 'hidden' name ='FilterArgs' id="FilterArgs" value="JournalDate:Date*3;JournalEntryUser:User*0;JournalEntryDebit:Debit*2;JournalEntryCredit:Credit*2">
				<input type = 'hidden' name ='SortArgs' id="SortArgs" value="JournalDate:Date;JournalEntryUser:User;JournalEntryDebit:Debit;JournalEntryCredit:Credit">
				<input type = 'hidden' name ='SearchArgs' id="SearchArgs" value="JournalDate:Date*3;JournalEntryUser:User*0;JournalEntryDebit:Debit*2;JournalEntryCredit:Credit*2">
				<input type = 'hidden' name ='dialogLayout' id="dialogLayout" value="height:36em;width:46em">
				
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