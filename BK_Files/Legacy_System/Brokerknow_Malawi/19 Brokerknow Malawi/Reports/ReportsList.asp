
<!--#include file="../libroutines.asp"-->
<%

Dim userID

UserID = Session("UserID")

If UserID = "" Then
	'session expired
	Response.Write ""
	Response.End 
End If


'======================= Begin_Alter_Across_Entities =================================
		
		const UDLName = "KBroker"
		const DataSource = "ReportsList"
		const DataEntity = "Report"
		const DataEntityPlural = "Reports"
		const ActionFolder = "Reports"
'======================= End_Alter_Across_Entities =================================	

		Dim conn 
		Dim sqlStr
		Dim rs
		Dim bcolor
		Dim rowIDs
		Dim rowData
		Dim quote
		Dim sortQryStr
		Dim mnuCatID
		
		mnuCatID = Request("mnuID")
		userID = Session("UserID")
		
		sortQryStr = Request.Form("SelectedSortArgs")
		filterStr = Request.Form("SelectedFilterArgs")
		searchStr = Request.Form("SelectedSearchArgs")
		
		

		sqlStrOrig = "SELECT * FROM Menus WHERE MainMenuID = " & mnuCatID & " AND IsReport = 1 AND EXISTS(SELECT     MenuGroups.ID " & _
				" FROM         UserGroups INNER JOIN " & _
				"                      MenuGroups ON UserGroups.GroupID = MenuGroups.groupID " & _
				"			WHERE     (UserGroups.UserID = " & userID & ") AND (MenuGroups.MenuID = Menus.menuID))  ORDER BY  mnuCaption"
		
		
		
		sqlStr = sqlStrOrig
		
		If filterStr <> "" Then
			sqlStr = Replace(sqlStr, " ORDER BY ", " AND  (" & filterStr & ")  ORDER BY ") 
		End If
		
		If searchStr <> "" Then
			sqlStr = Replace(sqlStr, " ORDER BY ", " AND  (" & searchStr & ")  ORDER BY ") 
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
				        
        Set rs = conn.Execute(sqlStr)
        If rs.EOF Or rs.BOF Then
        
		Else
			rs.MoveFirst	
        End If
        
%>
<html>
	<head>
		<title><%=DataEntityPlural%></title>
		<meta http-equiv="Content-Language" content="en-uk">
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
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
				<link href="../runtime/classic/activeui.css" rel="stylesheet" type="text/css">
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
					.active-column-0 {width: 200px;}
					.active-column-1 {width: 400px;}	
				</style>
										
</head>
	<body leftMargin=0 topMargin=0 marginheight="0" marginwidth="0"> 
	<form name='frm<%=DataSource%>' method='post' id='frmMain' action='Edit<%=DataEntity%>InPlace.asp'>
			<!-- grid data -->
		<% 'row data
		Dim rowCount
		quote = chr(34) 
		Do Until rs.EOF 

'======================= Begin_Alter_Across_Entities =================================
			'row ID 
			rowData = rowData & quote & rs.Fields("MenuID") & quote & " : " 
			
			'row data 
			rowData = rowData & "[" 
			rowData = rowData & quote & rs.Fields("mnuCaption") & quote & "," 
			rowData = rowData & quote & rs.Fields("mnuDescription") & quote
			rowData = rowData & "]" 
			'build the row IDs array 
			rowIDs = rowIDs & quote & rs.Fields("MenuID") & quote 
			rowCount = rowCount + 1
			
'======================= End_Alter_Across_Entities =================================

			rs.MoveNext 
			if not(rs.eof) then 
				rowIDs = rowIDs & "," 
				rowData = rowData & "," 
			end if 
		Loop
		
'======================= Begin_Alter_Across_Entities =================================%> 
		<script>
			//column titles 
			var colCount = 2;
			var colNames = ["", ""];
			
			var myColumns = ["Name", "Description"];
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
			
			var launchReport = function(src)
			{
				try{
					window.parent.frames["footer"].document.parentWindow.execScript('DoAdd()', 'JavaScript');
				}
				catch(e){}	
			}
/////==========================  End_CHanged   05 Apr 2004 =============================================================///////////////////////////////
		</script> 
		<script> 

			// create ActiveUI Grid javascript object 
			var grid = new Active.Controls.Grid; 
			
			//	set the first column template to image+text
			grid.setTemplate("column", new Active.Templates.Image, 0);

			
			// set rows ids 
			grid.setRowValues(myRowIDs); 
			

			// set number of columns 
			grid.setColumnCount(colCount); 

			// provide cells and headers text 
			grid.setDataText(function(i, j){return myData[i][j]}); 
			grid.setProperty("data/image", "txt");
			grid.setColumnText(function(i){return myColumns[i]}); 

			
			// set click action handler 
			grid.setAction("click", HandleClick); 
			grid.setAction("dblclick", launchReport); 
			grid.setAction("selectionChanged", RowChangeFn); 

			//stripes 
			//var alternate = function(){ return this.getProperty("row/order") % 2 ? "gainsboro" : "white";} 
			var row = new Active.Templates.Row;
			// row.setStyle("background", alternate); 
			//row.setEvent("onmouseover", "mouseover(this, 'active-row-highlight')"); 
			//row.setEvent("onmouseout", "mouseout(this, 'active-row-highlight')"); 			
			//grid.setTemplate("row", row); 
			//var column = new Active.Templates.Text; 
			//column.setStyle("border-right", "1px solid white");  
			//grid.setTemplate("column", column);  grid.setRowHeaderWidth("0px"); 
			
			grid.setRowHeaderWidth("0px");

			// write grid html to the page 
			document.write(grid); 
		</script> 
			
				<input type = 'hidden' name ='ID' id = 'ID'>
				<input type = 'hidden' name ='delAction' id = 'delAction' value="">
				<input type = 'hidden' name ='EditPage' id = "EditPage" value = "<%=ActionFolder%>/Edit<%=DataEntity%>.asp">
				<input type = 'hidden' name ='AddPage' id = "AddPage" value = "<%=ActionFolder%>/Add<%=DataEntity%>.asp"> 
				<input type = 'hidden' name ='DeletePage' id = "DeletePage" value = "Delete<%=DataEntity%>.asp">
				<input type = 'hidden' name ='ActionPage' id = "ActionPage" value = "<%=DataSource%>.asp">
				<input type = 'hidden' name ='mnuID' id = "mnuID" value = "<%= mnuCatID %>">
				

<%'======================= Begin_Alter_Across_Entities =================================%>
				
				<input type = 'hidden' name ='FilterArgs' id="FilterArgs" value="mnuCaption:Name*0;mnuDescription:Description*0">
				<input type = 'hidden' name ='SortArgs' id="SortArgs" value="mnuCaption:Name;mnuDescription:Description">
				<input type = 'hidden' name ='SearchArgs' id="SearchArgs" value="mnuCaption:Name*0;mnuDescription:Description*0">
				<input type = 'hidden' name ='dialogLayout' id="dialogLayout" value="height:28em;width:21em">
				
<%'======================= End_Alter_Across_Entities =================================%>
				
				<input type = 'hidden' name ='SelectedSortArgs' id="SelectedSortArgs" value="<%= sortQryStr %>">
				<input type = 'hidden' name ='SelectedFilterArgs' id="SelectedFilterArgs" value="<%= filterStr %>">
				<input type = 'hidden' name ='SelectedSearchArgs' id="SelectedSearchArgs" value="<%= searchStr %>">
			</form> 
			
			
			
</body>
  
</html>

