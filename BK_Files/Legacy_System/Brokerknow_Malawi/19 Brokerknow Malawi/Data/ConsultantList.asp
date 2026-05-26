<!--#include file="../libroutines.asp"-->
<%
'======================= Begin_Alter_Across_Entities =================================
		
		const UDLName = "kbroker"
		const DataSource = "ConsultantList"
		const ActionPage = "ConsultantList"
		const DataEntity = "Consultant"
		const DataEntityPlural = "Consultants"
		const ActionFolder = "ConsultantDB"

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
		Rs.Open  sqlStr, Conn.ConnectionString, adOpenKeySet, adLockOptimistic
        
        If rs.EOF Or rs.BOF Then
        
		 Else
			rs.MoveFirst	
	
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
					function DisplayConsultant()
							frm<%=ActionPage%>.submit
					end function
				</script>
				<!-- ActiveUI stylesheet and scripts -->
				<link href="../runtime/classic/activeui.css" rel="stylesheet" type="text/css">
				<script src="../runtime/activeui.js"></script>
				<!-- Include patches here -->
				<script src="../runtime/paging1.js"></script>
				<!-- grid format -->
        <style> .active-controls-grid {height: 100%; font: menu;}
					.active-row-highlight .active-row-cell {background-color: skyblue}
					.active-selection-true, .active-selection-true .active-row-cell {
						color: blue!important;
						background-color: bisque!important;
						}
				</style>
										
</head>
	<body leftMargin=0 topMargin=0 marginheight="0" marginwidth="0"> 
	<form name='frm<%=ActionPage%>' method='post' id='frmMain' action='../<%=ActionFolder%>/ShowConsultant.asp'>
			<!-- grid data -->
		<% 'row data
		Dim rowCount
		quote = chr(34) 
		intRecord = 1
		Do Until rs.EOF 

'======================= Begin_Alter_Across_Entities =================================
			'row ID 
			rowData = rowData & quote & rs.Fields("Consultant_DPA_") & quote & " : " 
			
			'row data 
			rowData = rowData & "[" 
			rowData = rowData & quote & "<a href='javascript:ShowDetails()'>" & rs.Fields("ConsultantName") & "</a>" & quote & "," 
			rowData = rowData & quote & rs.Fields("ConsultantCountry") & quote & "," 
			rowData = rowData & quote & rs.Fields("ConsultantContact") & quote & ","
			rowData = rowData & quote & rs.Fields("ConsultantEmail") & quote & "," 
			rowData = rowData & quote & rs.Fields("ConsultantPhone") & quote & "," 
			rowData = rowData & quote & rs.Fields("ConsultantRates") & quote & "," 
			rowData = rowData & quote & rs.Fields("ConsultantCV") & quote 

			rowData = rowData & "]" 
			'build the row IDs array 
			rowIDs = rowIDs & quote & rs.Fields("Consultant_DPA_") & quote 
			rowCount = rowCount + 1
			
'======================= End_Alter_Across_Entities =================================

			rs.MoveNext 
			if not(rs.eof) then 
				intRecord = intRecord + 1				
				rowIDs = rowIDs & "," 
				rowData = rowData & "," 
			end if 
			
			
			
			
		Loop
		
'======================= Begin_Alter_Across_Entities =================================%> 
		<script>
			//column titles 
			var colCount = 7;
			var colNames = ["", "", "", "", "", 
			"", ""];
			
			var myColumns = ["Name", "Country", "Contact", "Email", "Phone", 
			"Rates", "CV"];
		</script>
<%'======================= End_Alter_Across_Entities =================================%>			
		<script>
			//data
			var myData = {<%=rowData%>}; 
			var myRowIDs = [<%=rowIDs%>]; 
			
			//editing
			var inPlaceEdit = true;
			var clickedRowID = 0; 
			var dataChanged = true;
			var prevRow = -1;//the row currently under in-place edit
			
			var ShowDetails = function(src)
			{
				if(inPlaceEdit)
				{
					if(dataChanged)
					{
						//ItemSelected(prevRow);
						DisplayConsultant();
					}
				}
			}

			
			var HandleClick = function(src)
			{
				clickedRowID = src.getProperty("row/index");
				ItemSelected(clickedRowID);
				//alert("Here");
				//ShowDetails();
			}
			
							
/////==========================  End_CHanged   05 Apr 2004 =============================================================///////////////////////////////
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
			//grid.setAction("dblclick", ShowDetails); 
			//grid.setAction("selectionChanged", RowChangeFn); 

			//stripes 
			var alternate = function(){ return this.getProperty("row/order") % 2 ? "gainsboro" : "white";} 
			var row = new Active.Templates.Row; row.setStyle("background", alternate); 
			row.setEvent("onmouseover", "mouseover(this, 'active-row-highlight')"); 
			row.setEvent("onmouseout", "mouseout(this, 'active-row-highlight')"); 
			grid.setTemplate("row", row); 
			var column = new Active.Templates.Text; 
			column.setStyle("border-right", "1px solid white");  
			grid.setTemplate("column", column); 

			// write grid html to the page 
			document.write(grid);
		 
		</script> 
			
				<input type = 'hidden' name ='ID' id = 'ID'>
				<input type = 'hidden' name ='ActionPage' id = "ActionPage" value = "<%=ActionPage%>.asp">

<%'======================= Begin_Alter_Across_Entities =================================%>
				
				<input type = 'hidden' name ='FilterArgs' id="FilterArgs" value="ConsultantName:Consultant*0;ConsultantCountry:Country*0;ConsultantContact:Contact Name*0;ConsultantEmail:Email*0;ConsultantRates:Rates*0">
				<input type = 'hidden' name ='SortArgs' id="SortArgs" value="ConsultantName:Consultant;ConsultantCountry:Country;ConsultantContact:Contact Name;ConsultantEmail:Email;ConsultantRates:Rates">
				<input type = 'hidden' name ='SearchArgs' id="SearchArgs" value="ConsultantName:Consultant*0;ConsultantCountry:Country*0;ConsultantContact:Contact Name*0;ConsultantEmail:Email*0;ConsultantRates:Rates*0">
				<input type = 'hidden' name ='dialogLayout' id="dialogLayout" value="height:28em;width:35em">
				
<%'======================= End_Alter_Across_Entities =================================%>
				
				<input type = 'hidden' name ='SelectedSortArgs' id="SelectedSortArgs" value="<%= sortQryStr %>">
				<input type = 'hidden' name ='SelectedFilterArgs' id="SelectedFilterArgs" value="<%= filterStr %>">
				<input type = 'hidden' name ='SelectedSearchArgs' id="SelectedSearchArgs" value="<%= searchStr %>">
							

			<%			
			Set Rs = Nothing
			Set Conn = Nothing
			%>	
			</form> 
			

			
</body> 
</html>