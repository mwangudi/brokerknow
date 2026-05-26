<!--#include file="../libroutines.asp"-->
<%
'======================= Begin_Alter_Across_Entities =================================
		
		const UDLName = "KBroker"
		const DataSource = "ContractNewCertList" 'table or view in database
		const DataEntityList = "CertificateList" 'this page
		const DataEntity = "Certificate"
		const DataEntityPlural = "Certificates"
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
					.active-column-0 {width: 60px;}
					.active-column-1 {width: 100px;}
					.active-column-2 {width: 120px;}
					.active-column-3 {width: 50px;}
					.active-column-4 {width: 50px;}
					.active-column-5 {width: 70px;}
					.active-column-6 {width: 60px;}
					.active-column-7 {width: 70px; text-align: right;}	
					.active-column-8 {width: 60px; text-align: right;}	
					.active-column-9 {width: 70px;}	
					.active-column-10 {width: 100px;}		
					.active-column-11 {width: 1px;}	
					
					.active-grid-row,
					.active-grid-row.active-list-item,
					.active-scroll-left .active-list-item {height: 18px;}
				</style>
				<!--CALENDAR -->
				<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
				<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
										
</head>
	<body leftMargin=0 topMargin=0 marginheight="0" marginwidth="0"> 
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form name='frm<%=DataSource%>' method='post' id='frmMain' action='Edit<%=DataEntity%>InPlace.asp'>
			<!-- grid data -->
		<% 'row data
		Dim rowCount
		Dim entryID
		quote = chr(34) 
		intRecord = 1 
		Do Until rs.EOF 

'======================= Begin_Alter_Across_Entities =================================
			if rs.Fields("ContractNCertificate") = "" or isnull(rs.Fields("ContractNCertificate")) then
					displayColor = "DarkBlue"
			else
					displayColor = "Black"
			end if
			
			entryID = rs.Fields("Contract_DPA_")
			'row ID 
			rowData = rowData & quote & entryID & quote & " : " 
			
									
			'row data 
			rowData = rowData & "[" 
			rowData = rowData & quote & ApplyDisplayColor(rs.Fields("ContractNumber")) & quote & "," 
			rowData = rowData & quote & rs.Fields("ContractNCertificate") & quote & ","
			rowData = rowData & quote & FormatDate(rs.Fields("ContractNCDate")) & quote & "," 
			rowData = rowData & quote & ApplyDisplayColor(rs.Fields("BrokerCode")) & quote & "," 
			rowData = rowData & quote & ApplyDisplayColor(rs.Fields("SecurityCode")) & quote & ","
			rowData = rowData & quote & ApplyDisplayColor(rs.Fields("LotTDate")) & quote & ","
			rowData = rowData & quote & ApplyDisplayColor(rs.Fields("LotSlipNo")) & quote & "," 
			rowData = rowData & quote & ApplyDisplayColor(FormatNumCommasOnly(rs.Fields("LotQty"))) & quote & ","
			rowData = rowData & quote & ApplyDisplayColor(FormatNum(rs.Fields("LotPrice"))) & quote & ","
			rowData = rowData & quote & ApplyDisplayColor(rs.Fields("StatusDescription")) & quote & ","
			rowData = rowData & quote & ApplyDisplayColor(rs.Fields("OrdDetailClient")) & quote & ","
			rowData = rowData & quote & formatdate(rs.Fields("Settlementdate")) & quote 
			rowData = rowData & "]" 
			'build the row IDs array 
			rowIDs = rowIDs & quote & entryID & quote 
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
			var colNames = ["", "NewCert", "RDate", "","","", 
			"", "", "", "","",""];
			
			var myColumns = ["Contract", "New Certificate", "Date Received", "Broker","Security","Trade Date", 
			"Slip No", "Quantity", "Price", "Status","Client",""];
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
			
			var currentBrokerName = "";
			var currentDate = "";
			var RowEditFn = function(src)
			{
				var rowIndex = src.getProperty("row/index");
				var i;
				var inputType = "";
				
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
								if(colNames[i]=="RDate")
								{
									myData[prevRow][i] = currentDate;
								}
								else
								{
									myData[prevRow][i] = document.frmMain.elements(colNames[i]).value;
								}
							}
						}
						if(colNames[i]=="Broker")
						{	
							currentBrokerName =  myData[rowIndex][i];
							myData[rowIndex][i] = inPlaceList;
							
						}
						else
						{
							if(colNames[i]=="RDate")
							{
								currentDate =  myData[rowIndex][i];
								myData[rowIndex][i] = "<INPUT TYPE='text' NAME='" + colNames[i] + "' ID='" + colNames[i] + "' VALUE='"+ myData[rowIndex][11]+"' onChange = 'EditInPlaceDataChanged();event.cancelBubble=true;' OnClick='event.cancelBubble=true;'>";
							}
							else
							{
								myData[rowIndex][i] = "<INPUT TYPE='text' STYLE='WIDTH: 80px' NAME='" + colNames[i] + "' ID='" + colNames[i] + "' VALUE='" + myData[rowIndex][i] + "' onChange = 'EditInPlaceDataChanged();event.cancelBubble=true;' OnClick='event.cancelBubble=true;'>";
							}
						}						
					}
				}
				inPlaceEdit = true;
				prevRow = rowIndex;
				grid.refresh();
				//select the appropriate item
				/*var secList = document.frmMain.elements("cboBrokerInPlace");
				for (i=0; i < secList.options.length; i++) {
					if(secList.options(i).text == currentBrokerName)
					{
							secList.options(i).selected = true;
					}
				}*/
				
				//show the calendar
				changeDateInterface();
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
									if(colNames[i]=="Broker")
									{
										myData[prevRow][i] = currentBrokerName;
									}
									else
									{
										if(colNames[i]=="RDate")
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
/////==========================  End_CHanged   05 Apr 2004 =============================================================///////////////////////////////
			
			//get ready for in-place edit
			var inPlaceList = "<%=GetBrokerList("cboBrokerInPlace")%>"
			
			var calRDate;
			function changeDateInterface(){
				calRDate = new ctlSpiffyCalendarBox('calRDate', 'frm<%=DataSource%>', 'RDate', 'cmdRDate','<%= FormatDate(Date+7) %>', 1); 
				calRDate.returnOutStringOnWrite(); 
				var parentDiv = document.all.item("RDate").parentNode;
				parentDiv.innerHTML = calRDate.writeControl();
				parentDiv.style.zIndex = 10;
				parentDiv.childNodes(1).style.zIndex = 10	;
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
				<input type = 'hidden' name ='EditPage' id = "EditPage" value = "<%=ActionFolder%>/Edit<%=DataEntity%>.asp">
				<input type = 'hidden' name ='AddPage' id = "AddPage" value = "<%=ActionFolder%>/Add<%=DataEntity%>.asp"> 
				<input type = 'hidden' name ='DeletePage' id = "DeletePage" value = "Delete<%=DataEntity%>.asp">
				<input type = 'hidden' name ='ActionPage' id = "ActionPage" value = "<%=DataEntityList%>.asp">

<%'======================= Begin_Alter_Across_Entities =================================%>
				
				<input type = 'hidden' name ='FilterArgs' id="FilterArgs" value="ContractNumber:Contract*0;ContractNCertificate:New Certificate*0;ContractNCDate:Date Received*1;BrokerCode:Broker*0;LotSlipNo:Slip No*0;LotQty:Quantity*2;LotPrice:Price*2;StatusDescription:Status*0;OrdDetailClient:Client*0">
				<input type = 'hidden' name ='SortArgs' id="SortArgs" value="ContractNumber:Contract;ContractNCertificate:New Certificate;ContractNCDate:Date Received;BrokerCode:Broker;LotSlipNo:Slip No;LotQty:Quantity;LotPrice:Price;StatusDescription:Status;OrdDetailClient:Client">
				<input type = 'hidden' name ='SearchArgs' id="SearchArgs" value="ContractNumber:Contract*0;ContractNCertificate:New Certificate*0;ContractNCDate:Date Received*1;BrokerCode:Broker*0;LotSlipNo:Slip No*0;LotQty:Quantity*2;LotPrice:Price*2;StatusDescription:Status*0;OrdDetailClient:Client*0">			
				<input type = 'hidden' name ='dialogLayout' id="dialogLayout" value="height:25em;width:20em">
				
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
			
<%
       
  function GetBrokerList(listName)
		Dim secList
		Dim rs
		secList = "<select name = '" & listName & "' id = '" & listName & "' size='1' onChange = 'AddRowInProgress();event.cancelBubble=true;' OnClick='event.cancelBubble=true;'>"
		secList = secList & "<option selected value = ''></option>"
		
        sqlStr = "SELECT * FROM [BrokerList]"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF
                        secList = secList & "<option value = '" & rs.Fields("Broker_DPA_") & "'>" & rs.Fields("BrokerCode") & "</option>"
                        rs.MoveNext
                Loop
        End If
	    secList = secList & "</select>"
	    GetBrokerList = secList
  end function
  
 %>			
</body> 
</html>