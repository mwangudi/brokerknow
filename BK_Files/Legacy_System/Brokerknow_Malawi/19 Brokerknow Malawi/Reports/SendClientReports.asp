<!--#include file="../libroutines.asp"-->
<%

		Set conn = GetActiveConnection("KBroker")

		''DELETE FILES		       
		Dim FSO
		Set FSO = Server.CreateObject("Scripting.FileSystemObject")
		
		fpath = Server.MapPath("../") & "\Mailer\bin"
		
		Dim FL
		Set FL = FSO.GetFolder(fpath)

		If FL.Files.Count > 50 Then
			For each fle in FL.Files
				If Instr(1,fle,"htm") > 0 Or Instr(1,fle,"pdf") > 0 Or Instr(1,fle,"html") > 0 Then
					FSO.DeleteFile fle, True
				End If
			Next
		End If
        
        Set FL = Nothing
        Set FSO = Nothing
        
'======================= Begin_Alter_Across_Entities =================================
		
		const UDLName = "KBroker"
		const DataSource = "SendClientReports"
		const ActionPage = "SendClientReports"
		const DataEntity = "ClientReports"
		const DataEntityPlural = "Clients"
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
		
		SELECTstr = " SendClientReports.changedBy AS TheUser, SendClientReports.Timechanged AS TheTimeChanged, Client.Client_DPA_, Client.ClientEmail," & _
		" Client.ClientName, Client.ClientCDSNo, SendClientReports.ClientReportsID, SendClientReports.Client_DPA_, SendClientReports.ClientStatements," & _
		" SendClientReports.ClientContracts, SendClientReports.ClientContractsCompounded, SendClientReports.HoldingsValuation," & _
		" GenericSetting.GenericSettingDescription"
                      
		sqlStrOrig = "SELECT  "& SELECTstr &"  FROM [" & DataSource & "] INNER JOIN Client ON Client.Client_DPA_ = [" & DataSource & "].Client_DPA_ INNER JOIN GenericSetting ON SendClientReports.GenericSetting_DPA_2 = GenericSetting.GenericSetting_DPA_"
		
		sqlStr = "SELECT  "& SELECTstr &"  FROM [" & DataSource & "] INNER JOIN Client ON Client.Client_DPA_ = [" & DataSource & "].Client_DPA_ INNER JOIN GenericSetting ON SendClientReports.GenericSetting_DPA_2 = GenericSetting.GenericSetting_DPA_"
		
		If filterStr <> "" Then	
			sqlStr = "SELECT  "& SELECTstr &"  FROM [" & DataSource & "] INNER JOIN Client ON Client.Client_DPA_ = [" & DataSource & "].Client_DPA_ INNER JOIN GenericSetting ON SendClientReports.GenericSetting_DPA_2 = GenericSetting.GenericSetting_DPA_ WHERE " & filterStr
		End If
		
		If searchStr <> "" Then
			If InStr(1, sqlStr, " WHERE ") > 0 Then
				sqlStr = sqlStr & " AND " & searchStr
			Else
				sqlStr = "SELECT  "& SELECTstr &"  FROM [" & DataSource & "] INNER JOIN Client ON Client.Client_DPA_ = [" & DataSource & "].Client_DPA_ INNER JOIN GenericSetting ON SendClientReports.GenericSetting_DPA_2 = GenericSetting.GenericSetting_DPA_ WHERE " & searchStr
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
		<title>Send Client Reports</title>
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
				    .active-column-0 {width: 40px;}
					.active-column-1 {width: 200px;}
					.active-column-2 {width: 120px;}
					.active-column-3 {width: 120px;}
					.active-column-4 {width: 120px;}
					.active-column-5 {width: 120px;}
					.active-column-6 {width: 120px;}
					.active-column-7 {width: 120px;}	
					.active-column-8 {width: 120px;}
					.active-column-9 {width: 130px;}
					.active-column-10 {width: 130px;}
					
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
			'Format date if specified
			
			if Len(rs.Fields("TheTimeChanged")) > 0 then 
				TimeChanged = FormatDate(rs.Fields("TheTimeChanged")) & " " & FormatDateTime(rs.Fields("TheTimeChanged"),vbLongTime)
			else
				TimeChanged = ""
			end if
			
			if rs.Fields("ClientStatements") = 1 then 
				theClientStatement = "checked"
			else
				theClientStatement = ""
			end if
			StrClientStatement = "<input type=checkbox disabled "& theClientStatement &" name=ClientStatement"& rs.Fields("Client_DPA_") &">"
			
			if rs.Fields("ClientContracts") = 1 then 
				theClientContract = "checked"
			else
				theClientContract = ""
			end if
			StrClientContract = "<input type=checkbox disabled "& theClientContract &" name=ClientContract"& rs.Fields("Client_DPA_") &">"
			
			if rs.Fields("ClientContractsCompounded") = 1 then 
				theClientContractCompounded = "checked"
			else
				theClientContractCompounded = ""
			end if
			StrClientContractCompounded = "<input type=checkbox disabled "& theClientContractCompounded &" name=ClientContractCompounded"& rs.Fields("Client_DPA_") &">"
			
			if rs.Fields("HoldingsValuation") = 1 then 
				theHoldingsValuation = "checked"
			else
				theHoldingsValuation = ""
			end if
			StrHoldingsValuation = "<input type=checkbox disabled "& theHoldingsValuation &" name=HoldingsValuation"& rs.Fields("Client_DPA_") &">"
			
			'row ID 
			rowData = rowData & quote & rs.Fields("ClientReportsID") & quote & " : " 
			
			'row data 
			rowData = rowData & "[" 
			rowData = rowData & quote & rs.Fields("Client_DPA_") & quote & "," 
			rowData = rowData & quote & rs.Fields("ClientName") & quote & ","
			rowData = rowData & quote & rs.Fields("ClientCDSNo") & quote & "," 
			rowData = rowData & quote & rs.Fields("ClientEmail") & quote & "," 
			rowData = rowData & quote & rs.Fields("GenericSettingDescription") & quote & "," 
			rowData = rowData & quote & StrClientStatement & quote & "," 
			rowData = rowData & quote & StrClientContract & quote & ","
			rowData = rowData & quote & StrClientContractCompounded & quote & "," 
			rowData = rowData & quote & StrHoldingsValuation & quote & ","
			rowData = rowData & quote & rs.Fields("theUser") & quote  & ","
			rowData = rowData & quote & TimeChanged & quote 
			
			rowData = rowData & "]" 
			'build the row IDs array 
			rowIDs = rowIDs & quote & rs.Fields("ClientReportsID") & quote 
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
			var colNames = ["", "", "", "", "", "","", "", "", "", ""];
			
			var myColumns = ["Code", "Name", "CDS No", "Email", "Frequency", "Client Statements", "Client Contracts", "Client Contracts Compounded", "Holdings Valuation", "Changed By", "Time Changed"];
			
						
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
				<input type = 'hidden' name ='delAction' id = 'delAction' value="Execute">
				<input type = 'hidden' name ='EditPage' id = "EditPage" value = "<%=ActionFolder%>/Edit<%=DataEntity%>.asp">
				<input type = 'hidden' name ='AddPage' id = "AddPage" value = "<%=ActionFolder%>/Add<%=DataEntity%>.asp"> 
				<input type = 'hidden' name ='DeletePage' id = "DeletePage" value = "<%=ActionFolder%>/Delete<%=DataEntity%>.asp?del=1">
				<input type = 'hidden' name ='ActionPage' id = "ActionPage" value = "<%=ActionPage%>.asp">

<%'======================= Begin_Alter_Across_Entities =================================%>
				
				<input type = 'hidden' name ='FilterArgs' id="FilterArgs" value="Client_DPA_: Code*2;ClientName: Name*0;ChangedBy: Last Modified By*0;TimeChanged: Time Modified*1">
				<input type = 'hidden' name ='SortArgs' id="SortArgs" value="Client_DPA_: Code;ClientName: Name;ChangedBy: Last Modified By;TimeChanged: Time Modified">
				<input type = 'hidden' name ='SearchArgs' id="SearchArgs" value="Client_DPA_: Code*2;ClientName: Name*0;ChangedBy: Last Modified By*0;TimeChanged: Time Modified*1">
				<input type = 'hidden' name ='dialogLayout' id="dialogLayout" value="height:20em;width:40em">
				
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