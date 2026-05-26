<!--#include file="../libroutines.asp"-->
<%
'======================= Begin_Alter_Across_Entities =================================
		
		const UDLName = "KBroker"
		const DataSource = "PurchaseOrdersList"
		const DataEntity = "PurchaseOrder"
		const DataEntityPlural = "PurchaseOrders"
		const ActionFolder = "Operations"
'======================= End_Alter_Across_Entities =================================		
		
		Dim conn 
		Dim sqlStr
		Dim sqlStr1
		Dim sqlStr2		
		Dim rs		
		Dim rstotal
		Dim rsCurrent
		Dim rsUpdateLimit
		Dim bcolor
		Dim rowIDs
		Dim rowData
		Dim quote
		Dim sortQryStr
		Dim DatastreamTotal
		
		sortQryStr = Request.Form("SelectedSortArgs")
		filterStr = Request.Form("SelectedFilterArgs")
		searchStr = Request.Form("SelectedSearchArgs")
			
		Set conn = GetActiveConnection(UDLName)

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
		
        Set Conn = GetActiveConnection(UDLName)
			
		Set rstotal = Server.CreateObject("ADODB.Recordset")
		Set rsCurrent = Server.CreateObject("ADODB.Recordset")
		Set rsUpdateLimit = Server.CreateObject("ADODB.Recordset")
		
		rstotal.CursorLocation = adUseClient
		rsCurrent.CursorLocation = adUseClient
		
		todaydate=holiday(Date)
		
		Conn.Execute("Delete From TotalKnown")
				
        Set rsCurrent=Conn.execute(sqlStr)
        
        if not (rsCurrent.EOF and rsCurrent.BOF) then
			do while rsCurrent.EOF=false
			if (rsCurrent("Best")=true) then
				if(rsCurrent("Amount")<>0) then
				LimitAmount=rsCurrent("Amount")/(rsCurrent("Price")*1.105)
				else
				LimitAmount=rsCurrent("ordDetailQty")
				end if
				
			dim sqlstrLimit
			sqlstrLimit="update ordDetail set ordDetailQty=" & LimitAmount & ",ordDetailPrice=" & rsCurrent("Price")*1.105 & " where(ordDetail_DPA_=" & rsCurrent("ordDetail_DPA_") & ")"
			Conn.execute(sqlstrLimit)
			end if
			
			DatastreamTotal=GetDataStreamPrice(rsCurrent("SecurityCode"))
						
			if(DatastreamTotal="UnKnown") then
			sqlStr1="Select * From TotalKnown where(Client_DPA_=" & rsCurrent("Client_DPA_") & ")"
						
			Set rstotal=Conn.Execute(sqlStr1)
				if Not(rstotal.EOF and rstotal.BOF) then
				sqlStr2="Update TotalKnown Set Total=1 where(Client_DPA_=" & rsCurrent("Client_DPA_") & ")"
				else				
				sqlStr2="insert INTO TotalKnown(Client_DPA_,Total) VALUES(" & rsCurrent("Client_DPA_") & ",1)"
				end if
			Conn.Execute(sqlStr2)			
			end if
		
			rsCurrent.MoveNext 
			loop
        end if
        
        Set rstotal=nothing
        
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
					
					for colIndex = 14 to 18
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
						
					.active-column-0 {width: 60px;}
					.active-column-1 {width: 80px;}
					.active-column-2 {width: 40px; text-align: center}
					.active-column-3 {width: 60px;}
					.active-column-4 {width: 200px;}
					.active-column-5 {width: 60px;}
					.active-column-6 {width: 30px;}
					.active-column-7 {width: 80px;text-align: right;}
					.active-column-8 {width: 60px;text-align: right;}		
					.active-column-9 {width: 90px;text-align: right;}
					.active-column-10 {width: 90px;text-align: right;}		
					.active-column-11 {width: 90px;text-align: right;}	
					.active-column-12 {width: 90px;text-align: right;}						
					.active-column-13 {width: 120px;text-align: right;}						
					.active-column-14 {width: 90px;text-align: right;}	
					.active-column-15 {width: 90px;text-align: right;}						
					.active-column-16 {width: 120px;text-align: right;}											
					.active-column-17 {width: 150px;text-align: right;}	
					.active-column-18 {width: 90px;text-align: right;}	
					
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
	<form name='frm<%=DataSource%>' method='post' id='frmMain' action='Edit<%=DataEntity%>InPlace.asp'>
			<!-- grid data -->
		<% 'row data
		Dim rowCount
		Dim entryID
		Dim displayColor
		dim clientid1
		dim clientid2
		Dim Total
		
		clientid1="0"
		clientid2="0"
		Total=0		
		
		quote = chr(34) 
		intRecord = 1
		Do Until rs.EOF 
		clientid1=rs("Client_DPA_")
'======================= Begin_Alter_Across_Entities =================================
			if rs.Fields("BalanceQty") > 0 then
					displayColor = "DarkBlue"
			else
					displayColor = "Black"
			end if
			entryID = rs.Fields("OrdDetail_DPA_") '& "<->" & rs.Fields("Lot_DPA_")

			'row ID 
			rowData = rowData & quote & entryID & quote & " : " 			
			
			'row data 
			rowData = rowData & "["
			DatastreamPrice=GetDataStreamPrice(rs.Fields("SecurityCode")) 
			rowData = rowData & quote & ApplyDisplayColor(rs.Fields("Order_DPA_")) & quote & ","
			rowData = rowData & quote & ApplyDisplayColor(FormatDate(rs.Fields("OrderDate"))) & quote & ","
			rowData = rowData & quote & ApplyDisplayColor(IIf(CBool(rs.Fields("OrderTypeSale")) = True, "S", "P")) & quote & "," 
			rowData = rowData & quote & ApplyDisplayColor(rs.Fields("Client_DPA_")) & quote & "," 
			rowData = rowData & quote & ApplyDisplayColor(rs.Fields("OrdDetailClient")) & quote & ","
			if(trim(rs.Fields("ordDetailSecType"))="Fixed") then
			dim SecurityCodes
			SecurityCodes=rs.Fields("SecurityCode")
						
			SecurityCodes=split(SecurityCodes," ")			
			
			boundi=UBound(SecurityCodes)
			
			if(boundi=1) then
			SecurityCode=rs.Fields("SecurityCode")
			end if
			if(boundi=2) then
			SecurityCode=SecurityCodes(1) & " " & SecurityCodes(2) 			
			end if
			if(boundi=3) then
			SecurityCode=SecurityCodes(1) & " " & SecurityCodes(2) & " " & SecurityCodes(3)
			end if
			
			rowData = rowData & quote & SecurityCode & quote & ","
			'rowData = rowData & quote & endStrPos & quote & ","
			else
			rowData = rowData & quote & rs.Fields("SecurityCode") & quote & ","
			end if			
			rowData = rowData & quote & IIf(rs.Fields("Best") =True, "Y", "N") & quote & "," 
			rowData = rowData & quote & ApplyDisplayColor(FormatNum(rs.Fields("Amount"))) & quote & "," 
			
			if(rs.Fields("Best")=true) and (rs.Fields("Amount")<>0) then			
				if(DatastreamPrice="UnKnown") then
				DatastreamTotal=1
				BestQuantity=DatastreamPrice
				rowData = rowData & quote & ApplyDisplayColor(BestQuantity) & quote & "," 
				else													
				BestQuantity=rs.Fields("Amount")/(rs.Fields("Price")*1.105)
				rowData = rowData & quote & ApplyDisplayColor(FormatNumEx(BestQuantity,0)) & quote & "," 
				end if			
			else			
			rowData = rowData & quote & ApplyDisplayColor(FormatNumEx(rs.Fields("OrdDetailQty"),0)) & quote & "," 
			end if
					
			if(DatastreamPrice="UnKnown" and rs.Fields("Best")=true) then
			rowData = rowData & quote & ApplyDisplayColor(DatastreamPrice) & quote & ","
			DatastreamPrice="UnKnown"
			else			
			rowData = rowData & quote & ApplyDisplayColor(FormatNumEx(rs.Fields("BalanceQty"),0)) & quote & ","
			end if
			
			if(rs.Fields("Best")=true) then				
				'Response.Write(DatastreamPrice)										
				if(trim(DatastreamPrice)="UnKnown") then
				rowData = rowData & quote & ApplyDisplayColor(DatastreamPrice) & quote & ","  
				DatastreamPrice="UnKnown"
				else
				rowData = rowData & quote & ApplyDisplayColor(FormatNum(rs.Fields("Price")*1.105)) & quote & ","  
				end if
			else
			rowData = rowData & quote & ApplyDisplayColor(FormatNum(rs.Fields("OrdDetailPrice"))) & quote & ","  
			end if
			
			if(rs.Fields("Best")=true) then
				if(DatastreamPrice="UnKnown") then				
				rowData = rowData & quote & ApplyDisplayColor(DatastreamPrice) & quote & "," 
				DatastreamPrice="UnKnown"
				else										
				rowData = rowData & quote & ApplyDisplayColor(FormatNum(rs.Fields("BalanceQty")*rs.Fields("Price")*1.105)) & quote & "," 
				end if
			else
			rowData = rowData & quote & ApplyDisplayColor(FormatNum(rs.Fields("BalanceQty")*rs.Fields("OrdDetailPrice"))) & quote & "," 
			end if
			
			if(clientid1<>clientid2) then
			'sqlStr1="SELECT SUM(ISNULL(BalanceQty * CONVERT(numeric,(replace(IsNull(case(isNumeric(OrdDetailPrice)) when 1 then OrdDetailPrice else '0' end,0),',',''))), 0)) AS Total FROM dbo.LotList WHERE(RTRIM(OrdDetailType) LIKE '%Purchase%') and (Client_DPA_=" & rs.Fields("Client_DPA_") & ")"			

			sqlStr1="SELECT SUM(ISNULL(BalanceQty * CONVERT(numeric, OrdDetailPrice), 0)) AS Total FROM  LotList WHERE (RTRIM(OrdDetailType) LIKE '%Purchase%') AND (Client_DPA_ = " & rs.Fields("Client_DPA_") & ")"
			sqlStr2="SELECT SUM(ISNULL(Credit-Debit,0)) AS CurrentBal,CreditLimit From ClientStatement inner join Client on ClientStatement.Client_DPA_=Client.Client_DPA_" & _
			        " where(ClientStatement.Client_DPA_=" & rs("Client_DPA_") & ") Group By CreditLimit"
			
			'Response.Write(sqlStr1)
			'Response.End 
			
			set rstotal=Conn.Execute(sqlStr1)
			if not(rstotal.EOF and rstotal.BOF) then
			Total=rstotal.Fields("Total")			
			Total=Cdbl(Total)
			end if
			
			set rstotal=nothing
			'Calc Current Bal						
						
			sqlStr1="Select * from TotalKnown where(Client_DPA_=" & rs("Client_DPA_") & ")"
			Set rstotal=Conn.Execute(sqlStr1)
			if Not(rstotal.EOF and rstotal.BOF) then
			Total="UnKnown"					
			end if
			
			set rsCurrent=Conn.Execute(sqlStr2)
			if not(rsCurrent.EOF and rsCurrent.BOF) then
			Current=rsCurrent.Fields("CurrentBal")
			Current=Cdbl(Current)
			CreditLimit=rsCurrent.Fields("CreditLimit")			
			CreditLimit=Cdbl(CreditLimit)
			end if
			
			Set rsCurrent=nothing			
			end if
			
			if(Total="UnKnown") then							
			rowData = rowData & quote & ApplyDisplayColor(Total) & quote & ","
			Total="UnKnown"
			else
			rowData = rowData & quote & ApplyDisplayColor(FormatNum(Total)) & quote & ","
			end if
			
			rowData = rowData & quote & ApplyDisplayColor(FormatNum(Current)) & quote & ","  			
			
			if(Total="UnKnown") then							
			rowData = rowData & quote & ApplyDisplayColor(Total) & quote & ","
			Total="UnKnown"
			else
			rowData = rowData & quote & ApplyDisplayColor(FormatNum(Current-Total)) & quote & ","
			end if
					
			rowData = rowData & quote & ApplyDisplayColor(FormatNum(-CreditLimit)) & quote & ","  
			
			if(Total="UnKnown") then							
			rowData = rowData & quote & ApplyDisplayColor(Total) & quote & ","
			Total="UnKnown"
			else
			rowData = rowData & quote & ApplyDisplayColor(FormatNum(-CreditLimit-(Current-Total))) & quote & ","  
			end if
						
			rowData = rowData & quote & ApplyDisplayColor(FormatDateFull(todaydate)) & quote & ","
						
			if(trim(DatastreamPrice)="UnKnown") then			
			rowData = rowData & quote & DatastreamPrice & quote & ","
			else
			rowData = rowData & quote & ApplyDisplayColor(FormatNum(DatastreamPrice)) & quote & ","  
			end if
			rowData = rowData & "]" 
			'build the row IDs array 
			rowIDs = rowIDs & quote & entryID & quote 
			rowCount = rowCount + 1
			clientid2=clientid1
'======================= End_Alter_Across_Entities =================================

			rs.MoveNext 
			if not(rs.eof) then 
				If  intRecord >= intPageSize Then	Exit Do				
				intRecord = intRecord + 1	
				rowIDs = rowIDs & "," 
				rowData = rowData & "," 
			end if 
		Loop
set rstotal=nothing
'======================= Begin_Alter_Across_Entities =================================%> 
		<script>
			//column titles 
			var colCount = 19;
			var colNames = ["", "", "", "", "", "","", 
			"","","","","","","","","","","","","",""];
			
			var myColumns = ["Order No","Date", "Type", "Code", "Client", "Security","Best","Amount", 
			 "Limit","Balance","Price", "Cost","Total","Current Balance","Working Balance","Credit Limit","Excess","Yesterday","Data Stream Price"];
			
		
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
								if(colNames[i]=="TDate")
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
							if(colNames[i]=="TDate")
							{
								currentDate =  myData[rowIndex][i];
								myData[rowIndex][i] = "<INPUT TYPE='text' NAME='" + colNames[i] + "' ID='" + colNames[i] + "' VALUE='<%=FormatDate(Date)%>' onChange = 'EditInPlaceDataChanged();event.cancelBubble=true;' OnClick='event.cancelBubble=true;'>";
							}
							else
							{
								myData[rowIndex][i] = "<INPUT TYPE='text' NAME='" + colNames[i] + "' ID='" + colNames[i] + "' VALUE='" + myData[rowIndex][i] + "' onChange = 'EditInPlaceDataChanged();event.cancelBubble=true;' OnClick='event.cancelBubble=true;'>";
							}
						}						
					}
				}
				inPlaceEdit = true;
				prevRow = rowIndex;
				grid.refresh();
				//select the appropriate item
				var secList = document.frmMain.elements("cboBrokerInPlace");
				for (i=0; i < secList.options.length; i++) {
					if(secList.options(i).text == currentBrokerName)
					{
							secList.options(i).selected = true;
					}
				}
				
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
										if(colNames[i]=="TDate")
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
			
			var calTDate;
			function changeDateInterface(){
				calTDate = new ctlSpiffyCalendarBox('calTDate', 'frm<%=DataSource%>', 'TDate', 'cmdTDate','<%= FormatDate(Date) %>', 1); 
				calTDate.returnOutStringOnWrite(); 
				var parentDiv = document.all.item("TDate").parentNode;
				parentDiv.innerHTML = calTDate.writeControl();
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
				<input type = 'hidden' name ='ActionPage' id = "ActionPage" value = "PurchaseOrders.asp">

<%'======================= Begin_Alter_Across_Entities =================================%>
				
				<input type = 'hidden' name ='FilterArgs' id="FilterArgs" value="Order_DPA_:Order No*2;OrderDate:Date*1;OrderTypeSale:Type*0;Client_DPA_:Client Code*0;OrdDetailClient:Client*0;OrdDetailSecurity:Security*0;OrdDetailQty:Limit*2;BalanceQty:Balance*2;LotPrice:Price*2">
				<input type = 'hidden' name ='SortArgs' id="SortArgs" value="Order_DPA_:Order No;OrderDate:Date;OrderTypeSale:Type;Client_DPA_:Client Code;OrdDetailClient:Client;OrdDetailSecurity:Security;OrdDetailQty:Limit;BalanceQty:Balance;LotPrice:Price">
				<input type = 'hidden' name ='SearchArgs' id="SearchArgs" value="Order_DPA_:Order No*2;OrderDate:Date*1;OrderTypeSale:Type*0;Client_DPA_:Client Code*0;OrdDetailClient:Client*0;OrdDetailSecurity:Security*0;OrdDetailQty:Limit*2;BalanceQty:Balance*2;LotPrice:Price*2">
				<input type = 'hidden' name ='dialogLayout' id="dialogLayout" value="height:32em;width:46em">
				
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
		secList = "<select name = '" & listName & "' id = '" & listName & "' size='1' onChange = 'event.cancelBubble=true;' OnClick='event.cancelBubble=true;'>"
		secList = secList & "<option selected value = ''></option>"
		
        sqlStr = "SELECT * FROM [BrokerList] Order By BrokerCode"
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
  
  function GetDataStreamPrice(Sec)
		Dim secList
		Dim rs
		Dim sqlStr5
        sqlStr5 = "SELECT Price FROM dbo.datastream_SecurityPriceList " & _
				  "WHERE (SecurityCode='" & Sec & "')"
		
		'Response.Write(sqlStr5)
		'Response.End 
				 
        Set rs = conn.Execute(sqlStr5)
        If Not (rs.EOF AND rs.BOF) Then
        GetDataStreamPrice=rs("Price")
        else
        GetDataStreamPrice="UnKnown"
        End If	    
	    if(isNull(GetDataStreamPrice)) then
	    GetDataStreamPrice="UnKnown"
	    end if
  end function
 %>			
</body> 
</html>