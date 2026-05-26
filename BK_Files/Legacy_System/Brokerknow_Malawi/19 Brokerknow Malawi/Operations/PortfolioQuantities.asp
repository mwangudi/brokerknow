<!--#include file="../libroutines.asp"-->
<%
'======================= Begin_Alter_Across_Entities =================================
		
		const UDLName = "KBroker"
		const DataSource = "PortfolioQuantities"
		const DataEntity = "Quantity"
		const DataEntityPlural = "Quantities"
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
		Dim UserId
		
		sortQryStr = Request.Form("SelectedSortArgs")
		filterStr = Request.Form("SelectedFilterArgs")
		searchStr = Request.Form("SelectedSearchArgs")
		
		UserId=Session("Userid")
				
		sqlStrOrig = "SELECT * From LatestPortfolios where(userid=" & UserId & ")"
				
		sqlStr = "SELECT * From LatestPortfolios where(userid=" & UserId & ")"
		
		sqlStr1 = "SELECT * From [" & DataSource & "]"		
		
		If filterStr <> "" Then
			sqlStr = "SELECT Distinct LotList.*,ordDetail.Amount,Security.SecurityMktPrice FROM LotList inner join ordDetail on LotList.ordDetail_DPA_=ordDetail.ordDetail_DPA_ " & _
					 " inner join Security on LotList.Security_DPA_=Security.Security_DPA_ WHERE LotList." & filterStr & " order by LotList.Order_DPA_ Desc,ContractNumber Desc	"
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
				sqlStr = sqlStr & " , LotList." & sortQryStr	
				'hoping that the last clause in the sql will always be an order by
			Else
				sqlStr = sqlStr & " ORDER BY " & sortQryStr	
			End	If
		End If 
				
        Set conn = GetActiveConnection(UDLName)		        
		
		Set Rs = Server.CreateObject("ADODB.Recordset")
        Rs.CursorLocation = adUseClient  
		
		Set rstotal = Server.CreateObject("ADODB.Recordset")
        rstotal.CursorLocation = adUseClient  
		
		'Set rsUpdateLimit = Server.CreateObject("ADODB.Recordset")
		Set rsCurrent = Server.CreateObject("ADODB.Recordset")
		Set rsUpdateLimit = Server.CreateObject("ADODB.Recordset")
		
		rsUpdateLimit.CursorLocation=adUseClient
		rsCurrent.CursorLocation = adUseClient
		
        Dim Security1
		Dim Security2
		Dim Balance
		
		Dim rowCount
		Dim entryID
		Dim displayColor
				
		intRecord = 1
		
		clientdpa=0
		clientname=""
		tradedate=""
		slipno=""
		contract=""
		securitydpa=0
		security=""
		stype=""
		quantity=0
		price=0
		sbalance=0
		saverageprice=0
		sentryquantity=0
		sentryprice=0
		valuedate=""
		cash=0
		marketprice=0
		bookvalue=0
		currentvalue=0
		pl=0
		pl_=0
		securitybookvalue=0		
		sclientbookvalue=0
		sclientcurrentvalue=0
		stotalportfolio=0
		cost=0
		sales=0
		sprofit=0
		profit_=0
		slastprice=0
		slastquantity=0
		NetAmount=0
		slastpl=0
		slastpl_=0
		clientpl=0
		clientpl_=0
		
		Security1=0
		Security2=0
		Client1=0
		Client2=0
		
		Balance =0
		Balance1=0
		PreviousPrice=0
		AveragePrice=0
		PreviousBalance=0
		PreviousBookValue=0
		CurrentBookValue=0
		ClientBookValue=0
		EntryQuantity=0
		EntryPrice=0
		SecurityCurrentValue=0
		ClientCurrentValue=0
		TotalPortfolio=0
		LastQuantity=0
		LastPrice=0
		Profit=0
		CurrentBookValue=""
		
		i=0
		
		first=0
		j=0
		
		'Response.Write(sqlStr1)
		'Response.End 
		Conn.execute("Delete from PortfoliosQuantities where(userid=" & UserId & ")")
		
		Set Rs=Conn.execute(sqlStr1)
		
		if(Rs.EOF and Rs.BOF) then
		Response.End 
		end if
				
		rs.MoveFirst 
		Do Until rs.EOF 
		i=i+1
			entryID = rs.Fields("OrdDetail_DPA_") & "<->" & rs.Fields("Lot_DPA_")
			Security1=rs.Fields("Security_DPA_")			
			Client1=rs.Fields("Client_DPA_")			
			'rowData = rowData & quote & entryID & quote & " : " 					

			'rowData = rowData & "[" 
			sqlStr2="SELECT SUM(LevyAmount) AS Levies, LotGrossAmount FROM dbo.SaleContracts " & _
                    "WHERE     (SystemMaintained <> 8) AND (SystemMaintained <> 12) GROUP BY Lot_DPA_, LotGrossAmount HAVING (Lot_DPA_ = " & rs.Fields("Lot_DPA_") & ")"			

			set rstotal=conn.execute(sqlStr2)
			if not (rstotal.EOF and rstotal.BOF) then
				If(CBool(rs.Fields("OrderTypeSale"))= True)  then
				NetAmount=rs("LotGrossAmount")-rstotal("Levies")
				else
				NetAmount=rs("LotGrossAmount") + rstotal("Levies")
				end if
			else
			NetAmount=rs("LotGrossAmount")
			end if
			quantity=rs.Fields("LotQty")
			
			if(quantity=0) then
			price
			else
			price=NetAmount/quantity
			end if
			
			clientdpa=rs.Fields("Client_DPA_")			
			clientname=rs.Fields("OrdDetailClient")			 			
			tradedate=rs.Fields("LotTDate")
			slipno=rs.Fields("LotSlipNo")
			contract=rs.Fields("ContractNumber")
			securitydpa=rs.Fields("Security_DPA_")
			security=rs.Fields("SecurityCode")
			If(CBool(rs.Fields("OrderTypeSale"))= True)  then
			stype="S"
			else
			stype="P"
			end if 
			
			'price=rs.Fields("LotPrice")
			
			if(Client1<>Client2) then
			sqlStr2="SELECT SUM(ISNULL(Credit-Debit,0)) AS CurrentBal,CreditLimit From ClientStatement inner join Client on ClientStatement.Client_DPA_=Client.Client_DPA_" & _
			        " where(ClientStatement.Client_DPA_=" & rs("Client_DPA_") & ") Group By CreditLimit"
			
			set rsCurrent=Conn.Execute(sqlStr2)
				if not(rsCurrent.EOF and rsCurrent.BOF) then
				Current=rsCurrent.Fields("CurrentBal")
				Current=Cdbl(Current)
				CreditLimit=rsCurrent.Fields("CreditLimit")			
				CreditLimit=Cdbl(CreditLimit)
				end if
			end if
			
			'CurrentBookVlaue=PreviousBookValue
			if(Security1<>Security2) or (Client1<>Client2) then
			Balance=0
			EntryPrice=0
			EntryQuantity=0
			
			PreviousPrice=0			
			first=1			
			end if
			If(CBool(rs.Fields("OrderTypeSale"))= True)  then
			Balance=Balance-rs.Fields("LotQty")
			AveragePrice=PreviousPrice
			else
			Balance=Balance+rs.Fields("LotQty")
			
				if(cint(first)<>1) then
				AveragePrice=((PreviousBalance*PreviousPrice)+(rs.Fields("LotQty")*price))/(PreviousBalance+rs.Fields("LotQty"))
				else
				AveragePrice=price
				first=0
				end if
			
			end if
			if(Balance>0) then						
			sbalance=Balance
			else
			Balance=0
			AveragePrice=0
			sbalance=0
			end if
			PreviousBalance=Balance
			
			saverageprice=AveragePrice
			sentryquantity=EntryQuantity
			sentryprice=EntryPrice
			valuedate=Date
			cash=Current
			marketprice=rs.Fields("SecurityMktPrice")
			bookvalue=AveragePrice*Balance
			currentvalue=rs.Fields("SecurityMktPrice")*Balance
			pl=(rs.Fields("SecurityMktPrice")*Balance)-(AveragePrice*Balance)
			if (AveragePrice*Balance)<>0 then
			pl_=(((rs.Fields("SecurityMktPrice")*Balance)-(AveragePrice*Balance))/(AveragePrice*Balance))*100
			else
			pl_=0
			end if
						
			'Response.Write()
			if(Cint(Security1)<>Cint(Security2)) then
				if(Cint(j)=1) then
				securitybookvalue=PreviousBookValue
				SecurityCurrentValue=SecurityCurrentValue
				slastquantity=LastQuantity
				slastprice=LastPrice
				slastpl=LastPL
				slastpl_=LastPL_
				
				sqlStr1="Update PortfoliosQuantities set Securitybookvalue=" & securitybookvalue & ",SecurityCurrentValue=" & SecurityCurrentValue & _
				        ",LastQuantity=" & slastquantity & ",LastPrice=" & slastprice & ",LastPL=" & slastpl & ",LastPL_=" & slastpl_ & _
				        " where(Security_DPA_=" & Security2 & " and Client_DPA_=" & Client2 & ")"
				
				conn.execute(sqlStr1)
								
				ClientBookValue=ClientBookValue+PreviousBookValue
				ClientCurrentValue=ClientCurrentValue+SecurityCurrentValue								
				else
				end if
			else
			end if
			
			if(client1<>Client2 or i>=rs.RecordCount) then
				if(Cint(j)=1) then
				sclientbookvalue=ClientBookValue
				sclientcurrentvalue=ClientCurrentValue
				TotalPortfolio=ClientCurrentValue+Cash
				
				clientpl=ClientCurrentValue-ClientBookValue
				
				if(ClientBookValue<>0) then
				clientpl_=(clientpl/ClientBookValue)*100
				else
				clientpl_=0
				end if
				
				sqlStr1="Update PortfoliosQuantities set ClientBookValue=" & sclientbookvalue & ",ClientCurrentValue=" & sclientcurrentvalue & _
				        ",ClientPL=" & clientpl & ",ClientPL_=" & clientpl_ & ",TotalPortfolio=" & TotalPortfolio & " where(Client_DPA_=" & Client2 & ")"
				
				conn.execute(sqlStr1)
				
				ClientBookValue=0
				else
				j=1
				end if
			else
			end if
						
			'rowData = rowData & quote & ApplyDisplayColor(FormatNum(TotalPortfolio)) & quote & ","
			
			If(CBool(rs.Fields("OrderTypeSale"))= True)  then
			cost=EntryPrice*rs.Fields("LotQty")
			sales=rs.Fields("LotPrice")*rs.Fields("LotQty")
			
			sprofit=sales-cost
			
				if(cost<>0) then
				profit_=(sprofit/cost)*100
				else
				profit_=0				
				end if			
			else
			cost=0
			sales=0
			sprofit=0
			profit_=0
			end if
						
			sqlStr2="insert into PortfoliosQuantities(Client_DPA_,TradeDate,SlipNo,ContractNumber," & _
			        "Security_DPA_,SecurityCode,Type,Quantity,Price,Balance,AveragePrice,EntryQuantity,EntryPrice," & _
			        "ValueDate,Cash,MarketPrice,BookValue,CurrentValue,PL,PL_," & _
			        "Cost,Sales,Profit,Profit_,entryid,userid,NetAmount)" & _
			        "Values(" & clientdpa & ",'" & tradedate & "','" & slipno & "','" & Contract & "'" & _
			        "," & securitydpa & ",'" & security & "','" & stype & "'," & quantity & "," & price & "," & sbalance & "," & saverageprice & "," & sentryquantity & "," & sentryprice & _
			        ",'" & date & "'," & cash & "," & marketprice & "," & bookvalue & ", " & currentvalue & "," & pl & "," & pl_ & _
			        "," & cost & "," & sales & "," & sprofit & "," & profit_ & ",'" & entryid & "'," & UserId & "," & NetAmount & ")"
			
			Conn.execute(sqlStr2)
'======================= End_Alter_Across_Entities =================================
			
			PreviousBookValue=AveragePrice*Balance
			SecurityCurrentValue=CurrentValue
			EntryPrice=AveragePrice
			EntryQuantity=Balance
			PreviousPrice=AveragePrice			
			LastQuantity=Balance
			LastPrice=AveragePrice			
			Security2=Security1
			Client2=Client1
			LastPL=PL
			lastPL_=PL_			
			rs.MoveNext 
						
		Loop
		
		set rs=nothing        
        
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
				<style> 					
					.active-controls-grid {height: 100%; font: menu;}
					.active-row-highlight .active-row-cell {background-color: skyblue}
					.active-selection-true, .active-selection-true .active-row-cell {
						color: blue!important;
						background-color: bisque!important;
						}
						
						<%=styleStr%>
						
					.active-column-0 {width: 30px;}
					.active-column-1 {width: 100px;}
					.active-column-2 {width: 80px;}
					.active-column-3 {width: 60px;}
					.active-column-4 {width: 70px;}
					.active-column-5 {width: 60px;}
					.active-column-6 {width: 30px; text-align: center}
					.active-column-7 {width: 80px;text-align: center;}
					.active-column-8 {width: 80px;text-align: right;}	
					.active-column-9 {width: 80px;text-align: right;}	
					.active-column-10 {width: 80px;text-align: right;}	
					.active-column-11 {width: 80px;text-align: right;}	
					.active-column-12 {width: 80px;text-align: right;}
					.active-column-13 {width: 100px;}		
					.active-column-14 {width: 100px;text-align: right;}	
					.active-column-15 {width: 100px;text-align: right;}	
					.active-column-16 {width: 100px;text-align: right;}	
					.active-column-17 {width: 100px;text-align: right;}	
					.active-column-18 {width: 100px;text-align: right;}	
					.active-column-19 {width: 100px;text-align: right;}	
					.active-column-20 {width: 100px;text-align: right;}	
					.active-column-21 {width: 150px;text-align: right;}	
					.active-column-22 {width: 100px;text-align: right;}	
					.active-column-23 {width: 100px;text-align: right;}	
					.active-column-24 {width: 150px;text-align: right;}	
					.active-column-24 {width: 150px;text-align: right;}	
					.active-column-24 {width: 100px;text-align: right;}	
					.active-column-24 {width: 100px;text-align: right;}	
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
		'Dim rowCount
		quote = chr(34) 
		intRecord = 1
		Do Until rs.EOF 

'======================= Begin_Alter_Across_Entities =================================
			'row ID 
			rowData = rowData & quote & rs.Fields("Portfolio_DPA_") & quote & " : " 
			
			'row data 
			rowData = rowData & "[" 
			rowData = rowData & quote & rs.Fields("Client_DPA_") & quote & "," 
			rowData = rowData & quote & rs.Fields("ClientName") & quote & ","			
			rowData = rowData & quote & FormatDate(rs.Fields("TradeDate")) & quote & ","			
			rowData = rowData & quote & rs.Fields("SlipNo") & quote & "," 
			rowData = rowData & quote & rs.Fields("ContractNumber") & quote & ","  
			rowData = rowData & quote & rs.Fields("SecurityCode") & quote & "," 
			rowData = rowData & quote & rs.Fields("Type") & quote & "," 
			rowData = rowData & quote & FormatNumCommasOnly(rs.Fields("Quantity")) & quote & ","			 						
			rowData = rowData & quote & FormatNum(rs.Fields("Price")) & quote & ","
			rowData = rowData & quote & FormatNum(rs.Fields("Balance")) & quote & ","
			rowData = rowData & quote & FormatNum(rs.Fields("AveragePrice")) & quote & ","
			rowData = rowData & quote & FormatNumEx(rs.Fields("EntryQuantity"),0) & quote & ","
			rowData = rowData & quote & FormatNum(rs.Fields("EntryPrice")) & quote & ","
			rowData = rowData & quote & FormatDate(rs.Fields("ValueDate")) & quote & ","			
			rowData = rowData & quote & FormatNum(rs.Fields("Cash")) & quote & ","
			rowData = rowData & quote & FormatNum(rs.Fields("MarketPrice")) & quote & ","
			rowData = rowData & quote & FormatNum(rs.Fields("BookValue")) & quote & ","
			rowData = rowData & quote & FormatNum(rs.Fields("CurrentValue")) & quote & ","
			rowData = rowData & quote & FormatNum(rs.Fields("PL")) & quote & ","
			rowData = rowData & quote & FormatNum(rs.Fields("PL_")) & quote & ","
			rowData = rowData & quote & FormatNum(rs.Fields("SecurityBookValue")) & quote & ","
			rowData = rowData & quote & FormatNum(rs.Fields("SecurityCurrentValue")) & quote & ","
			rowData = rowData & quote & FormatNum(rs.Fields("LastQuantity")) & quote & ","
			rowData = rowData & quote & FormatNum(rs.Fields("LastPrice")) & quote & ","
			rowData = rowData & quote & FormatNum(rs.Fields("ClientBookValue")) & quote & ","
			rowData = rowData & quote & FormatNum(rs.Fields("ClientCurrentValue")) & quote & ","
			rowData = rowData & quote & FormatNum(rs.Fields("TotalPortfolio")) & quote & ","
			rowData = rowData & quote & FormatNum(rs.Fields("Cost")) & quote & ","
			rowData = rowData & quote & FormatNum(rs.Fields("Sales")) & quote & ","
			rowData = rowData & quote & FormatNum(rs.Fields("Profit")) & quote & ","
			rowData = rowData & quote & FormatNum(rs.Fields("Profit_")) & quote & ","
			rowData = rowData & quote & FormatNum(rs.Fields("Netamount")) & quote & ","
			rowData = rowData & quote & FormatNum(rs.Fields("LastPL")) & quote & ","
			rowData = rowData & quote & FormatNum(rs.Fields("LastPL_")) & quote & ","
			rowData = rowData & quote & FormatNum(rs.Fields("ClientPL")) & quote & ","
			rowData = rowData & quote & FormatNum(rs.Fields("ClientPL_")) & quote & ","
			
			rowData = rowData & "]" 
			'build the row IDs array 
			rowIDs = rowIDs & quote & rs.Fields("Portfolio_DPA_") & quote 
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
		'Response.Write(rowdata)
		'Response.End 
		
'======================= Begin_Alter_Across_Entities =================================%> 
		<script>
			//column titles 
			var colCount = 36;
			var colNames = ["", "", "", "", "", "", "", "", "", "", "", "", "",
			                "","","","","","","","","","","","","","","","","","","","","","",""];
			
			var myColumns = ["Code","Client","Trade Date","Slip No", "Contract", "Security", "Type", "Quantity","Price",
			                  "Balance","Average Price","Entry Quantity","Entry Price","Value Date","Cash",
			                  "Market Price","Book Value","Current Value","P/L","% P/L",
			                  "Security Book Value","Security Current Value","Last Quantity","Last Price",
			                  "Client Book Value","Client Current Value","Total Portfolio","Cost","Sales","Profit","%Profit","NetAmount","Last P/L","Last P/L %","Client P/L","Client P/L%"];			
		
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
						myData[prevRow][i] = document.frmMain.elements(colNames[i]).value;
								
						}
						myData[rowIndex][i] = "<INPUT TYPE='text' NAME='" + colNames[i] + "' ID='" + colNames[i] + "' VALUE='" + myData[rowIndex][i] + "' onChange = 'EditInPlaceDataChanged();event.cancelBubble=true;' OnClick='event.cancelBubble=true;'>";
											
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
/////==========================  End_CHanged   05 Apr 2004 =============================================================///////////////////////////////
			
			//get ready for in-place edit
			
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
				<input type = 'hidden' name ='ActionPage' id = "ActionPage" value = "<%=DataSource%>.asp">

<%'======================= Begin_Alter_Across_Entities =================================%>
				
				<input type = 'hidden' name ='FilterArgs' id="FilterArgs" value="Order_DPA_:Order No*2;OrderDate:Date*1;OrderTypeSale:Type*0;Client_DPA_:Client Code*0;OrdDetailClient:Client*0;OrdDetailSecurity:Security*0;OrdDetailQty:Limit*2;BalanceQty:Balance*2;LotSlipNo:Slip No*0;LotQty:Quantity*2;LotPrice:Price*2;BrokerCode:Broker*0;LotTDate:Trade Date*1;ContractNumber:Contract*0;StatusDescription:Status*0">
				<input type = 'hidden' name ='SortArgs' id="SortArgs" value="Order_DPA_:Order No;OrderDate:Date;OrderTypeSale:Type;Client_DPA_:Client Code;OrdDetailClient:Client;OrdDetailSecurity:Security;OrdDetailQty:Limit;BalanceQty:Balance;LotSlipNo:Slip No;LotQty:Quantity;LotPrice:Price;BrokerCode:Broker;LotTDate:Trade Date;ContractNumber:Contract;StatusDescription:Status">
				<input type = 'hidden' name ='SearchArgs' id="SearchArgs" value="Order_DPA_:Order No*2;OrderDate:Date*1;OrderTypeSale:Type*0;Client_DPA_:Client Code*0;OrdDetailClient:Client*0;OrdDetailSecurity:Security*0;OrdDetailQty:Limit*2;BalanceQty:Balance*2;LotSlipNo:Slip No*0;LotQty:Quantity*2;LotPrice:Price*2;BrokerCode:Broker*0;LotTDate:Trade Date*1;ContractNumber:Contract*0;StatusDescription:Status*0">
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
			
</body> 
</html>