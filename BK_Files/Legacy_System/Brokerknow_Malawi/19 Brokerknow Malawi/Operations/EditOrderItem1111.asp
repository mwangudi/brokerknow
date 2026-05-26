<!--#include file="../libroutines.asp"-->
<%
	const UDLName = "KBroker"
	const DataSource = "EditOrder"
	const DataEntity = "Order"
	const DataEntityPlural = "Orders"
	const ActionFolder = "Operations"
	
	const LinkedIndependent = 1
    const LinkedDependent = 2
	
   Dim action
   Dim conn 
   Dim sqlStr
   Dim rs
   Dim ID
   Dim rsEdit   
   dim validdate	
   Dim UserId
   Dim sqlStr1
   
   UserId=Session("UserID")
   Set conn = GetActiveConnection("KBroker")	
   action = ucase(Request.Form("action"))
   from = Request.QueryString("action")
	
	ItemID = Request.Form("ItemID")

	ID = Request("ID")
	ID1=Request("ID")

		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for editing"
                		
                </script>
                <% response.end
        End If

        if action <> "" then
				If Trim(ItemID) = "" Then%>
						<script language = 'vbscript'>
                				ShowMessage "No item specified"
                				
						</script>
						<%response.end
				End If
		end if

	select case action 
		case "EXECUTE_DETAIL"
   			Dim security
			Dim qty
			Dim price
			Dim validity
			Dim bond
			Dim Amount 
			Dim Best
			Dim chkprice
			Dim Limit
						
			 
			if(Ucase(Trim(Request.Form("orderType")))="SALE") then
			orderType1 = 2
			else
			orderType1 = 1
			end  if

	        if itemID = "-1" then
					security = Request.Form("cboSecurity")
					qty = Request.Form("txtQty")
					price = Request.Form("txtPrice")
					if(Cint(Request.Form("txtCalendar1"))=1) then
					validity = Request.Form("txtValidity")
					else
					validity = null
					end if
					
					bond = Request.Form("cboBond")
					Amount=Request.Form("txtAmount")
					chkprice=Request.Form("cboprice")
					best=Request.Form("txtbest")
					Limit = Request.Form("txtLimit")

			else
					security = Request.Form("cboSecurityInPlace")
					qty = Request.Form("txtQtyInPlace")
					price = Request.Form("txtPriceInPlace")
					if(Cint(Request.Form("txtCalendar1"))=1) then
					validity = Request.Form("txtCalendar")
					else
					validity = null
					end if
					bond = Request.Form("cboBondInPlace")
					Amount=Request.Form("txtAmountInPlace")
					best=Request.Form("txtbestinplace")
					chkprice=Request.Form("cboprices")
					Limit = Request.Form("txtLimitInPlace")

			end if			
			
			if(Ucase(trim(price))="BEST") then
				best=1
			end if
                  
			sectype1  = Request.Form("sectype")
			
			'validate Security 
			If (Trim(Security) = "" AND sectype1 = 2) Then%>
                	<script language = 'vbscript'>
                			ShowMessage "Please specify the Security."
                			
                	</script>
                	<% response.end
			End If

            'Validate Bond
			if (trim(bond) = "" AND sectype1 = 1) then
			 %>
                	<script language = 'vbscript'>
                			ShowMessage "Please specify the Bond."
                			
                	</script>
                	<% response.end

			end if

            'validate Order Limit	
			If (Trim(Limit) = "" and Cint(best)= 1 and Cint(orderType1)=1) Then%>
                	<script language = 'vbscript'>
                			ShowMessage "Please specify the Limit. "
                	</script>
                	<% response.end
			End If

             'validate Estimated Quantity			
			If Trim(qty)= "" and ((Cint(best)=1 and Cint(orderType1)=2) OR Cint(best)=0) Then%>
                	<script language = 'vbscript'>
                			ShowMessage "Please specify the Quantity "
                			
                	</script>
                	<% response.end
			End If
            
			'ensure Order Detail Estimated Quantity is numeric
			If (Qty <> "") And (Not IsNumeric(Qty)) Then%>
                	<script language = 'vbscript'>
					ShowMessage "Order Detail Estimated Quantity must be numeric"					
                	</script>
                	<% response.end
			End If

			'validate Estimated Price
			If Trim(Price) = "" and (chkprice="P") Then%>
                	<script language = 'vbscript'>
                			ShowMessage "Please specify the Price "
                			
                	</script>
                	<% response.end
			End If
			
			'Check for Amount to be more than 0
			if(Cint(best)=1 and Cint(orderType1)=1 and Cint(sectype1)=1 and Cdbl(amount)<=0) then
			%>
			<script language = 'vbscript'>
			         		ShowMessage "Order Detail Amount Must be more than Zero"
			         </script>
			         <% response.end
			end if
			
			'ensure Order Detail Amount is numeric
			If (Amount <> "") And (Not IsNumeric(Amount)) Then%>
                	<script language = 'vbscript'>
					ShowMessage "Order Detail Amount must be numeric"
					
                	</script>
                	<% response.end
			End If
		
			'validate date
			if(Cint(Request.Form("txtCalendar1"))=1) then
				If Not IsDate(validity) Then%>
						<script language = 'vbscript'>
							ShowMessage "Validity must be a valid date"
						</script>
				
						<% response.end
				End If
			end if
			
			if(trim(bond)="") then
				bond=0
			end if

			if(trim(security)="") then
				security=0
			end if 
 
            If Trim(Limit) = "" then
               Limit = 0
			end if

			if itemID = "-1" then
					'save detail data
					
					'Check for duplicate b4 u enter the detail item
					
					'get the client Code there 
					set rsDuplicate = server.createobject("Adodb.Recordset")
						sqlDuplicate = "SELECT dbo.tbOrder.Client_DPA_ FROM  dbo.tbOrder  WHERE dbo.tbOrder.Detail_DPA_ = "& ID1 &")"

					sqlDuplicate="SELECT Client_DPA_ FROM dbo.tbOrder WHERE (Order_DPA_ = "& ID &")"
					Set rsDuplicate = conn.Execute(SQLServerFormat(HandleQuote(sqlDuplicate)))
					client=rsDuplicate("client_DPA_")					
	
					sqlDuplicate = "SELECT * FROM DB_OrderList WHERE (Client_DPA_ = "& client &") AND (Security_DPA_ = "& security &") AND (BalanceQty > 0)"	
					
					Set rsDuplicate = conn.Execute(SQLServerFormat(HandleQuote(sqlDuplicate)))
					If  Not (rsDuplicate.EOF Or rsDuplicate.BOF) Then
					  %>
							<script language = 'vbscript'>
									ShowMessage "This is a duplicate order. Please follow up the original Order."
									
							</script>
					   <% 
					   rsDuplicate.end
					   set rsDuplicate= nothing
					   response.end
					End If

						
					if(isnull(validity)) then
					sqlStr = "INSERT INTO [OrdDetail] (OrdDetailPrice,Amount,Best,Bond_DPA_,OrdDetailQty" & _
							",OrdDetail_DPA_,Order_DPA_, Limit,Security_DPA_) SELECT " & "'" & price & "'" & " as OrdDetailPrice" & _
							"," & " " & CDbl(Amount) & " " & " as Amount" & _
							"," & " " & Best & " " & " as Amount" & _
							"," & " " & bond & " " & " as Bond_DPA_" & _
							"," & " " & CDbl(qty) & " " & " as OrdDetailQty" & _							
							"," & " " & "iif(isnull(max([OrdDetail_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'OrdDetail'),max([OrdDetail_DPA_]) + 1)" & " " & " as OrdDetail_DPA_" & _
							"," & " " & ID & " " & " as Order_DPA_" & _
							"," & " " & Limit & " " & " as Limit" & _
							"," & " " & security & " " & " as Security_DPA_" & _
							" FROM [OrdDetail]"					
					else
					sqlStr = "INSERT INTO [OrdDetail] (OrdDetailPrice,Amount,Best,Bond_DPA_,OrdDetailQty,OrdDetailValidity" & _
							",OrdDetail_DPA_,Order_DPA_, Limit, Security_DPA_) SELECT " & "'" & price & "'" & " as OrdDetailPrice" & _
							"," & " " & CDbl(Amount) & " " & " as Amount" & _
							"," & " " & Best & " " & " as Best" & _
							"," & " " & bond & " " & " as Bond_DPA_" & _
							"," & " " & CDbl(qty) & " " & " as OrdDetailQty" & _
							"," & "#" & validity & "#" & " as OrdDetailValidity" & _
							"," & " " & "iif(isnull(max([OrdDetail_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'OrdDetail'),max([OrdDetail_DPA_]) + 1)" & " " & " as OrdDetail_DPA_" & _
							"," & " " & ID & " " & " as Order_DPA_" & _
							"," & " " & Limit & " " & " as Limit" & _
							"," & " " & security & " " & " as Security_DPA_" & _
							" FROM [OrdDetail]"					
					end if
		    else
					'edit detail data
					if(isnull(validity)) then
					sqlStr = "UPDATE OrdDetail SET OrdDetailPrice = '" & price & "', Amount = " & CDbl(Amount) & ",Best = " & Best & ",OrdDetailQty = " & CDbl(qty) & "," & _
							" OrdDetailValidity = null ," & _
							" Bond_DPA_ = " & bond & "," & _
							" Security_DPA_ = " & security & "," & _
							" Limit = " & Limit & _ 
							" WHERE OrdDetail_DPA_=" & itemID					
					else
					sqlStr = "UPDATE OrdDetail SET OrdDetailPrice = '" & price & "', Amount = " & CDbl(Amount) & ", Best = " & Best & ",OrdDetailQty = " & CDbl(qty) & "," & _
							" OrdDetailValidity = #" & validity & "#," & _
							" Bond_DPA_ = " & bond & "," & _
							" Security_DPA_ = " & security & "," & _
							" Limit = " & Limit & _ 
							" WHERE OrdDetail_DPA_=" & itemID					
					end if
		    end if
		    
			sqlStr1="UPDATE tbOrder Set ChangedBy=" & UserId & ",TimeChanged=GetDate() Where( Order_DPA_=" & ID & ")"
			
       							
			
			conn.BeginTrans
					conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))			
					conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr1))			
			conn.CommitTrans		
			
			 'retrieve the item ID
				sqlStr = "SELECT OrdDetail.OrdDetail_DPA_  FROM OrdDetail WHERE OrdDetail.Order_DPA_=" & ID
				Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
				ID = Rs.Fields("OrdDetail_DPA_")
			
			conn.execute ("ClientTotalsDelete")
			conn.execute ("ClientTotalsProcedure")
			conn.execute ("ClientBalancesDelete")
			conn.execute ("ClientBalancesProcedure")
			Conn.Execute("BestBalanceQtys")

			conn.Close
			Set conn = Nothing
			WriteDialogRelocateScript "EditOrderItem.asp?ID=" & ID
			Response.End
			
		case "EXECUTE_DELETE"
			'ensure at least one detail record is left over	
			sqlStr = "SELECT COUNT(OrdDetail_DPA_) as Total FROM [OrdDetail] WHERE Order_DPA_=" & ID
			Set conn = GetActiveConnection("KBroker")
	        
			Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
			If rs.EOF Or rs.BOF Then%>
					<script language = 'vbscript'>
                			ShowMessage "The database is corrupted"
                			
					</script>
					<%response.end
			End If
			If (CInt(rs.Fields("Total")) < 2) Then%>
					<script language = 'vbscript'>
                			ShowMessage "There must be at least one Order Detail"
                			
					</script>
					<%response.end
			End If
	        
			'find out whether any child records exist
			sqlStr = "SELECT Child,DeletionMessage,ParentKey FROM [_Parent_Child_Links_] WHERE (Parent = 'OrdDetail') AND (ChildType = " & LinkedIndependent & ")"
			Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
			If Not (rs.BOF Or rs.EOF) Then
					Dim childRS
					Dim tableName
	                
					rs.MoveFirst
					Do Until rs.EOF
                				tableName = rs.Fields("Child")
							sqlStr = "SELECT TOP 1 * FROM [" & tableName & "] WHERE OrdDetail_DPA_ = " & ItemID & " and Deleted=0"
							Set childRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
							If Not (childRS.BOF Or childRS.EOF) Then%>
                					<script language = 'vbscript'>
                						ShowMessage "<%=rs.Fields("DeletionMessage")%>"
                						
                					</script>
                					<%response.end
							End If
							rs.MoveNext
					Loop
			End If
			
			 'retrieve the item ID
				sqlStr = "SELECT OrdDetail.OrdDetail_DPA_  FROM OrdDetail WHERE Deleted=0 and OrdDetail.Order_DPA_=" & ID & " AND OrdDetail_DPA_ <> " & ItemID
				Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
				if (Rs.eof and Rs.bof) then
				%>
					<script language = 'vbscript'>
                			ShowMessage "There must be at least one Order Detail"
                			
					</script>
					<%response.end
					else
				otherItemID = Rs.Fields("OrdDetail_DPA_")
				end if								
			
			'Set Rs=nothing
			'delete from database
			sqlStr = "Update [OrdDetail] Set Deleted=1 WHERE OrdDetail_DPA_ = " & ItemID
			conn.Execute SQLServerFormat(HandleQuote(sqlStr))			
			
			conn.execute ("ClientTotalsDelete")
			conn.execute ("ClientTotalsProcedure")
			conn.execute ("ClientBalancesDelete")
			conn.execute ("ClientBalancesProcedure")
			Conn.Execute("BestBalanceQtys")

			conn.Close
			Set conn = Nothing
			WriteDialogRelocateScript "EditOrderItem.asp?ID=" & otherItemID
			Response.End
   	end select
   		
   	
%>
<%
 Set conn = GetActiveConnection("KBroker")
 
	Dim rowCount
	Dim securityList
	Dim quote 
	Dim secType
	dim CalendarChk
	Dim Ordertype

	if(trim(from)="") then
	sqlstr = "SELECT  isnull(dbo.OrdDetail.Limit,0) as Limit, dbo.Security.SecurityCode, dbo.OrdDetail.Best, dbo.OrdDetail.Amount, dbo.Bond.BondIssue AS bondDescription,  " & _
				 "                       dbo.OrdDetail.OrdDetailPrice, dbo.OrdDetail.Limit, dbo.OrdDetail.OrdDetailQty, dbo.OrdDetail.OrdDetailValidity, dbo.OrdDetail.OrdDetail_DPA_,  " & _
				 "                       dbo.Security.SecurityName, dbo.Security.Security_DPA_, dbo.tbOrder.OrderSecType_DPA_, dbo.tbOrder.OrderType_DPA_, '' AS cboprice,  " & _
				 "                       '' AS CalendarChk, dbo.OrderSecType.OrderSecTypeDisplayName AS ordDetailSecType, dbo.OrderType.OrderTypeDescription AS ordDetailType,  " & _
				 "                       ISNULL(dbo.ClientHoldingsStatus.CurrentHoldings, 0) AS CurrentHoldings, ISNULL(dbo.ClientHoldingsStatus.AvailableHoldings, 0)  " & _
				 "                       AS AvailableHoldings , dbo.tbOrder.Client_DPA_ " & _
				 " FROM         dbo.OrdDetail LEFT OUTER JOIN " & _
				 "                       dbo.Bond ON dbo.OrdDetail.Bond_DPA_ = dbo.Bond.Bond_DPA_ INNER JOIN " & _
				 "                       dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ INNER JOIN " & _
				 "                       dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN " & _
				 "                       dbo.OrderSecType ON dbo.tbOrder.OrderSecType_DPA_ = dbo.OrderSecType.OrderSecType_DPA_ INNER JOIN " & _
				 "                       dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ LEFT OUTER JOIN " & _
				 "                       dbo.ClientHoldingsStatus ON dbo.tbOrder.Client_DPA_ = dbo.ClientHoldingsStatus.Client_DPA_ AND  " & _
				 "                       dbo.OrdDetail.Security_DPA_ = dbo.ClientHoldingsStatus.Security_DPA_ " & _
				 " WHERE     (dbo.OrdDetail.Order_DPA_ IN " & _
				 "                           (SELECT     OrdDetail.Order_DPA_ " & _
				 "                             FROM          OrdDetail " & _
	 "                             WHERE      OrdDetail.OrdDetail_DPA_ = " & ID & ")) AND (dbo.OrdDetail.Deleted = 0)"
  
	    else

    sqlstr = "SELECT DISTINCT  " & _
			 " isnull(dbo.OrdDetailList.Limit,0) as Limit, dbo.OrdDetailList.SecurityCode, dbo.OrdDetailList.OrderTypeSale,  dbo.OrdDetailList.BondDescription,  " & _
			 "                       dbo.OrdDetailList.OrdDetailPrice, dbo.FullOrderList.Amount AS Amount, dbo.OrdDetailList.OrdDetailQty, dbo.OrdDetailList.OrdDetailValidity,  " & _
			 "                       dbo.OrdDetailList.OrdDetail_DPA_, dbo.OrdDetailList.SecurityName, dbo.OrdDetailList.Security_DPA_, dbo.OrdDetailList.OrderSecType_DPA_,  " & _
			 "                       '' AS cboprice, dbo.OrdDetailList.Client_DPA_, '' AS CalendarChk, dbo.OrdDetailList.limit, dbo.OrdDetailList.OrdDetailSecType, dbo.OrdDetailList.OrdDetailType,  " & _
			 "                       ISNULL(dbo.ClientHoldingsStatus.CurrentHoldings, '0') AS CurrentHoldings, ISNULL(dbo.ClientHoldingsStatus.AvailableHoldings, 0)  " & _
			 "                       AS AvailableHoldings " & _
			 " FROM         dbo.OrdDetailList INNER JOIN " & _
			 "                       dbo.FullOrderList ON dbo.OrdDetailList.OrdDetail_DPA_ = dbo.FullOrderList.OrdDetail_DPA_ LEFT OUTER JOIN " & _
			 "                       dbo.ClientHoldingsStatus ON dbo.OrdDetailList.Client_DPA_ = dbo.ClientHoldingsStatus.Client_DPA_ AND  " & _
			 "                       dbo.OrdDetailList.Security_DPA_ = dbo.ClientHoldingsStatus.Security_DPA_ " & _
			 " WHERE     (dbo.OrdDetailList.Order_DPA_ = " & ID & ")"		 
    end if
     
	Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
	
	If Not(rs.EOF Or rs.BOF) Then   

	secType = rs.Fields("OrderSecType_DPA_")
	
	quote =  chr(34)

	securityList = GetSecurityList("cboSecurity")	    
	
	bondList = GetBondList("cboBond")       
	
	ClientID = rs.Fields("Client_DPA_")
    
		<!-- grid data -->
		 
		 'row data		 
		rs.MoveFirst
		rowCount = 0
		Do Until rs.EOF 
		
			if(isnull(rs.Fields("OrdDetailValidity")) or rs.Fields("OrdDetailValidity")="") then
			   validdate = 0
			   CalendarChk = "<input type = 'hidden' name ='chkdate' id='chkdate' size='1' value='0'>"		
			else
			   validdate = 1
			   CalendarChk = "<input type = 'hidden' name ='chkdate' id='chkdate' size='1' value='1'>"		
			end if

'======================= Begin_Alter_Across_Entities =================================
			'row ID 
			rowData = rowData & quote & rs.Fields("OrdDetail_DPA_") & quote & " : " 
			
			'row data 
			rowData = rowData & "[" 
			rowData = rowData & quote & rs.Fields("OrdDetail_DPA_") & quote & "," 
			
			if(trim(rs.Fields("ordDetailSecType"))="Fixed") then
			 CurrentHoldings = "N/A"
			 AvailableHoldings = "N/A"

			dim SecurityCodes
			SecurityCodes=rs.Fields("SecurityCode")
			SecurityCodes=split(SecurityCodes," ")
			SecurityCode=SecurityCodes(0)
			rowData = rowData & quote & SecurityCode & quote & ","
			else
			 CurrentHoldings = FormatNumCommasOnly(rs.Fields("CurrentHoldings"))
			 AvailableHoldings = FormatNumCommasOnly(rs.Fields("AvailableHoldings"))
			rowData = rowData & quote & rs.Fields("SecurityCode") & quote & ","
			end if			
			
			rowData = rowData & quote & rs.Fields("BondDescription") & quote & ","		
			
			rowData = rowData & quote & CurrentHoldings & quote & ","
			rowData = rowData & quote &  AvailableHoldings & quote & ","
			Ordertype=rs("ordDetailType")
			
				if(rs("Best")=True) then
				   BestCboPrice = "<input type = 'hidden' name ='BestCboPrice' id='BestCboPrice' size='1' value='01'>"
				 rowData = rowData & quote & BestCboPrice & quote & ","
				else
				   BestCboPrice="<input type = 'hidden' name ='BestCboPrice' id='BestCboPrice' size='1' value='00'>"
				  rowData = rowData & quote & BestCboPrice & quote & ","				
				end if

			rowData = rowData & quote & FormatNumCommasOnly(rs.Fields("Limit")) & quote & ","
			rowData = rowData & quote & FormatNumCommasOnly(rs.Fields("OrdDetailQty")) & quote & ","
			
			if(rs.Fields("Best")=True) then
				BestPrice="BEST"
				rowData = rowData & quote & BestPrice & quote & ","
			else
				if(Ordertype="Purchase") then			
				rowData = rowData & quote & FormatNumEx(rs.Fields("OrdDetailPrice"),4) & quote & ","
				else
				rowData = rowData & quote & FormatNum(rs.Fields("OrdDetailPrice")) & quote & ","
				end if
			end if		
			
			rowData = rowData & quote & FormatNum(rs.Fields("Amount")) & quote & ","
			rowData = rowData & quote & CalendarChk & quote & ","
			rowData = rowData & quote & FormatDate(rs.Fields("OrdDetailValidity")) & quote & "," 
			rowData = rowData & quote & " " & quote
			rowData = rowData & "]" 
			
			'build the row IDs array
			rowIDs = rowIDs & quote & rs.Fields("OrdDetail_DPA_") & quote 
			rowCount = rowCount + 1
			
'======================= End_Alter_Across_Entities =================================

			rs.MoveNext 
			
			'build the row IDs array
			rowIDs = rowIDs & "," 
			rowData = rowData & ","	
		Loop
		rs.MoveFirst 
		
	End if

	' Build new line to add record
	dim cboprice

	if(trim(rs.Fields("ordDetailSecType"))="Security") then
	 cboprice="<select name='cboprice' onchange='ChangePrice(this,1);'> " & _			
			 "<option selected SearchCode = '0' SearchText = 'P' value = 'P'>P</option>" & _			
			 "<option value='BF'>BF</option>" & _			
		     "</select>"

	else
	 cboprice="<select name='cboprice' disabled onchange='ChangePrice(this,1);'> " & _			
			 "<option selected SearchCode = '0' SearchText = 'P' value = 'P'>P</option>" & _			
			 "<option value='BF'>BF</option>" & _			
		     "</select>"	
	end if

	'row ID 
	
	rowData = rowData & quote & 0 & quote & " : " 
	 
	 CalendarChk="<input type ='checkbox' name ='chkdate' id='chkdate' size='1' class='BorderLess' onClick = 'ChangeCalendar1(this);'>"		
	 
	'row data 
	rowData = rowData & "[" 
	rowData = rowData & quote & "New Line" & quote & "," 
	rowData = rowData & quote & securityList & quote & ","
	rowData = rowData & quote & BondList & quote & ","
	
	rowData = rowData & quote & "<input type = 'text'  readonly id = 'txtCurrentHoldings' name ='txtCurrentHoldings' style='BORDER-RIGHT: SILVER 1px OUTSET ridge;BORDER-TOP: SILVER 1px  SOLID;ONT-WEIGHT: normal;FONT-SIZE: 8pt;BORDER-LEFT: SILVER 1px  SOLID;COLOR: NAVY; BORDER-BOTTOM: SILVER 1px OUTSET;BACKGROUND-COLOR: #C0C0C0;FONT-FAMILY: verdana, arial, helvetica, sans-serif;' OnKeyUp='JavaScript: updateBestPrice(this)' OnBlur='JavaScript: format2Number(this)' size='9' OnClick='event.cancelBubble=true;'>" & quote & ","
	rowData = rowData & quote & "<input type = 'text'  readonly id = 'txtAvailableHoldings' name ='txtAvailableHoldings' style='BORDER-RIGHT: SILVER 1px OUTSET ridge;BORDER-TOP: SILVER 1px  SOLID;ONT-WEIGHT: normal;FONT-SIZE: 8pt;BORDER-LEFT: SILVER 1px  SOLID;COLOR: NAVY; BORDER-BOTTOM: SILVER 1px OUTSET;BACKGROUND-COLOR: #C0C0C0;FONT-FAMILY: verdana, arial, helvetica, sans-serif;' OnKeyUp='JavaScript: updateBestPrice(this)' OnBlur='JavaScript: format2Number(this)' size='9' OnClick='event.cancelBubble=true;'>" & quote & ","
	rowData = rowData & quote & cboprice & quote & ","
	rowData = rowData & quote & "<input type = 'text' name ='txtLimit' disabled STYLE='width: 95px;text-align: right' id = 'txtLimit' size='9' OnClick='event.cancelBubble=true;'  OnBlur='JavaScript: format2NumberCommasOnly(this)'>" & quote & ","
	rowData = rowData & quote & "<input type = 'text' name ='txtQty' STYLE='width: 95px;text-align: right' id = 'txtQty' size='9' OnClick='event.cancelBubble=true;'  OnBlur='JavaScript: format2NumberCommasOnly(this)'>" & quote & ","
	rowData = rowData & quote &  "<input type = 'text' name ='txtPrice' STYLE='width: 95px;text-align: right' id = 'txtPrice' OnKeyUp='JavaScript: updateBestPrice(this)' OnBlur='JavaScript: format2Number(this)' size='9' OnClick='event.cancelBubble=true;'>" & quote & ","		
	rowData = rowData & quote & "<input type = 'text' name ='txtAmount' STYLE='width: 95px;text-align: right' id = 'txtAmount' disabled size='9'>" & quote & ","	
	rowData = rowData & quote & CalendarChk & quote & ","
	rowData = rowData & quote & "<input type='text' name='txtValidity' size=40 value='" & FormatDate(Date) & "' disabled=true OnClick='event.cancelBubble=true;'>" & quote & ","
	rowData = rowData & quote & "<input type=button value='Add' Class=Buttons OnClick='JavaScript: AddRowInProgress();'>&nbsp;&nbsp;<input type='reset' value='Cancel' Class=Buttons>" & quote 
	
	rowData = rowData & "]" 
	

	'build the row IDs array 
	rowIDs = rowIDs & quote & 0 & quote 
	rowCount = rowCount + 1
		
'======================= Begin_Alter_Across_Entities =================================%> 
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit <%=DataEntity%> Item</title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
 <script language='vbscript'>
			function ItemSelected(itemID)
					
 					frm<%=DataSource%>Item.elements("ItemID").value = itemID
			end function
			
			
			function SaveInPlaceEdit()
				    Dim myOwnerFrame	
					Dim my
					UpdateID
					Set window.parent.dialogArguments.opener.parent.frames("footer").editDocOpener = window.self
					frm<%=DataSource%>Item.target = "deleteFrame" 
					frm<%=DataSource%>Item.submit
					restoreID
			end function
		</script>
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
		
		<!-- ActiveUI stylesheet and scripts -->
		<link href="../runtime/classic/activeui.css" rel="stylesheet" type="text/css">
		<script src="../runtime/activeui.js"></script>
		<!-- Include patches here -->
		<script src="../runtime/paging1.js"></script>
		<!-- grid format -->
		<%	Dim colIndex
			
			Dim bondColumn
			Dim SecondLastColumn
			Dim LastColumn
			
			colIndex = 2
			if rs.Fields("OrderSecType_DPA_") = 1 then 'fixed security
				bondColumn = ".active-column-" & colIndex & " {width: 100px;}" & chr(13)
			else
				bondColumn = ".active-column-" & colIndex & " {width: 0px;}" & chr(13)
			end if

			colIndex = 6
			if rs.Fields("OrderSecType_DPA_") = 2 and lcase(trim(rs.Fields("ordDetailType"))) = "purchase" then  
				limitColumn = ".active-column-" & colIndex & " {width: 70px;}" & chr(13)
			else
				limitColumn = ".active-column-" & colIndex & " {width: 0px;}" & chr(13)
			end if

		%>
		<style>
			.active-controls-grid {height: 100%; font: menu;}
			.active-row-highlight .active-row-cell {background-color: skyblue}
			.active-selection-true, .active-selection-true .active-row-cell {
				color: blue!important;
				background-color: bisque!important;
				}
					
			.active-column-0 {width: 50px;}
			.active-column-1 {width: 100px;}
			<%=bondColumn%>
			.active-column-3 {width: 60px; text-align: right;}
			.active-column-4 {width: 60px; text-align: right;}
			.active-column-5 {width: 45px;}
			<%=limitColumn%>
			.active-column-7 {width: 70px;}	
			.active-column-8 {width: 70px; text-align: right;}
			.active-column-9 {width: 70px; text-align: right;}		
			.active-column-10 {width: 25px; text-align: centre;}
			.active-column-11 {width: 110px;}			
			
			.active-grid-row,
			.active-grid-row.active-list-item,
			.active-scroll-left .active-list-item {height: 22px;}	
			
		</style>

</head>

<body Class="Dialog" SCROLL="NO">	
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<form name = 'frm<%=DataSource%>Item' id='frmMain' method = 'post' action = '<%=DataSource%>Item.asp' OnSubmit="UpdateID();">

<SCRIPT language="JavaScript">
	var calValidity;
	var cal1Validity;	
	function changeDateInterface(){
		calValidity = new ctlSpiffyCalendarBox('calValidity', 'frm<%=DataSource%>Item', 'txtValidity', 'cmdValDate','<%= FormatDate(Date) %>', 1); 
		
		calValidity.readonly = false;
		calValidity.returnOutStringOnWrite(); 		
		var parentDiv = document.all.item("txtValidity").parentNode;
		parentDiv.innerHTML = calValidity.writeControl();
		parentDiv.style.zIndex = 10;
		parentDiv.childNodes(1).style.zIndex = 10	;		
				
		document.all.item('cboSecurity').focus();
		document.all.item('txtValidity').disabled=true;
	}	
	
	document.body.onload = changeDateInterface;	
</SCRIPT>
		<script language="javascript">
		
			//column titles 
			var colCount = 13;
			var colNames = ["", "Security","Bond", "txtCurrentInPlace", "txtAvailableInPlace","Price",  
					 "txtLimitInPlace", "txtQtyInPlace", "txtPriceInPlace", "txtAmountInPlace", "Chk", "txtCalendar", ""];
			
			var myColumns = ["Line No", "Security","Bond","Current","Available","",  
					 "Limit", "Quantity", "Price", "Amount",  "","Validity", ""];
		</script>

<%'======================= End_Alter_Across_Entities =================================%>			
		<script language="javascript">
			//data			
			var myData = {<%=rowData%>}; 
			var myRowIDs = [<%=rowIDs%>]; 			
			var doubleclicked = 0;			
			var inPlaceEdit = false;
			var addInProgress = false;
			var clickedRowID = -1; 
			var dataChanged = false;
			var prevRow = -1;//the row currently under in-place edit
			
			function EditInPlaceDataChanged()
			{
				dataChanged = true;
			}
			
			function AddRowInProgress()
			{
				addInProgress = true;
			}
			
			var currentSecurityName = "";
			var currentBondName = "" ;

var RowEditFn = function(src)
			{
				var rowIndex = src.getProperty("row/index");
				var ordertypa = document.frmMain.elements("Ordertype").value;
                var sectype = document.frmMain.elements("sectype").value;
     
				var i;
				var status;
				var best=0;
				var LimitDisabled = true;
				
				if (rowIndex ==0) return;
				
				if (doubleclicked==0)
				{
				doubleclicked=1
				}
				else
				{
				return;
				}
				
				for(i = 0; i < colCount; i++)
				{
					if(colNames[i] != "")
					{
						if(prevRow >= 0)
						{
						
							if(colNames[i]=="Security")
							{
								myData[prevRow][i] = currentSecurityName;								
							}
							else
								if(colNames[i]=="Bond")
								{
									myData[prevRow][i] = currentBondName;
								}
								else
								{
                                   if(colNames[i]=="Price" || colNames[i]=="Chk") {
								    // Do nothing
								   }
								   else {
									    myData[prevRow][i] = document.frmMain.elements(colNames[i]).value;			
									}
								}
						}

						if(colNames[i]=="Security")
						{	
							currentSecurityName =  myData[rowIndex][i];
							myData[rowIndex][i] = inPlaceSecurityList;
							
						}
						else 
							 if(colNames[i]=="Bond")
							  {	
								currentBondName =  myData[rowIndex][i];
								myData[rowIndex][i] = inPlaceBondList;							
							  }
							  else
								if(colNames[i]=="Price")
								{
								   
									if(myData[rowIndex][i]=="<input type = 'hidden' name ='BestCboPrice' id='BestCboPrice' size='1' value='01'>")
									{
									   LimitDisabled  = false;
									   best = 1
									   document.all.item("txtbestinplace").value = best;

									   myData[rowIndex][i]="<select name='cboprices' onchange='ChangePrice(this,2);'><option selected SearchCode = '0' SearchText = 'BF' value = 'BF'>BF</option><option value='P'>P</option></select>";						
									}	
									
									if(myData[rowIndex][i]=="<input type = 'hidden' name ='BestCboPrice' id='BestCboPrice' size='1' value='00'>")
									{ 
									  best = 0
									  document.all.item("txtbestinplace").value = best;

									  myData[rowIndex][i]="<select name='cboprices' onchange='ChangePrice(this,2);'><option selected SearchCode = '0' SearchText = 'P' value = 'P'>P</option><option value='BF'>BF</option></select>";				
									}
							    }
						        else
									if(colNames[i]=="Chk")
									{														
										if(myData[rowIndex][i]=="<input type = 'hidden' name ='chkdate' id='chkdate' size='1' value='1'>")
										{							
										  myData[rowIndex][i]="<input type ='checkbox' name ='chkdate' id='chkdate' size='1' checked class='BorderLess' onClick = 'ChangeCalendar(this);'>"		
										}
										else
										{
										  myData[rowIndex][i]="<input type ='checkbox' name ='chkdate' id='chkdate' size='1' class='BorderLess' onClick = 'ChangeCalendar(this);'>"		
										}
								   }						
						           else						
						           {    
										
										if(colNames[i]=="txtCurrentInPlace")
										{
											myData[rowIndex][i] = "<INPUT TYPE='text' NAME='" + colNames[i] + "' ID='" + colNames[i] + "' VALUE='" + myData[rowIndex][i] + "' onChange = 'EditInPlaceDataChanged();event.cancelBubble=true;' style='BORDER-RIGHT: SILVER 1px OUTSET ridge;BORDER-TOP: SILVER 1px  SOLID;ONT-WEIGHT: normal;FONT-SIZE: 8pt;BORDER-LEFT: SILVER 1px  SOLID;COLOR: NAVY; BORDER-BOTTOM: SILVER 1px OUTSET;BACKGROUND-COLOR: #C0C0C0;FONT-FAMILY: verdana, arial, helvetica, sans-serif;' align='right' readonly OnClick='event.cancelBubble=true;'>";
										}

										else if(colNames[i]=="txtAvailableInPlace")

										{
											myData[rowIndex][i] = "<INPUT TYPE='text' NAME='" + colNames[i] + "' ID='" + colNames[i] + "' VALUE='" + myData[rowIndex][i] + "' onChange = 'EditInPlaceDataChanged();event.cancelBubble=true;' style='BORDER-RIGHT: SILVER 1px OUTSET ridge;BORDER-TOP: SILVER 1px  SOLID;ONT-WEIGHT: normal;FONT-SIZE: 8pt;BORDER-LEFT: SILVER 1px  SOLID;COLOR: NAVY; BORDER-BOTTOM: SILVER 1px OUTSET;BACKGROUND-COLOR: #C0C0C0;FONT-FAMILY: verdana, arial, helvetica, sans-serif;'  align='right' readonly OnClick='event.cancelBubble=true;'>";
										}

										else if(colNames[i]=="txtLimitInPlace")
										{ 
											  if (LimitDisabled == true) 
											   {
												 myData[rowIndex][i] = "<INPUT TYPE='text' NAME='" + colNames[i] + "' ID='" + colNames[i] + "' VALUE='" + myData[rowIndex][i] + "' onChange = 'EditInPlaceDataChanged();event.cancelBubble=true;' disabled OnClick='event.cancelBubble=true;'>";
											   }
											  else
											  {
												 myData[rowIndex][i] = "<INPUT TYPE='text' NAME='" + colNames[i] + "' ID='" + colNames[i] + "' VALUE='" + myData[rowIndex][i] + "' onChange = 'EditInPlaceDataChanged();event.cancelBubble=true;' OnClick='event.cancelBubble=true;'>";
											  }
										}

										else if(colNames[i]=="txtAmountInPlace")
										{
											if(ordertypa=="Sale") 
											{
												myData[rowIndex][i] = "<INPUT TYPE='text' NAME='" + colNames[i] + "' ID='" + colNames[i] + "' VALUE='" + myData[rowIndex][i] + "' onChange = 'EditInPlaceDataChanged();event.cancelBubble=true;' disabled OnClick='event.cancelBubble=true;'>";
											}
											else
											{																		
												if(best==1)							
												{
												myData[rowIndex][i] = "<INPUT TYPE='text' NAME='" + colNames[i] + "' ID='" + colNames[i] + "' VALUE='" + myData[rowIndex][i] + "' onChange = 'EditInPlaceDataChanged();event.cancelBubble=true;' OnClick='event.cancelBubble=true;'>";
												}
												else
												{
												myData[rowIndex][i] = "<INPUT TYPE='text' NAME='" + colNames[i] + "' ID='" + colNames[i] + "' VALUE='" + myData[rowIndex][i] + "' onChange = 'EditInPlaceDataChanged();event.cancelBubble=true;' disabled OnClick='event.cancelBubble=true;'>";
												}
											}
										}
							            else
											if(colNames[i]=="txtQtyInPlace")
											{
											
												if(ordertypa=="Purchase") 
												{
													if(best==1)							
													{
													
													myData[rowIndex][i] = "<INPUT TYPE='text' NAME='" + colNames[i] + "' ID='" + colNames[i] + "' VALUE='" + myData[rowIndex][i] + "' onChange = 'EditInPlaceDataChanged();event.cancelBubble=true;' disabled OnClick='event.cancelBubble=true;'>";
													}
													else
													{
													
													myData[rowIndex][i] = "<INPUT TYPE='text' NAME='" + colNames[i] + "' ID='" + colNames[i] + "' VALUE='" + myData[rowIndex][i] + "' onChange = 'EditInPlaceDataChanged();event.cancelBubble=true;' OnClick='event.cancelBubble=true;'>";
													}
												}
												else
												{
												   myData[rowIndex][i] = "<INPUT TYPE='text' NAME='" + colNames[i] + "' ID='" + colNames[i] + "' VALUE='" + myData[rowIndex][i] + "' onChange = 'EditInPlaceDataChanged();event.cancelBubble=true;' OnClick='event.cancelBubble=true;'  >";
												}
								           }
								           else
										   {
                                              status = "";

										      if (myData[rowIndex][i]=="BEST")
											  {
											    status = "Disabled" ;
											  }
											  
											   myData[rowIndex][i] = "<INPUT TYPE='text' NAME='" + colNames[i] + "' ID='" + colNames[i] + "' VALUE='" + myData[rowIndex][i] + "' onChange = 'EditInPlaceDataChanged();event.cancelBubble=true;' OnClick='event.cancelBubble=true;' " + status + " >  ";
										   }
							            }
					               }					
				                }

				myData[rowIndex][colCount - 1] = "<INPUT TYPE='button' class='Buttons' VALUE='Save' onClick = 'SaveInPlaceEdit();event.cancelBubble=true;'>&nbsp;<INPUT TYPE='button' class='Buttons' VALUE='Cancel' onClick = 'cancelEditRow();event.cancelBubble=true;'>";
				inPlaceEdit = true;
				prevRow = rowIndex;
				grid.refresh();
		
				var secList = document.frmMain.elements("cboSecurityInPlace");
				var bonList =document.frmMain.elements("cboBondInPlace");
				secList.focus();
				for (i=0; i < secList.options.length; i++) {
					
					if(secList.options(i).text == currentSecurityName)
					{
							secList.options(i).selected = true;
					}
				}
				for (i=0; i < bonList.options.length; i++) {					
					if(bonList.options(i).text == currentBondName)
					{
							bonList.options(i).selected = true;
					}
				}							
				
				cal1Validity = new ctlSpiffyCalendarBox('cal1Validity', 'frm<%=DataSource%>Item', 'txtCalendar', 'cmdDate','<%= FormatDate(Date) %>', 1); 
		
				cal1Validity.readonly = false;
				cal1Validity.returnOutStringOnWrite(); 		
				var parentDiv = document.all.item("txtCalendar").parentNode;
				parentDiv.innerHTML = cal1Validity.writeControl();
				parentDiv.style.zIndex = 10;
				parentDiv.childNodes(1).style.zIndex = 10	;		
								
			}
			
			function cancelEditRow(){
						var i;
							for(i = 0; i < colCount; i++)
							{
								if(colNames[i] != "")
								{
									if(colNames[i]=="Security")
									{
										myData[prevRow][i] = currentSecurityName;
									}
									else
									if(colNames[i]=="Bond")
									{
										myData[prevRow][i] = currentBondName;
									}
									else
									if(colNames[i]=="Price")
									{
										if(myData[prevRow][i]=="<select name='cboprices' onchange='ChangePrice(this,2);'><option selected SearchCode = '0' SearchText = 'BF' value = 'BF'>BF</option><option value='P'>P</option></select>")
										{
										myData[prevRow][i]="<input type = 'hidden' name ='BestCboPrice' id='BestCboPrice' size='1' value='01'>";										
										}						
										if(myData[prevRow][i]=="<select name='cboprices' onchange='ChangePrice(this,2);'><option selected SearchCode = '0' SearchText = 'P' value = 'P'>P</option><option value='BF'>BF</option></select>")
										{
										myData[prevRow][i]="<input type = 'hidden' name ='BestCboPrice' id='BestCboPrice' size='1' value='00'>";																
										}								
									}																		
									else
									if(colNames[i]=="Chk")
									{
										
										if(myData[prevRow][i]=="<input type ='checkbox' name ='chkdate' id='chkdate' size='1' checked class='BorderLess' onClick = 'ChangeCalendar(this);'>")
										{							
										myData[prevRow][i]="<input type = 'hidden' name ='chkdate' id='chkdate' size='1' value='1'>"		
										}
										else
										{
										myData[prevRow][i]="<input type = 'hidden' name ='chkdate' id='chkdate' size='1' value='0'>"		
										}
									}
									else
									{
										myData[prevRow][i] = document.frmMain.elements(colNames[i]).value;
									}
								}
							}								
							
								calValidity = new ctlSpiffyCalendarBox('calValidity', 'frm<%=DataSource%>Item', 'txtValidity', 'cmdValDate','<%= FormatDate(Date) %>', 1); 
		                        
								calValidity.readonly = false;
								calValidity.returnOutStringOnWrite(); 		
								var parentDiv = document.all.item("txtValidity").parentNode;
								parentDiv.innerHTML = calValidity.writeControl();
								parentDiv.style.zIndex = 10;
								parentDiv.childNodes(1).style.zIndex = 10	;		
								
								document.frmMain.elements("txtValidity").disabled=true
							myData[prevRow][colCount - 1] = "";						
							inPlaceEdit = false;
							prevRow = -1;
							grid.refresh();
			}
			
			var RowChangeFn = function(src)
			{
				doubleclicked=0;
				if(inPlaceEdit || addInProgress)
				{
					if(dataChanged || addInProgress)
					{
						ItemSelected(prevRow);
						SaveInPlaceEdit();
					}
					else
					{
						if(prevRow != clickedRowID)
						{
							cancelEditRow();
						}
					}
				}
			}
			
			var HandleClick = function(src)
			{
				clickedRowID = src.getProperty("row/index");
				ItemSelected(clickedRowID);
			}
			
			var headerID;
			
			function UpdateID(){
				headerID = window.parent.frames["header"].document.all.item("ID").value;
				document.all.item("ID").value = headerID;
				
			}
			
			function restoreID(){
				document.all.item("ID").value = headerID;
			}
			
			function HandleDeleteAction()
			{
					document.frmMain.elements("action").value = "Execute_Delete"
					SaveInPlaceEdit();
					document.frmMain.elements("action").value = "Execute_Detail"
			}
			
			//get ready for in-place edit
			var inPlaceSecurityList = "<%=GetSecurityList("cboSecurityInPlace")%>"
			var inPlaceBondList= "<%=GetBondList("cboBondInPlace")%>"
		
function UpdateClientHoldings(securityctrl){

		 var client = <%=ClientID%>
         var security = securityctrl.value
         //Clean up
		 if (securityctrl.name == 'cboSecurityInPlace') 
		  {
			 document.getElementById("txtCurrentInPlace").innerText = "";
			 document.getElementById("txtAvailableInPlace").innerText = "";
		  }
		  else
		  {
			document.getElementById("txtCurrentHoldings").innerText = "";
			document.getElementById("txtAvailableHoldings").innerText = "";
		  }

			xmlhttp = createXMLHTTPObj();
			
			url="GetList.asp?client="+client+"&security="+security+"&action=GetClientHoldings";
			 
			xmlhttp.open("GET",url,true);

			xmlhttp.onreadystatechange=function() 
			  {
						if (xmlhttp.readyState==4) 
						{
							returnStr = xmlhttp.responseText;
							returnStr = getBodyHTML(returnStr);
							
							myArray = returnStr.split(";");
							  
							  if (securityctrl.name == 'cboSecurityInPlace') 
							  {
								 document.getElementById("txtCurrentInPlace").innerText = myArray[0]
								 document.getElementById("txtAvailableInPlace").innerText = myArray[1]
							  }
							  else
							  {
								document.getElementById("txtCurrentHoldings").innerText = myArray[0]
								document.getElementById("txtAvailableHoldings").innerText = myArray[1]
							  }
							 
						}
				}
			 
			 xmlhttp.setRequestHeader('Accept','message/x-jl-formresult');
			 xmlhttp.send(); 
		}

		function ChangePrice(thecbo,place)
			{
			var ordertype=document.frmMain.elements("Ordertype").value;			
			if(place==1)
				{				
				if(thecbo.value == "P")
					{
					    
                        document.frmMain.elements("txtPrice").disabled=false;
						document.frmMain.elements("txtPrice").value='';
						document.frmMain.elements("txtPrice").focus();

						document.frmMain.elements("txtQty").disabled=false;
						document.frmMain.elements("txtAmount").disabled=true;
						document.frmMain.elements("txtAmount").value='0';
						document.frmMain.elements("txtbest").value='0';

						document.frmMain.elements("txtLimit").value='0';
						document.frmMain.elements("txtLimit").disabled=true;
										
					}
				else
					{   							
						document.frmMain.elements("txtAmount").value='0';
						document.frmMain.elements("txtbest").value='1';
						document.frmMain.elements("txtPrice").value='BEST';	
						document.frmMain.elements("txtPrice").disabled=true;
						
							if(ordertype=='Sale')
							{   
							    document.frmMain.elements("txtLimit").value='0';
						        document.frmMain.elements("txtLimit").disabled=true;
								document.frmMain.elements("txtAmount").disabled=true;
								document.frmMain.elements("txtQty").disabled=false;
							}				
							else
							{
								document.frmMain.elements("txtLimit").value='';
						        document.frmMain.elements("txtLimit").disabled=false;
								document.frmMain.elements("txtLimit").focus();
							    
								document.frmMain.elements("txtQty").disabled=true;	
								document.frmMain.elements("txtAmount").disabled=false;					
							}
					}
				}
			else
				{
				//in-place edit	
				if(thecbo.value == "P")
					{
						document.frmMain.elements("txtPriceInPlace").disabled=false
						document.frmMain.elements("txtQtyInPlace").disabled=false
						document.frmMain.elements("txtPriceInPlace").value=""
						document.frmMain.elements("txtPriceInPlace").focus();
						document.frmMain.elements("txtAmountInPlace").disabled=true;
						document.frmMain.elements("txtAmountInPlace").value='0';
						document.frmMain.elements("txtbestinplace").value='0';
						document.frmMain.elements("txtLimitInPlace").value='0';
						document.frmMain.elements("txtLimitInPlace").disabled=true;
					}
				else
					{
						document.frmMain.elements("txtPriceInPlace").value="Best"
						document.frmMain.elements("txtPriceInPlace").disabled=true
						document.frmMain.elements("txtAmountInPlace").disabled=false;
						document.frmMain.elements("txtAmountInPlace").focus();
						document.frmMain.elements("txtAmountInPlace").value='';
						document.frmMain.elements("txtbestinplace").value='1';
					
					if(ordertype=='Sale')
						{
							document.frmMain.elements("txtAmountInPlace").disabled=true;
							document.frmMain.elements("txtQtyInPlace").disabled=false;
							document.frmMain.elements("txtLimitInPlace").value='0';
							document.frmMain.elements("txtLimitInPlace").disabled=true;
						}				
						else
						{
							document.frmMain.elements("txtLimitInPlace").value='';
							document.frmMain.elements("txtLimitInPlace").disabled=false;
							document.frmMain.elements("txtQtyInPlace").disabled=true;	
							document.frmMain.elements("txtAmountInPlace").disabled=false;
					
						}

					}
				}
			}

		function FetchAccounts(theList)
			{
		
			var i = 0;
			var entity = theList.value;			
			var toList = document.frmMain.cboBond;			
									
			frm = document.frmMain;				
			xmlhttp = createXMLHTTPObj();
			
			url="GetList.asp?ID="+entity+"&action=GetBondList";
			xmlhttp.open("GET",url,true);
			xmlhttp.onreadystatechange=function() {
				if (xmlhttp.readyState==4) {
				returnStr = xmlhttp.responseText;
				returnStr = getBodyHTML(returnStr);
			
				var secList = "<select name = '" + toList.name + "' id = '" + toList.name + "' size='1' ";
				secList += "OnClick='event.cancelBubble=true;' " ;
				secList += "onChange='event.cancelBubble=true;' " ;
				secList += "onKeypress='return (dodefaultaction()==\"\"); ' "  ;
				secList += "onKeydown='return (dodefaultaction()==\"\");' " ; 
				secList += "onKeyup='return (change(" + toList.name + "));' " ; 
				secList += "onfocus='txtval = \"\";inputIsItemCode = 1;' "  ;
				secList += "onblur='txtval = \"\";inputIsItemCode = 1;'>" ;
				secList += returnStr ;
				secList += "</select>";
				
				toList.outerHTML = secList;
				
				}
				}
			xmlhttp.setRequestHeader('Accept','message/x-jl-formresult');
			xmlhttp.send(); 
		
		
		}
		function FetchAccounts1(theList)
		{		
			var i = 0;
			var entity = theList.value;			
			var toList = document.frmMain.cboBondInPlace;
									
			frm = document.frmMain;				
			xmlhttp = createXMLHTTPObj();
			
			url="GetList.asp?ID="+entity+"&action=GetBondList";
			xmlhttp.open("GET",url,true);
			xmlhttp.onreadystatechange=function() {
				if (xmlhttp.readyState==4) {
				returnStr = xmlhttp.responseText;
				returnStr = getBodyHTML(returnStr);
			
				var secList = "<select name = '" + toList.name + "' id = '" + toList.name + "' size='1' ";
				secList += "OnClick='event.cancelBubble=true;' " ;
				secList += "onChange='event.cancelBubble=true;' " ;
				secList += "onKeypress='return (dodefaultaction()==\"\"); ' "  ;
				secList += "onKeydown='return (dodefaultaction()==\"\");' " ; 
				secList += "onKeyup='return (change(" + toList.name + "));' " ; 
				secList += "onfocus='txtval = \"\";inputIsItemCode = 1;' "  ;
				secList += "onblur='txtval = \"\";inputIsItemCode = 1;'>" ;
				secList += returnStr ;
				secList += "</select>";
				
				toList.outerHTML = secList;
				
				}
				}
			xmlhttp.setRequestHeader('Accept','message/x-jl-formresult');
			xmlhttp.send(); 
		
		
		}
		
		function ChangeCalendar(thechk)
		{
		//ShowMessage 'Please specify the Price';
		if(thechk.checked)
		{
		document.frmMain.elements("txtCalendar").disabled=false
		document.frmMain.elements("txtcalendar1").value="1"
		}
		else
		{
		document.frmMain.elements("txtCalendar").disabled=true
		document.frmMain.elements("txtcalendar1").value="0"
		}
		}
		
		function ChangeCalendar1(thechk)
		{
		//ShowMessage 'Please specify the Price';
		if(thechk.checked)
		{
		document.frmMain.elements("txtValidity").disabled=false
		document.frmMain.elements("txtcalendar1").value="1"
		}
		else
		{
		document.frmMain.elements("txtValidity").disabled=true
		document.frmMain.elements("txtcalendar1").value="0"
		}
		}
		</script> 
		
		<script language="javascript"> 

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
			
			function ResetObjects()
			{
				alert('working...')
				txtval = '';
				inputIsItemCode = 1;
			}
		</script> 
        <%       
  function GetSecurityList(listName)
		Dim secList
		Dim rsSecurity
						
		secList = "<select name = '" & listName & "' id = '" & listName & "' size='1' "
		secList = secList & "OnClick='event.cancelBubble=true;' "  
		if(trim(listName)="cboSecurity") then
		secList = secList & "onChange='FetchAccounts(this);UpdateClientHoldings(this);' "
		else
		secList = secList & "onChange='FetchAccounts1(this);UpdateClientHoldings(this);' "
		end if
		secList = secList & "onKeypress='return (dodefaultaction()==\""\""); ' "  
		secList = secList & "onKeydown='return (dodefaultaction()==\""\"");event.cancelBubble=true;' "  
		secList = secList & "onKeyup='return (change(" & listName & "));' "  
		secList = secList & "onfocus='txtval = \""\"";inputIsItemCode = 1;' "  
		secList = secList & "onblur='txtval = \""\"";inputIsItemCode = 1;'>"		
		
		secList = secList & "<option selected SearchCode = '0' SearchText = ''  value = ''></option>"
		    
        
        sqlStr = "SELECT * FROM [SecurityList] WHERE OrderSecType_DPA_ = " & secType & " ORDER BY SecurityName"
        Set rsSecurity = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rsSecurity.EOF Or rsSecurity.BOF) Then
                rsSecurity.MoveFirst
                Do Until rsSecurity.EOF
                        secList = secList & "<option SearchCode = '" & rsSecurity.Fields("SecurityCode") & "' SearchText = '" & rsSecurity.Fields("SecurityName") & "'  value = '" & rsSecurity.Fields("Security_DPA_") & "'>" & rsSecurity.Fields("SecurityCode") & "</option>"
                        rsSecurity.MoveNext
                Loop
        End If
	    secList = secList & "</select>"
	    GetSecurityList = secList
  End function

  function GetBondList(listName)
		Dim secList
		Dim rsSecurity
		
		secList = "<select name = '" & listName & "' id = '" & listName & "' size='1' "
		secList = secList & "OnClick='event.cancelBubble=true;' "  
		secList = secList & "onChange='event.cancelBubble=true;' " 
		secList = secList & "onKeypress='return (dodefaultaction()==\""\""); ' "  
		secList = secList & "onKeydown='return (dodefaultaction()==\""\"");event.cancelBubble=true;' "  
		secList = secList & "onKeyup='return (change(" & listName & "));' "  
		secList = secList & "onfocus='txtval = \""\"";inputIsItemCode = 1;' "  
		secList = secList & "onblur='txtval = \""\"";inputIsItemCode = 1;'>"
		
		secList = secList & "<option selected SearchCode = '0' SearchText = ''  value = ''></option>"		
		sqlStr = "SELECT BondIssue AS Issue,Security_DPA_,SecurityCode,Bond_DPA_ FROM [IssueList] "
    
        Set rsSecurity = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rsSecurity.EOF Or rsSecurity.BOF) Then
                rsSecurity.MoveFirst
                Do Until rsSecurity.EOF
                        secList = secList & "<option SearchCode = '" & rsSecurity.Fields("SecurityCode") & "' SearchText = '" & rsSecurity.Fields("Issue") & "'  value = '" & rsSecurity.Fields("Bond_DPA_") & "'>" & rsSecurity.Fields("Issue") & "</option>"
                        rsSecurity.MoveNext
                Loop
        End If
	    
        secList = secList & "</select>"
	    GetBondList = secList
  end function
 %>
 <table border="0" width="100%" ID="Table1">
<tr><td>
<input type = 'hidden' name ='ItemID' id = 'ItemID'>
<input type = 'hidden' name ='ID' id = 'ID' value="<%= ID %>">
<input type = 'hidden' name ='txtcalendar1' id = 'txtcalendar1' value='0'>
<input type = 'hidden' name ='txtbest' id = 'txtbest' value='0'>
<input type = 'hidden' name ='txtbestinplace' id = 'txtbestinplace' value='0'>
<input type = 'hidden' name ='Ordertype' id = 'Ordertype' value='<%=Ordertype%>'>
<input type = 'hidden' name ='sectype' id = 'sectype' value='<%=secType%>'>
<input type = 'hidden' name ='action' id = 'action' value="Execute_Detail">
</td>
</tr>
</table>
</form>
</body>

</html>