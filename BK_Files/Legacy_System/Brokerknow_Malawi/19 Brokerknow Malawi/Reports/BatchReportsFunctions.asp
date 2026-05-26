<%
''CONTRACTS SCHEDULE
Sub ContractScheduleGenerate()
	%>
	<form id="repToPDF0" name="repToPDF0">
	<%
	'response.write "CONTRACT SCHEDULE"
	'headerDescription = FormatDateFull(selectedContractDate)
	%>
	<!--<i id="landRem">Remember to select landscape settings while printing.</i>-->

	<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
	     <tr>
			<td nowrap><b><font face="Arial Narrow" size="4">Contracts Schedule</font></b></td>
			<td nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>
		</tr>	
	    <tr>
		   <td COLSPAN=2><font face="Arial" size="2">for Deals traded on:  <%= headerDescription %></font></td>
		</tr>
	    <tr>
			  <td COLSPAN=2><font face="Arial" size="2">&nbsp;</font></td>
		</tr>
	</table>				

	<table border="0" width="100%" cellPadding="2" cellSpacing=0>
		<tr bgColor="#000000">
			<td bgColor="#000000" nowrap align="center"><b><font color="#FFFFFF">Traded</font></b></td>
			<td bgColor="#000000" nowrap align="center"><b><font color="#FFFFFF">Client</font></b></td>
			<td bgColor="#000000" nowrap align="center"><b><font color="#FFFFFF">Security</font></b></td>
			<td bgColor="#000000" nowrap align="center"><b><font color="#FFFFFF">Br</font></b></td>
			<td bgColor="#000000" nowrap align="center"><b><font color="#FFFFFF">Contr</font></b></td>
			<td bgColor="#000000" nowrap align="center"><b><font color="#FFFFFF">CDS Ref</font></b></td>
			<td bgColor="#000000" nowrap align="right"><b><font color="#FFFFFF">Price</font></b></td>
			<td bgColor="#000000" nowrap align="right"><b><font color="#FFFFFF">Quantity</font></b></td>
			<td bgColor="#000000" nowrap>&nbsp;</td>
			<%
			Dim fld
			Dim conn 
			Dim sqlStr
			Dim rs
			Dim i
			Dim dailyTotalsArray1()
			
			beginLeviesCol1=15	
			
			Set conn = GetActiveConnection("KBroker")
	 		sqlStr = "ContractLeviesCrossTab"	
			Set Rs = CreateObject("ADODB.Recordset")		
			Rs.CursorLocation = adUseClient		
			Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
			Rs.Filter = "LotTDate = '" & FormatDate(selectedContractDate) & "'"
			
			'i = 0
			'fldCount = fldCount + 2 'this is hard coding just to skip some unwanted columns 
			for i = beginLeviesCol1 to rs.fields.count - 2
				Redim Preserve dailyTotalsArray1(i - beginLeviesCol1)
				dailyTotalsArray1(i - beginLeviesCol1) = 0
				%>
					<td bgColor="#000000" nowrap  align="center"><b><font color="#FFFFFF"><%=rs.fields(i).name%></font></b></td>		
				<%
				'i = i + 1
			next
			
			'add gross, and net amount
			Redim Preserve dailyTotalsArray1((i) - beginLeviesCol1)
			dailyTotalsArray1((i) - beginLeviesCol1) = 0
			
			Redim Preserve dailyTotalsArray1((i + 1) - beginLeviesCol1)
			dailyTotalsArray1((i + 1) - beginLeviesCol1) = 0
			
			'Redim Preserve dailyTotalsArray1((i + 2) - beginLeviesCol1)
			'dailyTotalsArray1((i + 2) - beginLeviesCol1) = 0
			%>
			<td bgColor="#000000" nowrap align="center"><b><font color="#FFFFFF">&nbsp;&nbsp;Gross&nbsp;&nbsp;</font></b></td>
			<td bgColor="#000000" nowrap align="center"><b><font color="#FFFFFF">Net Amount</font></b></td>
	    </tr>
		<%
		If Not(rs.EOF Or rs.BOF) Then
			'rs.MoveFirst
			Do Until rs.EOF
					%>
		    		<tr>
						<td nowrap><font size="1"><%= Day(rs.Fields("LotTDate")) & " " & MonthName(Month(rs.Fields("LotTDate")), True) %></font></td>
						<td nowrap><font size="1">
							<%
							If Len(rs.Fields("ClientName")) > 25 Then 
								Response.Write Mid(rs.Fields("ClientName"), 1, 25)
							Else
								Response.Write rs.Fields("ClientName")	
							End If
							%>
						</font></td>
						<td nowrap><font size="1"><%=rs.Fields("SecurityCode")%></font></td>
						<td nowrap align="center"><font size="1"><%=rs.Fields("BrokerCode")%></font></td>
						<td nowrap><font size="1"><%=rs.Fields("ContractNumber")%></font></td>
						<td nowrap><font size="1"><%=rs.Fields("LotSlipNo")%></font></td>
						<td nowrap align="right"><font size="1"><%=FormatNum(rs.Fields("LotPrice"))%></font></td>
						<td nowrap align="right"><font size="1"><%=FormatNumCommasOnly(rs.Fields("LotQty"))%></font></td>       
						<td style="BORDER-RIGHT: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent">&nbsp;</td>
						<%	
						levyTotals = 0
					
						for i = beginLeviesCol1 to rs.fields.count - 2
							If i = rs.Fields.Count - 2 Then
								myStyle =  "BORDER-RIGHT: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent"
							Else
								myStyle = ""	
							End If
							%>
								<td  nowrap align=right  Style="<%= myStyle %>"><font size="1"><%=FormatNum(rs.fields(i).value)%></font></td>		
							<%
							'i = i + 1						
							levyTotals = levyTotals + rs.fields(i).value					
							dailyTotalsArray1(i - beginLeviesCol1) = dailyTotalsArray1(i - beginLeviesCol1) + rs.fields(i).value
						next
					
						'check whether this contract has the
						'levy of agent commission type, which is not
						'really a levy, but a rate of "broker commission" levy
						'if so, minus this value from the levyTotals variable
						sqlStr = "SELECT * FROM LevyContract WHERE Contract_DPA_ = " & rs.Fields("Contract_DPA_").Value & " AND SystemMaintained = 12"  
						Set tmpRs = Conn.Execute (SQLServerFormat(HandleQuote(sqlStr)))
						If Not (tmpRs.EOF OR tmpRs.BOF) Then
							'agent commission exists
							levyTotals = levyTotals - tmpRs.Fields("LevyAmount").Value
						End If
						Set tmpRs = Nothing		
					
						grossAmt = rs.Fields("LotGrossAmount") 'rs.Fields("LotPrice") * rs.Fields("LotQty") 
					
						If rs.Fields("OrderTypeSale").Value = 0 Then 
							netAmt = grossAmt + levyTotals 
						Else
							netAmt = grossAmt - levyTotals 
						End If
					
						dailyTotalsArray1((i) - beginLeviesCol1) = dailyTotalsArray1((i) - beginLeviesCol1) + grossAmt
						dailyTotalsArray1((i + 1) - beginLeviesCol1) = dailyTotalsArray1((i + 1) - beginLeviesCol1) + netAmt
						%>
		         		<td nowrap align="right"><font size="1"><%= FormatNum(grossAmt) %></font></td>       
						<td nowrap align="right"><font size="1"><%= FormatNum(netAmt) %></font></td> 
					</tr>      
					<%
		         rs.MoveNext
			Loop
		
		    %>
		    
		    <tr>
				<td colspan="<%= 9 + UBound(dailyTotalsArray1)%>">&nbsp;</td>
		    </tr>
		  
			<tr height="30px">
				<td colspan=9 align=right><b>Daily totals:</b></td>	
				<%
				for i = 0 to UBound(dailyTotalsArray1)
					If i = 0 Then 
						myStyle = "BORDER-LEFT: #C0C0C0 1px inset; BORDER-TOP: #C0C0C0 1px inset; BORDER-BOTTOM: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent; "
					ElseIf (i <> 0 And i <> UBound(dailyTotalsArray1)) Then
						myStyle = "BORDER-TOP: #C0C0C0 1px inset; BORDER-BOTTOM: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent;"
					Else
						myStyle = "BORDER-RIGHT: #C0C0C0 1px inset;BORDER-TOP: #C0C0C0 1px inset; BORDER-BOTTOM: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent;"
					End If
					%>
					<td  nowrap align=right style="<%= myStyle %>"><font size="1"><%= FormatNum(Trim(dailyTotalsArray1(i))) %></font></td>		
					<%
				next
				%>
			</tr>
		<%End If%>
	</table>
	</form>
	<%
End Sub

''CONTRACTS
Sub ContractsGenerate()
	%>
	<br class="newpage">
	<form id="repToPDF1" name="repToPDF1">
	<%
	Set conn = GetActiveConnection("KBroker")
	Set Rs = CreateObject("ADODB.Recordset")						        
	sqlStr = "SELECT * FROM SaleContracts WHERE LotTDate = '" & FormatDate(selectedContractDate) & "' ORDER BY ContractNumber"		
		
	Rs.CursorLocation = adUseClient	
	Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
		
	If Not rs.EOF Or rs.BOF Then
		Set groupRs = Conn.Execute("SELECT ContractNumber FROM SaleContracts WHERE LotTDate = '" & FormatDate(selectedContractDate) & "' GROUP BY ContractNumber")
		Set levyOrderRs = Conn.Execute("SELECT * FROM LevyOrderList ORDER BY LevyOrder")
			
		pageNumber = 0

		Do Until groupRs.EOF
			Rs.Filter = "ContractNumber = '" & groupRs.Fields("ContractNumber").Value & "'"
				
			totalGross = Rs.Fields("LotGrossAmount").Value
				
			sqlStr = "SELECT tbOrder.*, Client.ClientAddr FROM Lot INNER JOIN " & _
				" OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN " & _
				" tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN " & _
				" Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ " & _
				" WHERE Lot.Lot_DPA_ = " & Rs.Fields("Lot_DPA_").Value
			Set temprs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
				
			If Not (temprs.EOF Or temprs.BOF) Then
				orderRef = temprs.Fields("orderRef").Value
				if isnull(orderRef) or trim(orderRef) = "" then
					orderRef = temprs.Fields("Order_DPA_").Value
				else
					orderRef = temprs.Fields("Order_DPA_").Value & "/" & orderRef
				end if
				clientAddress = Replace(temprs.Fields("ClientAddr").Value, Chr(13), ",")		
			End If
				
			Set temprs = Nothing
				
			If Trim(UCase(Rs.Fields("OrdDetailType").Value)) = "PURCHASE" Then
				IsPurchase = True
			Else
				IsPurchase = False
			End If
				
			pageNumber = pageNumber + 1
			
			If Not pageNumber = 1 Then
				%>
				<br class="newpage">
				<%
			End If
			
			%>
			<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" align ="center" width="80%"> 
				<tr class="pageNumbering">
					<td align="left">
						&nbsp;<!--<FONT FACE=ARIAL SIZE=2><B>Page <%=pageNumber%>	</B></FONT>	-->
					</td>		
				</tr>
				<tr>
					<td align="center">
						<Img Src="../data/photos/aaprintlogo.jpg">			
					</td>		
				</tr>
				<tr>
					<td align="center">
						<FONT FACE=ARIAL SIZE=3><B>
							<%= UCase(Rs.Fields("OrdDetailType").Value) %> CONTRACT <%= Rs.Fields("ContractNumber").Value %>
						</B></FONT>
					</td>		
				</tr>  
				<tr>
					<td colspan=2>&nbsp;</td>
				</tr>
			</table>

			<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" align ="center" width="80%"> 
				<tr>
					<td nowrap valign="top" width="70%">			
						<table cellspacing=0 cellpadding=0 border=0>
							<tr>
								<td nowrap><%= Rs.Fields("OrdDetailClient").Value %> </td>					
							</tr>
							<tr>					
								<td nowrap><%= clientAddress %> </td>
							</tr>
						</table>
				   	</td>
					<td valign="top" align="center" width="30%">
						<table cellspacing=0 cellpadding=0 border=0>
							<tr>
								<td nowrap>Date:&nbsp;<%= FormatDate(Rs.Fields("LotTDate").Value) %></td>
							</tr>
							<tr>
								<td nowrap>Order Ref:&nbsp;<%= orderRef %></td>					
							</tr>
							<tr>
								<td nowrap>CDS Ref:&nbsp;<%= Rs.Fields("LotSlipNo").Value %></td>					
							</tr>
						</table>
					</td>
			    </tr>
				<tr>
					<td colspan=2>Dear Sir/Madam</td>
				</tr>
				<tr>
					<td colspan=2>&nbsp;</td>
				</tr>
				<tr>
					<td colspan=2><U><b>RE: <%= UCase(Rs.Fields("OrdDetailType").Value) %> OF SECURITIES</b></U></td>
				</tr>
				<tr>
					<td colspan=2>&nbsp;</td>
				</tr>
				<tr>
					<td colspan=2><font face=Arial size=2>We wish to advise that we have <% If IsPurchase Then Response.Write "bought" Else Response.Write "sold"  %> the following securities as per your instruction. Kindly arrange to complete the transaction as per the details given below: </font></td>
				</tr>
				<tr>
					<td colspan=2 style="BORDER-BOTTOM: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent">&nbsp;</td>
				</tr>
				<!--<tr>
					<td colspan=2>&nbsp;</td>
				</tr>-->
			</table>	

			<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" align ="center" width="80%">	
				<tr>
					<td nowrap width="30%">Security</td>
					<td nowrap width="70%" align=right><%= Rs.Fields("OrdDetailSecurity").Value %></td>
				</tr>
				<tr>
					<td nowrap width="30%"><%if lcase(rs.Fields("OrdDetailSecType")) = "fixed" Then%>
								Face Value
						<%else%>
								Quantity
						<%end if%></td>
					<td nowrap width="70%" align=right> <%= FormatNum(Rs.Fields("LotQty").Value) %> </td>
				</tr>
				<tr>
					<td nowrap width="30%">Price</td>
					<td nowrap width="70%" align=right>
					<%if lcase(rs.Fields("OrdDetailSecType")) = "fixed" Then%>
							<%= FormatNumEx(Rs.Fields("LotPrice").Value,4) %>
					<%else%>
							<%= FormatNum(Rs.Fields("LotPrice").Value) %>
					<%end if%>
					  </td>
				</tr>
			    <tr>
					<td nowrap width="30%">Gross</td>
					<td nowrap width="70%" align=right><b> <%= FormatNum(totalGross) %> </b> </td>
				</tr>
				<tr>
					<TD nowrap width="100%" COLSPAN=2 VALIGN="TOP">
						<TABLE align ="center" align ="center" width="100%" CELLSPACING="0" CELLPADDING="4">
								<%
								totalLevies = 0
								totalContractStamps = 0
								levyArray = SortLevies(rs, levyOrderRs)
								transferFeeVal = 0
								contractStampsVal = 0
								
								For i = 1 To UBound(levyArray)
									thisLevyName = Trim(levyArray(i, 0))
									If levyArray(i, 2) = 10 Then
										totalContractStamps = levyArray(i, 1)
										totalContractStampsName = levyArray(i, 0)
									End If
									
									'ignore the below line as the stored levy percentage
									'is not very reliable....
									'levyPerc = levyArray(i, 3)
												
									'recalculate percentage (really getting tired with this...)
									levyPerc = (levyArray(i, 1) / totalGross) * 100
									levyPerc = FormatNum(levyPerc) & "%"
									
									'do not display agent commission (staff commission too) HARD CODING!!!!!!
									If (levyArray(i, 2) = 12) or (levyArray(i, 2) = 8) Then
										thisLevyName = ""
									End If
									
									If  levyArray(i, 2) = 13  Then				
										transferFeeDesc = thisLevyName
										transferFeeVal = levyArray(i, 1)
										thisLevyName = ""
									End If
									
									If levyArray(i, 2) = 10 Then
										contractStampsDesc = thisLevyName
										contractStampsVal = levyArray(i, 1)
										thisLevyName = ""
									End If
									
									If thisLevyName <> "" Then
										totalLevies = totalLevies + levyArray(i, 1) 
										%>
										<TR>		
											<td nowrap width="10%" STYLE="PADDING-LEFT: 0PX"><%= thisLevyName %></td>
											<td ALIGN="LEFT">&nbsp;&nbsp;<%= levyPerc %></td>
											<td align=right><%= FormatNum(levyArray(i, 1)) %></td>		
										</TR>
										<%	
									End If
								Next
								%>
						</TABLE>	
					</TD>
				</tr>
				
				<tr>
					<td nowrap width="30%">&nbsp;</td>
					<td nowrap width="70%" align=right>
						<hr width=70px height=1px>
						<%= Replace(Space(Len(FormatNum(totalLevies)) * 4), Space(1), "&nbsp;") %>
					</td>
				</tr>
				<tr>
					<td nowrap width="30%">Total Commission and Levies</td>
					<td nowrap width="70%" align=right>	
						<%If IsPurchase Then%>
							<b><%= FormatNum(totalLevies) %></b>
						<%Else%>	
							<b>(<%= FormatNum(totalLevies) %>)</b>
						<%End If%>
					</td>
				</tr>
				<%if trim(transferFeeDesc) <> "" then%>
						<tr>
							<td nowrap width="30%"><%= transferFeeDesc %></td>
							<td nowrap width="70%" align=right>	
								 <%= FormatNum(transferFeeVal) %>		
							</td>
						</tr>
				<%end if%>
				
				<%if trim(contractStampsDesc) <> "" then%>
						<tr>
							<td nowrap width="30%"><%= contractStampsDesc %></td>
							<td nowrap width="70%" align=right>	
								 <%= FormatNum(contractStampsVal) %>		
							</td>
						</tr>
				<%end if%>
				<tr>
					<td nowrap width="30%">&nbsp;</td>
					<td nowrap width="70%" align=right> 
						<%= Replace(Space(Len(FormatNum(totalLevies)) * 4), Space(1), "&nbsp;") %>
						<hr width=70px height=1px>
					</td>
				</tr>
				<tr>
					<td nowrap width="30%">TOTAL AMOUNT PAYABLE IN KSHS</td>
					<td nowrap width="70%" align=right><b>
						<%
						totalLevies = totalLevies + transferFeeVal + contractStampsVal
						If IsPurchase Then
							Response.Write FormatNum(totalGross + totalLevies) 
						Else
							Response.Write FormatNum(totalGross - totalLevies) 
						End If%>
						
						</b></td>
				</tr>
				<tr>
					<td nowrap width="100%" colspan=2>&nbsp;</td>
				</tr>
				<tr>
					<td nowrap width="100%" colspan=2>&nbsp;</td>
				</tr>
				<tr>
					<td nowrap width="100%" colspan=2>&nbsp;</td>
				</tr>
				<tr>
					<td nowrap width="100%">
					<font face=Arial size=2>
			       <i>For and on behalf of</i>
			African Alliance Uganda Securities<br><br>

			Sign..........................................
					</FONT>			
					</td>
					<td nowrap width="70%" align="left" valign="top">			
						<img src="../images/stamp.gif" border="0" style="position: absolute;z-index: 2">
						<BR>
						&nbsp;<SPAN style="position: absolute;z-Index: 10"><b>KShs&nbsp;&nbsp;<%= FormatNum(totalContractStamps) %></b></SPAN>
						<br><br><br>
						Revenue Stamps prepaid
						
					</td>
				</tr>
			</table>
			<%
			Rs.Cancel
			groupRs.MoveNext
		Loop
	End If
	%></form><%
End Sub

''CONTRACTS COMPOUNDED
Sub ContractsCompoundedGenerate()
	%>
	<form id="repToPDF2" name="repToPDF2">
	<%
	Set conn = GetActiveConnection("KBroker")
	Set Rs = CreateObject("ADODB.Recordset")	
	Set groupRs = CreateObject("ADODB.Recordset")	
	
	sqlStr = "SELECT * FROM ClientContractCompounded WHERE (LotTDate = '" & FormatDate(selectedContractDate) & "') ORDER BY LotSlipNo"
	groupRs.Open sqlStr, conn.ConnectionString, 0, 1

	If Not groupRs.EOF Or groupRs.BOF Then
		Set OrderRs = CreateObject("ADODB.Recordset")						        
		sqlStr = "SELECT DISTINCT Order_DPA_ FROM ClientContractCompounded WHERE (LotTDate = '" & FormatDate(selectedContractDate) & "') ORDER BY Order_DPA_"
	
		OrderRs.CursorLocation = adUseClient	
		OrderRs.Open sqlStr, conn.ConnectionString, adOpenKeyset, adLockOptimistic
	
		Set levyOrderRs = Conn.Execute("SELECT * FROM LevyOrderList ORDER BY LevyOrder")
	
		blnEndFound = False
	
		pageNumber = 0
		
		Do Until OrderRs.EOF
			groupRs.Filter = "Order_DPA_ = '" & OrderRs.Fields("Order_DPA_").Value & "'"
				
			prevSecurity = groupRs.Fields("OrdDetailSecurity").Value
			prevDate = groupRs.Fields("LotTDate").Value
			prevOrderTypeSale = groupRs.Fields("OrderTypeSale").Value
			slipNos = groupRs.Fields("LotSlipNo").Value
			prevSlipNo = slipNos
			prevOrderNo = groupRs.Fields("Order_DPA_").Value
			prevClient = groupRs.Fields("OrdDetailClient").Value
			
			Do Until groupRs.EOF
				If prevSlipNo <> groupRs.Fields("LotSlipNo").Value  Then
					If slipNos = "" Then 
						slipNos = groupRs.Fields("LotSlipNo").Value
					Else 
						slipNos = slipNos & ", " & groupRs.Fields("LotSlipNo").Value
					End If
					prevSlipNo = groupRs.Fields("LotSlipNo").Value					
				End If
				groupRs.MoveNext
			loop

			Set Rs = CreateObject("ADODB.Recordset")						        
			'the number 13 in the SQL below represents 'Transfer Fee'					        
			sqlStr = "SELECT     Order_DPA_, OrdDetailSecType, LevyName, LevyShortName, CASE SystemMaintained WHEN 13 THEN MAX(LevyAmount) ELSE SUM(LevyAmount) " & _
			    "END AS LevyAmount, SUM(GrossAmount) AS GrossAmount, MIN(CAST(ClientAddr AS nvarchar(4000))) AS ClientAddr, " & _	
				"MIN(OrderRef) AS orderRef, MIN(CAST(OrderTypeSale AS INT)) AS OrderTypeSale, MIN(OrdDetailType) AS OrdDetailType, MIN(ContractNumber) " & _
				"AS ContractNumber, OrdDetailSecurity, SUM(LotQty) AS LotQty, SUM(LotPrice) AS LotPrice, OrdDetailClient, SystemMaintained, LevyRatePercentage " & _
				"FROM         dbo.ClientContractCompounded " & _
				"WHERE     (LotSlipNo IN (" & slipNos  & ")) And LevyName <> '' AND OrdDetailClient = '" & prevClient & "' AND (Order_DPA_ = " & OrderRs.Fields("Order_DPA_").Value & ")" & _
				"GROUP BY LevyName, LevyShortName, OrdDetailSecurity, OrdDetailClient, SystemMaintained, LevyRatePercentage,OrdDetailSecType,Order_DPA_"
							
			Rs.CursorLocation = adUseClient	
		
			on error resume next
			
			Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
							
			totalGross = Rs.Fields("GrossAmount").Value 
			orderRef = Rs.Fields("orderRef").Value
			if isnull(orderRef) or trim(orderRef) = "" then
				orderRef = Rs.Fields("Order_DPA_").Value
			else
				orderRef = Rs.Fields("Order_DPA_").Value & "/" & orderRef
			end if
			clientAddress = Replace(Rs.Fields("ClientAddr").Value, Chr(13), ",")		
						
			If Trim(UCase(Rs.Fields("OrdDetailType").Value)) = "PURCHASE" Then
				IsPurchase = True
			Else
				IsPurchase = False
			End If  
				
			pageNumber = pageNumber + 1
			%>
			<br class="newpage">
			<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" align ="center" width="80%">
				<THEAD>
				<tr class="pageNumbering">
					<td align="left">
						&nbsp;<!--<FONT FACE=ARIAL SIZE=2><B>Page <%=pageNumber%>	</B></FONT>	-->
					</td>		
				</tr>
				<tr>
					<td align="center">
						<Img Src="../data/photos/aaprintlogo.jpg">			
					</td>		
				</tr>
				<THEAD>  
				<tr>
					<td align="center">
						<FONT FACE=ARIAL SIZE=3><B>
							<%= UCase(Rs.Fields("OrdDetailType").Value) %> CONTRACT <%= Rs.Fields("ContractNumber").Value %>
						</B></FONT>
					</td>
								
				</tr>  
				<!--<tr>
					<td colspan=2>&nbsp;</td>
				</tr>-->	
			</table>

			<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" align ="center" width="80%"> 
				<tr>
					<td  nowrap valign="top" width="70%">			
						<table cellspacing=0 cellpadding=0 border=0>
							<tr>
								<td nowrap><%= Rs.Fields("OrdDetailClient").Value %> </td>					
							</tr>
							<tr>					
								<td nowrap><%= clientAddress %> </td>
							</tr>
						</table>
					</td>
					
					<td valign="top" align="center" width="30%">
						<table cellspacing="0" cellpadding="0" border="0">
							<tr>
								<td  nowrap colspan="2">Date:&nbsp;<%= FormatDate(selectedContractDate) %></td>
							</tr>
							<tr>
								<td  colspan="2" nowrap>Order Ref:&nbsp;<%= orderRef %></td>					
							</tr>
							
						</table>
					</td>
				</tr>
				<tr>
				<tr>
					<td colspan=2>
					      <table cellspacing="0" cellpadding="0" border="0"  align ="center" width="100%">
							<tr>
							<td nowrap valign="top" width="10%">CDS Ref:&nbsp;</td>
							<td valign="top" width="60%">
									<table border="0" cellpadding="1" cellspacing="0" align ="center" width="100%">
													
												<%
												If InStr(1, slipNos, ",") > 0 Then
													maxSlipNosPerRow = 11
													slipNosArray = Split(slipNos, ",")
													maxSlipCount = UBound(slipNosArray) 
																
													For k = 0 To maxSlipCount	%>
														<tr><td nowrap>
														<%For l = 1 To maxSlipNosPerRow
															If k <= maxSlipCount Then											
																thisSlipNo = Trim(slipNosArray(k))
																If k < maxSlipCount Then
																	thisSlipNo = thisSlipNo & ", "
																End If																						
																Response.Write thisSlipNo											
																If l <> maxSlipNosPerRow Then
																	k = k + 1
																End If	
															Else
																Exit For
															End If																																									
														Next%>									
														</td></tr>
													<%Next
															
												Else%>
												<tr><td nowrap><%= slipNos %></td></tr>
												<%End If%>	
									</table>	
							</td>
			               </tr>
						</table>
					</td>
				</tr>
				<tr>
					<td colspan=2>&nbsp;</td>
				</tr>
				<tr>
					<td colspan=2>Dear Sir/Madam</td>
				</tr>
				<tr>
					<td colspan=2>&nbsp;</td>
				</tr>
				<tr>
					<td colspan=2><U><b>RE: <%= UCase(Rs.Fields("OrdDetailType").Value) %> OF SECURITIES</b></U></td>
				</tr>
				<tr>
					<td colspan=2>&nbsp;</td>
				</tr>
				<tr>
					<td colspan=2><font face=Arial size=2>We wish to advise that we have <% If IsPurchase Then Response.Write "bought" Else Response.Write "sold"  %> the following securities as per your instruction. Kindly arrange to complete the transaction as per the details given below: </font></td>
				</tr>
				
				<tr>
					<td colspan=2>&nbsp;</td>
				</tr>
			</table>	

			<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable"  align ="center" width="80%">	
				<tr>
					<td>Security</td>
					<td align=right><%= Rs.Fields("OrdDetailSecurity").Value %></td>
				</tr>
				<tr>
					<td>
						<%if lcase(rs.Fields("OrdDetailSecType")) = "fixed" Then%>
								Face Value
						<%else%>
								Quantity
						<%end if%>
					</td>
					<td align=right> <%= FormatNum(Rs.Fields("LotQty").Value) %> </td>
				</tr>
				<tr>
					<td>Average Price</td>
					<td align=right> 
						<%if lcase(rs.Fields("OrdDetailSecType")) = "fixed" Then%>
								<%= FormatNumEx((totalGross / Rs.Fields("LotQty").Value) * 100,4) %>
						<%else%>
								<%= FormatNum(totalGross / Rs.Fields("LotQty").Value) %>
						<%end if%>
				 </td>
				</tr>
			    <tr>
					<td>Gross</td>
					<td align=right><b> <%= FormatNum(totalGross) %> </b> </td>
				</tr>
				<tr>
					<TD COLSPAN=2 VALIGN="TOP">
						<TABLE align ="center" width="100%"  CELLSPACING="0" CELLPADDING="4">
								<%
								totalLevies = 0
								totalContractStamps = 0
								levyArray = SortLevies(rs, levyOrderRs)
								transferFeeVal = 0
								contractStampsVal = 0
								'careful, returned array is one-based
								For i = 1 To UBound(levyArray) 
												 
									thisLevyName = Trim(levyArray(i, 0))
									If levyArray(i, 2) = 10 Then
										totalContractStamps = levyArray(i, 1)
									End If
												
									levyPerc = levyArray(i, 3)
									If InStr(1, levyPerc, "%") > 0 Then
										'recalculate average percentage
										levyPerc = (levyArray(i, 1) / totalGross) * 100
										levyPerc = FormatNum(levyPerc) & "%"
									End If
												
									'skip agent commission info
									If levyArray(i, 2) = 12 Then
										thisLevyName = ""
									End If
												
									'transfer fee
									If levyArray(i, 2) = 13  Then				
										transferFeeDesc = thisLevyName
										transferFeeVal = levyArray(i, 1)
										thisLevyName = ""
									End If
												
									If levyArray(i, 2) = 10 Then
										contractStampsDesc = thisLevyName
										contractStampsVal = levyArray(i, 1)
										thisLevyName = ""
									End If
												
												
									If thisLevyName <> "" Then
										totalLevies = totalLevies + levyArray(i, 1)%>
										<TR>		
											<td nowrap width="10%" STYLE="PADDING-LEFT: 0PX"><%= thisLevyname %></td>
											<td ALIGN="LEFT">&nbsp;&nbsp;<%= levyPerc  %></td>
											<td align=right><%= FormatNum(levyArray(i, 1)) %></td>		
										</TR>
								<% End If
								Next%>
						</TABLE>	
					</TD>
				<tr>
					<td>&nbsp;</td>
					<td align=right>
						<hr width=70px height=1px>
						<%= Replace(Space(Len(FormatNum(totalLevies)) * 4), Space(1), "&nbsp;") %>
					</td>
				</tr>
				<tr>
					<td>Total Commission and Levies</td>
					<td align=right>	
						<%If IsPurchase Then%>
							<b><%= FormatNum(totalLevies) %></b>
						<%Else%>	
							<b>(<%= FormatNum(totalLevies) %>)</b>
						<%End If%>
					</td>
				</tr>
				<%if trim(transferFeeDesc) <> "" then%>
						<tr>
							<td><%= transferFeeDesc %></td>
							<td align=right>	
								 <%= FormatNum(transferFeeVal) %>		
							</td>
						</tr>
				<%end if%>
				
				<%if trim(contractStampsDesc) <> "" then%>
						<tr>
							<td><%= contractStampsDesc %></td>
							<td align=right>	
								 <%= FormatNum(contractStampsVal) %>		
							</td>
						</tr>
				<%end if%>
				<tr>
					<td>&nbsp;</td>
					<td align=right> 
						<%= Replace(Space(Len(FormatNum(totalLevies)) * 4), Space(1), "&nbsp;") %>
						<hr width=70px height=1px>
					</td>
				</tr>
				
				<tr>
					<td>TOTAL AMOUNT PAYABLE IN KSHS</td>
					<td align=right>
						<b>
						<%
						totalLevies = totalLevies + transferFeeVal + contractStampsVal
						If IsPurchase Then
							Response.Write FormatNum(totalGross + totalLevies) 
						Else
							Response.Write FormatNum(totalGross - totalLevies) 
						End If%>
									
						</b></td>
				</tr>
				<!--<tr>
					<td colspan=2>&nbsp;</td>
				</tr>
				<tr>
					<td colspan=2>&nbsp;</td>
				</tr>-->
				<tr>
					<td colspan=2>&nbsp;</td>
				</tr>
				<tr>
					<td>
					<font face=Arial size=2>
			       <i>For and on behalf of</i>
			African Alliance Uganda Securities<br><br>

			Sign..........................................
					</FONT>	
					</td>
					<td align="left" valign="top">			
								
						<img src="../images/stamp.gif" border="0" style="position: absolute;z-index: 2">
						<BR>
						&nbsp;<SPAN style="position: absolute;z-Index: 10"><b>KShs&nbsp;&nbsp;<%= FormatNum(totalContractStamps) %></b></SPAN>
						<br><br><br>
						Revenue Stamps prepaid
									
					</td>
				</tr>
			</table>
			<%			
			slipNos = ""
			OrderRs.MoveNext
		Loop
	End If
	%></form><%
End Sub

''AGENT CONTRACTS
Sub AgentContractsGenerate()
	%>
	<form id="repToPDF3" name="repToPDF3">
	<%
	Set conn = GetActiveConnection("KBroker")
	Set Rs = CreateObject("ADODB.Recordset")	
	Set groupRs = CreateObject("ADODB.Recordset")
	
	sqlStr = "SELECT * FROM AgentContracts WHERE LotTDate = '" & FormatDate(selectedContractDate) & "' ORDER BY ContractNumber"		
		
	Rs.CursorLocation = adUseClient	
	Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
	
	If Not rs.EOF Or rs.BOF Then
		Set groupRs = Conn.Execute("SELECT ContractNumber FROM AgentContracts WHERE LotTDate = '" & FormatDate(selectedContractDate) & "' GROUP BY ContractNumber ")
		
		Set levyOrderRs = Conn.Execute("SELECT * FROM LevyOrderList ORDER BY LevyOrder")
		
		pageNumber = 0
		
		Do Until groupRs.EOF
		
		Rs.Filter = "ContractNumber = '" & groupRs.Fields("ContractNumber").Value & "'"
		
		totalGross = Rs.Fields("LotGrossAmount").Value
		
		sqlStr = "SELECT tbOrder.*, Client.ClientAddr FROM Lot INNER JOIN " & _
			" OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN " & _
			" tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN " & _
			" Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ " & _
			" WHERE Lot.Lot_DPA_ = " & Rs.Fields("Lot_DPA_").Value
		Set temprs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		
		If Not (temprs.EOF Or temprs.BOF) Then
			orderRef = temprs.Fields("orderRef").Value
			if isnull(orderRef) or trim(orderRef) = "" then
				orderRef = temprs.Fields("Order_DPA_").Value
			else
				orderRef = temprs.Fields("Order_DPA_").Value & "/" & orderRef
			end if
			clientAddress = Replace(temprs.Fields("ClientAddr").Value, Chr(13), ",")		
		End If
		
		Set temprs = Nothing
		
		If Trim(UCase(Rs.Fields("OrdDetailType").Value)) = "PURCHASE" Then
			IsPurchase = True
		Else
			IsPurchase = False
		End If
		
		pageNumber = pageNumber + 1
		%>
		<br class="newpage">
		<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" align ="center" width="80%"> 
			<THEAD>
			<tr class="pageNumbering" colspan=2>
				<td align="left">
					&nbsp;<!--<FONT FACE=ARIAL SIZE=2><B>Page <%=pageNumber%>	</B></FONT>	-->
				</td>		
			</tr>
			<tr>
				<td align="center" colspan=2>
					<Img Src="../data/photos/aaprintlogo.jpg">			
				</td>		
			</tr>
			<THEAD>  
			<tr>
				<td align="center" colspan=2>
					<FONT FACE=ARIAL SIZE=3><B>
						<%= UCase(Rs.Fields("OrdDetailType").Value) %> CONTRACT <%= Rs.Fields("ContractNumber").Value %>
					</B></FONT>
				</td>
				
			</tr>  
			<!--<tr>
				<td colspan=2>&nbsp;</td>
			</tr>  	-->
			<tr>
				<td nowrap width="100%" valign="top">			
					<table cellspacing=0 cellpadding=0 border=0  align ="center" width="100%">
						<tr>
							<td nowrap><%= Rs.Fields("AgentName").Value %> </td>					
						</tr>
						<tr>					
							<td nowrap><%= Replace(Rs.Fields("AgentAddr").Value, Chr(13), ",") %></td>
						</tr>
						<tr>					
							<td nowrap><b>Account:</b> <%= Rs.Fields("OrdDetailClient").Value %></td>
						</tr>
					</table>
				</td>
				<td valign="top" width="20%" align="right">
					<table cellspacing=0 cellpadding=0 border=0  align ="center" width="100%">
						<tr>
							<td nowrap>Date:&nbsp;<%= FormatDate(Rs.Fields("LotTDate").Value) %></td>
						</tr>
						<tr>
							<td nowrap>Order Ref:&nbsp;<%= orderRef %></td>					
						</tr>
						<tr>
							<td nowrap>CDS Ref:&nbsp;<%= Rs.Fields("LotSlipNo").Value %></td>					
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td colspan=2>Dear Sir/Madam</td>
			</tr>
			<tr>
				<td colspan=2>&nbsp;</td>
			</tr>
			<tr>
				<td colspan=2><U><b>RE: <%= UCase(Rs.Fields("OrdDetailType").Value) %> OF SECURITIES<b></U></td>
			</tr>
			<tr>
				<td colspan=2>&nbsp;</td>
			</tr>
			<tr>
				<td colspan=2><font face=Arial size=2>We wish to advise that we have <% If IsPurchase Then Response.Write "bought" Else Response.Write "sold"  %> the following securities as per your instruction. Kindly arrange to complete the transaction as per the details given below: </font></td>
			</tr>
			
			<tr>
				<td colspan=2>&nbsp;</td>
			</tr>
		</table>	
		<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" align ="center" width="80%">	
			
			<tr>
				<td>Security</td>
				<td align=right><%= Rs.Fields("OrdDetailSecurity").Value %></td>
			</tr>
			<tr>
				<td><%if lcase(rs.Fields("OrdDetailSecType")) = "fixed" Then%>
							Face Value
					<%else%>
							Quantity
					<%end if%></td>
				<td align=right> <%= FormatNum(Rs.Fields("LotQty").Value) %> </td>
			</tr>
			<tr>
				<td>Price</td>
				<td align=right> 
				<%if lcase(rs.Fields("OrdDetailSecType")) = "fixed" Then%>
						<%= FormatNumEx(Rs.Fields("LotPrice").Value,4) %>
				<%else%>
						<%= FormatNum(Rs.Fields("LotPrice").Value) %>
				<%end if%></td>
			</tr>
		    <tr>
				<td>Gross</td>
				<td align=right><b> <%= FormatNum(totalGross) %> </b> </td>
			</tr>
			<tr>
				<TD COLSPAN=2 VALIGN="TOP">
				<TABLE WIDTH="100%" CELLSPACING="0" CELLPADDING="4">
				<%
				totalLevies = 0
				totalContractStamps = 0
				levyArray = SortLevies(rs, levyOrderRs)
				transferFeeVal = 0
				contractStampsVal = 0
				
				
				For i = 1 To UBound(levyArray)			
							
					thisLevyName = Trim(levyArray(i, 0))
					
					
					If levyArray(i, 2) = 10 Then
						totalContractStamps = levyArray(i, 1)
						totalContractStampsName = levyArray(i, 0)
					End If			
					
					
					'ignore the below line as the stored levy percentage
					'is not very reliable....
					'levyPerc = levyArray(i, 3)
					
					'recalculate percentage (really getting tired with this...)
					levyPerc = (levyArray(i, 1) / totalGross) * 100			
					
					MinLevyPerc = ""
					If levyArray(i, 2) = 11 Then
						If FormatNum(levyArray(i, 1)) <= 100 then
							MinLevyPerc = "Minimum"
						End If
					End If
								
					'test = levyArray(i, 1)
					'response.write "test =" & test
					
					'grab agent commission here: to be used laters..
					If levyArray(i, 2) = 12 Then
						AgentCommission = levyArray(i, 1)
						AgentRate = levyPerc 'levyArray(i, 3)
						AgentRateDescription = levyArray(i, 0)
						'empty the levy name variable so as not to display
						thisLevyName = ""
					End If
					
					'do not display staff commission  HARD CODING!!!!!!
					If (levyArray(i, 2) = 8) Then
						thisLevyName = ""
					End If
					
					'grab the broker rate here
					If levyArray(i, 2) = 11 Then
						brokerRate = levyPerc
					End If
					
					'transfer fee
					If levyArray(i, 2) = 13  Then				
						transferFeeDesc = thisLevyName
						transferFeeVal = levyArray(i, 1)
						thisLevyName = ""
					End If
					
					If levyArray(i, 2) = 10 Then
						contractStampsDesc = thisLevyName
						contractStampsVal = levyArray(i, 1)
						thisLevyName = ""
					End If
					
					
					If thisLevyName <> "" Then
						'total up levies only when displayed
						totalLevies = totalLevies + levyArray(i, 1) 
						levyPerc = FormatNum(levyPerc) & "%"
						%>
						<TR>		
							<td nowrap width="10%" STYLE="PADDING-LEFT: 0PX"><%= thisLevyName %></td>
							
							<%If Len(MinLevyPerc) > 0 Then%>
								<td ALIGN="LEFT">&nbsp;&nbsp;<%= MinLevyPerc %></td>
							<%Else%>
								<td ALIGN="LEFT">&nbsp;&nbsp;<%= levyPerc %></td>
							<%End If%>
							
							<td align=right><%= FormatNum(levyArray(i, 1)) %></td>		
						</TR>
						<%	
					End If
				Next
				%>
				</TABLE>	
				</TD>
			</tr>
			
			<tr>
				<td>&nbsp;</td>
				<td align=right>
					<hr width=70px height=1px>
					<%= Replace(Space(Len(FormatNum(totalLevies)) * 4), Space(1), "&nbsp;") %>
				</td>
			</tr>
			<tr>
				<td>Total Commission and Levies</td>
				<td align=right>	
					<%If IsPurchase Then%>
						<b><%= FormatNum(totalLevies) %></b>
					<%Else%>	
						<b>(<%= FormatNum(totalLevies) %>)</b>
					<%End If%>
				</td>
			</tr>
			<%if trim(transferFeeDesc) <> "" then%>
					<tr>
						<td><%= transferFeeDesc %></td>
						<td align=right>	
							 <%= FormatNum(transferFeeVal) %>		
						</td>
					</tr>
			<%end if%>
			
			<%if trim(contractStampsDesc) <> "" then%>
					<tr>
						<td><%= contractStampsDesc %></td>
						<td align=right>	
							 <%= FormatNum(contractStampsVal) %>		
						</td>
					</tr>
			<%end if%>
			<tr>
				<td>&nbsp;</td>
				<td align=right> 
					<%= Replace(Space(Len(FormatNum(totalLevies)) * 4), Space(1), "&nbsp;") %>
					<hr width=70px height=1px>
				</td>
			</tr>
			<tr>
				<td>GROSS AMOUNT</td>
				<td align=right><b>
					<%
					totalLevies = totalLevies + transferFeeVal + contractStampsVal
					If IsPurchase Then
						grossAmount = totalGross + totalLevies				
					Else
						grossAmount = totalGross - totalLevies				
					End If
					
					Response.Write FormatNum(grossAmount) 
					%>
					
					</b></td>
			</tr>
			<tr>	
				<td COLSPAN=2 VALIGN="TOP">
					<TABLE align ="left" width="100%" CELLSPACING="0" CELLPADDING="4">
						<TR>		
							<td nowrap width="10%" STYLE="PADDING-LEFT: 0PX">Returnable Commission</td>
							<td ALIGN="LEFT">&nbsp;&nbsp;<%
							'Trim to nearest whole number. Change By Muchiri
							if brokerRate= 0 then 
								Response.Write "0 %"
							else
								Response.write FormatNum(FormatNumEx(AgentRate * 100/brokerRate,0)) & "%"
							End if	
							 %>
							
							
							</td>
							<td align=right><%= FormatNum(AgentCommission) %></td>		
						</TR>		
				</table>
				</td>	
			</tr>
			<tr>
				<td>NET AMOUNT</td>
				<td align=right><b>
					<%
						If IsPurchase Then
							netAmount = grossAmount - AgentCommission				
						Else
							netAmount = grossAmount + AgentCommission
						End If
					%>
					
					 <%=FormatNum(netAmount) %>
					
					</b></td>
			</tr>
			<!--<tr>
				<td colspan=2>&nbsp;</td>
			</tr>
			<tr>
				<td colspan=2>&nbsp;</td>
			</tr>-->
			<tr>
				<td>
				<font face=Arial size=2>
		       <i>For and on behalf of</i>
		African Alliance Uganda Securities

		Sign..........................................
				</FONT>			
				</td>
				<td align="left" valign="top">			
					<img src="../images/stamp.gif" border="0" style="position: absolute;z-index: 2">
					<BR>
					&nbsp;<SPAN style="position: absolute;z-Index: 10"><b>KShs&nbsp;&nbsp;<%= FormatNum(totalContractStamps) %></b></SPAN>
					<br><br>
					Revenue Stamps prepaid
					
				</td>
			</tr>
		</table>
		<%
		groupRs.MoveNext
		Loop
	End if
	%></form><%
End Sub

''AGENT CONTRACTS COMPOUNDED
Sub AgentContractsCompoundedGenerate()
	%>
	<form id="repToPDF4" name="repToPDF4">
	<%
	Set conn = GetActiveConnection("KBroker")
	Set Rs = CreateObject("ADODB.Recordset")	
	Set groupRs = CreateObject("ADODB.Recordset")
	
	sqlStr = "SELECT * FROM AgentContractCompounded WHERE (LotTDate = '" & FormatDate(selectedContractDate) & "') ORDER BY LotSlipNo"
	
	groupRs.CursorLocation = adUseClient	
	groupRs.Open sqlStr, conn.ConnectionString, adOpenKeyset, adLockOptimistic
		
	If Not groupRs.EOF Or groupRs.BOF Then
		Set OrderRs = CreateObject("ADODB.Recordset")						        
		sqlStr = "SELECT DISTINCT Order_DPA_ FROM AgentContractCompounded WHERE (LotTDate = '" & FormatDate(selectedContractDate) & "') ORDER BY Order_DPA_"
		OrderRs.CursorLocation = adUseClient	
		OrderRs.Open sqlStr, conn.ConnectionString, adOpenKeyset, adLockOptimistic
			
		Set levyOrderRs = Conn.Execute("SELECT * FROM LevyOrderList ORDER BY LevyOrder")
	
		pageNumber = 0
		
		Do Until OrderRs.EOF
			groupRs.Filter = "Order_DPA_ = '" & OrderRs.Fields("Order_DPA_").Value & "'"
			
			prevSecurity = groupRs.Fields("OrdDetailSecurity").Value
			prevDate = groupRs.Fields("LotTDate").Value
			prevOrderTypeSale = groupRs.Fields("OrderTypeSale").Value
			slipNos = groupRs.Fields("LotSlipNo").Value
			prevSlipNo = slipNos
			prevOrderNo = groupRs.Fields("Order_DPA_").Value
			prevClient = groupRs.Fields("OrdDetailClient").Value
		
			Do Until groupRs.EOF
				If prevSlipNo <> groupRs.Fields("LotSlipNo").Value  Then
					If slipNos = "" Then 
						slipNos = groupRs.Fields("LotSlipNo").Value
					Else 
						slipNos = slipNos & ", " & groupRs.Fields("LotSlipNo").Value
					End If
														
					prevSlipNo = groupRs.Fields("LotSlipNo").Value					
				End If
						
				groupRs.MoveNext
			loop
			
			Set Rs = CreateObject("ADODB.Recordset")
			'the number 13 in the SQL below represents 'Transfer Fee'							        
			sqlStr = "SELECT   Order_DPA_, OrdDetailSecType,Agent_DPA_, AgentName, AgentAddr, LevyName, LevyShortName, CASE SystemMaintained WHEN 13 THEN MAX(LevyAmount) ELSE SUM(LevyAmount) " & _
			        "END AS LevyAmount, SUM(GrossAmount) AS GrossAmount, MIN(CAST(ClientAddr AS nvarchar(4000))) AS ClientAddr, " & _	
					"MIN(OrderRef) AS orderRef, MIN(CAST(OrderTypeSale AS INT)) AS OrderTypeSale, MIN(OrdDetailType) AS OrdDetailType, MIN(ContractNumber) " & _
					"AS ContractNumber, OrdDetailSecurity, SUM(LotQty) AS LotQty, SUM(LotPrice) AS LotPrice, OrdDetailClient, SystemMaintained, LevyRatePercentage " & _
					"FROM         dbo.AgentContractCompounded " & _
					"WHERE     (LotSlipNo IN (" & slipNos  & ")) And LevyName <> '' AND OrdDetailClient = '" & prevClient & "' AND (Order_DPA_ = " & OrderRs.Fields("Order_DPA_").Value & ")" & _
					"GROUP BY Agent_DPA_, AgentName, AgentAddr, LevyName, LevyShortName, OrdDetailSecurity, OrdDetailClient, SystemMaintained, LevyRatePercentage,OrdDetailSecType,Order_DPA_"
				
			Rs.CursorLocation = adUseClient	
			Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
		
			totalGross = Rs.Fields("GrossAmount").Value 
			orderRef = Rs.Fields("orderRef").Value
			if isnull(orderRef) or trim(orderRef) = "" then
				orderRef = Rs.Fields("Order_DPA_").Value
			else
				orderRef = Rs.Fields("Order_DPA_").Value & "/" & orderRef
			end if
			agentAddress = Replace(Rs.Fields("AgentAddr").Value, Chr(13), ",")		
				
			If Trim(UCase(Rs.Fields("OrdDetailType").Value)) = "PURCHASE" Then
				IsPurchase = True
			Else
				IsPurchase = False
			End If  
				
			pageNumber = pageNumber + 1
			%>
			<br class="newpage">
				
			<table border="0" cellspacing=0 cellpadding=0 class="ReportsTable" width="80%" align="center">
				<tr class="pageNumbering">
					<td align="left" colspan=2>
						&nbsp;<!--<FONT FACE=ARIAL SIZE=2><B>Page <%=pageNumber%>	</B></FONT>	-->
					</td>		
				</tr>
				<tr>
					<td align="center" colspan=2>
						<Img Src="../data/photos/aaprintlogo.jpg">			
					</td>		
				</tr>
				<tr>
					<td align="center" colspan=2>
						<FONT FACE=ARIAL SIZE=3><B>
							<%= UCase(Rs.Fields("OrdDetailType").Value) %> CONTRACT <%= Rs.Fields("ContractNumber").Value %>
						</B></FONT>
					</td>		
				</tr>  
				<!--<tr>
					<td colspan=2>&nbsp;</td>
				</tr>  	-->
				<tr>
					<td  nowrap width="70%">			
						<table cellspacing=0 cellpadding=0 border=0  width="100%">
							<tr>
								<td nowrap><%= Rs.Fields("AgentName").Value %> </td>					
							</tr>
							<tr>					
								<td nowrap><%= agentAddress %></td>
							</tr>
							<tr>					
								<td nowrap><b>Account:</b> <%= Rs.Fields("OrdDetailClient").Value %></td>
							</tr>
						
						</table>
					</td>
					<td width="30%">
						<table cellspacing=0 cellpadding=0 border=0 width="100%">
							<tr>
								<td  nowrap colspan="2">Date:&nbsp;<%= FormatDate(selectedContractDate) %></td>
							</tr>
							<tr>
								<td  colspan="2" nowrap>Order Ref:&nbsp;<%= orderRef %></td>					
							</tr>
							<tr>
								<td nowrap valign="top">CDS Ref:&nbsp;</td>
								<td valign="top">
								<table border=0 cellpadding=1 cellspacing=0>
									
										<%
										If InStr(1, slipNos, ",") > 0 Then
											Dim slipNosArray
											maxSlipNosPerRow = 3
											slipNosArray = Split(slipNos, ",")
											maxSlipCount = UBound(slipNosArray) 
											
											For k = 0 To maxSlipCount	%>
												<tr><td nowrap>
												<%For l = 1 To maxSlipNosPerRow
													If k <= maxSlipCount Then											
														thisSlipNo = Trim(slipNosArray(k))
														If k < maxSlipCount Then
															thisSlipNo = thisSlipNo & ", "
														End If																						
														Response.Write thisSlipNo											
														If l <> maxSlipNosPerRow Then
															k = k + 1
														End If	
													Else
														Exit For
													End If																																									
												Next%>									
												</td></tr>
											<%Next
										
										Else%>
										<tr><td nowrap><%= slipNos %></td></tr>
										<%End If%>
									
								</table>
								
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td colspan=2>Dear Sir/Madam</td>
				</tr>
				<tr>
					<td colspan=2>&nbsp;</td>
				</tr>
				<tr>
					<td colspan=2><U><b>RE: <%= UCase(Rs.Fields("OrdDetailType").Value) %> OF SECURITIES<b></U></td>
				</tr>
				<tr>
					<td colspan=2>&nbsp;</td>
				</tr>
				<tr>
					<td colspan=2><font face=Arial size=2>We wish to advise that we have <% If IsPurchase Then Response.Write "bought" Else Response.Write "sold"  %> the following securities as per your instruction. Kindly arrange	to complete the transaction as per the details given below: </font></td>
				</tr>
				<tr>
					<td colspan=2 style="BORDER-BOTTOM: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent">&nbsp;</td>
				</tr>
				<tr>
					<td colspan=2>&nbsp;</td>
				</tr>
			</table>	

			<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable"  width="80%" align="center">	
				<tr>
					<td>Security</td>
					<td align=right><%= Rs.Fields("OrdDetailSecurity").Value %></td>
				</tr>
				<tr>
					<td><%if lcase(rs.Fields("OrdDetailSecType")) = "fixed" Then%>
								Face Value
						<%else%>
								Quantity
						<%end if%></td>
					<td align=right> <%= FormatNum(Rs.Fields("LotQty").Value) %> </td>
				</tr>
				<tr>
					<td>Average Price</td>
					<td align=right> 
						<%if lcase(rs.Fields("OrdDetailSecType")) = "fixed" Then%>
								<%= FormatNumEx((totalGross / Rs.Fields("LotQty").Value) * 100,4) %>
						<%else%>
								<%= FormatNum(totalGross / Rs.Fields("LotQty").Value) %>
						<%end if%> </td>
				</tr>
			    <tr>
					<td>Gross</td>
					<td align=right><b> <%= FormatNum(totalGross) %> </b> </td>
				</tr>
				<tr>
					<TD COLSPAN=2 VALIGN="TOP">
					<TABLE WIDTH="100%"  CELLSPACING="0" CELLPADDING="4">
					<%
					totalLevies = 0
					totalContractStamps = 0
					levyArray = SortLevies(rs, levyOrderRs)
					transferFeeVal = 0
					contractStampsVal = 0
					'careful, returned array is one-based
					For i = 1 To UBound(levyArray) 
						 
						thisLevyName = Trim(levyArray(i, 0))
						If levyArray(i, 2) = 10 Then
							totalContractStamps = levyArray(i, 1)
						End If
						
						'levyPerc = levyArray(i, 3)
						levyPerc = (levyArray(i, 1) / totalGross) * 100			
						
						'grab agent commission here: to be used laters..
						If levyArray(i, 2) = 12 Then
							AgentCommission = levyArray(i, 1)
							AgentRate = Replace(levyPerc, "%", "") '
							
							AgentRateDescription = levyArray(i, 0)
							'empty the levy name variable so as not to display
							thisLevyName = ""
						End If
						
						'do not display staff commission  HARD CODING!!!!!!
						If (levyArray(i, 2) = 8) Then
							thisLevyName = ""
						End If
						
						'grab the broker rate here
						If levyArray(i, 2) = 11 Then
							brokerRate = Replace(levyPerc, "%","")
							If brokerRate = "" Then
								brokerRate = "0"
							End If	
							
						End If
						
						If InStr(1, levyPerc, "%") > 0 Then
							'recalculate average percentage
							levyPerc = (levyArray(i, 1) / totalGross) * 100
							levyPerc = FormatNum(levyPerc) & "%"
						End If
						
						'transfer fee
						If levyArray(i, 2) = 13  Then				
							transferFeeDesc = thisLevyName
							transferFeeVal = levyArray(i, 1)
							thisLevyName = ""
						End If
						
						If levyArray(i, 2) = 10 Then
							contractStampsDesc = thisLevyName
							contractStampsVal = levyArray(i, 1)
							thisLevyName = ""
						End If
						
						If thisLevyName <> "" Then
							totalLevies = totalLevies + levyArray(i, 1)
							levyPerc = FormatNum(levyPerc) & "%"
							
							
							%>
							<TR>		
								<td nowrap width="10%" STYLE="PADDING-LEFT: 0PX"><%= thisLevyname %></td>
								<td ALIGN="LEFT">&nbsp;&nbsp;<%= levyPerc  %></td>
								<td align=right><%= FormatNum(levyArray(i, 1)) %></td>		
							</TR>
					<% End If
					Next%>
					</TABLE>	
					</TD>
				<tr>
					<td>&nbsp;</td>
					<td align=right>
						<hr width=70px height=1px>
						<%= Replace(Space(Len(FormatNum(totalLevies)) * 4), Space(1), "&nbsp;") %>
					</td>
				</tr>
				<tr>
					<td>Total Commission and Levies</td>
					<td align=right>	
						<%If IsPurchase Then%>
							<b><%= FormatNum(totalLevies) %></b>
						<%Else%>	
							<b>(<%= FormatNum(totalLevies) %>)</b>
						<%End If%>
					</td>
				</tr>
				<%if trim(transferFeeDesc) <> "" then%>
						<tr>
							<td><%= transferFeeDesc %></td>
							<td align=right>	
								 <%= FormatNum(transferFeeVal) %>		
							</td>
						</tr>
				<%end if%>
				
				<%if trim(contractStampsDesc) <> "" then%>
						<tr>
							<td><%= contractStampsDesc %></td>
							<td align=right>	
								 <%= FormatNum(contractStampsVal) %>		
							</td>
						</tr>
				<%end if%>
				<tr>
					<td>&nbsp;</td>
					<td align=right> 
						<%= Replace(Space(Len(FormatNum(totalLevies)) * 4), Space(1), "&nbsp;") %>
						<hr width=70px height=1px>
					</td>
				</tr>
				<tr>
					<td>GROSS AMOUNT</td>
					<td align=right><b>
						<%
						totalLevies = totalLevies + transferFeeVal + contractStampsVal
						If IsPurchase Then
							grossAmount = totalGross + totalLevies				
						Else
							grossAmount = totalGross - totalLevies				
						End If
						
						Response.Write FormatNum(grossAmount) 
						%>
						
						</b></td>
				</tr>
				<tr>	
					<td COLSPAN=2 VALIGN="TOP">
						<TABLE WIDTH="100%" CELLSPACING="0" CELLPADDING="4">
							<TR>		
								<td nowrap width="10%" STYLE="PADDING-LEFT: 0PX">Returnable Commission</td>
								<%
								 if brokerRate = 0 or brokerRate = "" then
								   ReturnCommission = 0
								 else
								   ReturnCommission = (AgentRate * 100/brokerRate)
								 end if
								
								%>
								<td ALIGN="LEFT">&nbsp;&nbsp;<%= FormatNum(FormatNumEx(ReturnCommission,0)) & "%" %></td>
								<td align=right><%= FormatNum(AgentCommission)  %></td>		
							</TR>		
					</table>
					</td>	
				</tr>
				<tr>
					<td>NET AMOUNT</td>
					<td align=right><b>
						<%
							If IsPurchase Then
								netAmount = grossAmount - AgentCommission				
							Else
								netAmount = grossAmount + AgentCommission
							End If
						%><%=FormatNum(netAmount) %>
						</b></td>
				</tr>
				<tr>
					<td colspan=2>&nbsp;</td>
				</tr>
				<tr>
					<td colspan=2>&nbsp;</td>
				</tr>
				<tr>
					<td>
					<font face=Arial size=2>
			       <i>For and on behalf of</i>
			African Alliance Uganda Securities<br><br>

			Sign..........................................
					</FONT>	
					</td>
					<td align="left" valign="top">			
					
						<img src="../images/stamp.gif" border="0" style="position: absolute;z-index: 2">
						<BR>
						&nbsp;<SPAN style="position: absolute;z-Index: 10"><b>KShs&nbsp;&nbsp;<%= FormatNum(totalContractStamps) %></b></SPAN>
						<br><br><br>
						Revenue Stamps prepaid
						
					</td>
				</tr>
			</table>
			<%			
			slipNos = ""
			OrderRs.MoveNext
		Loop
	End If
	%></form><%
End Sub

''PAYMENT REQUEST
Sub PaymentRequest()
	%>
	<form id="repToPDF5" name="repToPDF5">
	<%
	'if abc then
	Set conn = GetActiveConnection("KBroker")
	Set Rs = CreateObject("ADODB.Recordset")	
	
	sqlStr = "SELECT PaymentRequest_DPA_ AS ID FROM PaymentRequestList WHERE (PaymentPDate = '" & FormatDate(selectedContractDate) & "') ORDER BY ClientName"
	Set rs2 = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))

	if not (Rs2.bof or Rs2.eof) then
		do until Rs2.eof
			
			sqlStr = "SELECT * FROM PaymentRequestList WHERE PaymentRequest_DPA_=" & Rs2("ID")
			Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))

			if not (Rs.bof or Rs.eof) then
				%>
				<br class="newpage">
			
				<TABLE cellSpacing=3 cellPadding=2 border=0 width="80%" align="center">
					<TR>	  
					    <TD colspan=5 width="100%" align="center"><Img Src="../data/photos/aaprintlogo.jpg"><TD>
				    </TR>
						
					<tr><td colspan=5>&nbsp;</td></tr>
						
					<TR>	  
						<td style="border:1 solid black;" colspan=5 align="center" height=30><b>TRADE SETTLEMENT PAYMENT REQUISITION</b></td>
				    </TR>
						
					<tr><td colspan=5>&nbsp;</td></tr>
						
					<TR>
						<td nowrap width="5%" align=left><b>DATE:</b></td>
						<td nowrap style="border: 1px solid #000000;" width="35%" height=25><%=formatDate(rs("PaymentPDate"))%></td>
						<td nowrap width="20%" height=25>&nbsp;</td>
						<td nowrap width="5%" align=right><b>OUR REFERENCE:</b></td>
						<td nowrap style="border: 1px solid #000000;" width="35%" height=25>&nbsp;</td>
					</TR>
						
					<tr><td colspan=5 height=10>&nbsp;</td></tr>
						
					<tr>
						<td nowrap width="5%"><b>ISSUE CHEQUE IN FAVOUR OF:</b>&nbsp;&nbsp;</td>
						<td align=center nowrap width="95%" colspan="4" style="border: 1 solid black;" height="30"><%=rs.Fields("ClientName")%>&nbsp;</td>
					<tr>
						
					<tr><td colspan=5 height=10>&nbsp;</td></tr>
						
					<tr>
						<td nowrap width="5%"><b>AMOUNT IN WORDS:</b>&nbsp;&nbsp;</td>
						<td align=center nowrap width="95%" colspan="4" style="border: 1 solid black;" height="80"><%=ConvertTowords(rs("PaymentAmount"),"")%></td>
					<tr>
						
					<tr><td colspan=5 height=10>&nbsp;</td></tr>
						
					<tr>
						<td nowrap width="5%"><b>AMOUNT IN FIGURES:</b>&nbsp;&nbsp;</td>
						<td nowrap width="60%" colspan="3">&nbsp;</td>
						<td align=center nowrap style="border: 1px solid #000000;" width="35%" height=25><%=FormatNum(rs("PaymentAmount"))%></td>
					<tr>
						
					<tr><td colspan=5 height=10>&nbsp;</td></tr>
						
					<tr>
						<td nowrap width="5%"><b>CURRENCY:</b>&nbsp;&nbsp;</td>
						<td nowrap width="60%" colspan="3">&nbsp;</td>
						<td align=center nowrap style="border: 1px solid #000000;" width="35%" height=25><b>USHS</b></td>
					<tr>
						
					<tr><td colspan=5 height=10>&nbsp;</td></tr>
						
					<tr>
						<td nowrap width="5%"><b>CHEQUE NUMBER:</b>&nbsp;&nbsp;</td>
						<td nowrap width="60%" colspan="3">&nbsp;</td>
						<td align=center nowrap style="border: 1px solid #000000;" width="35%" height=25><%=rs("PaymentReference")%></td>
					<tr>
						
					<tr><td colspan=5 height=10>&nbsp;</td></tr>
						
					<tr>
						<td nowrap width="5%"><b>DEBIT A/C NUMBER:</b>&nbsp;&nbsp;</td>
						<td nowrap width="60%" colspan="3">&nbsp;</td>
						<%
						If Len(rs("BankAccountName"))>0 Then
							bnk = split(rs("BankAccountName")," ")
							BankAccountName = bnk(Ubound(bnk))
						End If
						%>
						<td align=center nowrap style="border: 1px solid #000000;" width="35%" height=25><%=BankAccountName%></td>
					<tr>
						
					<tr><td colspan=5 height=10>&nbsp;</td></tr>
						
					<tr>
						<td nowrap width="5%"><b>PREPARED BY:</b>&nbsp;&nbsp;</td>
						<td align=center nowrap width="95%" colspan="4" style="border: 1 solid black;" height="25"><%=rs("ChangedBy")%></td>
					<tr>
						
					<tr><td colspan=5>&nbsp;</td></tr>
						
					<tr>
						<td colspan=5>
							<table cellpadding="1" cellspacing="1" border="0" width="100%">
								<tr>
								<td width="10%" nowrap><b>APPROVED BY:</b>&nbsp;&nbsp;</td>
								<td align=right width="36%" nowrap style="border: 1 solid black;" height=25>&nbsp;</td>
								<td width="8%" nowrap>&nbsp;</td>
								<td width="10%" nowrap><b>APPROVED BY:</b>&nbsp;&nbsp;</td>
								<td width="36%" nowrap style="border: 1 solid black;" height=25>&nbsp;</td>
								</tr>
							</table>
						</td>
					</tr>
						
					<tr><td colspan=5>&nbsp;</td></tr>
						
					<tr>
						<td colspan=5>
							<table cellpadding="1" cellspacing="1" border="0" width="100%">
								<tr>
								<td width="10%" nowrap><b>SIGNATURE:</b>&nbsp;&nbsp;</td>
								<td align=right width="36%" nowrap style="border: 1 solid black;" height=50>&nbsp;</td>
								<td width="8%" nowrap>&nbsp;</td>
								<td width="10%" nowrap><b>SIGNATURE:</b>&nbsp;&nbsp;</td>
								<td width="36%" nowrap style="border: 1 solid black;" height=50>&nbsp;</td>
								</tr>
							</table>
						</td>
					</tr>
						
					<tr><td colspan=5>&nbsp;</td></tr>
						
					<tr>
						<td colspan=5>
							<table cellpadding="1" cellspacing="1" border="0" width="100%">
								<tr>
								<td width="10%" nowrap><b>DATE:</b>&nbsp;&nbsp;</td>
								<td align=right width="36%" nowrap style="border: 1 solid black;" height=25>&nbsp;</td>
								<td width="8%" nowrap>&nbsp;</td>
								<td width="10%" nowrap><b>DATE:</b>&nbsp;&nbsp;</td>
								<td width="36%" nowrap style="border: 1 solid black;" height=25>&nbsp;</td>
								</tr>
							</table>
						</td>
					</tr>
				</TABLE>  
				<%
			end if
			
			Rs2.MoveNext
		loop
	else
		Response.End
	end if
	'end if
	%></form><%
End Sub

''SETTLEMENT
Sub SettlementGenerate()
	%>
	<form id="repToPDF6" name="repToPDF6">
	<%
	Set conn = GetActiveConnection("KBroker")
	Set Rs = CreateObject("ADODB.Recordset")	
	Set groupRs = CreateObject("ADODB.Recordset")
	
	''SETTLEMENT SLIPS
	selectedTradeDate = selectedContractDate

	sqlStr = "SELECT * FROM [TimeLimitList] WHERE TimeLimit_DPA_ = 1" 'Settlement (Brokers)
	Set tempRs = Conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
	        
	If Not (tempRs.EOF Or tempRs.BOF) Then
		If timeLimit = "0" Then
			'use NSE
			timeRange = tempRs.Fields("TimeLimitNSE").Value
		Else
			'use internal
			timeRange = tempRs.Fields("TimeLimitInternal").Value
		End If		
	Else
		Set tempRs = Nothing
		Set Conn = Nothing
		Response.End 
	End If
        
	sqlStr = "SELECT * FROM [Holidays]"
	Set holidayRs = CreateObject("ADODB.Recordset")

	holidayRs.CursorLocation = adUseClient
	holidayRs.Open SQLServerFormat(HandleQuote(sqlStr)), Conn.ConnectionString, 1, 3
		                
	'get upper filter date
	originalselectedTradeDate = selectedTradeDate
	settlementDate = FormatDate(selectedTradeDate)
	theSelDay = WeekDay(selectedTradeDate)
		        
	'get to a monday if selected date is on weekend
	settlementDate = JumpUpFromWeekendToWeek (settlementDate)
		        
	'check upwards until upper date is not a holiday
	upperLimitDate = settlementDate
        
	For i = 1 To timeRange
		upperLimitDate = DateAdd("d", 1, upperLimitDate)
		upperLimitDate = JumpUpFromWeekendToWeek(upperLimitDate)
				
		'count only working days
		'check upwards until is not a holiday
		isProperDate = False
			
		Do Until isProperDate = True
			holidayRs.Cancel		
			holidayRs.Filter = "Holiday = '" & upperLimitDate & "'"		
			If Not (holidayRs.EOF Or holidayRs.BOF) Then
				upperLimitDate = DateAdd("d", 1, upperLimitDate)
			Else
				formerSelectedTradeDate = upperLimitDate
				upperLimitDate = JumpBackFromWeekendToWeek (upperLimitDate)
				If DateDiff("d", FormatDate(upperLimitDate), FormatDate(formerSelectedTradeDate)) = 0 Then
					isProperDate = True
					holidayRs.Cancel
				End If	
			End If
		Loop
	Next
        
	Set holidayRs = Nothing
		
	sqlStr = "SELECT * FROM [SettlementSlips] WHERE SettlementDate = '" & FormatDate(selectedTradeDate) & "' ORDER BY BrokerName"
	        
	Set rs = CreateObject("ADODB.Recordset")
	rs.CursorLocation = adUseClient
	Rs.Open SQLServerFormat(HandleQuote(sqlStr)), Conn.ConnectionString, 1, 3
	        
	If Not Rs.EOF or RS.BOF Then
		Do Until rs.EOF
			currBrokerCode = Rs.Fields("BrokerCode").Value
			%>
			<br class="newpage">
			
			<table width=80% align="center" class="ReportsTable">
				<tr>
					<td height=100>&nbsp;</td>
				</tr>	
			</table>
				
			<table width=80% align="center" class="ReportsTable">
				<tr>
					<td><b><font face="Arial Narrow" size="4">Settlement Slip</font></b></td>
					<td align=right><b><font face="Arial Narrow" size="4"><%= Session("CompanyName") %></font></b></td>
				</tr>		
				<tr>
					<td colspan=2><font face="Arial" size="2">printed for each individual broker&nbsp;</font></td>
				</tr>
				<tr>
					<td colspan=2><font face="Arial" size="2">&nbsp;</font></td>
				</tr>	
			</table>

			<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="80%" align="center">
			    <tr>
			      <td width="20%" style="border-top-style: solid; border-top-width: 1"><b><font face="Arial Narrow" size="2">TO:</font></b></td>
			      <td style="border-top-style: solid; border-top-width: 1"><b><font face="Arial Narrow" size="2">UGANDA STOCK EXCHANGE</font></b></td>
			    </tr>
				<tr>
			      <td width="20%"><b><font face="Arial Narrow" size="2">BROKER:</font></b></td>
			      <td><b><font face="Arial Narrow" size="2"><%= Rs.Fields("BrokerName").Value %></font></b></td>
			    </tr>
			    <tr>
			      <td width="20%" style="border-bottom-style: solid; border-bottom-width: 1">
			        <p align="left"><b><font size="2"  face="Arial Narrow">SETTLEMENT DATE:</font></b></td>

			      <td style="border-bottom-style: solid; border-bottom-width: 1">
			        <p align="left"><font face="Arial" size="2"><%= FormatDateFull(upperLimitDate) %></font></td>
			    </tr>
			</table>
	
			<table border="0" cellspacing=0 cellpadding=3 width="80%" align="center">
				<tr>
					<td colspan=7>	
					
					<font face="Arial" size="2">
					<b>Payment to Brokers</b></font>			
					<b><font face="Arial" size="2">		
					<BR>Please prepare payment payable to brokers as per the details shown on this voucher.
					Kindly note that all contracts on the list were traded on the same date shown below:</font></b>
					</td>
				</tr>
				<tr bgcolor="#000000">
					<td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial" size="2">Slip</font></b></td>
					<td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial" size="2">Traded</font></b></td>	
					<td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial" size="2">Security</font></b></td>	
					<td align=right bgcolor="#000000"><b><font color="#FFFFFF" face="Arial" size="2">Quantity</font></b></td>
					<td align=right bgcolor="#000000"><b><font color="#FFFFFF" face="Arial" size="2">Price</font></b></td>
					<td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial" size="2">Contract</font></b></td>
					<td align=right bgcolor="#000000"><b><font color="#FFFFFF" face="Arial" size="2">Gross</font></b></td>
				</tr>
				<%		
				totalLevyAmount = 0
				isDifferentBroker = False
				
				Do Until isDifferentBroker
				   totalLevyAmount = totalLevyAmount + FormatNum(rs.Fields("LevyAmount")) %>
			        <tr>
			             <td><%=rs.Fields("LotSlipNo")%></td>
			             <td><%= FormatDate(rs.Fields("LotTDate")) %></td>
			             <td><%=rs.Fields("OrdDetailSecurity")%></td>
			             <td align=right><%= FormatNum(rs.Fields("LotQty")) %></td>
			             <td align=right><%= FormatNum(rs.Fields("LotPrice")) %></td>
			             <td style="BORDER-RIGHT: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent"><%=rs.Fields("ContractNumber")%></td>                   
			             <td align=right><%= FormatNum(rs.Fields("LevyAmount")) %></td>
			                        
			        </tr>
			        <%
			        Rs.MoveNext   

			        If Not (Rs.EOF Or Rs.BOF) Then
				   		If Rs.Fields("BrokerCode").Value <> currBrokerCode Then					
				   			Rs.Move -1
				   			isDifferentBroker = True
				   		End If
			        Else
				   		isDifferentBroker = True
				   		Rs.Move -1
			        End If   
			    Loop      
			    %>
				<tr>
					<td colspan=7 align=right>&nbsp;</td>
				</tr>               
				        
				<tr>
					<td colspan=6 align=right><b>Broker Totals:</b></td>
					<td align=right style="border-style: solid; border-color: #000000; border-width: 1" height="30px"><%= FormatNum(totalLevyAmount) %></td>
				</tr>        
			</table>
			<%
			rs.MoveNext
		Loop
	End If
	        
	''SETTLEMENT REMINDERS	
	sqlStr = "SELECT * FROM [TimeLimitList] WHERE TimeLimit_DPA_ = 1" 'Settlement (Brokers)
	Set tempRs = Conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
	        
	If Not (tempRs.EOF Or tempRs.BOF) Then
		If timeLimit = "0" Then
			'use NSE
			timeRange = tempRs.Fields("TimeLimitNSE").Value
		Else
			'use internal
			timeRange = tempRs.Fields("TimeLimitInternal").Value
		End If		
	Else
		Set tempRs = Nothing
		Set Conn = Nothing
		Response.End 
	End If
        
	sqlStr = "SELECT * FROM [Holidays]"
	Set holidayRs = CreateObject("ADODB.Recordset")

	holidayRs.CursorLocation = adUseClient
	holidayRs.Open SQLServerFormat(HandleQuote(sqlStr)), Conn.ConnectionString, 1, 3
	        
	'get upper filter date
	originalselectedTradeDate = selectedTradeDate
	settlementDate = FormatDate(selectedTradeDate)
	theSelDay = WeekDay(selectedTradeDate)
	        
	'get to a monday if selected date is on weekend
	settlementDate = JumpUpFromWeekendToWeek (settlementDate)

	'check upwards until upper date is not a holiday
	upperLimitDate = settlementDate
	        
	For i = 1 To timeRange
		upperLimitDate = DateAdd("d", 1, upperLimitDate)
		upperLimitDate = JumpUpFromWeekendToWeek(upperLimitDate)
				
		'count only working days
		'check upwards until is not a holiday
		isProperDate = False
			
		Do Until isProperDate = True
			holidayRs.Cancel		
			holidayRs.Filter = "Holiday = '" & upperLimitDate & "'"		
			If Not (holidayRs.EOF Or holidayRs.BOF) Then
				upperLimitDate = DateAdd("d", 1, upperLimitDate)
			Else
				formerSelectedTradeDate = upperLimitDate
				upperLimitDate = JumpBackFromWeekendToWeek (upperLimitDate)
				If DateDiff("d", FormatDate(upperLimitDate), FormatDate(formerSelectedTradeDate)) = 0 Then
					isProperDate = True
					holidayRs.Cancel
				End If	
			End If
		Loop
	Next

	sqlStr = "SELECT * FROM [SettlementReminder] WHERE TRADED = '" & FormatDate(originalselectedTradeDate) & "' ORDER BY Traded DESC"
		
	Set rs = CreateObject("ADODB.Recordset")
	rs.CursorLocation = adUseClient
	Rs.Open SQLServerFormat(HandleQuote(sqlStr)), Conn.ConnectionString, 1, 3

	If Not rs.EOF Or rs.BOF Then
		%>
		<br class="newpage">
		
		<table width=80% align="center" class="ReportsTable">
			<tr>
				<td height=100>&nbsp;</td>
			</tr>	
		</table>
			
		<table width="80%" align="center">
			<tr>
				<td><b><font face="Arial Narrow" size="4">Settlement Reminder</font></b></td>
				<td align=right><b><font face="Arial Narrow" size="4"><%= Session("CompanyName") %> </font></b></td>  
			</tr>	
			<tr>
				<td colspan=2><font face="Arial" size="2">for funds from or to brokers on: <%= FormatDateFull(upperLimitDate)%></font></td>
			</tr>
			<tr>
				<td  colspan=2><font face="Arial" size="2">&nbsp;</font></td>
			</tr>	
		</table>

		<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="80%" align="center">
			<tr>
				<td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow" size="4">Traded</font></b></td>
				<td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow" size="4">Broker</font></b></td>
				<td bgcolor="#000000"><p align="right"><b><font color="#FFFFFF" face="Arial Narrow" size="4">Gross</font></b></p></td>
			</tr>
	
			<tr>
				<td colspan="3">&nbsp;</td>
			</tr>
			<%
			dailyTotals = 0
			myTotals = 0
    
			Rs.Filter = "OrderTypeSale = 0"
    
			If Not (Rs.EOF Or Rs.BOF) Then%>
				<tr>
				  <td colspan="3"><font size="3" face="Impact">Amounts to be paid to brokers</font></td>
				</tr>
				<%	
				Do Until Rs.EOF
					myTotals = myTotals + Rs.Fields("GrossAmt").Value%>
					<tr>
					  <td><%= FormatDate(Rs.Fields("Traded").Value) %></td>
					  <td><%= Rs.Fields("Broker").Value  %></td>
					     <td>
					    <p align="right"><%= FormatNum(Rs.Fields("GrossAmt").Value)%></td>
					</tr>
					<%		
					pageNumber = pageNumber + 1
					Rs.MoveNext
				Loop
				%>
				<tr>
				  <td>&nbsp;</td>
				  <td>&nbsp;</td>	
				  <td style="border-top: solid; border-width: 1" align=right valign=bottom height="25px"><%= FormatNum(myTotals) %></td>
				</tr>
				<%
			End If
		
		    dailyTotals = dailyTotals + myTotals
			myTotals = 0
    
			Rs.Filter = "OrderTypeSale = 1"
    
			If Not (Rs.EOF Or Rs.BOF) Then
				%>    
				<tr>
				  <td colspan="3">&nbsp; </td>
				</tr>
    
				<tr>
				  <td colspan="3"><font size="3" face="Impact">Amounts expected from brokers</font></td>
				</tr>
			    <%	
				Do Until Rs.EOF
					myTotals = myTotals + Rs.Fields("GrossAmt").Value%>
					<tr>
						<td><%= FormatDate(Rs.Fields("Traded").Value) %></td>
						<td><%= Rs.Fields("Broker").Value  %></td>
						<td><p align="right"><%= FormatNum(Rs.Fields("GrossAmt").Value)%></td>
					</tr>
					<%
					Rs.MoveNext
				Loop
				%>
				<tr>
				  <td>&nbsp;</td>
				  <td>&nbsp;</td>		
				  <td style="border-top: solid; border-width: 1; border-color: #000000;" align=right valign=bottom height="25px"><%= FormatNum(myTotals) %></td>
				</tr>
				<%
			End If
			dailyTotals = dailyTotals + myTotals
			%>
			<tr>
				<td colspan="3">&nbsp; </td>
			</tr>
			<tr>
				<td colspan="3">&nbsp; </td>
			</tr>   
			<tr>
				<td></td>
				<td><p align="right">Daily Totals:&nbsp;&nbsp; </td>
				<td style="border-style: solid; border-width: 1"><p align="right"><%= FormatNum(dailyTotals) %></p></td>
			</tr>
		</table>
		<%
	End If
	%></form><%
End Sub

Function SortLevies(srcRs, orderRs)
	Dim returnArray()
				
	ReDim returnArray(srcRs.RecordCount, 3)
				
	'careful, recordcount is 1-based
	arrCounter = 0 
	If Not (orderRs.EOF Or orderRs.BOF) Then
		Do Until orderRs.EOF
			isExistingLevy = False
			Do Until srcRs.EOF
				If StrComp(orderRs.Fields("LevyName").Value, srcRs.Fields("LevyShortName").Value, vbTextCompare) = 0 Then
					arrCounter = arrCounter + 1
					returnArray(arrCounter, 0) = srcRs.Fields("LevyName").Value
					returnArray(arrCounter, 1) = srcRs.Fields("LevyAmount").Value
					returnArray(arrCounter, 2) = srcRs.Fields("SystemMaintained").Value
					returnArray(arrCounter, 3) = srcRs.Fields("LevyRatePercentage").Value
					Exit Do	
				End If
				srcRs.MoveNext
			Loop	
			srcRs.MoveFirst					
			orderRs.MoveNext
		Loop	
		orderRs.MoveFirst
					
		If arrCounter <> UBound(returnArray) Then
			'take care of the rest unordered items
			Do Until srcRs.EOF
				thisName = srcRs.Fields("LevyName").Value
							
				'check if it's in array first
				existsInArray = False
				For k = 1 To arrCounter
					If StrComp(thisName, returnArray(k, 0), vbTextCompare) = 0 Then
						existsInArray = True
						Exit For
					End If
				Next
							
				If Not existsInArray Then
					arrCounter = arrCounter + 1
					returnArray(arrCounter, 0) = srcRs.Fields("LevyName").Value
					returnArray(arrCounter, 1) = srcRs.Fields("LevyAmount").Value
					returnArray(arrCounter, 2) = srcRs.Fields("SystemMaintained").Value	
					returnArray(arrCounter, 3) = srcRs.Fields("LevyRatePercentage").Value	
				End If
				srcRs.MoveNext
			Loop	
			srcRs.MoveFirst
		End If
	Else
		Do Until srcRs.EOF
			arrCounter = arrCounter + 1
			returnArray(arrCounter, 0) = srcRs.Fields("LevyName").Value
			returnArray(arrCounter, 1) = srcRs.Fields("LevyAmount").Value
			returnArray(arrCounter, 2) = srcRs.Fields("SystemMaintained").Value		
			returnArray(arrCounter, 3) = srcRs.Fields("LevyRatePercentage").Value
			srcRs.MoveNext
		Loop	
		srcRs.MoveFirst
	End If
	SortLevies = returnArray
End Function

function stripFormatting(theVal, delimiter)
	Dim i
	Dim tempVal, char
			  
	for i = 1 to len(theVal)
		char = mid(theVal,i,1)
		if isnumeric(char) OR  char = delimiter then tempVal = tempVal & char
	next 
			  
	stripFormatting = tempVal
end function
		 
function ConvertTowords(theVal, delimiter)
	Dim tempVal, strVal,mainpart, decimalpart
	Dim i,d
	theVal = trim(theVal)
			 
	'******************* Validate parameter **************
			 
	if not isnumeric(theVal) or theVal = "" then
	 ConvertTowords = theVal
	 exit function
	end if  
			 
	if delimiter = "" then delimiter = "." 'default delimiter: decimal
			 
	'There should only be one delimiter 
	d=0
	for i = 1 to len(theVal)
	 char = mid(theVal,i,1)
	 if char = delimiter then d = d + 1
	next
			 
	if d > 1 then
	 ConvertTowords = theVal
	 exit function
	end if 
			 
	theVal = stripFormatting(theVal, delimiter)
		  
	pos = 0
	pos = instr(1,theVal, delimiter,1)
			 
	if pos <> 0 then ' Separate main and decimal parts of the number
	 mainpart = left(theVal,pos-1)
	 decimalpart = right(theVal,len(theVal)- pos)
	else
	 mainpart = theVal
	 decimalpart = ""
	end if
			 
	strVal = translate(mainpart,0)
			  
	if strVal <> "" then  
	  if trim(strVal) = "ONE" then 
	   strVal = strVal & " SHILLING "
	  else
	   strVal = strVal & " SHILLINGS "
	  end if
	end if
			 
	decimalpart = left(decimalpart,2) 'Default decimal Places: 2
			 
	if decimalpart <> "" then
	 strVal2 = translate(decimalpart,1)
		 
	 	if strVal2 <> "" AND strVal <> "" then 
	 	  strVal = strval & " AND " & strVal2 & " CENTS "
	 	elseif strVal2 <> "" then 
	 	  strVal = strval & " " & strVal2 & " CENTS "
	 	end if
	end if
			 
	ConvertTowords = strVal
end function
		 
function translate(tempNum, deci)
	 Dim k,convnum,strNum
	 Dim noElements, noCategory
			  
	if deci = "" then deci = 0 'default
	strNum = "" 'initialise empty string to hold amount in words
	convnum  = int(tempNum) 'Remove leading zeros
	noElements = len(tempNum)
			 
	extra = 0
	if noElements > 3 then
	 if (Abs(noElements / 3) - int(noElements / 3)) > 0 then extra = 1
	 noCategory = int(noElements/3) + extra
	else
	 noCategory=1
	end if
			 
	 'Translate given number
			    
	    for k = 1 to  noCategory
			    
	      if len(convnum) >= 3 then
	       num = right(convnum,3) 'fetch number to be translated 
	       convnum = left(convnum,len(convnum)-3) 'strip number from original  
	      else
	       num = convnum
	       convnum = ""
	      end if
			 
	      strEquivalent = wordEquivalent(num,deci)'Get word equivalent
			       
	      if strNum <> "" then comma = "," 'Format output
						
	      if k = 1 then
	       strNum = strEquivalent & strNum
	      elseif k = 2 AND strEquivalent <> "" then
	          if strNum <> "" then
	  		 if mid(tempNum,len(tempNum)-2,1) = 0 then 
	  		  comma = " AND " 'Format output
	  		 else
	  		  comma = "," 'Format output
	  		 end if
	  	    end if
	       strNum = strEquivalent & " THOUSAND" & comma & " " & strNum 
	      elseif k = 3 AND strEquivalent <> "" then
	       strNum = strEquivalent & " MILLION" & comma & " " & strNum
	      elseif k = 4 AND strEquivalent <> "" then
	       strNum = strEquivalent & " BILLION" & comma & " " & strNum
	      elseif k = 5 AND strEquivalent <> "" then
	       strNum = strEquivalent & " TRILLION" & comma & " " & strNum
	      end if  
		 
	   next
			 
	translate = strNum
end function
		 
function wordEquivalent(Num,deci)
	Dim tempNum, strNum, lenNum
			 
	if deci = "" then deci =0
			 
	'**** Validate *****
	if Num = "" then
	 wordEquivalent = ""
	 exit function
	end if
			 
	lenNum = len(Num)
			 
	'Force leading zeros as appropriate
	if deci = 0 then     ' Main number part  
	  if lenNum = 1 then
	   num = "00" & num
	  elseif lenNum = 2 then
	   num = "0" & num
	  end if
	elseif deci =1 then ' Decimal part
	  if lenNum = 1 then
	   num = "0" & num & "0"
	  elseif lenNum = 2 then
	   num = "0" & num
	  end if

	end if
			 
	'wordEquivalent = num
	'exit function        
	strNum = "" 'initialise empty string to hold amount in words
	isNonZero = left(Num,1) <> 0 'first number not a zero
	isNotOne = mid(Num,2,1) <> 1 'Second number not a one
		 
	 for i = 1 to 3
			  
	   tempNum = mid(Num, i, 1)
			    
	    select case tempNum
	       case "0"
	          if i = 2  OR i = 3 then strNum = strNum & ""
	       case 1
	          if i = 1 then
	           strNum = " ONE HUNDRED "
	          elseif i = 2  then
	             if isNonZero then strNum = strNum & " AND "
	             select case right(Num,1)
	  				case 0 
	  				 strNum = strNum & " TEN "
	  				case 1
	  				  strNum = strNum & " ELEVEN "
	  				case 2
	  				  strNum = strNum & " TWELVE "
	  				case 3
	  				  strNum = strNum & " THIRTEEN "
	  				case 4
	  				  strNum = strNum & " FOURTEEN "
	  				case 5
	  				  strNum = strNum & " FIFTEEN "
	  				case 6
	  				  strNum = strNum & " SIXTEEN "
	  				case 7
	  				  strNum = strNum & " SEVENTEEN "
	  				case 8
	  				  strNum = strNum & " EIGHTEEN "
	  				case 9
	  				  strNum = strNum & " NINETEEN "
	             end select 
	          elseif i = 3 AND isNotOne then
	           if mid(Num,2,1) = 0 AND isNonZero then  strNum = strNum & " AND "
	           strNum = strNum &  " ONE "
	          end if 
	       case 2
	          if i = 1 then
	           strNum = " TWO HUNDRED "
	          elseif i = 2  then
	           if isNonZero then strNum = strNum & " AND "
	           strNum = strNum  & " TWENTY "
	          elseif i = 3 AND isNotOne then
	           if mid(Num,2,1) = 0 AND isNonZero then  strNum = strNum & " AND "
	           strNum = strNum & " TWO "
	          end if 
	       case 3
	          if i = 1 then
	           strNum = " THREE HUNDRED "
	          elseif i = 2  then
	           if isNonZero then strNum = strNum & " AND "
	           strNum = strNum  & " THIRTY "
	          elseif i = 3 AND isNotOne then
	           if mid(Num,2,1) = 0 AND isNonZero then  strNum = strNum & " AND "
	           strNum = strNum & " THREE "
	          end if 
	       case 4
	          if i = 1 then
	           strNum = " FOUR HUNDRED "
	          elseif i = 2  then
	           if isNonZero then strNum = strNum & " AND "
	           strNum = strNum  & " FOURTY "
	          elseif i = 3 AND isNotOne then
	           if mid(Num,2,1) = 0 AND isNonZero then  strNum = strNum & " AND "
	           strNum = strNum & " FOUR "
	          end if 
	       case 5
	          if i = 1 then
	           strNum = " FIVE HUNDRED "
	          elseif i = 2  then
	           if isNonZero then strNum = strNum & " AND "
	           strNum = strNum  & " FIFTY "
	          elseif i = 3 AND isNotOne then
	           if mid(Num,2,1) = 0 AND isNonZero then  strNum = strNum & " AND "
	           strNum = strNum & " FIVE "
	          end if 
	       case 6
	          if i = 1 then
	           strNum = " SIX HUNDRED "
	          elseif i = 2  then
	           if isNonZero then strNum = strNum & " AND "
	           strNum = strNum  & " SIXTY "
	          elseif i = 3 AND isNotOne then
	           if mid(Num,2,1) = 0 AND isNonZero then  strNum = strNum & " AND "
	           strNum = strNum & " SIX "
	          end if 
	       case 7
	          if i = 1 then
	           strNum = " SEVEN HUNDRED "
	          elseif i = 2  then
	           if isNonZero then strNum = strNum & " AND "
	           strNum = strNum  & " SEVENTY "
	          elseif i = 3 AND isNotOne then
	           if mid(Num,2,1) = 0 AND isNonZero then  strNum = strNum & " AND "
	           strNum = strNum & " SEVEN "
	          end if
	       case 8
	          if i = 1 then
	           strNum = " EIGHT HUNDRED "
	          elseif i = 2  then
	           if isNonZero then strNum = strNum & " AND "
	           strNum = strNum  & " EIGHTY "
	          elseif i = 3 AND isNotOne then
	           if mid(Num,2,1) = 0 AND isNonZero then  strNum = strNum & " AND "
	           strNum = strNum & " EIGHT "
	          end if
	       case 9
	          if i = 1 then
	           strNum = " NINE HUNDRED "
	          elseif i = 2  then
	           if isNonZero then strNum = strNum & " AND "
	           strNum = strNum  & " NINETY "
	          elseif i = 3 AND isNotOne then
	           if mid(Num,2,1) = 0 AND isNonZero then  strNum = strNum & " AND "
	           strNum = strNum & " NINE "
	          end if
	    end select
			   
	 next
	 wordEquivalent = strNum
End function
%>