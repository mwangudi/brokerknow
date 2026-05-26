<!--#include file="../libroutinesTEST.asp"-->
<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Agent Contract</title>  
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
<SCRIPT language=VBScript src="../scripts/reports.vbs"></SCRIPT>
<SCRIPT language=Javascript src="../scripts/reports.js"></SCRIPT>


	<style media="print">
		@page {
			@top{font-family: Helvetica, Arial, sans-serif;
				font-size: 150%;
				font-weight: bolder;
				text-align: left;
				content: "<%= FormatDate(Date) %>";			
			}
			
			margin-left: 2cm;
			margin-right: 2cm;
			margin-top: 1cm;    
			margin-bottom: 2cm;
			size: portrait;
			
			br.newpage{
				page-break-before:always;
			}
			
			tr.pageNumbering{
				display:none;
			}
		}

	</style>
</head>

<body Class="Reports">

<%
genReport = Request("genReport")
DocNo = Request("DocNo")
Agent = Request("Agent")
selectedFromDate = Request("transFromDate")
selectedToDate = Request("transToDate") 

DrawPageFunctions True, True, True, True
%>

<p id="toPDFOrient" name="toPDFOrient" value="P" style="display:none;">P
<p id="toPDF" name="toPDF">

<%
	Set conn = GetActiveConnection("KBroker")
	Set groupRs = CreateObject("ADODB.Recordset")						        
	
	If Len(DocNo)>0 Then
		sqlStr = "SELECT * FROM AgentContractCompounded WHERE (ContractNumber LIKE '%"& DocNo &"%')"
		sqlStr = "SELECT     * " & _
				 " FROM         AgentContractCompounded " & _
				 " WHERE     (Order_DPA_ IN " & _
				 "                           (SELECT DISTINCT Order_DPA_ " & _
				 "                             FROM          AgentContractCompounded " & _
				 "                             WHERE      (ContractNumber LIKE N'%"& DocNo &"%')))"
	End If
	
	If Len(Agent)>0 Then
		sqlStr = "SELECT * FROM AgentContractCompounded WHERE (Agent_DPA_ = "& Agent &") AND (LotTDate BETWEEN '" & FormatDate(selectedFromDate) & "' AND '"& FormatDate(selectedToDate)  &"')"
	End If
	
	groupRs.CursorLocation = adUseClient	
	groupRs.Open sqlStr, conn.ConnectionString, adOpenKeyset, adLockOptimistic
		
	If groupRs.EOF Or groupRs.BOF Then%>
		<Script Language="JavaScript">
			alert("There are no contracts based on the specified criterion.")
			window.parent.history.go(-1);			
		</Script>
		<%Set groupRs = Nothing
		Set Conn = Nothing
		Response.End
	End If	
	
	Set OrderRs = CreateObject("ADODB.Recordset")						        
	
	If Len(DocNo)>0 Then
		sqlStr = "SELECT DISTINCT Order_DPA_,OrdDetail_DPA_,LotTdate FROM AgentContractCompounded WHERE (ContractNumber LIKE '%"& DocNo &"%')"
	End If
	
	If Len(Agent)>0 Then
		sqlStr = "SELECT DISTINCT Order_DPA_,OrdDetail_DPA_,LotTdate FROM AgentContractCompounded WHERE (Agent_DPA_ = "& Agent &") AND (LotTDate BETWEEN '" & FormatDate(selectedFromDate) & "' AND '"& FormatDate(selectedToDate)  &"')"
	End If
	
	OrderRs.CursorLocation = adUseClient	
	OrderRs.Open sqlStr, conn.ConnectionString, adOpenKeyset, adLockOptimistic
		
	If OrderRs.EOF Or OrderRs.BOF Then%>
		<Script Language="JavaScript">
			alert("There are no contracts based on the specified criterion.")
			window.parent.history.go(-1);			
		</Script>
		<%Set groupRs = Nothing
		Set Conn = Nothing
		Response.End
	End If
	
	Set levyOrderRs = Conn.Execute("SELECT * FROM LevyOrderList ORDER BY LevyOrder")
	Dim pageNumber
	
	pageNumber = 0
	Do Until OrderRs.EOF
			groupRs.Filter = "OrdDetail_DPA_ = '" & OrderRs.Fields("OrdDetail_DPA_").Value & "' and LotTDate= '" & trim(OrderRs.Fields("LotTDate").Value) & "'"

		
			prevSecurity = groupRs.Fields("OrdDetailSecurity").Value
			prevDate = groupRs.Fields("LotTDate").Value
			prevOrderTypeSale = groupRs.Fields("OrderTypeSale").Value
			slipNos = "'" & groupRs.Fields("LotSlipNo").Value & "'"
			prevSlipNo = slipNos
			prevOrderNo = groupRs.Fields("Order_DPA_").Value
			prevClient = groupRs.Fields("OrdDetailClient").Value
		
			Do Until groupRs.EOF
					'If prevSlipNo <> "'" & groupRs.Fields("LotSlipNo").Value & "'"  Then
					'		If slipNos = "" Then 
					'			slipNos = "'" & groupRs.Fields("LotSlipNo").Value & "'"
					'		Else 
					'			slipNos = slipNos & ",'" & groupRs.Fields("LotSlipNo").Value &"'"
					'		End If
					'								
					'		prevSlipNo = ",'" & groupRs.Fields("LotSlipNo").Value	&"'"				
					'					
					'End If			
					If instr(1, slipNos,"'" & trim(groupRs.Fields("LotSlipNo").Value) & "'") =0  Then
							If slipNos = "" Then 
								slipNos = "'" & groupRs.Fields("LotSlipNo").Value & "'"
							Else 
								slipNos = slipNos & ",'" & groupRs.Fields("LotSlipNo").Value &"'"
							End If
													
							prevSlipNo = groupRs.Fields("LotSlipNo").Value					
										
					End If
									
					
					groupRs.MoveNext
			loop
			
			Set Rs = CreateObject("ADODB.Recordset")
			'the number 13 in the SQL below represents 'Transfer Fee'							        
			sqlStr = "SELECT   Order_DPA_, Client_DPA_, OrdDetailSecType,Agent_DPA_, AgentName, AgentAddr, LevyName, LevyShortName, CASE SystemMaintained WHEN 13 THEN MAX(LevyAmount) ELSE SUM(LevyAmount) " & _
                      "END AS LevyAmount, SUM(GrossAmount) AS GrossAmount, MIN(CAST(ClientAddr AS nvarchar(4000))) AS ClientAddr, " & _	
					  "MIN(OrderRef) AS orderRef, MIN(CAST(OrderTypeSale AS INT)) AS OrderTypeSale, MIN(OrdDetailType) AS OrdDetailType, MIN(ContractNumber) " & _
					  "AS ContractNumber, OrdDetailSecurity, SUM(LotQty) AS LotQty, SUM(LotPrice) AS LotPrice, OrdDetailClient, SystemMaintained, LevyRatePercentage ,ClientCDSNo,SUM(LevyVATAmount) AS LevyVATAmount,ContractSettlementDate  " & _
					  "FROM         dbo.AgentContractCompounded " & _
					  "WHERE     (LotSlipNo IN (" & slipNos  & ")) And LevyName <> '' AND OrdDetailClient = '" & prevClient & "' AND (OrdDetail_DPA_ = " & OrderRs.Fields("OrdDetail_DPA_").Value & ")" & _
					  "GROUP BY Agent_DPA_, AgentName, AgentAddr, LevyName, LevyShortName, OrdDetailSecurity, OrdDetailClient, SystemMaintained, LevyRatePercentage,OrdDetailSecType,Order_DPA_,Client_DPA_,ClientCDSNo,ContractSettlementDate"
			'Response.Write SQLSTR
			'Response.End 
			
			Rs.CursorLocation = adUseClient	
			Rs.Open SQLServerFormat(sqlStr), conn.ConnectionString, adOpenKeyset, adLockOptimistic
		
			
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
<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" align ="center" width="90%">
	<THEAD>
	<tr class="pageNumbering">
		<td align="left" colspan=2>
			<FONT FACE=ARIAL SIZE=2><B>Page <%=pageNumber%>	</B></FONT>	
		</td>		
	</tr>
	<tr>
		<td align="center" colspan="2" valign="top">
			<!--#include file="Header.asp"-->			
		</td>		
	</tr>
	<THEAD>  
	<tr>
		<td align="center" colspan="2">
			<FONT FACE=ARIAL SIZE=3><B>
				<%= UCase(Rs.Fields("OrdDetailType").Value) %> CONTRACT <%= Rs.Fields("ContractNumber").Value %>
			</B></FONT>
		</td>		
	</tr>  
	<tr>
		<td colspan=2>&nbsp;</td>
	</tr>  	
	<tr>
		<td  nowrap width="70%">			
			<table cellspacing=0 cellpadding=0 border=0  align ="left" width="90%">
				<tr>
					<td nowrap><%= Rs.Fields("AgentName").Value %> </td>					
				</tr>
				<tr>					
					<td nowrap><%= agentAddress %></td>
				</tr>
				<tr>					
					<td nowrap><b>Account:</b> <%= Rs.Fields("OrdDetailClient").Value %>  [<%= Rs.Fields("Client_DPA_").Value %>] [<%= Rs.Fields("ClientCDSNo").Value %>] </td>					
				</tr>
			
			</table>
		</td>
		<td width="30%">
			<table cellspacing=0 cellpadding=0 border=0 align ="center" width="90%">
				<tr>
					<td  nowrap colspan="2">Trade Date:&nbsp;&nbsp;&nbsp;<%= FormatDate(prevDate) %></td>
				</tr>
				<tr>
					<td  colspan="2" nowrap>Order Ref:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%= orderRef %></td>					
				</tr>
				<tr>
					<td nowrap valign="top">CDS Ref:&nbsp;&nbsp;&nbsp;</td>
					<td valign="top">
					<table border=0 cellpadding=1 cellspacing=0>
						
							<%
							If InStr(1, slipNos, ",") > 0 Then
								Dim slipNosArray
								maxSlipNosPerRow = 6
								slipNosArray = Split(slipNos, ",")
								maxSlipCount = UBound(slipNosArray) 
								
								For k = 0 To maxSlipCount	%>
									<tr><td >
									<%For l = 1 To maxSlipNosPerRow
										If k <= maxSlipCount Then											
											thisSlipNo = Trim(slipNosArray(k))
											If k < maxSlipCount Then
												thisSlipNo = thisSlipNo & ", "
											End If																						
											Response.Write replace(thisSlipNo,"'","")											
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
							<tr><td nowrap><%= replace(slipNos,"'","") %></td></tr>
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
		<td colspan=2><font face=Arial size=2>We wish to advise that we have <% If IsPurchase Then Response.Write "bought" Else Response.Write "sold"  %> the following securities as per your instruction. Kindly arrange to complete the transaction as per the details given below: </font></td>
	</tr>
	<tr>
		<td colspan=2 style="BORDER-BOTTOM: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent">&nbsp;</td>
	</tr>
</table>	

<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable"  align ="center" width="90%">	
	
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
		
		<td><font face=Arial >Settlement Date  </font></td>
		<td align=right><font face=Arial ><b><%= formatdate(Rs.Fields("ContractSettlementDate").Value )%> </b></font> </td>

	</tr>
	<tr>
		<TD COLSPAN=2 VALIGN="TOP">
		<TABLE align ="center" width="100%"  CELLSPACING="0" CELLPADDING="4">

				<TR>		
					<td nowrap  STYLE="PADDING-LEFT: 0PX">&nbsp;</td>
					<td ALIGN="LEFT">&nbsp;</td>
					<td ALIGN="Right"><B>&nbsp;&nbsp;VAT</B></td>
					<td align=right><B>VAT INCL</B></td>		
				</TR>
		<%
		totalLevies = 0
		totalContractStamps = 0
		levyArray = SortLevies(rs, levyOrderRs)
		transferFeeVal = 0
		contractStampsVal = 0

		TotalVat=0
		TotalVatInclusive=0
		'careful, returned array is one-based
		For i = 1 To UBound(levyArray) 
			 
			thisLevyName = Trim(levyArray(i, 0))
			'response.write thisLevyName & "<BR>"
			If levyArray(i, 2) = 10 Then
				totalContractStamps = levyArray(i, 1)
			End If
			
			levyPerc = levyArray(i, 3)
			If InStr(1, levyPerc, "%") > 0 Then
				'recalculate average percentage
				levyPerc = (levyArray(i, 1) / totalGross) * 100
				levyPerc = FormatNumber(levyPerc,2) & " %"
			End If

			' Vat Percentage and MSE percentage
			if levyArray(i, 2) = 99 then
				levyPerc="16.5%"
				thisLevyName = ""
			elseif levyArray(i, 2) = 25  then
				levyPerc="5%"
				thisLevyName = ""
			end if	
			
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
				
				
				TotalVat = TotalVat + levyArray(i, 4)
				VatInclusive = levyArray(i, 1) + levyArray(i, 4)
				totalLevies = totalLevies + VatInclusive
				'Dont Display the Broker Commission Percentage
				If  levyArray(i, 2) = 11  Then		levyPerc=""	
				
				%>
				<TR>		
								<td nowrap width="10%" STYLE="PADDING-LEFT: 0PX"><%= thisLevyName %></td>
								<td ALIGN="LEFT" width="40%">&nbsp;&nbsp;<%= levyPerc %></td>
								<td ALIGN="Right" width="20%">&nbsp;&nbsp;<%= FormatNum(levyArray(i, 4)) %></td>
								<td align=right width="30%"><%= FormatNum(VatInclusive) %></td>		
							</TR>
		<% End If
		Next%>

			<TR>		
				<td nowrap width="10%" STYLE="PADDING-LEFT: 0PX">Total Commission and Levies</td>
				<td ALIGN="LEFT" width="40%">&nbsp;</td>
				<td ALIGN="Right" width="20%"><hr width=70px height=1px>
					
					<%If IsPurchase Then%>
						<b><%= FormatNum(TotalVat) %></b>
					<%Else%>	
						<b>(<%= FormatNum(TotalVat) %>)</b>
					<%End If%></td>
				<td align=right width="30%">
					<hr width=70px height=1px>
					<%If IsPurchase Then%>
						<b><%= FormatNum(totalLevies) %></b>
					<%Else%>	
						<b>(<%= FormatNum(totalLevies) %>)</b>
					<%End If%>
				</td>		
			</TR>
		</TABLE>	
		</TD>
	<!--<tr>
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
	</tr>-->
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
	<%
					 if brokerRate = 0 or brokerRate = "" then
					   ReturnCommission = 0
					 else
					   ReturnCommission = (AgentRate * 100/brokerRate)
					 end if
					
					%>
		<td>Returnable Commission &nbsp;&nbsp;&nbsp;<%= FormatNum(FormatNumEx(ReturnCommission,0)) & "%" 
					
					%></td>
		
		<td align=right> 
			<%= FormatNum(AgentCommission)  %></td>
	</tr>
	<tr>
		<td>NET AMOUNT</td>
		<td align=right><b>
			<% Dim netAmount
				
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
		<PRE><font face=Arial size=2>
       <i>For and on behalf of</i>
African Alliance Malawi Securities

Sign..........................................
		</FONT></PRE>	
		</td>
		<td align="left" valign="top">			
		<!--
			<img src="../images/stamp.gif" border="0" style="position: absolute;z-index: 2">
			<BR>
			&nbsp;<SPAN style="position: absolute;z-Index: 10"><b>KShs&nbsp;&nbsp;<%= FormatNum(totalContractStamps) %></b></SPAN>
			<br><br><br>
			Revenue Stamps prepaid-->&nbsp;
			
		</td>
	</tr>
	</tr>	
	<tr>
		<td valign=top>
		<small><pre><font face=Arial size=1.5>
Please verify that all details are correct. AAMS must be alerted of any discrepancies within 48 hrs of receipt.<BR>TPIN No:&nbsp;&nbsp;20180816.</font>
</pre>
			</small>
		</td>
	</tr>
</table>
				<P class="newpage">
		
<%			
		slipNos = ""
		OrderRs.MoveNext
	Loop
	
	Set groupRs = Nothing
	Set rs = Nothing
	Set Conn = Nothing
	

	Function SortLevies(srcRs, orderRs)
		Dim returnArray()
		
		ReDim returnArray(srcRs.RecordCount, 4)
		
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
						returnArray(arrCounter, 4) = srcRs.Fields("LevyVATAmount").Value
				srcRs.MoveNext
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
						returnArray(arrCounter, 4) = srcRs.Fields("LevyVATAmount").Value
				srcRs.MoveNext
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
				returnArray(arrCounter, 4) = srcRs.Fields("LevyVATAmount").Value
				srcRs.MoveNext
				srcRs.MoveNext
			Loop	
			srcRs.MoveFirst
		End If
		
		SortLevies = returnArray
		
	End Function
%>

</body>

</html>