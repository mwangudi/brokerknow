<html>

<head>

<title>Client Contract</title>  

<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
<SCRIPT language=VBScript src="../scripts/reports.vbs"></SCRIPT>
<SCRIPT language=Javascript src="../scripts/reports.js"></SCRIPT>

<style media="print">
		@page {
				size: landscape;
				margin-left: 0cm;
				margin-right: 0cm;
				margin-top: 0cm;    
				margin-bottom: 0cm;
				writing-mode: tb-rl;
				br.newpage{
					page-break-before:always;
				}		
			}		 
	</style>
</head>

<body Class="Reports">
<!--#include file="../libroutinesVB.asp"-->
<form name=frm1 id=frm1>
	<%
	selectedClient = Request.QueryString("a")
	selectedContractDate = Request.QueryString("b")
	selectedEndContractDate = Request.QueryString("c")
	
	Set conn = GetActiveConnection("KBroker")
 	Set Rs = CreateObject("ADODB.Recordset")						        

	'sqlStr = "SELECT * FROM SaleContracts WHERE LotTDate = '" & FormatDate(selectedContractDate) & "' ORDER BY ContractNumber"		
	
	If Len(selectedEndContractDate) = 0 Then
		'DAILY
		sqlStr = "SELECT SaleContracts.*, SaleContracts.ContractNumber as ContractNumber" & _
			" FROM SaleContracts INNER JOIN" & _
			" Lot ON SaleContracts.Lot_DPA_ = Lot.Lot_DPA_ INNER JOIN" & _
			" OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN" & _
			" tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_" & _
			" WHERE (tbOrder.Client_DPA_ = "& selectedClient &")" & _
			" AND SaleContracts.LotTDate = '" & FormatDate(selectedContractDate) & "' ORDER BY SaleContracts.ContractNumber"
	Else
		'WEEKLY, MONTHLY, QUARTERLY
		sqlStr = "SELECT SaleContracts.*, SaleContracts.ContractNumber as ContractNumber" & _
			" FROM SaleContracts INNER JOIN" & _
			" Lot ON SaleContracts.Lot_DPA_ = Lot.Lot_DPA_ INNER JOIN" & _
			" OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN" & _
			" tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_" & _
			" WHERE (tbOrder.Client_DPA_ = "& selectedClient &")" & _
			" AND SaleContracts.LotTDate BETWEEN '" & FormatDate(selectedContractDate) & "' AND '" & FormatDate(selectedEndContractDate) & "' ORDER BY SaleContracts.ContractNumber"
	End If
	
	'Response.Write sqlstr
	'Response.End 
	
	Rs.CursorLocation = adUseClient	
	Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
	
	If rs.EOF Or rs.BOF Then
		''NO CLIENT RECORDS
		Set Rs = Nothing
		Set Conn = Nothing
		Response.End
	Else
		If Len(selectedEndContractDate) = 0 Then
			'DAILY
			sqlStr = "SELECT DISTINCT SaleContracts.ContractNumber AS THEContractNumber" & _
				" FROM SaleContracts INNER JOIN" & _
				" Lot ON SaleContracts.Lot_DPA_ = Lot.Lot_DPA_ INNER JOIN" & _
				" OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN" & _
				" tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_" & _
				" WHERE (tbOrder.Client_DPA_ = "& selectedClient &")" & _
				" AND SaleContracts.LotTDate = '" & FormatDate(selectedContractDate) & "' ORDER BY SaleContracts.ContractNumber"
		Else
			'WEEKLY, MONTHLY, QUARTERLY
			sqlStr = "SELECT DISTINCT SaleContracts.ContractNumber AS THEContractNumber" & _
				" FROM SaleContracts INNER JOIN" & _
				" Lot ON SaleContracts.Lot_DPA_ = Lot.Lot_DPA_ INNER JOIN" & _
				" OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN" & _
				" tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_" & _
				" WHERE (tbOrder.Client_DPA_ = "& selectedClient &")" & _
				" AND SaleContracts.LotTDate BETWEEN '" & FormatDate(selectedContractDate) & "' AND '" & FormatDate(selectedEndContractDate) & "' ORDER BY SaleContracts.ContractNumber"
		End If
		
		'Response.Write sqlstr
		'Response.End 
		
		'Set groupRs = Conn.Execute("SELECT ContractNumber FROM SaleContracts WHERE LotTDate = '" & FormatDate(selectedContractDate) & "' GROUP BY ContractNumber")
		Set groupRs = Conn.Execute(sqlStr)
	
		Set levyOrderRs = Conn.Execute("SELECT * FROM LevyOrderList ORDER BY LevyOrder")
	
		Dim pageNumber
	
		pageNumber = 0
		
		If Not groupRs.EOF or groupRs.BOF Then
		
		Do Until groupRs.EOF
	
				Rs.Filter = "ContractNumber = '" & groupRs.Fields("THEContractNumber").Value & "'"
	
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
					clientAddress = Replace(temprs.Fields("ClientAddr").Value, vbCrLf, ",")		
				End If
	
				Set temprs = Nothing
	
				If Trim(UCase(Rs.Fields("OrdDetailType").Value)) = "PURCHASE" Then
					IsPurchase = True
				Else
					IsPurchase = False
				End If
	
				pageNumber = pageNumber + 1
				
				
				If len(selectedEndContractDate) = 0 then
				If pageNumber > 1 and groupRs.EOF = false Then 
					%>
					<br class="newpage">
					<%
				End If
				end if
				%>
				
				<table border=0 cellspacing=2 cellpadding=2 align=center width="90%">
					<tr>
						<td align=center><img src="aaprintlogo.jpg" width="482" height="178"></td>
					</tr>
				</table>
		
				<table border="0" cellspacing=2 cellpadding=2 style="FONT-SIZE: 10pt;BACKGROUND: #FFFFFF;FONT-FAMILY: Arial, Arial Narrow, Tahoma, Verdana;" align ="center" width="90%"> 
					<THEAD>
					<tr class="pageNumbering">
						<td align="left" colspan=2>
							<FONT FACE=ARIAL SIZE=2><B>Page <%=pageNumber%>	</B></FONT>	
						</td>		
					</tr>
					<THEAD>  
					<tr>
						<td colspan="2"  align="center">
							<FONT FACE=ARIAL SIZE=3><B>
								<%= UCase(Rs.Fields("OrdDetailType").Value) %> CONTRACT <%= Rs.Fields("ContractNumber").Value %>
							</B></FONT>
						</td>		
					</tr>  
					<tr>
						<td colspan=2>&nbsp;</td>
					</tr>  
				</table>
				
				<table border="0" cellspacing=2 cellpadding=2 style="FONT-SIZE: 10pt;BACKGROUND: #FFFFFF;FONT-FAMILY: Arial, Arial Narrow, Tahoma, Verdana;" align ="center" width="90%"> 
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
									<td nowrap>Date:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%= FormatDate(Rs.Fields("LotTDate").Value) %></td>
								</tr>
								<tr>
									<td nowrap>Order Ref:&nbsp;&nbsp;&nbsp;<%= orderRef %></td>					
								</tr>
								<tr>
									<td nowrap>REF Ref:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%= Rs.Fields("LotSlipNo").Value %></td>					
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
						<td colspan=2><PRE><font face=Arial size=2>We wish to advise that we have <% If IsPurchase Then Response.Write "bought" Else Response.Write "sold"  %> the following securities as per your instruction. Kindly arrange to complete the transaction as per the details given below: </font></PRE></td>
					</tr>
					<tr>
						<td colspan=2 style="BORDER-BOTTOM: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent">&nbsp;</td>
					</tr>
				</table>	

				<table border="0" cellspacing=2 cellpadding=2 style="FONT-SIZE: 10pt;BACKGROUND: #FFFFFF;FONT-FAMILY: Arial, Arial Narrow, Tahoma, Verdana;" align ="center" width="90%">	
					
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
						<%end if%>
						  </td>
					</tr>
				    <tr>
						<td>Gross</td>
						<td align=right><b> <%= FormatNum(totalGross) %> </b> </td>
					</tr>
					<tr>
						<TD COLSPAN=2 VALIGN="TOP">
							<TABLE align ="center" width="100%" CELLSPACING="0" CELLPADDING="4">
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
										If (levyArray(i, 2) = 11) and (levyArray(i, 1) <= 100) Then
											levyPerc = "Minimum"
										End If
										
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
									<%	End If
									Next%>
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
						<td>TOTAL AMOUNT PAYABLE IN KSHS</td>
						<td align=right><b>
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
						<td colspan=2>&nbsp;</td>
					</tr>
					
					<tr>
						<td colspan=2>&nbsp;</td>
					</tr>
					<tr>
						<td>
							<font face=Arial size=2>
							<i>For and on behalf of</i><br><br>
							African Alliance Kenya Securities<br>
							<br>
							Sign..........................................
							</font></pre>			
						</td>
						<td align="left" valign="top">			
							<img src="stamp.gif" border="0" style="position: absolute;z-index: 2">
							<BR>
							<SPAN style="position: absolute;z-Index: 10"><b>KShs&nbsp;&nbsp;<%= FormatNum(totalContractStamps) %></b></SPAN>
							<br>
							Revenue Stamps prepaid
						</td>
					</tr>
				</table>
				</div>
				<%
				Rs.Cancel
			groupRs.MoveNext
			
			If len(selectedEndContractDate) > 0 then
			If pageNumber > 1 and groupRs.EOF = false Then 
				%>
				<br class="newpage">
				<%
			End If
			end if
		Loop
	
	End if
	
end if
	
	Set groupRs = Nothing
	Set rs = Nothing
	Set Conn = Nothing
	
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
%>

<input type="hidden" value="" id="hidData" name="hidData">
<input type="hidden" value="<%=selectedClient%>" id="hidClient" name="hidClient">
<input type="hidden" value="Contract" id="hidCategory" name="hidCategory">
	
</form>
</body>
</html>

<% Response.End %>
<script language="vbscript">
	theData = frm1.innerHTML 
	
	frm1.hidData.value = theData
	frm1.method = "post"
	frm1.action = "PDFMail.asp"
	frm1.submit 
</script>

