<!--#include file="../libroutinesTEST.asp"-->
<html>
<head>
	<meta http-equiv="Content-Language" content="en-uk">
	<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
	<title>Sale Contract</title>  
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
				margin-top: 0cm;    
				margin-bottom: 0cm;
				size: portrait;
				
				br.newpage{
					page-break-before: always;
				}
				
				tr.pageNumbering{
					display:none;
				}
			}
	
		</style>
</head>

<body Class="Reports">

<%

genReport = Request.Form("genReport")
selectedContractDate = Request.Form("txtDate")

If genReport <> "1" Or selectedContractDate = "" Then%>
		<Script Language="JavaScript">
		report_SetBodyClass();
		function validateForm(frm){			
			if (frm.txtDate.value==''){
				alert("Select a date");
				frm.txtDate.focus();
				return;
			}
			
			frm.target = '_self';			
			frm.submit();
		}
		
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "txtDate","cmdDate","<%= FormatDate(Date) %>",1);
	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="ClientCompounded.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">	
		<table>
			<tr>
				<td>Select date:</td>
				<td>
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>
			</tr>
			
			<tr>
				<td colspan=2><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... ">&nbsp;&nbsp;</td>
			</tr>
		</table>
		
	</form>
	
	<%
	Response.End
End If

%>

<% DrawPageFunctions True, True, True, True %>

<p id="toPDFOrient" name="toPDFOrient" value="P" style="display:none;">P
<p id="toPDF" name="toPDF">

<%
	Set conn = GetActiveConnection("KBroker")
	Set groupRs = CreateObject("ADODB.Recordset")						        
	sqlStr = "SELECT * FROM ClientContractCompounded WHERE (LotTDate = '" & FormatDate(selectedContractDate) & "') ORDER BY LotSlipNo"
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
	sqlStr = "SELECT DISTINCT Order_DPA_ FROM ClientContractCompounded WHERE (LotTDate = '" & FormatDate(selectedContractDate) & "') ORDER BY Order_DPA_"
	
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
	
	blnEndFound = False
	Dim pageNumber
	
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
					  "AS ContractNumber, OrdDetailSecurity, SUM(LotQty) AS LotQty, SUM(LotPrice) AS LotPrice, OrdDetailClient, SystemMaintained, LevyRatePercentage, Client_DPA_ " & _
					  "FROM         dbo.ClientContractCompounded " & _
					  "WHERE     (LotSlipNo IN (" & slipNos  & ")) And LevyName <> '' AND OrdDetailClient = '" & prevClient & "' AND (Order_DPA_ = " & OrderRs.Fields("Order_DPA_").Value & ")" & _
					  "GROUP BY LevyName, LevyShortName, OrdDetailSecurity, OrdDetailClient, SystemMaintained, LevyRatePercentage,OrdDetailSecType,Order_DPA_, Client_DPA_"

			
			'Response.Write(sqlStr)
			'Response.end			
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
<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" align ="center" width="90%">
	<THEAD>
	<tr class="pageNumbering">
		<td align="left" colspan=2>
			<FONT FACE=ARIAL SIZE=2><B>Page <%=pageNumber%>	</B></FONT>	
		</td>		
	</tr>
	<tr>
		<td align="center" colspan=2 height="200px">
			<Img Src="../data/photos/aaprintlogo.jpg">			
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
</table>

<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" align ="center" width="90%"> 
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
					<td  nowrap colspan="2">Date:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%= FormatDate(selectedContractDate) %></td>
				</tr>
				<tr>
					<td  colspan="2" nowrap>Order Ref:&nbsp;&nbsp;&nbsp;<%= orderRef %></td>					
				</tr>
				
			</table>
		</td>
	</tr>
	<tr>
	<!--<tr>
		<td colspan=2>
		      <table cellspacing="0" cellpadding="0" border="0"  align ="center" width="100%">
				<tr>
				<td nowrap valign="top" width="10%">CDS Ref:&nbsp;&nbsp;&nbsp;</td>
				<td valign="top" width="90%">
						<table border="0" cellpadding="1" cellspacing="0" align ="center" width="90%">
										
									<%
									If InStr(1, slipNos, ",") > 0 Then
										Dim slipNosArray
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
	</tr>-->
	<tr>
		<td  colspan="2" nowrap>&nbsp;</td>					
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

<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable"  align ="center" width="90%">	
				
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
					
					dim RST
					set RST = server.CreateObject("ADODB.Recordset")
					
					'careful, returned array is one-based
					For i = 1 To UBound(levyArray) 
						
						''RECALCULATE THE COMMISSION PERCENTAGE WHEN UPPER LIMIT > 100,000
						'SQL = "SELECT Commission.CommissionRate, Commission.UpperSecurityCommission, Commission.SecurityBoundary" & _
						'" FROM Client INNER JOIN Commission ON Client.Commission_DPA_ = Commission.Commission_DPA_" & _
						'" WHERE (Client.Client_DPA_ = "& Rs("Client_DPA_") &")"
						'RST.CursorLocation = adUseClient	
						'RST.Open SQLServerFormat(HandleQuote(SQL)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
					
						'ReDonePerc = FormatNum(0)
						'If Not (RST.EOF Or RST.BOF) Then
							'SecurityBoundary = RST("SecurityBoundary")
							'If totalGross > FormatNum(SecurityBoundary) then
								'ReDonePerc = FormatNum(RST("UpperSecurityCommission")) & "%"
							'End If
						'Else
							'SecurityBoundary = totalGross
						'End If
						
						thisLevyName = Trim(levyArray(i, 0))
						If levyArray(i, 2) = 10 Then
							totalContractStamps = levyArray(i, 1)
						End If
									
						levyPerc = levyArray(i, 3)
						If InStr(1, levyPerc, "%") > 0 Then
							'recalculate average percentage
							levyPerc = (levyArray(i, 1) / totalGross) * 100
							levyPerc = FormatNumber(levyPerc,3) & "%"
						End If
						if thisLevyName ="Value Added Tax" then
							levyPerc="10 %"
						end if			
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
								<%
								'If  (levyArray(i, 2) = 11) Then
									'Response.Write " [ " & thisLevyname & " ]**"
									'Response.write " levyPerc" & levyPerc & "**"
									'Response.write " ReDonePerc" & ReDonePerc & "**"
									'Response.write " totalGross" & totalGross & "**"
									'Response.write " SecurityBoundary" & SecurityBoundary
									
									'If levyPerc > ReDonePerc And totalGross >= SecurityBoundary Then
									'	thelevyPerc = ReDonePerc
									'Else
									'	thelevyPerc = levyPerc
									'End If
								'Else
								'	thelevyPerc = levyPerc
								'End If
								
								'If thelevyPerc = 0 Then
								'	thelevyPerc = levyPerc
								'End If
								%>
								
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
		<td>TOTAL AMOUNT PAYABLE IN BWP</td>
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
	<tr>
		<td colspan=2>&nbsp;</td>
	</tr>
	<tr>
		<td colspan=2>&nbsp;</td>
	</tr>
	<tr>
		<td colspan=2>&nbsp;</td>
	</tr>
	<tr>
		<td><font face=Arial size=2>
		<i>For and on behalf of</i>
		<br>African Alliance Malawi Securities
		<br><br>Sign..........................................</font>
		</td>
		<!--<td align="left" valign="top">			
			<img src="../images/stamp.gif" border="0" style="position: absolute;z-index: 2">
			<BR>
			&nbsp;<SPAN style="position: absolute;z-Index: 10"><b>KShs&nbsp;&nbsp;<%= FormatNum(totalContractStamps) %></b></SPAN>
			<br><br><br>
			Revenue Stamps prepaid
			
		</td>-->
		<td>&nbsp;</td>
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
				<br class="newpage">
						
					
<%			
			slipNos = ""
			OrderRs.MoveNext
	Loop
	
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

</body>

</html>
