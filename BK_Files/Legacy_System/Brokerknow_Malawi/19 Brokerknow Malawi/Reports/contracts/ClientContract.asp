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
				margin-right: 5cm;
				margin-top: 1cm;    
				margin-bottom: 2cm;
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

<!--#include file="../libroutinesTEST.asp"-->
<%

genReport = Request("genReport")
selectedContractDate = Request("txtDate")
selectedClient = Request("cboClient")
chkclient = Request("chkclient")

if len(chkclient) = 0 then chkclient = 0

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
	<form method="POST" action="ClientContract.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">	
		<table align=center width="90%">
			<tr>
				<td width="20%">Select Date:</td>
				<td width="80%">
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>
			</tr>
			
			<tr>
				<td width="20%">
				&nbsp;<input value='1' type="checkbox" name="chkclient" id="chkclient">&nbsp;
				Client:</td>
				<td width="80%" nowrap valign="bottom"><input type = 'text' name ='txtClientCode' id = 'txtClientCode' size="10" onBlur="txtval = this.value; selectItem(cboClient);"> 
				&nbsp;&nbsp;&nbsp;
				<select name = "cboClient" id = "cboClient" size="1" 
    				onchange = "UpdateCode(true,cboClient,txtClientCode)"
					onKeypress = "return (dodefaultaction()==''); "  
					onKeydown = "return (dodefaultaction()==''); " 
					onKeyup = "change(cboClient,0);"
					<!--onKeyup = "return (change(cboClient,0));"-->
				  onfocus = "txtval = '';inputIsItemCode = 0;" 
					onblur = "txtval = '';inputIsItemCode = 0;">
					<option selected SearchCode = "" SearchText = ""  value = ""></option>
					<%
					        Set conn = GetActiveConnection("KBroker")
					        
					        sqlStr = "SELECT * FROM Client WHERE Deleted<>1 ORDER BY ClientName"
					        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
					        If Not (rs.EOF Or rs.BOF) Then
					                rs.MoveFirst
					                Do Until rs.EOF%>
					                        <option SearchCode = "<%=rs.Fields("Client_DPA_")%>" SearchText = "<%=rs.Fields("ClientName")%>"  value = '<%=rs.Fields("Client_DPA_")%>'><%=mid(rs.Fields("ClientName"),1,50)%></option>
					                        <%rs.MoveNext
					                Loop
					        End If
					%>

					    </select>
				</td>
			</tr>
			
			<tr>
				<td colspan=2 width="100%"><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... " id=Button1 name=Button1>&nbsp;&nbsp; </td>
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
	Set Rs = CreateObject("ADODB.Recordset")						        
	
	if trim(chkclient) = 1 then
		sqlStr = "SELECT * FROM SaleContracts WHERE LotTDate = '" & FormatDate(selectedContractDate) & "' AND Client_DPA_ = "& selectedClient &" ORDER BY ContractNumber"		
	else
		sqlStr = "SELECT * FROM SaleContracts WHERE LotTDate = '" & FormatDate(selectedContractDate) & "' ORDER BY ContractNumber"		
	end if
	
	Rs.CursorLocation = adUseClient	
	Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
	
	If rs.EOF Or rs.BOF Then%>
		<Script Language="JavaScript">
			alert("There are no contracts based on the specified criterion.")
			window.parent.history.go(-1);			
		</Script>
		<%Set Rs = Nothing
		Set Conn = Nothing
		Response.End
	End If
	
	Set groupRs = Conn.Execute("SELECT ContractNumber FROM SaleContracts WHERE LotTDate = '" & FormatDate(selectedContractDate) & "' GROUP BY ContractNumber")
	
	if trim(chkclient) = 1 then
		Set groupRs = Conn.Execute("SELECT ContractNumber FROM SaleContracts WHERE LotTDate = '" & FormatDate(selectedContractDate) & "' AND Client_DPA_ = "& selectedClient &" GROUP BY ContractNumber")
	else
		Set groupRs = Conn.Execute("SELECT ContractNumber FROM SaleContracts WHERE LotTDate = '" & FormatDate(selectedContractDate) & "' GROUP BY ContractNumber")
	end if

	
	Set levyOrderRs = Conn.Execute("SELECT * FROM LevyOrderList ORDER BY LevyOrder")
	
	Dim pageNumber
	
	pageNumber = 0

	Do Until groupRs.EOF
	
	Rs.Filter = "ContractNumber = '" & groupRs.Fields("ContractNumber").Value & "'"
	
	totalGross = Rs.Fields("LotGrossAmount").Value
	
	sqlStr = "SELECT     tbOrder.*, Client.ClientAddr,Client.ClientCDSNo FROM Lot INNER JOIN " & _
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
		clientCDSNo = temprs.Fields("ClientCDSNo").Value
	End If
	
	Set temprs = Nothing
	
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
		<td align="center" colspan=2 height="250px">
			<Img Src="../data/photos/aaprintlogo.jpg">			
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

<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" align ="center" width="90%"> 
		
	<tr>
		<td nowrap valign="top" width="70%">			
			<table cellspacing=0 cellpadding=0 border=0>
				<tr>
					<td nowrap><%= Rs.Fields("OrdDetailClient").Value %>&nbsp;&nbsp;[<%= Rs.Fields("Client_DPA_").Value %>]  [<%=clientCDSNo%>] </td>					
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
				<!--<tr>
					<td nowrap>CDS Ref:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%= Rs.Fields("LotSlipNo").Value %></td>					
				</tr>-->
				<tr>
					<td  colspan="2" nowrap>&nbsp;</td>					
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
	<tr>
		<td colspan=2>&nbsp;</td>
	</tr>
</table>	

<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" align ="center" width="90%">	
	
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
						levyPerc = round((levyArray(i, 1) / totalGross) * 100,2)
						levyPerc = FormatNumber(levyPerc,2) & " %"
						levyPerc=replace(levyArray(i, 3),"%"," %")
						
						'do not display agent commission (staff commission too) HARD CODING!!!!!!
						If (levyArray(i, 2) = 11) and (levyArray(i, 1) <= 50) Then
						'response.write levyArray(i, 2) & "::" & levyArray(i, 1):response.end
							levyPerc = "Minimum"
						End If

						' Vat Percentage and MSE percentage
						if levyArray(i, 2) = 99 then
							levyPerc="16.5%"
							thisLevyName = ""
						elseif levyArray(i, 2) = 25  then
							levyPerc="5 %"
						end if	


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
							TotalVat = TotalVat + levyArray(i, 4)
							VatInclusive = levyArray(i, 1) + levyArray(i, 4)
							totalLevies = totalLevies + VatInclusive
					%>
							<TR>		
								<td nowrap width="10%" STYLE="PADDING-LEFT: 0PX"><%= thisLevyName %></td>
								<td ALIGN="LEFT" width="40%">&nbsp;&nbsp;<%= levyPerc %></td>
								<td ALIGN="Right" width="20%">&nbsp;&nbsp;<%= FormatNum(levyArray(i, 4)) %></td>
								<td align=right width="30%"><%= FormatNum(VatInclusive) %></td>		
							</TR>
					<%	End If
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
	</tr>
	
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
		<td>TOTAL AMOUNT PAYABLE IN MWK</td>
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
			
		</td>Client Verification disclaimer:-->
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
	
<%
		Rs.Cancel
	groupRs.MoveNext
	'important!
		If Not groupRs.EOF Then %>
			<P class="newpage">
	<%	End If
		
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
			Loop	
			srcRs.MoveFirst
		End If
		
		SortLevies = returnArray
		
	End Function
%>


</body>
</html>



