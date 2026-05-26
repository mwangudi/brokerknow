<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Batch Report [Portrait]</title>
  
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

'genReport = Request.Form("genReport")
'selectedContractDate = Request.Form("txtDate")

genReport = Request.QueryString("genReport")
selectedContractDate = Request.QueryString("selectedContractDate")

'Response.Write selectedContractDate
'Response.End
	
If genReport <> "1" Or selectedContractDate = "" Then%>
		<Script Language="JavaScript">
		report_SetBodyClass();
		function validateForm(frm){			
			if (frm.txtDate.value==''){
				alert("Select a date");
				frm.txtDate.focus();
				return;
			}
			
			frm.action = 'BatchReportsPortrait.asp?genReport=1&selectedContractDate='+frm.txtDate.value;
			frm.target = '_self';			
			frm.submit();
		}
		
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "txtDate","cmdDate","<%= FormatDate(Date) %>",1);
	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<form method="POST" action="BatchReportsPortrait.asp" Name="frmMain" id="frmMain">
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

<%
	Set conn = GetActiveConnection("KBroker")
	Set Rs = CreateObject("ADODB.Recordset")						        
	sqlStr = "SELECT * FROM SaleContracts WHERE LotTDate = '" & FormatDate(selectedContractDate) & "' ORDER BY ContractNumber"		
	
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
	
	If groupRs.EOF Or groupRs.BOF Then%>
		<Script Language="JavaScript">
			alert("There are no contracts based on the specified criterion.")
			window.parent.history.go(-1);			
		</Script>
		<%Set groupRs = Nothing
		Set Conn = Nothing
		Response.End
	End If
	
	Set levyOrderRs = Conn.Execute("SELECT * FROM LevyOrderList ORDER BY LevyOrder")
	
	If levyOrderRs.EOF Or levyOrderRs.BOF Then%>
		<Script Language="JavaScript">
			alert("There are no contracts based on the specified criterion.")
			window.parent.history.go(-1);			
		</Script>
		<%Set levyOrderRs = Nothing
		Set Conn = Nothing
		Response.End
	End If
	
	Dim pageNumber
	
	pageNumber = 0

	Do Until groupRs.EOF
	
	Rs.Filter = "ContractNumber = '" & groupRs.Fields("ContractNumber").Value & "'"
	
	totalGross = Rs.Fields("LotGrossAmount").Value
	
	sqlStr = "SELECT     tbOrder.*, Client.ClientAddr FROM Lot INNER JOIN " & _
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

<%If pageNumber=1 Then%>
	<p id="toPDFOrient" name="toPDFOrient" value="P" style="display:none;">P
	<p id="toPDF" name="toPDF">
	<P class="newpage">
	<BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR>
	<P align="center"><B>CONTRACTS</B></P>
	<BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR>
	<P class="newpage">
<%End If%>

<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" align ="center" width="60%"> 
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

<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" align ="center" width="60%"> 
		
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
					<td nowrap>CDS Ref:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%= Rs.Fields("LotSlipNo").Value %></td>					
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
		<td colspan=2><PRE><font face=Arial size=2>We wish to advise that we have <% If IsPurchase Then Response.Write "bought" Else Response.Write "sold"  %> the following securities as per your instruction. Kindly arrange
to complete the transaction as per the details given below: </font></PRE></td>
	</tr>
	<tr>
		<td colspan=2 style="BORDER-BOTTOM: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent">&nbsp;</td>
	</tr>
	<tr>
		<td colspan=2>&nbsp;</td>
	</tr>
</table>	

<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" align ="center" width="600px">	
	
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
					<%	End If
					Next%>
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
		<PRE><font face=Arial size=2>
       <i>For and on behalf of</i>
African Alliance Malawi Securities

Sign..........................................
		</FONT></PRE>			
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
	'important!
		If Not groupRs.EOF Then %>
			<P class="newpage">
			<%Else%>
			<P class="newpage">
	<%	End If
		
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

'Response.End 

' THIS IS CODE FOR THE CONTRACT NOTE COMPOUNDED 
dim  sqlStr2
Set conn = GetActiveConnection("KBroker")
Set Rs = CreateObject("ADODB.Recordset")	
Set groupRs = CreateObject("ADODB.Recordset")	
	
sqlStr = "SELECT * FROM ClientContractCompounded WHERE (LotTDate = '" & FormatDate(selectedContractDate) & "') ORDER BY LotSlipNo"

	'groupRs.CursorLocation = adUseClient	
	groupRs.Open sqlStr, conn.ConnectionString, 0, 1

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
			alert("There are no compounded contracts based on the specified criterion.")
			//window.parent.history.go(-1);			
		</Script>
		<%Set groupRs = Nothing
		Set Conn = Nothing
		'Response.End
	End If	
	
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

<%If pageNumber=1 Then%>
	<BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR>
	<P align="center"><B>CONTRACT NOTES COMPOUNDED</B></P>
	<BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR>
	<BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR>
	<P class="newpage">
<%End If%>

<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" align ="center" width="60%">
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

<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" align ="center" width="60%"> 
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
	<tr>
		<td colspan=2>
		      <table cellspacing="0" cellpadding="0" border="0"  align ="center" width="100%">
				<tr>
				<td nowrap valign="top" width="10%">CDS Ref:&nbsp;&nbsp;&nbsp;</td>
				<td valign="top" width="60%">
						<table border="0" cellpadding="1" cellspacing="0" align ="center" width="100%">
										
									<%
									If InStr(1, slipNos, ",") > 0 Then
										'Dim slipNosArray
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
		<td colspan=2><PRE><font face=Arial size=2>We wish to advise that we have <% If IsPurchase Then Response.Write "bought" Else Response.Write "sold"  %> the following securities as per your instruction. Kindly arrange
to complete the transaction as per the details given below: </font></PRE></td>
	</tr>
	
	<tr>
		<td colspan=2>&nbsp;</td>
	</tr>
</table>	

<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable"  align ="center" width="600px">	
				
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
		<td>
		<PRE><font face=Arial size=2>
       <i>For and on behalf of</i>
African Alliance Malawi Securities

Sign..........................................
		</FONT></PRE>	
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

	
'' AGENT CONTRACTS

	Set conn = GetActiveConnection("KBroker")
	Set Rs = CreateObject("ADODB.Recordset")
	Set groupRs = CreateObject("ADODB.Recordset")
	sqlStr = "SELECT * FROM AgentContracts WHERE LotTDate = '" & FormatDate(selectedContractDate) & "' ORDER BY ContractNumber"		
	
	
	Rs.CursorLocation = adUseClient	
	Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
	
	If rs.EOF Or rs.BOF Then%>
		<Script Language="JavaScript">
			alert("There are no agent contracts based on the specified criterion.")
		//	window.parent.history.go(-1);			
		</Script>
		<%Set Rs = Nothing
		Set Conn = Nothing
		'Response.End
	End If
	
	Set groupRs = Conn.Execute("SELECT ContractNumber FROM AgentContracts WHERE LotTDate = '" & FormatDate(selectedContractDate) & "' GROUP BY ContractNumber ")
	
	Set levyOrderRs = Conn.Execute("SELECT * FROM LevyOrderList ORDER BY LevyOrder")
	pageNumber = 0
	Do Until groupRs.EOF
	
		Rs.Filter = "ContractNumber = '" & groupRs.Fields("ContractNumber").Value & "'"
	
	totalGross = Rs.Fields("LotGrossAmount").Value
	
	sqlStr = "SELECT     tbOrder.*, Client.ClientAddr FROM Lot INNER JOIN " & _
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

<%If pageNumber=1 Then%>
	<P class="newpage">
	<BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR>
	<P align="center"><B>AGENT CONTRACTS</B></P>
	<BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR>
	<P class="newpage">
<%End If%>

<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" align ="center" width="60%"> 
	<THEAD>
	<tr class="pageNumbering">
		<td align="left" colspan=2>
			<FONT FACE=ARIAL SIZE=2><B>Page <%=pageNumber%>	</B></FONT>	
		</td>		
	</tr>
	<tr>
		<td align="center" colspan=2 height="190px">
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
	<tr>
		<td nowrap width="80%" valign="top">			
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
					<td nowrap>Date:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%= FormatDate(Rs.Fields("LotTDate").Value) %></td>
				</tr>
				<tr>
					<td nowrap>Order Ref:&nbsp;&nbsp;&nbsp;<%= orderRef %></td>					
				</tr>
				<tr>
					<td nowrap>CDS Ref:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%= Rs.Fields("LotSlipNo").Value %></td>					
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
		<td colspan=2><PRE><font face=Arial size=2>We wish to advise that we have <% If IsPurchase Then Response.Write "bought" Else Response.Write "sold"  %> the following securities as per your instruction. Kindly arrange
to complete the transaction as per the details given below: </font></PRE></td>
	</tr>
	
	<tr>
		<td colspan=2>&nbsp;</td>
	</tr>
</table>	
<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" align ="center" width="600px">	
	
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
		levyArray = SortLevies3(rs, levyOrderRs)
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
			<TABLE align ="center" width="60%" CELLSPACING="0" CELLPADDING="4">
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
			<% 'Dim netAmount
				
				If IsPurchase Then
					netAmount = grossAmount - AgentCommission				
				Else
					netAmount = grossAmount + AgentCommission
				End If
			%>
			
			 <%=FormatNum(netAmount) %>
			
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
			<img src="../images/stamp.gif" border="0" style="position: absolute;z-index: 2">
			<BR>
			&nbsp;<SPAN style="position: absolute;z-Index: 10"><b>KShs&nbsp;&nbsp;<%= FormatNum(totalContractStamps) %></b></SPAN>
			<br><br>
			Revenue Stamps prepaid
			
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
	
	Function SortLevies3(srcRs, orderRs)
		Dim returnArray()
		
		ReDim returnArray(srcRs.RecordCount, 3)
		
		'careful, recordcount is 1- based
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
		
		SortLevies3 = returnArray
		
	End Function


''AGENT CONTRACT COMPOUNDED

	Set conn = GetActiveConnection("KBroker")
	Set groupRs = CreateObject("ADODB.Recordset")						        
	sqlStr = "SELECT * FROM AgentContractCompounded WHERE (LotTDate = '" & FormatDate(selectedContractDate) & "') ORDER BY LotSlipNo"
	
	groupRs.CursorLocation = adUseClient	
	groupRs.Open sqlStr, conn.ConnectionString, adOpenKeyset, adLockOptimistic
		
	If groupRs.EOF Or groupRs.BOF Then%>
		<Script Language="JavaScript">
			//alert("There are no agent compounded contracts based on the specified criterion.")
			//window.parent.history.go(-1);			
		</Script>
		<%Set groupRs = Nothing
		Set Conn = Nothing
		'Response.End
	End If	
	
	Set OrderRs = CreateObject("ADODB.Recordset")						        
	sqlStr = "SELECT DISTINCT Order_DPA_ FROM AgentContractCompounded WHERE (LotTDate = '" & FormatDate(selectedContractDate) & "') ORDER BY Order_DPA_"
	OrderRs.CursorLocation = adUseClient	
	OrderRs.Open sqlStr, conn.ConnectionString, adOpenKeyset, adLockOptimistic
		
	If OrderRs.EOF Or OrderRs.BOF Then%>
		<Script Language="JavaScript">
			//alert("There are no contracts based on the specified criterion.")
		//	window.parent.history.go(-1);			
		</Script>
		<%Set groupRs = Nothing
		Set Conn = Nothing
		'Response.End
	else
	
	Set levyOrderRs = Conn.Execute("SELECT * FROM LevyOrderList ORDER BY LevyOrder")
	'Dim pageNumber
	
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

<%If pageNumber=1 Then%>
	<P class="newpage">
	<BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR>
	<P align="center"><B>AGENT CONTRACT COMPOUNDED</B></P>
	<BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR>
	<BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR>
	<P class="newpage">
<%End If%>

<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" width="60%" align="center">
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
					<td  nowrap colspan="2">Date:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%= FormatDate(selectedContractDate) %></td>
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
		<td colspan=2><PRE><font face=Arial size=2>We wish to advise that we have <% If IsPurchase Then Response.Write "bought" Else Response.Write "sold"  %> the following securities as per your instruction. Kindly arrange
to complete the transaction as per the details given below: </font></PRE></td>
	</tr>
	<tr>
		<td colspan=2 style="BORDER-BOTTOM: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent">&nbsp;</td>
	</tr>
	<tr>
		<td colspan=2>&nbsp;</td>
	</tr>
</table>	

<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable"  width="600px" align="center">	
	
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
		levyArray = SortLevies6(rs, levyOrderRs)
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
		
			<img src="../images/stamp.gif" border="0" style="position: absolute;z-index: 2">
			<BR>
			&nbsp;<SPAN style="position: absolute;z-Index: 10"><b>KShs&nbsp;&nbsp;<%= FormatNum(totalContractStamps) %></b></SPAN>
			<br><br><br>
			Revenue Stamps prepaid
			
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
	

	Function SortLevies6(srcRs, orderRs)
		Dim returnArray2()
		
		ReDim returnArray2(srcRs.RecordCount, 3)
		
		'careful, recordcount is 1-based
		arrCounter1 = 0 
		If Not (orderRs.EOF Or orderRs.BOF) Then
			Do Until orderRs.EOF
				isExistingLevy1 = False
				Do Until srcRs.EOF
					If StrComp(orderRs.Fields("LevyName").Value, srcRs.Fields("LevyShortName").Value, vbTextCompare) = 0 Then
						arrCounter1 = arrCounter1 + 1
						returnArray2(arrCounter1, 0) = srcRs.Fields("LevyName").Value
						returnArray2(arrCounter1, 1) = srcRs.Fields("LevyAmount").Value
						returnArray2(arrCounter1, 2) = srcRs.Fields("SystemMaintained").Value
						returnArray2(arrCounter1, 3) = srcRs.Fields("LevyRatePercentage").Value
						Exit Do	
					End If
					
					srcRs.MoveNext
				Loop	
				
				srcRs.MoveFirst					
				orderRs.MoveNext
			Loop	
			
			orderRs.MoveFirst
			
			If arrCounter1 <> UBound(returnArray2) Then
				'take care of the rest unordered items
				Do Until srcRs.EOF
					thisName = srcRs.Fields("LevyName").Value
					
					'check if it's in array first
					existsInArray1 = False
					For k = 1 To arrCounter1
						If StrComp(thisName, returnArray2(k, 0), vbTextCompare) = 0 Then
							existsInArray1 = True
							Exit For
						End If
					Next
					
					If Not existsInArray1 Then
						arrCounter1 = arrCounter1 + 1
						returnArray2(arrCounter1, 0) = srcRs.Fields("LevyName").Value
						returnArray2(arrCounter1, 1) = srcRs.Fields("LevyAmount").Value
						returnArray2(arrCounter1, 2) = srcRs.Fields("SystemMaintained").Value	
						returnArray2(arrCounter1, 3) = srcRs.Fields("LevyRatePercentage").Value	
					End If
					
					srcRs.MoveNext
				Loop	
				srcRs.MoveFirst
			End If
			
			
		Else
			Do Until srcRs.EOF
				arrCounter1 = arrCounter1 + 1
				returnArray2(arrCounter1, 0) = srcRs.Fields("LevyName").Value
				returnArray2(arrCounter1, 1) = srcRs.Fields("LevyAmount").Value
				returnArray2(arrCounter1, 2) = srcRs.Fields("SystemMaintained").Value		
				returnArray2(arrCounter1, 3) = srcRs.Fields("LevyRatePercentage").Value
				srcRs.MoveNext
			Loop	
			srcRs.MoveFirst
		End If
		
		SortLevies6 = returnArray2
		
	End Function



End If

''SETTLEMENT SLIPS

'response.write "SETTLEMENT SLIPS"

selectedTradeDate =selectedContractDate
 Set conn = GetActiveConnection("KBroker")
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
			
    Else%>
		<Script Language="JavaScript">
			alert("The time limits of type Settlement (Brokers) have not been set");
			window.parent.self.close();				
        </Script>
		<%
		Set tempRs = Nothing
		Set Conn = Nothing
		Response.End
    End If
        
    Set tempRs = Nothing
    
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
        If rs.EOF Or rs.BOF Then
               %>
                <Script Language="JavaScript">
					//alert("No settlement slips available");
					//window.parent.history.back();				
                </Script>
                <% Set Rs = Nothing
                Set Conn = Nothing
                'Response.End
        End If
        
        rs.MoveFirst
        
	'Dim pageNumber
	
	pageNumber = 0

Do Until rs.EOF
	pageNumber = pageNumber + 1
	
	currBrokerCode = Rs.Fields("BrokerCode").Value %>

<%If pageNumber=1 Then%>
	<P class="newpage">
	<BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR>
	<P align="center"><B>SETTLEMENT SLIPS</B></P>
	<BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR>
	<BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR>
	<P class="newpage">
<%End If%>

<table width=60% align="center">
     <tr class="pageNumbering">
		<td align="left" colspan=2>
			<FONT FACE=ARIAL SIZE=2><B>Page <%=pageNumber%>	</B></FONT>	
		</td>		
	</tr>
     <tr>
		     <td>
		        <b><font face="Arial Narrow" size="4">Settlement Slip</font></b></td>
				
			<td align=right>
				<b><font face="Arial Narrow" size="4"><%= Session("CompanyName") %></font></b></td>
					
			</td>
	</tr>		
     <tr>
		     <td colspan=2>
		        <font face="Arial" size="2">printed for each individual broker&nbsp;</font></td>
		</tr>
	<tr>
		     <td colspan=2>
		        <font face="Arial" size="2">&nbsp;</font></td>
		</tr>	
  
</table>

<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="60%" align="center">
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


<BR>

<center>
<table border="0" cellspacing=0 cellpadding=3 width="60%" align="center">
<tr>
	<td colspan=7>	
<PRE>
<font face="Arial" size="2">
			<b>Payment to Brokers</b></font>			
<b><font face="Arial" size="2">		
Please prepare payment payable to brokers as per the details shown on this voucher.
Kindly note that all contracts on the list were traded on the same date shown below:</font></b></PRE>

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
</center>
<%

			rs.MoveNext
			'important!		
			If Not Rs.EOF Then %>
				<BR class="newpage">
				<%Else%>
				<BR class="newpage">
		<%	End If
        Loop
        
        conn.Close
        Set conn = Nothing
        Set Rs = Nothing
		
		
		
	'SETTLEMENT REMINDERS	

	'response.write "SETTLEMENT REMINDERS"

		Set conn = GetActiveConnection("KBroker")
        
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
			
        Else%>
			<Script Language="JavaScript">
				alert("The time limits of type Settlement (Brokers) have not been set");
				window.parent.self.close();				
            </Script>
			<%
			Set tempRs = Nothing
			Set Conn = Nothing
			Response.End
        End If
        
        Set tempRs = Nothing
        
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
        
        
        sqlStr = "SELECT * FROM [SettlementReminder] WHERE TRADED = '" & FormatDate(originalselectedTradeDate) & "' ORDER BY Traded DESC"
		
	
        Set rs = CreateObject("ADODB.Recordset")
        rs.CursorLocation = adUseClient
        Rs.Open SQLServerFormat(HandleQuote(sqlStr)), Conn.ConnectionString, 1, 3
        If rs.EOF Or rs.BOF Then
               %>
                <Script Language="JavaScript">
					alert("No settlement reminders available");
					window.parent.history.go(-1);
                </Script>
                <% Set Rs = Nothing
                Set Conn = Nothing
                Response.End
        End If
        
        
        rs.MoveFirst
        pageNumber=1

%>

<%If pageNumber=1 Then%>
	<P class="newpage">
	<BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR>
	<P align="center"><B>SETTLEMENT REMINDERS</B></P>
	<BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR>
	<BR><BR><BR><BR><BR><BR><BR><BR><BR>
	<P class="newpage">
<%End If%>


<table width="60%" align="center">
     <tr>
		     <td>
		        <b><font face="Arial Narrow" size="4">Settlement Reminder</font></b></td>
		      <td align=right>
				<b><font face="Arial Narrow" size="4"><%= Session("CompanyName") %> </font></b></td>
					
			</td>  
		</tr>	
     <tr>
		     <td colspan=2> 
		        <font face="Arial" size="2">for funds from or to brokers on: <%= FormatDateFull(upperLimitDate)%></font></td>
		</tr>
	<tr>
		     <td  colspan=2>
		        <font face="Arial" size="2">&nbsp;</font></td>
		</tr>	
  
</table>


  <table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="60%" align="center">
    <tr>
      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow" size="4">Traded</font></b></td>
      <td bgcolor="#000000"><b><font color="#FFFFFF" face="Arial Narrow" size="4">Broker</font></b></td>

      <td bgcolor="#000000">
        <p align="right"><b><font color="#FFFFFF" face="Arial Narrow" size="4">Gross</font></b></p>
    </td>
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
		Loop%>
		
		<tr>
		  <td>&nbsp;</td>
		  <td>&nbsp;</td>	
		  <td style="border-top: solid; border-width: 1" align=right valign=bottom height="25px"><%= FormatNum(myTotals) %></td>
		</tr>
		  
		<%
		
	End If
    
    Rs.Cancel
    
    dailyTotals = dailyTotals + myTotals
    myTotals = 0
    
    Rs.Filter = "OrderTypeSale = 1"
    
    If Not (Rs.EOF Or Rs.BOF) Then%>    
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
		     <td>
		    <p align="right"><%= FormatNum(Rs.Fields("GrossAmt").Value)%></td>
		</tr>
		
		
	<%		Rs.MoveNext
		Loop%>
		
		<tr>
		  <td>&nbsp;</td>
		  <td>&nbsp;</td>		
		  <td style="border-top: solid; border-width: 1; border-color: #000000;" align=right valign=bottom height="25px"><%= FormatNum(myTotals) %></td>
		</tr>
		  
		<%
		
	End If
	
	dailyTotals = dailyTotals + myTotals%>
	
	 <tr>
      <td colspan="3">&nbsp; </td>
    </tr>
     <tr>
      <td colspan="3">&nbsp; </td>
    </tr>   
    <tr>
      <td></td>
  
    <td>
      <p align="right">Daily Totals:&nbsp;&nbsp; </td>
    <td style="border-style: solid; border-width: 1">
			<p align="right"><%= FormatNum(dailyTotals) %></p>
		</td>
    </tr>
  </table>


<%
Set Rs = Nothing
Set Conn = Nothing


''END THE REPORT HERE
Response.End 

'CONTRACT SCHEDULE

response.write "CONTRACT SCHEDULE"


headerDescription = FormatDateFull(selectedTradeDate) %>
<i id="landRem">Remember to select landscape settings while printing.</i>
<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%" align="center">
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

    <table border="0" width="100%" cellPadding="2" cellSpacing=0 align="center">
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
		Rs.Filter = "LotTDate = '" & FormatDate(selectedTradeDate) & "'"
		
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
        Do Until rs.EOF%>
        		<tr>
      <td nowrap><font size="1"><%= Day(rs.Fields("LotTDate")) & " " & MonthName(Month(rs.Fields("LotTDate")), True) %></font></td>
      <td nowrap><font size="1"><% If Len(rs.Fields("ClientName")) > 25 Then 
									Response.Write Mid(rs.Fields("ClientName"), 1, 25)
								   Else
									Response.Write rs.Fields("ClientName")	
								   End If	 %></font></td>
      <td nowrap><font size="1"><%=rs.Fields("SecurityCode")%></font></td>
      <td nowrap align="center"><font size="1"><%=rs.Fields("BrokerCode")%></font></td>
      <td nowrap><font size="1"><%=rs.Fields("ContractNumber")%></font></td>
      <td nowrap><font size="1"><%=rs.Fields("LotSlipNo")%></font></td>
      <td nowrap align="right"><font size="1"><%=FormatNum(rs.Fields("LotPrice"))%></font></td>
      <td nowrap align="right"><font size="1"><%=FormatNumCommasOnly(rs.Fields("LotQty"))%></font></td>       
      <td style="BORDER-RIGHT: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent">&nbsp;</td>
           <%	levyTotals = 0
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
             <%rs.MoveNext
        Loop%>
        
        <tr>
			<td colspan="<%= 9 + UBound(dailyTotalsArray1)%>">&nbsp;</td>
        </tr>
      
		<tr height="30px">
			<td colspan=9 align=right>
				<b>
					Daily totals:
				</b>	
			</td>	
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
        
     <%else%>
                <script language = 'javascript'>
                		alert ("No contracts found using the specified criteria");
                		window.parent.history.go(-1);          		
                </script>
                
                <%  Set Rs = Nothing
					Set Conn = Nothing
                Response.end
     
   End if

%>
		
		




</body>
</html>



