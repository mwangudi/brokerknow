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
				page-break-before:always;
			}
		}

	</style>
</head>

<body Class="Reports">

<!--#include file="../libroutines.asp"-->

<%

genReport = Request.Form("genReport")
selectedContractDate = Request.Form("txtDate")

If genReport <> "1" Or selectedContractDate = "" Then%>
		<Script Language="JavaScript">
		function validateForm(frm){			
			if (frm.txtDate.value==''){
				alert("Select a date");
				frm.txtDate.focus();
				return;
			}
			
			frm.target = '_self';			
			frm.submit();
		}
		document.body.className = 'dialog';
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "txtDate","cmdDate","<%= FormatDate(Date) %>",1);
	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="ClientCompoundedNew.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">	
		<table>
			<tr>
				<td>Select date:</td>
				<td>
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>
			</tr>
			
			<tr>
				<td colspan=2><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... ">&nbsp;&nbsp; <input type="Button" class="Buttons" Value=" Close " OnClick="JavaScript: window.parent.self.close();"></td>
			</tr>
		</table>
		
	</form>
	
	<%
	Response.End
End If

%>

<% DrawPageFunctions True, True, True %>


<%
	Set conn = GetActiveConnection("KBroker")
	Set groupRs = CreateObject("ADODB.Recordset")						        
	sqlStr = "SELECT * FROM CompoundedContractsList WHERE LotTDate = '" & FormatDate(selectedContractDate) & "'"
	groupRs.CursorLocation = adUseClient	
	groupRs.Open sqlStr, conn.ConnectionString, adOpenKeyset, adLockOptimistic
	'response.write groupRs.NextRecordset.Fields.Count
	'response.end
	'Set groupRs = groupRs.NextRecordset
	'groupRs.Filter = "LotTDate = '" & FormatDate(selectedContractDate) & "'"
		
	If groupRs.EOF Or groupRs.BOF Then%>
		<Script Language="JavaScript">
			alert("There are no contracts based on the specified criterion.")
			window.parent.history.go(-1);			
		</Script>
		<%Set groupRs = Nothing
		Set Conn = Nothing
		Response.End
	End If
	
	'Set Rs = CreateObject("ADODB.Recordset")						        
	'sqlStr = "SELECT * FROM ClientContractCompounded WHERE LotTDate = '" & FormatDate(selectedContractDate) & "' ORDER BY ContractNumber"		
	'Rs.CursorLocation = adUseClient	
	'Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
	
	Set levyOrderRs = Conn.Execute("SELECT * FROM LevyOrderList ORDER BY LevyOrder")
	
	prevSecurity = groupRs.Fields("OrdDetailSecurity").Value
	prevDate = groupRs.Fields("LotTDate").Value
	prevOrderTypeSale = groupRs.Fields("OrderTypeSale").Value
	slipNos = groupRs.Fields("LotSlipNo").Value
	prevSlipNo = slipNos
	
	blnEndFound = False
	
	Do Until groupRs.EOF
	
		'Rs.Filter = "ContractNumber = '" & groupRs.Fields("ContractNumber").Value & "'"
	
	If (prevSecurity <> groupRs.Fields("OrdDetailSecurity").Value) OR  (prevOrderTypeSale <> groupRs.Fields("OrderTypeSale").Value) Then
			
			Set Rs = CreateObject("ADODB.Recordset")						        
			sqlStr = "SELECT * FROM ClientContractCompounded WHERE LotSlipNo IN (" & slipNos  & ") ORDER BY ContractNumber"
			Rs.CursorLocation = adUseClient	
			Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
		
	
			totalGross = Rs.Fields("LotQty").Value * Rs.Fields("LotPrice").Value
	
			sqlStr = "SELECT     tbOrder.*, Client.ClientAddr FROM Lot INNER JOIN " & _
					" OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN " & _
					" tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN " & _
					" Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ " & _
					" WHERE Lot.Lot_DPA_ = " & Rs.Fields("Lot_DPA_").Value
				
	
			Set temprs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
	
			If Not (temprs.EOF Or temprs.BOF) Then
				orderRef = temprs.Fields("orderRef").Value
				clientAddress = Replace(temprs.Fields("ClientAddr").Value, Chr(13), ",")		
			End If
	
			Set temprs = Nothing
	
			If Trim(UCase(Rs.Fields("OrdDetailType").Value)) = "PURCHASE" Then
				IsPurchase = True
			Else
				IsPurchase = False
			End If
	
	
	
%>


<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" width="100%">
	<THEAD>
	<tr>
		<td align="center" colspan=2 height="200px">
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
		<td>&nbsp;</td>
	</tr>  
	<tr>
		<td colspan=2>&nbsp;</td>
	</tr>  	
	<tr>
		<td  nowrap width="100%">			
			<table cellspacing=0 cellpadding=0 border=0  width="100%">
				<tr>
					<td nowrap><%= Rs.Fields("OrdDetailClient").Value %> </td>					
				</tr>
				<tr>					
					<td nowrap><%= clientAddress %> </td>
				</tr>
			</table>
		</td>
		<td>
			<table cellspacing=0 cellpadding=0 border=0  width="100%">
				<tr>
					<td  nowrap>Date:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%= FormatDate(Rs.Fields("LotTDate").Value) %></td>
				</tr>
				<tr>
					<td  nowrap>Order Ref:&nbsp;&nbsp;&nbsp;<%= orderRef %></td>					
				</tr>
				<tr>
					<td  nowrap>Slip Nos:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%= slipNos %></td>					
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

<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable"  width="100%">	
	
	<tr>
		<td>Security</td>
		<td align=right><%= Rs.Fields("OrdDetailSecurity").Value %></td>
	</tr>
	<tr>
		<td>Quantity</td>
		<td align=right> <%= FormatNum(Rs.Fields("LotQty").Value) %> </td>
	</tr>
	<tr>
		<td>Average Price</td>
		<td align=right> <%= FormatNum(Rs.Fields("LotPrice").Value) %> </td>
	</tr>
    <tr>
		<td>Gross</td>
		<td align=right><b> <%= FormatNum(totalGross) %> </b> </td>
	</tr>
	<%
	totalLevies = 0
	levyArray = SortLevies(rs, levyOrderRs)
	For i = 1 To UBound(levyArray)
		totalLevies = totalLevies + levyArray(i, 1) %>
		<tr>
			<td><%= levyArray(i, 0) %></td>
			<td align=right><%= FormatNum(levyArray(i, 1)) %></td>
		</tr>
	<%Next%>	
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
			<%If IsPurchase Then
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
		<td colspan=2>
		<PRE><font face=Arial size=2>
       <i>For and on behalf of</i>
African Alliance Malawi Securities

Sign..........................................
		</FONT></PRE>	
		</td>
	</tr>
</table>
	
	
		<%
			'important!
			If Not groupRs.EOF Then %>
				<BR class="newpage">
		<%	End If
			
		
			prevSecurity = groupRs.Fields("OrdDetailSecurity").Value
			prevDate = groupRs.Fields("LotTDate").Value
			prevOrderTypeSale = groupRs.Fields("OrderTypeSale").Value
			slipNos = ""
		   	
		Else
			
			If prevSlipNo <> groupRs.Fields("LotSlipNo").Value  Then
							
					If slipNos = "" Then 
						slipNos = groupRs.Fields("LotSlipNo").Value
					Else 
						slipNos = slipNos & ", " & groupRs.Fields("LotSlipNo").Value
					End If
									
					prevSlipNo = groupRs.Fields("LotSlipNo").Value					
						
			End If
			
			
		End If
		
		groupRs.MoveNext
		
		If groupRs.EOF Then
			If blnEndFound = False Then
				blnEndFound = True
				prevSecurity = ""
				groupRs.Move -1
			End If	
		End If		
		
		
	Loop
	
	Set groupRs = Nothing
	Set rs = Nothing
	Set Conn = Nothing
	
	Function SortLevies(srcRs, orderRs)
		Dim returnArray()
		
		ReDim returnArray(srcRs.RecordCount, 1)
		
		'careful, recordcount is 1-based
		arrCounter = 0 
		If Not (orderRs.EOF Or orderRs.BOF) Then
			Do Until orderRs.EOF
				isExistingLevy = False
				Do Until srcRs.EOF
					If StrComp(orderRs.Fields("LevyName").Value, srcRs.Fields("LevyName").Value, vbTextCompare) = 0 Then
						arrCounter = arrCounter + 1
						returnArray(arrCounter, 0) = srcRs.Fields("LevyName").Value
						returnArray(arrCounter, 1) = srcRs.Fields("LevyAmount").Value
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
				srcRs.MoveNext
			Loop	
			srcRs.MoveFirst
		End If
		
		SortLevies = returnArray
		
	End Function
%>

</body>

</html>
