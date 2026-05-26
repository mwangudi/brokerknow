<!--#include file="../libroutines.asp"-->

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
			@top{font-family: Helvetica, Times New Roman, sans-serif;
				font-size: 150%;
				font-weight: bolder;
				text-align: left;
				content: "<%= FormatDate(Date) %>";			
			}
			
			margin-left: 0cm;
			margin-right: 0cm;
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
	<form method="POST" action="ClientContract.asp" Name="frmMain" id="frmMain">
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

<% DrawPageFunctions True, True, True %>


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
	
	Set levyOrderRs = Conn.Execute("SELECT * FROM LevyOrderList ORDER BY LevyOrder")
	
	Dim pageNumber
	
	pageNumber = 0
	Do Until groupRs.EOF
	
		Rs.Filter = "ContractNumber = '" & groupRs.Fields("ContractNumber").Value & "'"
	
	'totalGross = Rs.Fields("LotQty").Value * Rs.Fields("LotPrice").Value
	totalGross = Rs.Fields("LotGrossAmount").Value
	
	sqlStr = "SELECT     tbOrder.*, Client.ClientAddr FROM Lot INNER JOIN " & _
			" OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN " & _
			" tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ INNER JOIN " & _
			" Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ " & _
			" WHERE Lot.Lot_DPA_ = " & Rs.Fields("Lot_DPA_").Value
		'Response.Write sqlstr
		'Response.end
	
	Set temprs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
	
	Dim clientID
	If Not (temprs.EOF Or temprs.BOF) Then
		orderRef = temprs.Fields("orderRef").Value
		if isnull(orderRef) or trim(orderRef) = "" then
				orderRef = temprs.Fields("Order_DPA_").Value
		else
				orderRef = temprs.Fields("Order_DPA_").Value & "/" & orderRef
		end if
		clientAddress = Replace(temprs.Fields("ClientAddr").Value, Chr(13), ",")
		clientID = temprs.Fields("Client_DPA_").Value
	End If
	
	Set temprs = Nothing
	
	If Trim(UCase(Rs.Fields("OrdDetailType").Value)) = "PURCHASE" Then
		IsPurchase = True
	Else
		IsPurchase = False
	End If
	
	pageNumber = pageNumber + 1
%>
<table border="0" cellspacing=2 cellpadding=2  width="90%" class="ReportsTable" align="center"> 
<tr><td>
<table border="0" cellspacing=0 cellpadding=2 class="ReportsTable" width="100%"> 
	<THEAD>
	<tr class="pageNumbering">
		<td align="left" height="18">
			<FONT FACE=Times New Roman SIZE=2><B>Page <%=pageNumber%>	</B></FONT>	
		</td>		
	</tr>
	<tr>
		<td align="right" height="170" style="border-bottom: 2px inset #000000">
			<Img Src="../data/photos/aaprintlogo.bmp">			
		</td>		
	</tr>
	<tr>
		<td >
			&nbsp;<BR>&nbsp;<BR>&nbsp;		
		</td>		
	</tr>
	<THEAD>  
	<tr>
		<td  align="center" height="52">
            <table border="0" cellspacing=0 cellpadding=2  width="96%" align="right" >
              <tr>
                <td width="25%" style="border: 1px solid #000000" valign="top" align="left">
                  <table border="0" width="100%">
                    <tr>
                      <td width="41%" valign="top">NSE  SLIP NO</td>
                      <td width="59%" valign="top"><B><%= Rs.Fields("LotSlipNo").Value %></B></td>
                    </tr>
                  </table>
                  &nbsp;</td>
                <td width="6%"></td>
                <td width="30%" style="border: 1px solid #000000">
                  <table border="0" width="100%">
                    <tr>
                      <td width="64%" align="left">
				<%= UCase(Rs.Fields("OrdDetailType").Value) %> CONTRACT NO </td>
                      <td width="36%" align="right"><B><%= Rs.Fields("ContractNumber").Value %></B></td>
                    </tr>
                    <tr>
                      <td width="64%" align="left">TRADE DATE:</td>
                      <td width="36%" align="right"><B><%= FormatDate(Rs.Fields("LotTDate").Value) %></B></td>
                    </tr>
                  </table>
                </td>
              </tr>
			  <tr>
				<td align="right" height="8">
            &nbsp;			
		</td>		
	</tr>
	<tr>
		<td style="border: 1px solid #000000" valign="top" align="center" colspan="3"  height="90">
				<table cellspacing=0 cellpadding=2 border=0 width="100%">
				<tr>
					<td nowrap>[<%= clientID %>]&nbsp;<%= Rs.Fields("OrdDetailClient").Value %> </td>					
				</tr>
				<tr>					
					<td nowrap ><%= clientAddress %> </td>
				</tr>
			</table>
		</td>
	</tr>
            </table>
		</td>		
	</tr> 
	


  
</table>

<tr><td>
<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" width="100%">	
	<tr>
		<td width="25">&nbsp;</td>
		<td align=right colspan=2>KShs</td>
	</tr>
	<tr>
		<td width="25">&nbsp;</td>
		<td nowrap><B><%= FormatNumCommasOnly(Rs.Fields("LotQty").Value)%>&nbsp;&nbsp;<%= Rs.Fields("OrdDetailSecurity").Value %> @
          <%if lcase(rs.Fields("OrdDetailSecType")) = "fixed" Then%>
				<%= FormatNumEx(Rs.Fields("LotPrice").Value,4) %>
          <%else%>
				<%= FormatNum(Rs.Fields("LotPrice").Value) %>
		<%end if%>
		</B></td>
		<td align=right><b> <%= FormatNum(totalGross) %> </b></td>
	</tr>
	</table>
	</td></tr>

	<tr>
		<TD VALIGN="TOP"  >
		<TABLE WIDTH="100%" CELLSPACING="2" CELLPADDING="2" border=0 >
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
					<td width="10%">&nbsp;</td>
					<td nowrap width="20%" ><%= thisLevyName %></td>
					<td ALIGN="LEFT">&nbsp;&nbsp;<%if lcase(rs.Fields("OrdDetailSecType")) = "fixed" Then
														if(Cdbl(levyArray(i, 1))=Cdbl(rs("MinimumCommission")) and Trim(thisLevyName)="Broker Commission") then
														%>
														Minimum
														<%
														else
														%>
														<%= FormatNumEx(levyPerc,4) %>
													<%
														end if
													
												  else
														if(Cdbl(levyArray(i, 1))=Cdbl(rs("MinimumCommission")) and Trim(thisLevyName)="Broker Commission") then
														%>
														Minimum
														<%
														else
														%>
														<%= FormatNum(levyPerc) %>
													<%
														end if
													end if%></td>
					<td align=right><%= FormatNum(RoundPoint05(levyArray(i, 1))) %></td>		
				</TR>
		<%	End If
		Next%>
			<tr>
	<td width="10%">&nbsp;</td>
		<td colspan=2 >&nbsp;</td>
		<td align=right>
			<u>
			<%= Replace(Space(Len(FormatNum(RoundPoint05(totalLevies))) * 4), Space(1), "&nbsp;") %></u>
		</td>
	</tr>
	<tr>
	<td width="10%">&nbsp;</td>
		<td align="left" colspan=2>TOTAL CHARGES </td>
		<td align=right >	
			<%If IsPurchase Then%>
				<b><%= FormatNum(RoundPoint05(totalLevies)) %></b>
			<%Else%>	
				<b>(<%= FormatNum(RoundPoint05(totalLevies)) %>)</b>
			<%End If%>
		</td>
	</tr>
		</TABLE>	
		</TD>
	</tr>
	<tr><td>
	<Table border="0" cellspacing=2 cellpadding=2 width="100%" >

	<%if trim(transferFeeDesc) <> "" then%>
			<tr>
				<td width="10%">&nbsp;</td>
				<td><%= Ucase(transferFeeDesc) %></td>
				<td align=right>	
					 <%= FormatNum(RoundPoint05(transferFeeVal)) %>
				</td>
			</tr>
	<%end if%>
	
	<%if Ucase(trim(contractStampsDesc)) <> "" then%>
			<tr>
				<td width="10%">&nbsp;</td>
				<td><%= ucase(contractStampsDesc) %></td>
				<td align=right>	
					 <%= FormatNum(contractStampsVal) %>
				</td>
			</tr>
	<%end if%>
	<tr>
		<td width="10%">&nbsp;</td>
		<td>&nbsp;</td>
		<td align=right> 
		<u>
			<%= Replace(Space(Len(FormatNum(RoundPoint05(totalLevies))) * 4), Space(1), "&nbsp;") %></u>
			
		</td>
	</tr>
		</table>
	</td></tr>
	<tr><td>
	<table border="0" cellspacing=2 cellpadding=2  width="100%" >	
	<tr>
		<td width="25">&nbsp;</td>
		<td nowrap><B>TOTAL AMOUNT PAYABLE IN KSHS
		</B></td>
		<td align=right><b>
			<%
			totalLevies = totalLevies + transferFeeVal + contractStampsVal
			If IsPurchase Then
				Response.Write FormatNum(RoundPoint05(totalGross + totalLevies)) 
			Else
				Response.Write FormatNum(RoundPoint05(totalGross - totalLevies)) 
			End If%>
			</b></td>
	</tr>	
	<tr>
		<td colspan=3>&nbsp;</td>
	</tr>
	</table>
	<tr>
	<td>
	<table cellspacing =2 cellpaddin=2 width="100%"  border=0>
	<tr>
		<td width="10%">&nbsp;</td>
	
		<td align="left" valign="top" >			
					
			<img src="../images/stamp.gif" border="0" style="position: absolute;z-index: 2">
			<BR>
			&nbsp;<SPAN style="position: absolute;z-Index: 10"><b>KShs&nbsp;&nbsp;<%= FormatNum(RoundPoint05(totalContractStamps)) %></b></SPAN>
			<br><br><br>
            Revenue Stamps prepaid
						
		</td>

	</tr>
		</table>
	</td></tr>
	<tr>
		<td colspan=2 align="right" height="">
           &nbsp;<BR>&nbsp;<BR>&nbsp;<BR>&nbsp;		
		</td>		
	</tr>	
	<%Dim additionalText
	if lcase(rs.Fields("OrdDetailSecType")) = "fixed" Then
			additionalText = "Payment to be made in Kenya Shillings on receipt of this contract. " & _
							"Cheques should include exchange for Country Cheques and be made payable to Dyer & Blair Investment Bank Ltd."
	 else
			additionalText = "&nbsp;"
	end if%>
	<tr>
		<td align="left" colspan=2><small><small><%=additionalText%></small></small></td>
	</tr>
	
	
	<tr>
		<td colspan=2><!--#Include file="DirectorFooter.asp"--></td>
	</tr>
</table>
	
<%
		Rs.Cancel
	groupRs.MoveNext
	'important!
		If Not groupRs.EOF Then %>
			<BR class="newpage">
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
%>

</body>

</html>