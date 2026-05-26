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

<% DrawPageFunctions True, True, True %>


<%
	Set conn = GetActiveConnection("KBroker")
	
	Set groupRs = CreateObject("ADODB.Recordset")
	Set TotalRs = CreateObject("ADODB.Recordset")
						        
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
	sqlStr = "SELECT DISTINCT Order_DPA_,Security_DPA_ FROM ClientContractCompounded WHERE (LotTDate = '" & FormatDate(selectedContractDate) & "') ORDER BY Order_DPA_"
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
	Dim sqlStr7
	
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
					  "AS ContractNumber, OrdDetailSecurity, SUM(LotQty) AS LotQty, SUM(LotPrice) AS LotPrice, OrdDetailClient,Client_DPA_, SystemMaintained, LevyRatePercentage " & _
					  "FROM         dbo.ClientContractCompounded " & _
					  "WHERE     (LotSlipNo IN (" & slipNos  & ")) And LevyName <> '' AND OrdDetailClient = '" & prevClient & "' AND (Order_DPA_ = " & OrderRs.Fields("Order_DPA_").Value & ")" & _
					  "GROUP BY LevyName, LevyShortName, OrdDetailSecurity, OrdDetailClient,Client_DPA_, SystemMaintained, LevyRatePercentage,OrdDetailSecType,Order_DPA_"
						
			''Response.write(sqlStr)
			''Response.end
			
			Rs.CursorLocation = adUseClient	
			on error resume next
			Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
						
			sqlStr7="Select Gross,Quantity From ContractCompounded where(Order_DPA_=" & OrderRs.Fields("Order_DPA_").Value & ") and (Security_DPA_=" & OrderRs.Fields("Security_DPA_").Value & ") and (TransDate = '" & FormatDate(selectedContractDate) & "')"	
			
			
			Set Totalrs=Conn.Execute(sqlStr7)
			
			if not(Totalrs.eof and TotalRs.bof) then
			totalGross = TotalRs.Fields("Gross").Value
			Quantity= TotalRs.Fields("Quantity").Value

			end if
			
			orderRef = Rs.Fields("orderRef").Value
			
			Dim clientID
			
			if isnull(orderRef) or trim(orderRef) = "" then
					orderRef = Rs.Fields("Order_DPA_").Value
			else
					orderRef = Rs.Fields("Order_DPA_").Value & "/" & orderRef
			end if
			clientAddress = Replace(Rs.Fields("ClientAddr").Value, Chr(13), ",")
			clientID = Rs.Fields("Client_DPA_").Value		
					
			If Trim(UCase(Rs.Fields("OrdDetailType").Value)) = "PURCHASE" Then
				IsPurchase = True
			Else
				IsPurchase = False
			End If  
			
			pageNumber = pageNumber + 1
			
			%>


<table border="0" cellspacing=2 cellpadding=2  width="90%" align="center">
<tr><td>
<table border="0" cellspacing=0 cellpadding=2 class="ReportsTable" width="100%"> 
	<THEAD>
	<tr class="pageNumbering">
		<td align="left" colspan=2>
			<FONT FACE="Times New Roman" SIZE=2><B>Page <%=pageNumber%>	</B></FONT>	
		</td>		
	</tr>
	<tr>
		<td align="right" colspan=2 height="170" style="border-bottom: 2px inset #000000" >
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
          <table border="0" cellspacing=0 cellpadding=2  width="100%" align="center">
              <tr>
                <td width="50%" style="border: 1px solid #000000" valign="top" align="left">
                  <table border="0" width="100%">
                    <tr>
                      <td width="40%" valign="top" align="left">NSE 
                  SLIP NO</td>
                      <td width="60%" valign="top" align="left">
							<table border=0 cellpadding=1 cellspacing=0 width="100%">
											
									<%
									If InStr(1, slipNos, ",") > 0 Then
										Dim slipNosArray
										maxSlipNosPerRow = 3
										slipNosArray = Split(slipNos, ",")
										maxSlipCount = UBound(slipNosArray) 
													
										For k = 0 To maxSlipCount	%>
											<tr><td nowrap align="left">
											<%For l = 1 To maxSlipNosPerRow
												If k <= maxSlipCount Then											
													thisSlipNo = Trim(slipNosArray(k))
													If k < maxSlipCount Then
														thisSlipNo = thisSlipNo & ", "
													End If																						
													Response.Write "<B>" & thisSlipNo & "</B>"											
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
									<tr><td nowrap align="left"><B><%= slipNos %></B></td></tr>
									<%End If%>
											
							</table>
                      </td>
                    </tr>
                  </table>
                  &nbsp;</td>
                <td width="6%"></td>
                <td width="40%" valign="top" style="border: 1px solid #000000" >
                  <table border="0" width="100%" cellspacing="2" cellpadding=2>
                    <tr>
                      <td width="64%" align="left">
				<%= UCase(Rs.Fields("OrdDetailType").Value) %> CONTRACT NO </td>
                      <td width="36%" align="right"><B><%= Rs.Fields("ContractNumber").Value %></B></td>
                    </tr>
                    <tr>
                      <td width="64%" align="left">TRADE DATE:</td>
                      <td width="36%" align="right"><B><%= FormatDate(selectedContractDate) %></B></td>
                    </tr>
                  </table>
                </td>
              </tr>
      
	<tr>
		<td align="right" height="8" >
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
	<tr>
		<td align="right" height="8">
            &nbsp;			
		</td>		
	</tr>  	
</table>	

<tr><td>
<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" width="100%">	
	<tr>
		<td align=right colspan=2>KShs</td>
	</tr>

	<tr>
		<td nowrap><B><%= FormatNumCommasOnly(Quantity) %>&nbsp; <%= Rs.Fields("OrdDetailSecurity").Value %> @
          <%if lcase(rs.Fields("OrdDetailSecType")) = "fixed" Then%>
				<%= FormatNumEx((totalGross / Rs.Fields("LotQty").Value),4) %>&nbsp;%
		<%else%>
				<%= FormatNumEx(totalGross / Rs.Fields("LotQty").Value,4) %>
		<%end if%>
		</B></td>
		<td align=right><b> <%if lcase(rs.Fields("OrdDetailSecType")) = "fixed" Then%>
				<%= FormatNumEx(totalGross /100,2) %>
		<%else%>
				<%= FormatNumEx(totalGross,2) %>
		<%end if%>
 </b></td>
	</tr>
	</table>
	</td>
	</tr>
	<tr>
		<TD VALIGN="TOP" >
		<TABLE WIDTH="100%"  CELLSPACING="0" CELLPADDING="4" border=0>
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
				if lcase(rs.Fields("OrdDetailSecType")) = "fixed" Then
				levyPerc = (levyArray(i, 1) / (totalGross/100)) * 100
				else
				levyPerc = (levyArray(i, 1) / totalGross) * 100
				end if
				'levyPerc = FormatNum(levyPerc) & "%"
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
					<td width="8%">&nbsp;</td>
					<td nowrap width="10%" STYLE="PADDING-LEFT: 0PX"><%= thisLevyname %></td>
					<td ALIGN="LEFT">&nbsp;&nbsp;<%if lcase(rs.Fields("OrdDetailSecType")) = "fixed" OR trim(thisLevyname)="Broker Commission" Then%>
														<%= FormatNumEx(levyPerc,4) %>&nbsp;%
												  <%else%>
														<%= FormatNum(levyPerc) %>&nbsp;%
												<%end if%></td>
					<td align=right><%= FormatNum(RoundPoint05(levyArray(i, 1))) %></td>		
				</TR>
		<% End If
		Next%>
		<td width="8%">&nbsp;</td>
		<td colspan=2 >&nbsp;</td>
		<td align=right>
			<u>
			<%= Replace(Space(Len(FormatNum(RoundPoint05(totalLevies))) * 4), Space(1), "&nbsp;") %></u>
		</td>
	</tr>
	<tr>
		<td width="8%">&nbsp;</td>
		<td colspan=2 STYLE="PADDING-LEFT: 0PX">TOTAL CHARGES </td>
		<td align=right>	
			<%If IsPurchase Then%>
				<b><%= FormatNum(RoundPoint05(totalLevies)) %></b>
			<%Else%>	
				<b>(<%= FormatNum(RoundPoint05(totalLevies)) %>)</b>
			<%End If%>
		</td>
	</tr>
	<%if trim(transferFeeDesc) <> "" then%>
			<tr>
				<td width="8%">&nbsp;</td>
				<td align="left" colspan=2 STYLE="PADDING-LEFT: 0PX"><%= transferFeeDesc %></td>
				<td align=right>	
					 <%= FormatNum(transferFeeVal) %>		
				</td>
			</tr>
	<%end if%>
	
	<%if trim(contractStampsDesc) <> "" then%>
			<tr>
				<td width="8%">&nbsp;</td>
				<td colspan=2 STYLE="PADDING-LEFT: 0PX"><%=ucase(contractStampsDesc) %></td>
				<td align=right >	
					 <%= FormatNum(contractStampsVal) %>		
				</td>
			</tr>
	<%end if%>
	
	<tr>
		<td width="8%">&nbsp;</td>
		<td align=right	colspan="2"> 
		<%= Replace(Space(Len(FormatNum(RoundPoint05(totalLevies))) * 4), Space(1), "&nbsp;") %></u>
			
		</td>
	</tr>
		</TABLE>	
		</TD>
		</TR>
	
	<tr VALIGN="TOP"><td>
	<table border="0" cellspacing=0 cellpadding=4  width="100%" >	
	<tr>
		
		<td nowrap><B>TOTAL AMOUNT PAYABLE IN KSHS
		</B></td>
		<td align=right><b>
			<b>
			<%
			totalLevies = totalLevies + transferFeeVal + contractStampsVal
			
			if lcase(rs.Fields("OrdDetailSecType")) = "fixed" then
				If IsPurchase Then
					Response.Write FormatNum(RoundPoint05((totalGross/100) + totalLevies)) 
				Else
				Response.Write FormatNum(RoundPoint05((totalGross/100) - totalLevies)) 
				End If
			else
				If IsPurchase Then
				Response.Write FormatNum(RoundPoint05(totalGross + totalLevies)) 
				Else
				Response.Write FormatNum(RoundPoint05(totalGross - totalLevies)) 
				End If
			End If
			%>
						
			</b></td>
	</tr>	
	<tr>
		<td colspan=3>&nbsp;</td>
	</tr>
	</table>
	<tr>
	<td VALIGN="TOP">
	<table cellspacing =0 cellpadding=4 width="100%"  border=0>
	<tr>
		<td width="6%">&nbsp;</td>
	
		<td align="left" valign="top" STYLE="PADDING-LEFT: 0PX" colspan=2>			
					
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
		<td colspan=2 ><!--#Include file="DirectorFooter.asp"--></td>
	</tr>
</table>
				
	</td></tr>
</table>					
					
<%			
			slipNos = ""
			OrderRs.MoveNext
		if(OrderRs.eof=false) then
			%>
			<BR class="newpage">
			<%
		end if
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