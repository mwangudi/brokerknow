<%
''SETTLEMENT
Sub SettlementGenerate()
	%>
	<form id="repToPDF5" name="repToPDF5">
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
			      <td style="border-top-style: solid; border-top-width: 1"><b><font face="Arial Narrow" size="2">Malawi STOCK EXCHANGE</font></b></td>
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