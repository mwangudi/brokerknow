<html>

<head>
<title>Re-Do Compounded Commissions</title>

<style type="text/css">
	td { font-family:verdana; font-weight: bold; font-size:10pt; border: 0 solid gray; background-color:gainsboro; }
	input.buttons { font-family:verdana; font-weight: bold; font-size:10pt; border: 1 solid gray; }
	input.texts { font-family:verdana; font-weight: bold; font-size:10pt; border: 1 solid gray; background-color: lightyellow;}
</style>
</head>

<body Class="Reports">
<!--#include file="../libroutines.asp"-->
<%
genReport = Request.Form("genReport")

If genReport <> "1" Then
	%>
	<form method="POST" action="ReDoCompoundedCommissions.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">
		<br><br>
		<table width="50%" cellpadding=3 cellspacing=3 align=center style="border: 0 solid gray; background-color:gainsboro;" border=0>
			<tr>
				<td style="border: 0 solid gray; background-color:white;" colspan=2 width="100%">
					<b>Re-Do Compounded Commissions</b>
				</td>
			</tr>
			<tr>
				<td nowrap width="20%"><b>Order Number:</b></td>
				<td width="60%">
					<input type="text" class="texts" id="TxtOrder" name="TxtOrder" size="30">
				</td>
			</tr>
			<tr>
				<td style="border: 0 solid gray; background-color:white;" colspan=2 width="100%">
					<input type="submit" class="buttons" Value=" Recalculate ">
				</td>
			</tr>
		</table>
	</form>
	<%
	Response.End
End If
%>

<%
	theOrder = Trim(Replace(Request.Form("TxtOrder"),"'",""))
	
	If Len(theOrder) = 0 Then
		%>
		<script language = 'vbscript'>
			ShowMessage "Invalid Order Number"
			window.history.back(-1)
		</script>
		<%
		Response.End 
	End If
	
	If Not IsNumeric(theOrder) Then
		%>
		<script language = 'vbscript'>
			ShowMessage "Invalid Order Number"
			window.history.back(-1)
		</script>
		<%
		Response.End 
	End If
	
	Set conn = GetActiveConnection("KBroker")
	Set Rs = CreateObject("ADODB.Recordset")		
	Set rsGross = CreateObject("ADODB.Recordset")	
	
	sqlStr = "SELECT tbOrder.OrderCompounded, Contract.Contract_DPA_, Commission.CommissionRate AS ClientCommission, Lot.LotGrossAmount,  " & _
		" LevyContract.SystemMaintained, LevyContract.LevyName, LevyContract.LevyAmount, Commission_1.CommissionRate AS AgentCommission,  " & _
		" tbOrder.Order_DPA_, LevyContract.LevyContract_DPA_, Commission.UpperSecurityCommission, Commission.MinimumSecurityCommission, Commission.SecurityBoundary " & _
		" FROM LevyContract INNER JOIN " & _
		" Contract ON LevyContract.Contract_DPA_ = Contract.Contract_DPA_ INNER JOIN " & _
		" tbOrder INNER JOIN " & _
		" OrdDetail ON tbOrder.Order_DPA_ = OrdDetail.Order_DPA_ INNER JOIN " & _
		" Lot ON OrdDetail.OrdDetail_DPA_ = Lot.OrdDetail_DPA_ ON Contract.Contract_DPA_ = Lot.Contract_DPA_ INNER JOIN " & _
		" Client ON tbOrder.Client_DPA_ = Client.Client_DPA_ LEFT OUTER JOIN " & _
		" Agent ON Client.Agent_DPA_ = Agent.Agent_DPA_ LEFT OUTER JOIN " & _
		" Commission ON Client.Commission_DPA_ = Commission.Commission_DPA_ LEFT OUTER JOIN " & _
		" Commission Commission_1 ON Agent.Commission_DPA_ = Commission_1.Commission_DPA_ " & _
		" WHERE (tbOrder.OrderCompounded = 1) AND (LevyContract.SystemMaintained IN (11, 12)) AND (tbOrder.Order_DPA_ = "& theOrder &")"
	
	'Response.Write sqlstr
	'Response.End 
	
	Set Rs = conn.Execute(sqlStr)
	sqlStr2="SELECT     tbOrder.Order_DPA_, isnull(SUM(Lot.LotGrossAmount),0) AS Gross " & _
			 " FROM         Lot INNER JOIN " & _
			 "                       OrdDetail ON Lot.OrdDetail_DPA_ = OrdDetail.OrdDetail_DPA_ INNER JOIN " & _
			 "                       tbOrder ON OrdDetail.Order_DPA_ = tbOrder.Order_DPA_ " & _
			 " WHERE     (Lot.Deleted <> 1) AND (OrdDetail.Deleted <> 1) AND (tbOrder.Deleted <> 1) AND (tbOrder.OrderCompounded = 1)" & _
			 " GROUP BY tbOrder.Order_DPA_ " & _
			 " HAVING      (tbOrder.Order_DPA_ = "& theOrder &")"
	set rsGross = conn.execute(sqlStr2)
	if not (rsGross.eof or rsGross.bof) then
		TotalGrossPerOrder = rsGross("Gross")
	end if
	Set rsGross = nothing	
	'response.write TotalGrossPerOrder : response.end
	'Get Relevant/Correct Commission To Use
	If Not (Rs.EOF or rs.BOF) Then
		
		Dim CommArray()
		Dim i
		i = 0
		TotalLotGrossAmount = 0
			
		do until rs.EOF
			
				''Handle Values
				If Len(Rs("LotGrossAmount")) > 0 Then
					LotGrossAmount = Rs("LotGrossAmount")
				Else
					LotGrossAmount = 0
				End If
				
				If Len(Rs("ClientCommission")) > 0 Then
					ClientCommission = Rs("ClientCommission")
				Else
					ClientCommission = 0
				End If
				
				If Len(Rs("UpperSecurityCommission")) > 0 Then
					UpperSecurityCommission = Rs("UpperSecurityCommission")
				Else
					UpperSecurityCommission = 0
				End If
				
				If Len(Rs("SecurityBoundary")) > 0 Then
					SecurityBoundary = Rs("SecurityBoundary")
				Else
					SecurityBoundary = 0
				End If
				
				''Which Commission To Use?
				If UpperSecurityCommission = 0 Then
					UpperSecurityCommission	= ClientCommission
				End If
				
				If UpperSecurityCommission = ClientCommission Then
					UpperSecurityCommission	= ClientCommission
				End If
				
				If LotGrossAmount >= SecurityBoundary Then
					ClientCommission = UpperSecurityCommission
				End If
				
				ReDim Preserve CommArray(i)
				CommArray(i) = ClientCommission
				i = i + 1
				
				TotalLotGrossAmount = TotalLotGrossAmount + LotGrossAmount
				
			rs.MoveNext 
		loop
		
		If (TotalLotGrossAmount >= SecurityBoundary) Then
			ClientCommission = UpperSecurityCommission
		End If
		
		ReDim Preserve CommArray(i)
		CommArray(i) = ClientCommission
		
		'For i = 0 To Ubound(CommArray)
		'	Response.Write CommArray(i) & "<br>"
		'Next
		'Response.End 
		
		Dim aSorted
    
		aSorted = fnSort(CommArray, 1)
		Erase CommArray
		'response.write TotalGrossPerOrder & "<br>"
		'For i = 0 To Ubound(aSorted)
		'	Response.Write aSorted(i) & "<br>"
		'Next
    
		ClientCommission = aSorted(0)
		Erase aSorted
	End If
	
	Rs.Requery 
	
	If Not (Rs.EOF or rs.BOF) Then
		do until rs.EOF
			If Len(ClientCommission) = 0 Then
				ClientCommission = Rs("ClientCommission")
			End If
			
			''Broker Amount
			If Rs("SystemMaintained") = 11 Then
				''Handle Values
				If Len(Rs("LotGrossAmount")) > 0 Then
					LotGrossAmount = Rs("LotGrossAmount")
				Else
					LotGrossAmount = 0
				End If

				if LotGrossAmount <= 50000 then					
					ClientCommission = 2
				elseif LotGrossAmount>50000 and LotGrossAmount<=100000 then
					
					ClientCommission = 1.5
				elseif LotGrossAmount >100000  then
					
					ClientCommission = 1 					
				end if 
				'Response.Write LotGrossAmount & "<br>"
				'Response.Write ClientCommission & "<br>"

				'Broker Amount
				If (LotGrossAmount * ClientCommission) > 0 Then					
						BrokerAmount = (LotGrossAmount * ClientCommission) / 100					
				Else
					BrokerAmount = 0
				End If
				
				SqlStr = "UPDATE LevyContract SET LevyAmount = "& Replace(FormatNumber(BrokerAmount,2),",","") &" WHERE LevyContract_DPA_ = "& Rs("LevyContract_DPA_") &""
				'Response.Write Sqlstr & "<br>"
				conn.Execute(sqlStr)
			End If
			
			'AGENT COMMISSION
			If Rs("SystemMaintained") = 12 Then
				''Handle Values
				If Len(Rs("LotGrossAmount")) > 0 Then
					LotGrossAmount = Rs("LotGrossAmount")
				Else
					LotGrossAmount = 0
				End If
				
				''Broker Amount
				If (LotGrossAmount * ClientCommission) > 0 Then
					BrokerAmount = (LotGrossAmount * ClientCommission) / 100
				Else
					BrokerAmount = 0
				End If
				
				''Agent Commission
				If Len(Rs("AgentCommission")) > 0 Then
					AgentCommission = Rs("AgentCommission")
				Else
					AgentCommission = 0
				End If
				
				'Response.Write LotGrossAmount & "+<br>"
				'Response.Write ClientCommission & "+<br>"
				'Response.Write AgentCommission & "+<br>"
				
				If AgentCommission > 0 Then
					''Agent Amount
					If (BrokerAmount * AgentCommission) > 0 Then
						AgentAmount = (BrokerAmount * AgentCommission) / 100
					Else
						AgentAmount = 0
					End If
					
					SqlStr = "UPDATE LevyContract SET LevyAmount = "& Replace(FormatNumber(AgentAmount,2),",","") &" WHERE LevyContract_DPA_ = "& Rs("LevyContract_DPA_") &""
					'Response.Write Sqlstr & "<br>"
					conn.Execute(sqlStr)
				End If
			End If

			rs.MoveNext 
		loop
	End If

	'Response.End 

	Set Rs = Nothing
	Set Conn = Nothing
	
	Function fnSort(aSort, intAsc)
		Dim intTempStore
		Dim i, j 
		
		For i = 0 To UBound(aSort) - 1
			For j = i To UBound(aSort)
			'Sort Ascending
			if intAsc = 1 Then
				if aSort(i) > aSort(j) Then
					intTempStore = aSort(i)
					aSort(i) = aSort(j)
					aSort(j) = intTempStore
				End if 'i > j
			'Sort Descending
			Else
				if aSort(i) < aSort(j) Then
					intTempStore = aSort(i)
					aSort(i) = aSort(j)
					aSort(j) = intTempStore
				End if 'i < j
			End if 'intAsc = 1
			Next 'j
		Next 'i
		
		fnSort = aSort
    End function 'fnSort
%>

<script language = 'vbscript'>
	ShowMessage "Contracts For Order Number " & <%=theOrder%> & " Updated."
	window.location.href = "ReDoCompoundedCommissions.asp"
</script>

</body>
</html>
