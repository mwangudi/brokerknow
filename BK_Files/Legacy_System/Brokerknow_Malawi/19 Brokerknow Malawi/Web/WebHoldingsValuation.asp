<!--#include file="../libroutines.asp"-->

<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Holdings Valuation</title>  
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
				size: landscape;
				margin-left: 2cm;
				margin-right: 5cm;
				margin-top: 1cm;    
				margin-bottom: 2cm;
				writing-mode: tb-rl;
				height: 80%;
				margin: 10% 0%;						
				br.newpage{
					page-break-before:always;
				}		
			}		 
	</style>
</head>

<body Class="Reports">



<%
'FirstDay=DateSerial(Year(Date), Month(Date) + iOffset, 1)
selectedClient = Session("Client_DPA_")
SelectedFromDate=DateSerial(Year(Date), Month(Date)-3 + iOffset, 1)

if(SelectedFromDate<CDate(FormatDate("03/01/2005"))) then
	SelectedFromDate=CDate(FormatDate("03/01/2005"))
end if

SelectedToDate=Date
%>

<% DrawPageFunctions True, True, True %>


<%

If selectedClient="" Then%>
		<Script Language="JavaScript">
			alert("Please select The client")
			window.history.go(-1);
		</Script>
		<%
		Response.End
End If
	
	Set conn = GetActiveConnection("KBroker")
	Set Rs = CreateObject("ADODB.Recordset")						        
	sqlStr = "Select * From [DB_Portfolios] WHERE Client_DPA_ = " & selectedClient & " And TransDate between '" & CDate(selectedFromDate) & "' and '" & CDate(selectedToDate) & "' order by Security_DPA_,TransDate"
	
	'Response.Write sqlstr
	'Response.end
	
	Conn.execute("ClientTotalsDelete")		 
	Conn.execute("ClientBalancesDelete")		 
			
	Conn.execute("ClientTotalsProcedure")		 
	Conn.execute("ClientBalancesProcedure")		 
	
	Rs.CursorLocation = adUseClient	
	Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
	
	If rs.EOF Or rs.BOF Then%>
		<Script Language="JavaScript">
			alert("There are no records based on the specified criterion.")
			window.parent.history.go(-1);			
		</Script>
		<%Set Rs = Nothing
		Set Conn = Nothing
		Response.End
	End If
		
	Dim pageNumber
	
	pageNumber = 0	
	
		'Rs.Filter = "ContractNumber = '" & Rs.Fields("ContractNumber").Value & "'"
	
	sqlStr = "SELECT Client.ClientAddr,Client.ClientContact,'W' + ClientOfficeTel + '/' + Client.ClientEmail as Contacts,ClientName FROM Client " & _			
			" WHERE Client_DPA_ = " & selectedClient
		'Response.Write sqlstr
		'Response.end
	
	Set temprs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
	
	Dim clientID
	If Not (temprs.EOF Or temprs.BOF) Then		
		clientAddress = Replace(temprs.Fields("ClientAddr").Value, Chr(13), ",")
		clientID = selectedClient
		ClientContact=temprs("ClientContact")
		Contacts=temprs("Contacts")
		ClientName=temprs("ClientName")
	End If
	
	Set temprs = Nothing
		
	pageNumber = pageNumber + 1
%>
<table border="0" cellspacing="2" cellpadding="2" style="font-family: Arial Narrow" width="100%">
    <tr>
	  <td bgcolor="#000000" width="80%" nowrap align="left"><font color="#FFFFFF" face="Impact" size="2">HOLDINGS&nbsp;&nbsp;&nbsp;&nbsp;VALUATION&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></td>
      <td bgcolor="#000000" width="20%" nowrap align=right><font color="#FFFFFF" face="Impact" size="2"><%= Session("CompanyName") %></font></td>
      
    </tr>

  </table>
<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" width="100%" > 
	<tr>      
      			<td ><b>Date:&nbsp;&nbsp;<%= FormatDate(Date) %></b></td>      			
    </tr>
	<tr>
		<td style="border: 1px solid #000000" valign="top">
				<table cellspacing=0 cellpadding=0 border=0>		

					<tr>
						<td nowrap>CLIENT&nbsp;&nbsp;NAME&nbsp;:&nbsp;<%= ClientName %> </td>					
					</tr>
					<tr>
					<td>CLIENT&nbsp;&nbsp;CODE&nbsp;:&nbsp;&nbsp;<%= clientID %></td>
					</tr>
					<tr>					
					<tr>
					<td>CONTACT&nbsp;&nbsp;PERSON&nbsp;:&nbsp;&nbsp;<%=ClientContact  %></td>
					</tr>
					<tr>										
					<td nowrap>CLIENT&nbsp;&nbsp;ADDRESS&nbsp;:&nbsp;&nbsp;<%= clientAddress %> </td>
					</tr>
					<tr>										
					<td nowrap>CONTACT TEL No&nbsp;:<%= Contact %> </td>
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
<table border="0" cellspacing=2 cellpadding=2 width="400">		
      <tr>
	  <td align="left" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>Code&nbsp;&nbsp;</b></font></td>      
      <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>Security&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b></font></td>                        
      <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>Quantity</b></font></td>      
      <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>&nbsp;&nbsp;&nbsp;Buy</b></font></td>
      <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>Last&nbsp;Spot&nbsp;Price</b></font></td>
      <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>Book&nbsp;Value</b></font></td>
      <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>Current&nbsp;Value</b></font></td>
      <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;P/L</b></font></td>
      <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;%P/L</b></font></td>            
    </tr>
	
    	    <%
Dim Security1
Dim Security2
Dim QuantityPrice

Security1=0
Security2=0
SecurityName=""
SecurityCode=""
BookQty=0
BookPrice=0
PreviousBookQty=0
PreviousBookPrice=0
BookValue=0
PreviousBookValue=0
QtySold=0
BookPriceSold=0
SellPrice=0
BookValueSold=0
SoldValue=0
PL=0
PL_=0
ClientValue=0

PreviousSecurity=""
PreviousCode=""

PreviousBookaValue=0
PreviousBuy=0

TotalBookValue=0	
first=1
    
Do Until Rs.EOF
Security1=rs("Security_DPA_")	
	if(Security1<>Security2) then
		if(first=0) then
			if(PreviousBookQty<>0) then
				ClientValue=PreviousBookQty*Price
				
				PL=ClientValue-PreviousBookaValue
				
				if(price=0) then
				PL=0
				end if
					if(PreviousBookaValue<>0) then
					PL_=(PL/PreviousBookaValue)*100
					else
					PL_=0
					end if				
				%>
				<tr>      		
				<td><font face="Arial Narrow" size="2"><%=PreviousCode%></font></td>      
				<td><font face="Arial Narrow" size="2"><%=PreviousSecurity%></font></td>		
				<td align="Right"><font face="Arial Narrow" size="2"><%=FormatNumEx(PreviousBookQty,0)%></font></td>      
				<td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(PreviousBuy)%></font></td>
				<td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(Price)%></font></td>
				<td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(PreviousBookaValue)%></font></td>      
				<td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(ClientValue)%></font></td>
				<td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(PL)%></font></td>
				<td align="Right"><font face="Arial Narrow" size="2"><%=FormatNumEx(PL_,2)%>%</font></td>      
				</tr>
				<%
				Equityvalue=Equityvalue+PreviousBookaValue		
				Equityvalue1=Equityvalue1+ ClientValue
			end if		
		end if		
	
		Quantity=0
		QuantityPrice=0
		
	end if	
	
	      if(Security1<>Security2) then
			if(rs("OrderTypeSale")=1) then
			BookQty=0
			else
			BookQty=rs("LotQty")
			end if
          else
            if(rs("OrderTypeSale")=1) then
					if((PreviousBookQty-rs("LotQty"))<=0) then
					BookQty=0
					else
					BookQty=PreviousBookQty-rs("LotQty")
					end if
				else
				BookQty=PreviousBookQty+rs("LotQty")
				end if
		 end if

    
			if(Security1<>Security2) then
				if(rs("OrderTypeSale")=1) then
				BookPrice=0
				else
				BookPrice=rs("NetAmount")/rs("LotQty")
				end if
			else
				if((BookQty)=0) then
						BookPrice=0				
				else		
					if(rs("OrderTypeSale")=1) then
						BookPrice=PreviousBookPrice
					else
					BookPrice=((PreviousBookPrice*PreviousBookQty)+((rs("NetAmount")/(rs("LotQty"))*rs("LotQty"))))/(rs("LotQty")+PreviousBookQty)
					end if
				end if
			end if
			BookValue=BookQty*BookPrice
			
			Quantity=BookQty + rs("LotQty")
			QuantityPrice =QuantityPrice + rs("LotQty")*rs("LotPrice")
		
			CashOnHand=rs("CurrentBal")    
    
			PreviousBookPrice=BookPrice
			PreviousBookQty=BookQty
			PreviousBookValue=Bookvalue
		
			PreviousBookaValue=BookValue
			PreviousBuy=BookPrice
			
			PreviousSecurity=rs("SecurityName")
			PreviousCode=rs("SecurityCode")
			Price=rs("SecurityMktPrice")
			
			Security2=Security1
			first=0
		Rs.MoveNext
	
	Loop
		if(PreviousBookQty<>0) then			
			ClientValue=PreviousBookQty*Price
			PL=ClientValue-PreviousBookaValue
			
			if(price=0) then
			PL=0
			end if
				
			if(PreviousBookaValue<>0) then
			PL_=(PL/PreviousBookaValue)*100
			else
			PL_=0
			end if
			%>
			<tr>      		
			<td><font face="Arial Narrow" size="2"><%=PreviousCode%></font></td>      
			<td><font face="Arial Narrow" size="2"><%=PreviousSecurity%></font></td>		
			<td align="Right"><font face="Arial Narrow" size="2"><%=FormatNumEx(PreviousBookQty,0)%></font></td>      
			<td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(PreviousBuy)%></font></td>
			<td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(Price)%></font></td>
			<td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(PreviousBookaValue)%></font></td>
			<td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(ClientValue)%></font></td>
			<td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(PL)%></font></td>
			<td align="Right"><font face="Arial Narrow" size="2"><%=FormatNumEx(PL_,2)%>%</font></td>            
			</tr>
			<%
		end if
			Equityvalue=Equityvalue+PreviousBookaValue
		Equityvalue1=Equityvalue1+ ClientValue	
		
		Total=Equityvalue1 + CashOnHand
		
		Set rs = Nothing
		Set Conn = Nothing
		
		if(EquityValue<>0) then
		TotalPL	=((EquityValue1-EquityValue)/EquityValue)*100
		else
		TotalPL=0
		end if
	
		'Response.write(TotalPL)
	
		if(Abs(TotalPL)<0.01) then
		TotalPL=0
		end if
		%>
		<tr>
<td colspan="9">&nbsp;</td>
</tr>
	<tr>
      <td><font face="Arial Narrow" size="2">&nbsp;</font></td>
      <td><font face="Arial Narrow" size="2">&nbsp;</font></td>
      <td><font face="Arial Narrow" size="2">&nbsp;</font></td>
      <td><font face="Arial Narrow" size="2">Equity&nbsp;Value</font></td>
      <td><font face="Arial Narrow" size="2">&nbsp;</font></td>
      <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(EquityValue)%></font></td>
      <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(EquityValue1)%></font></td>
      <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(EquityValue1-EquityValue)%></font></td>
      <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNumEx(TotalPL,2)%>%</font></td>      
    </tr>
    <!--
    <tr>
	<td colspan="9">&nbsp;</td>
	</tr>
	-->
	<tr>
      <td><font face="Arial Narrow" size="2">&nbsp;</font></td>
      <td><font face="Arial Narrow" size="2">&nbsp;</font></td>
      <td><font face="Arial Narrow" size="2">&nbsp;</font></td>
      <td><font face="Arial Narrow" size="2">Cash&nbsp;In&nbsp;Hand</font></td>
      <td><font face="Arial Narrow" size="2">&nbsp;</font></td>
      <td><font face="Arial Narrow" size="2">&nbsp;</font></td>
      <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(CashOnHand)%></font></td>
      <td><font face="Arial Narrow" size="2">&nbsp;</font></td>
      <td><font face="Arial Narrow" size="2">&nbsp;</font></td>      
    </tr>
    
	<tr>
      <td><font face="Arial Narrow" size="2">&nbsp;</font></td>
      <td><font face="Arial Narrow" size="2">&nbsp;</font></td>
      <td><font face="Arial Narrow" size="2">&nbsp;</font></td>
      <td style="BORDER-BOTTOM: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent"><font face="Arial Narrow" size="2">Total</font></td>
      <td><font face="Arial Narrow" size="2">&nbsp;</font></td>
      <td><font face="Arial Narrow" size="2">&nbsp;</font></td>
      <td align="Right" style="BORDER-BOTTOM: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent"><font face="Arial Narrow" size="2"><%=FormatNum(Total)%></font></td>
      <td><font face="Arial Narrow" size="2">&nbsp;</font></td>
      <td><font face="Arial Narrow" size="2">&nbsp;</font></td>      
    </tr>	
	
</table>	
</body>

</html>

<%
function convertSign(balance)
	dim sign
	sign = sgn(balance) '1 indicates positive, -1 indicates negative
	if sign = 1 then
		convertSign = "-" & balance	'// make this positive number a negative number
	elseif sign = -1 then
		convertSign = Abs(balance)	'// make this negative number a positive number. Abs removes the - element.
	end if
end function
%>

