<!--#include file="../libroutines.asp"-->

<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Portifolio Valuation Statement</title>  
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
FirstDay=CDate("3/1/2005")
genReport = Request.Form("genReport")
selectedClient = Request.Form("cboClient")
selectedFromDate = Request.Form("txtFromDate")
selectedToDate = Request.Form("txtToDate")
If genReport <> "1" Then%>
		<Script Language="JavaScript">
		report_SetBodyClass();
		function validateForm(frm){			
			//if (frm.txtDate.value==''){
			//	alert("Select a date");
			//	frm.txtDate.focus();
			//	return;
			//}
			
			frm.target = '_self';			
			frm.submit();
		}
		
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "txtFromDate","cmdDate","<%= FormatDate(FirstDay) %>",1);
		var cal1=new ctlSpiffyCalendarBox("cal1", "frmMain", "txtToDate","cmdDate","<%= FormatDate(Date) %>",1);
	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="HoldingsValuation.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">	
		<table>			
			<tr>
				<td>Client: </td>
				<td><input type = 'text' name ='txtClientCode' id = 'txtClientCode' size="10" onBlur="txtval = this.value; selectItem(cboClient);"></td>
				<td><select name = 'cboClient' id = "cboClient" size="1" 
    				onchange='UpdateCode(true,cboClient,txtClientCode)'
					onKeypress="return (dodefaultaction()==''); " 
					onKeydown="return (dodefaultaction()==''); " 
					onKeyup="return (FilterData(this,1,UpdateCode(change(cboClient,0),cboClient,txtClientCode)));" 
					onfocus="txtval = '';inputIsItemCode = 1;" 
					onblur="txtval = '';inputIsItemCode = 1;">
					<option selected SearchCode = "" SearchText = ""  value = ''></option>
					<%
					dim ClientName
					dim NameClient
					        Set conn = GetActiveConnection("KBroker")
					        
					        sqlStr = "SELECT * FROM FullClientList order by ClientName"
					        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
					        If Not (rs.EOF Or rs.BOF) Then
					                rs.MoveFirst
					                Do Until rs.EOF
					                ClientName=rs.Fields("ClientName")
					                NameClient=Mid(ClientName,1,30)
					                %>					                        
					                        <option SearchCode = "<%=rs.Fields("Client_DPA_")%>" SearchText = "<%=NameClient%>" value = '<%=rs.Fields("Client_DPA_")%>'><%=NameClient%></option>

					                        <%rs.MoveNext
					                Loop
					        End If
					%>

					    </select>
				</td>
			</tr>
			<tr>
				<td>From Date:</td>
				<td>
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>
			</tr>
			<tr>
				<td>To date:</td>
				<td>
					<SCRIPT language="JavaScript">cal1.writeControl();</SCRIPT>	
				</td>
			</tr>
			
			<tr>
				<td colspan=2><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... " id=Button1 name=Button1>&nbsp;&nbsp;</td>
			</tr>
		</table>
		
	</form>
	
	<%
	Response.End
End If

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
	sqlStr = "Select * From [DB_Portfolios] WHERE Client_DPA_ = " & selectedClient & " And TransDate between '" & CDate(selectedFromDate)-1 & "' and '" & CDate(selectedToDate)+1 & "' order by Security_DPA_,TransDate"
	
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

<table border="0" cellspacing=2 cellpadding=2 width="800" > 
	<tr><td style="border: 2px solid #000000" align="center" height="100"><font size="5"><b>Portifolio Valuation Statement</b></font>
	<br><font size="2"><b><% tarehe = split(formatdatefull(date),",")
		response.write replace(tarehe(1),"-"," ")
	%></b></font></td></tr>
</table>
<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" width="100%" > 
	<tr>
		<td align="right" height="8">
            &nbsp;			
		</td>		
	</tr>  
</table>
<table border="0" cellspacing=2 cellpadding=2 width="800" > 
	<tr><td style="border: 1px solid #000000" align="center" height="25">
	<table border="0" cellspacing=2 cellpadding=2 width="800">		
		  <tr>
			   
		  <td align="left" style="border: 0px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>Instrument&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b></font></td>                        
		  <td align="Right" style="border: 0px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>Holding</b></font></td>      
		  <td align="Right" style="border: 0px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>&nbsp;&nbsp;&nbsp;Price</b></font></td>
		  <td align="Right" style="border: 0px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>Book&nbsp;Value</b></font></td>
		  <td align="Right" style="border: 0px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>Market&nbsp;Value</b></font></td>
		  <td align="Right" style="border: 0px solid #000000" valign="top" nowrap><font face="Arial Narrow" size="2"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Profit/(Loss)</b></font></td>
		  <td align="Right" style="border: 0px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Retuns</b></font></td>            
		</tr>
	</td></tr></table>
	</td></tr></table>
	<table border="0" cellspacing=0 cellpadding=2 width="800">		
	<td align="Left" style="border: 0px solid #000000" valign="top"><font face="Arial Narrow" size="2" ><b>Equity&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b></font></td>                        
		  <td align="Right" style="border: 0px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>&nbsp;&nbsp;</b></font></td>      
		  <td align="Right" style="border: 0px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>&nbsp;&nbsp;&nbsp;</b></font></td>
		  <td align="Right" style="border: 0px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>&nbsp;&nbsp;&nbsp</b></font></td>
		  <td align="Right" style="border: 0px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>&nbsp;&nbsp;&nbsp;</b></font></td>
		  <td align="Right" style="border: 0px solid #000000" valign="top" nowrap><font face="Arial Narrow" size="2"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp</b></font></td>
		  <td align="Right" style="border: 0px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp</b></font></td>            
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
				<td><font face="Arial Narrow" size="2">&nbsp;&nbsp;&nbsp<%=PreviousSecurity%></font></td>		
				<td align="Right"><font face="Arial Narrow" size="2">&nbsp;&nbsp;&nbsp<%=FormatNumEx(PreviousBookQty,0)%></font></td>      
				<td align="Right"><font face="Arial Narrow" size="2">&nbsp;&nbsp;&nbsp<%=FormatNum(PreviousBuy)%></font></td>
				<td align="Right"><font face="Arial Narrow" size="2">&nbsp;&nbsp;&nbsp<%=FormatNum(PreviousBookaValue)%></font></td>      
				<td align="Right"><font face="Arial Narrow" size="2">&nbsp;&nbsp;&nbsp<%=FormatNum(ClientValue)%></font></td>
				<td align="Right"><font face="Arial Narrow" size="2">&nbsp;&nbsp;&nbsp<%=FormatNum(PL)%></font></td>
				<td align="Right"><font face="Arial Narrow" size="2">&nbsp;&nbsp;&nbsp<%=FormatNumEx(PL_,2)%>%</font></td>      
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
			<td><font face="Arial Narrow" size="2">&nbsp;&nbsp;&nbsp<%=PreviousSecurity%></font></td>		
			<td align="Right"><font face="Arial Narrow" size="2">&nbsp;&nbsp;&nbsp<%=FormatNumEx(PreviousBookQty,0)%></font></td>      
			<td align="Right"><font face="Arial Narrow" size="2">&nbsp;&nbsp;&nbsp<%=FormatNum(PreviousBuy)%></font></td>
			<td align="Right"><font face="Arial Narrow" size="2">&nbsp;&nbsp;&nbsp<%=FormatNum(PreviousBookaValue)%></font></td>
			<td align="Right"><font face="Arial Narrow" size="2">&nbsp;&nbsp;&nbsp<%=FormatNum(ClientValue)%></font></td>
			<td align="Right"><font face="Arial Narrow" size="2">&nbsp;&nbsp;&nbsp<%=FormatNum(PL)%></font></td>
			<td align="Right"><font face="Arial Narrow" size="2">&nbsp;&nbsp;&nbsp<%=FormatNumEx(PL_,2)%>%</font></td>            
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
	<tr height="25">
     <td style="BORDER-BOTTOM: #000000 1px inset; BORDER-LEFT: #000000 1px inset; BORDER-TOP: #000000 1px inset; BACKGROUND-COLOR: transparent"><font face="Arial Narrow" size="2" ><b><i>Total Equity Value</i></b></font></td>
      <td style="BORDER-BOTTOM: #000000 1px inset; BORDER-TOP: #000000 1px inset; BACKGROUND-COLOR: transparent"><font face="Arial Narrow" size="2" >&nbsp;</font></td>
      <td style="BORDER-BOTTOM: #000000 1px inset; BORDER-TOP: #000000 1px inset; BACKGROUND-COLOR: transparent"><font face="Arial Narrow" size="2">&nbsp;</font></td>
      <td align="Right" style="BORDER-BOTTOM: #000000 1px inset; BORDER-TOP: #000000 1px inset; BACKGROUND-COLOR: transparent"><font face="Arial Narrow" size="2"><%=FormatNum(EquityValue)%></font></td>
      <td align="Right" style="BORDER-BOTTOM: #000000 1px inset; BORDER-TOP: #000000 1px inset; BACKGROUND-COLOR: transparent"><font face="Arial Narrow" size="2"><%=FormatNum(EquityValue1)%></font></td>
      <td align="Right" style="BORDER-BOTTOM: #000000 1px inset; BORDER-TOP: #000000 1px inset; BACKGROUND-COLOR: transparent"><font face="Arial Narrow" size="2"><%=FormatNum(EquityValue1-EquityValue)%></font></td>
      <td align="Right" style="BORDER-BOTTOM: #000000 1px inset; BORDER-TOP: #000000 1px inset; BORDER-RIGHT: #000000 1px inset; BACKGROUND-COLOR: transparent"><font face="Arial Narrow" size="2"><%=FormatNumEx(TotalPL,2)%>%</font></td>      
    </tr>
    <tr>
		<td colspan="9">&nbsp;&nbsp;</td>
	</tr>
	<tr>
		<td colspan="9">&nbsp;&nbsp;</td>
	</tr>
    <tr>
		<td colspan="9">&nbsp;&nbsp;<b><i>Money Market</i></b></td>
	</tr>
	
	<tr>
     <td ><font face="Arial Narrow" size="2">&nbsp;&nbsp;&nbsp;&nbsp;Cash&nbsp;In&nbsp;Hand</font></td>
      <td ><font face="Arial Narrow" size="2">&nbsp;</font></td>
      <td><font face="Arial Narrow" size="2">&nbsp;</font></td>
      <td><font face="Arial Narrow" size="2">&nbsp;</font></td>
      <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(CashOnHand)%></font></td>
      <td><font face="Arial Narrow" size="2">&nbsp;</font></td>
      <td><font face="Arial Narrow" size="2">&nbsp;</font></td>      
    </tr>
    
	<tr>
		<td colspan="9">&nbsp;&nbsp;</td>
	</tr>
   	<tr>
       <td style="BORDER-BOTTOM: #000000 1px inset; BORDER-LEFT: #000000 1px inset; BORDER-TOP: #000000 1px inset; BACKGROUND-COLOR: transparent"><font face="Arial Narrow" size="2">&nbsp;&nbsp;<b><i>Total Money Market Value</i></b></font></td>
      <td style="BORDER-BOTTOM: #000000 1px inset; BORDER-TOP: #000000 1px inset; BACKGROUND-COLOR: transparent"><font face="Arial Narrow" size="2">&nbsp;</font></td>
      <td style="BORDER-BOTTOM: #000000 1px inset; BORDER-TOP: #000000 1px inset; BACKGROUND-COLOR: transparent"><font face="Arial Narrow" size="2">&nbsp;</font></td>
      <td style="BORDER-BOTTOM: #000000 1px inset; BORDER-TOP: #000000 1px inset; BACKGROUND-COLOR: transparent"><font face="Arial Narrow" size="2">&nbsp;</font></td>
      <td align="Right" style="BORDER-BOTTOM: #000000 1px inset; BORDER-TOP: #000000 1px inset; BACKGROUND-COLOR: transparent"><font face="Arial Narrow" size="2"><%=FormatNum(Total)%></font></td>
      <td style="BORDER-BOTTOM: #000000 1px inset; BORDER-TOP: #000000 1px inset; BACKGROUND-COLOR: transparent"><font face="Arial Narrow" size="2">&nbsp;</font></td>
      <td style="BORDER-BOTTOM: #000000 1px inset; BORDER-TOP: #000000 1px inset; BORDER-RIGHT: #000000 1px inset; BACKGROUND-COLOR: transparent"><font face="Arial Narrow" size="2">&nbsp;</font></td>      
    </tr>	
	<tr>
		<td colspan="9">&nbsp;&nbsp;</td>
	</tr>
   	<tr>
       <td style="BORDER-BOTTOM: #000000 1px inset; BORDER-LEFT: #000000 1px inset; BORDER-TOP: #000000 1px inset; BACKGROUND-COLOR: transparent"><font face="Arial Narrow" size="2">&nbsp;&nbsp;<b><i>Total Portifolio Value</i></b></font></td>
      <td style="BORDER-BOTTOM: #000000 1px inset; BORDER-TOP: #000000 1px inset; BACKGROUND-COLOR: transparent"><font face="Arial Narrow" size="2">&nbsp;</font></td>
      <td style="BORDER-BOTTOM: #000000 1px inset; BORDER-TOP: #000000 1px inset; BACKGROUND-COLOR: transparent"><font face="Arial Narrow" size="2">&nbsp;</font></td>
      <td style="BORDER-BOTTOM: #000000 1px inset; BORDER-TOP: #000000 1px inset; BACKGROUND-COLOR: transparent"><font face="Arial Narrow" size="2">&nbsp;</font></td>
      <td align="Right" style="BORDER-BOTTOM: #000000 1px inset; BORDER-TOP: #000000 1px inset; BACKGROUND-COLOR: transparent"><font face="Arial Narrow" size="2"><%=FormatNum(Total)%></font></td>
      <td style="BORDER-BOTTOM: #000000 1px inset; BORDER-TOP: #000000 1px inset; BACKGROUND-COLOR: transparent"><font face="Arial Narrow" size="2">&nbsp;</font></td>
      <td style="BORDER-BOTTOM: #000000 1px inset; BORDER-TOP: #000000 1px inset; BORDER-RIGHT: #000000 1px inset; BACKGROUND-COLOR: transparent"><font face="Arial Narrow" size="2">&nbsp;</font></td>      
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

