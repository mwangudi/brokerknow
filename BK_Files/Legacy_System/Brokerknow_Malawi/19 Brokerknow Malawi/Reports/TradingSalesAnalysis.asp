<!--#include file="../libroutines.asp"-->

<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Trading History</title>  
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
FirstDay=Date-90
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
	<form method="POST" action="TradingSalesAnalysis.asp" Name="frmMain" id="frmMain">
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
	sqlStr = "Select * From [DB_Portfolios] WHERE Client_DPA_ = " & selectedClient & " And TransDate between #" & formatdate(dateadd("d",-1,selectedFromDate)) & "# and #" & formatdate(dateadd("d",1,selectedToDate)) & "# order by Security_DPA_,TransDate"
	
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
	  <td bgcolor="#000000" width="80%" nowrap align="left"><font color="#FFFFFF" face="Impact" size="2">TRADING&nbsp;&nbsp;&nbsp;&nbsp;SALES&nbsp;&nbsp;&nbsp;&nbsp;ANALYSIS&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></td>
      <td bgcolor="#000000" width="20%" nowrap align=right><font color="#FFFFFF" face="Impact" size="2"><%= Session("CompanyName") %></font></td>
      
    </tr>

  </table>
<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" width="100%" > 
	<!--
	<THEAD>
	
	<tr class="pageNumbering">
		<td align="left" height="18">
			<FONT FACE=ARIAL SIZE=2><B>Page <%=pageNumber%>	</B></FONT>	
		</td>		
	</tr>
	
	<tr>
		<td align="right" height="200" style="border-bottom: 2px inset #000000">
			<Img Src="../data/photos/aaprintlogo.jpg">			
		</td>		
	</tr>	
	<THEAD>   
	-->
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
  <tr>      
      			<td ><b>From&nbsp;&nbsp;<%= FormatDate(selectedFromDate) %>&nbsp;&nbsp;&nbsp;&nbsp;To&nbsp;&nbsp;<%= FormatDate(selectedToDate) %></b></td>      			
    </tr>
</table>
<table border="0" cellspacing=2 cellpadding=2 width="400">		
    	    <%
Dim Security1
Dim Security2

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

SecurityQtyTotal=0
SecurityBuyTotal=0
SecurityAmtTotal=0
SecurityPLTotal=0
SecurityPL_Total=0

PreviousBookaValue=0
PreviousBuy=0

TotalAMT=0
TotalPL=0
TotalPL_=0	

first=1
    
Do Until Rs.EOF
Security1=rs("Security_DPA_")	
	if(Security1<>Security2) then
		if(first=0) then
			if(SecurityQtyTotal<>0) then
			SecurityBuyTotal=SecurityAmtTotal/SecurityQtyTotal
			else
			SecurityBuyTotal=0
			end if		
			
			if(SecurityAmtTotal<>0) then
			SecurityPL_Total=(SecurityPLTotal/SecurityAmtTotal)*100
			else
			SecurityPL_Total=0
			end if
			
		%>
		<tr>      
		<td align="right" colspan="5"><font face="Arial Narrow" size="3"><b>Security&nbsp;&nbsp;Total</b></font></td>      
		<td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b><%=FormatNumEx(SecurityQtyTotal,0)%></b></font></td>      
		<td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b><%=FormatNum(SecurityBuyTotal)%></font></b></td>
		<td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b><%=FormatNum(SecurityAmtTotal)%></b></font></td>      
		<td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b><%=FormatNum(SecurityPLTotal)%></font></b></td>
		<td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b><%=FormatNum(SecurityPL_Total)%>&nbsp;%</b></font></td>      
		</tr>
		<%		
		TotalAMT=TotalAMT+SecurityAmtTotal
		TotalPL=TotalPL+SecurityPLTotal
		
		SecurityQtyTotal=0
		SecurityBuyTotal=0
		SecurityAmtTotal=0
		SecurityPLTotal=0
		SecurityPL_Total=0

		end if
	%>
	<tr>      
      <td align="left" colspan="10"><font face="Arial Narrow" size="3"><b>Security:&nbsp;<%=rs("SecurityName")%></b></font></td>      
    </tr>
    <tr>      
      <td align="left" colspan="5" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="3"><b>TRADES</b></font></td>      
      <td align="left" colspan="3" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="3"><b>HOLDINGS&nbsp;&nbsp;SOLD</b></font></td>            
      <td align="left" colspan="2" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="3"><b>BOOK&nbsp;&nbsp;PROFIT</b></font></td>
    </tr>	
	<tr>      
      <td align="left" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>Date&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b></font></td>
      <td align="left" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>Ref&nbsp;&nbsp;</b></font></td>      
      <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>&nbsp;&nbsp;&nbsp;Quantity</b></font></td>      
      <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>&nbsp;Price</b></font></td>
      <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>Net&nbsp;Amount</b></font></td>
      <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>Quantity</b></font></td>      
      <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>&nbsp;&nbsp;&nbsp;Buy</b></font></td>
      <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>Book&nbsp;AMT</b></font></td>      
      <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>&nbsp;&nbsp;&nbsp;P/(L)</b></font></td>
      <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>&nbsp;P/(L)&nbsp;%</b></font></td>      
    </tr>
	<%
	end if
	%>
	<tr>      
      <td><font face="Arial Narrow" size="2"><%=FormatDate(rs("TransDate"))%></font></td>
      <td><font face="Arial Narrow" size="2"><%=rs("ContractNumber")%></font></td>      
      <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNumEx(rs("LotQty"),0)%></font></td>      
      <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(rs("LotPrice"))%></font></td>
      <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(rs("NetAmount"))%></font></td>
	<% 
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
					BookPrice=((PreviousBookPrice*PreviousBookQty)+(rs("NetAmount")/(rs("LotQty"))*rs("LotQty")))/(rs("LotQty")+PreviousBookQty)
					end if
				end if
			end if
			
			BookValue=BookQty*BookPrice
			
	if(rs("OrderTypeSale")=false and BookPrice=0) then
      QtySold=0
      else
		if(Security1<>Security2) then
		QtySold=0
		else
			if(rs("OrderTypeSale")=1) then
				if(rs("LotQty")>PreviousBookQty) then
				QtySold=PreviousBookQty
				else
				Qtysold=Rs("LotQty")
				end if
			else
			QtySold=0	
			end if
		end if
      end if

            %>
      <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNumEx(QtySold,0)%></font></td>      
	<%
			if(QtySold=0) then
      		BookPriceSold=0
      		else
				if(Security1<>Security2) then
				BookPriceSold=0
				else
				BookPriceSold=PreviousBookPrice
				end if
      		end if


			if(rs("OrderTypeSale")=1) then
			SellPrice=rs("NetAmount")/rs("LotQty")						
			else
			SellPrice=0
			end if

      SoldValue=Qtysold*SellPrice
      
      %>
      <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(BookPriceSold)%></font></td>
	<%BookValueSold=QtySold*BookPriceSold%>
      <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(BookValueSold)%></font></td>	
	<%
	if(BookPriceSold=0) then
      PL=0
      else
		if(QtySold=0) then
		PL=0
		else
		PL=SoldValue-BookValueSold
		end if
      end if
      %>     
      <td align="Right"><font face="Arial Narrow" size="2"><%=Formatnum(PL)%></font></td>
	<%
      if(PL=0) then
      PL_=0
      else
		if(BookValueSold=0) then
		PL_=0
		else
		PL_=(PL/BookValueSold)*100
		end if
      end if
      %>
      
      <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(PL_)%>%</font></td>          
	</tr>

	<%
	PreviousBookPrice=BookPrice
	PreviousBookQty=BookQty
	PreviousBookValue=Bookvalue
		
	PreviousBookaValue=BookValue
	PreviousBuy=BookPrice
	
	SecurityQtyTotal=SecurityQtyTotal+QtySold
	SecurityAmtTotal=SecurityAmtTotal+BookValueSold	
	
	SecurityPLTotal=SecurityPLTotal+PL

	Security2=Security1
	first=0
	Rs.MoveNext
	
	Loop
	
	TotalAMT=TotalAMT+SecurityAmtTotal
	TotalPL=TotalPL+SecurityPLTotal
	
			if(SecurityQtyTotal<>0) then
			SecurityBuyTotal=SecurityAmtTotal/SecurityQtyTotal
			else
			SecurityBuyTotal=0
			end if		
			
			if(SecurityAmtTotal<>0) then
			SecurityPL_Total=(SecurityPLTotal/SecurityAmtTotal)*100
			else
			SecurityPL_Total=0
			end if
		
		if(TotalAMT<>0) then
		TotalPL_=(TotalPL/TotalAMT)*100
		else
		TotalPL_=0
		end if	
		%>
		<tr>      
		<td align="right" colspan="5"><font face="Arial Narrow" size="3"><b>Security&nbsp;&nbsp;Total</b></font></td>      
		<td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b><%=FormatNumEx(SecurityQtyTotal,0)%></b></font></td>      
		<td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b><%=FormatNum(SecurityBuyTotal)%></font></b></td>
		<td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b><%=FormatNum(SecurityAmtTotal)%></b></font></td>      
		<td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b><%=FormatNum(SecurityPLTotal)%></font></b></td>
		<td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b><%=FormatNum(SecurityPL_Total)%>&nbsp;%</b></font></td>      
		</tr>
		<tr>      
		<td align="right" colspan="8"><font face="Arial Narrow" size="3">&nbsp;</font></td>      				
	</tr>
	<tr>      
		<td align="right" colspan="7"><font face="Arial Narrow" size="2"><b>TRADING&nbsp;TOTAL</b></font></td>      		
		<td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b><%=FormatNum(TotalAMT)%></b></font></td>
		<td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b><%=FormatNum(TotalPL)%></b></font></td>      
		<td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b><%=FormatNum(TotalPL_)%>&nbsp;%</b></font></td>            
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

