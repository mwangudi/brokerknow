<!--#include file="../libroutines.asp"-->

<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Portfolio summary</title>  
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
	<form method="POST" action="PortfolioSummary.asp" Name="frmMain" id="frmMain">
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
	sqlStr = "Select * From [DB_Portfolios] WHERE Client_DPA_ = " & selectedClient & " And TransDate between #" & formatdate(dateadd("d",-1,selectedFromDate))  & "# and '" & formatdate(dateadd("d",1,selectedToDate)) & "' order by Security_DPA_,TransDate"
	
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
	  <td bgcolor="#000000" width="80%" nowrap align="left"><font color="#FFFFFF" face="Impact" size="2">PORTFOLIO&nbsp;&nbsp;&nbsp;&nbsp;SUMMARY&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></td>
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
      			<td ><b>From :&nbsp;&nbsp;<%= FormatDate(selectedFromDate) %>&nbsp;&nbsp;&nbsp;&nbsp;To :&nbsp;&nbsp;<%= FormatDate(selectedToDate) %></b></td>      			
    </tr>
</table>
<table border="0" cellspacing=2 cellpadding=2 width="700">		    	    
	<tr>      
      <td align="left" colspan="2" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="3"><b>TRADES</b></font></td>
      <td align="left" colspan="3" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="3"><b>QUANTITIES</b></font></td>            
      <td align="left" colspan="3" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="3"><b>AMOUNTS</b></font></td>            
    </tr>	
	<tr>      
      <td align="left" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>Code&nbsp;&nbsp;</b></font></td>      
      <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>Security&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b></font></td>                        
      <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>Bought</b></font></td>      
      <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>Sold</b></font></td>
      <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>Balance</b></font></td>      
      <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>Bought</b></font></td>      
      <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>Sold</b></font></td>
      <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b>Balance</b></font></td>      
    </tr>
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

BoughtQty=0
SoldQty=0
BalanceQty=0
BoughtAmount=0
SoldAmount=0
BalanceAmount=0

PreviousBoughtQty=0
PreviousSoldQty=0
PreviousBalanceQty=0
PreviousBoughtAmount=0
PreviousSoldAmount=0
PreviousBalanceAmount=0

SecurityBoughtQty=0
SecurityBoughtAmount=0
SecuritySoldQty=0
SecuritySoldAmount=0

TotalBoughtAmount=0
TotalSoldAmount=0
TotalBalanceAmount=0
	
PreviousBookaValue=0
PreviousBuy=0

TotalBookValue=0	
first=1
    
Do Until Rs.EOF
Security1=rs("Security_DPA_")	
	if(Security1<>Security2) then
		if(first=0) then		
		%>
		<tr>      
		<td><font face="Arial Narrow" size="2"><%=PreviousCode%></font></td>      
		<td><font face="Arial Narrow" size="2"><%=PreviousSecurity%></font></td>						
		<td align="Right"><font face="Arial Narrow" size="2"><%=FormatNumEx(SecurityBoughtQty,0)%></font></td>      
		<td align="Right"><font face="Arial Narrow" size="2"><%=FormatNumEx(SecuritySoldQty,0)%></font></td>
		<td align="Right"><font face="Arial Narrow" size="2"><%=FormatNumEx(PreviousBalanceQty,0)%></font></td>      
		<td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(SecurityBoughtAmount)%></font></td>      
		<td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(SecuritySoldAmount)%></font></td>
		<td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(PreviousBalanceAmount)%></font></td>      		
		</tr>
		<%
		TotalBookValue=TotalBookValue+PreviousBookaValue
		
		TotalBoughtAmount=TotalBoughtAmount+SecurityBoughtAmount
		TotalSoldAmount=TotalSoldAmount+SecuritySoldAmount
		TotalBalanceAmount=TotalBalanceAmount+PreviousBalanceAmount
		
		TotalBoughtQty=TotalBoughtQty+SecurityBoughtQty
		TotalSoldQty=TotalSoldQty+SecuritySoldQty
		TotalBalanceQty=TotalBalanceQty+PreviousBalanceQty
		
		SecurityBoughtQty=0	
		SecurityBoughtAmount=0
		SecuritySoldQty=0
		SecuritySoldAmount=0
	
		end if
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
		 
		 'Response.Write(rs("OrderTypeSale"))
		 
		 if(rs("OrderTypeSale")=1) then
			BoughtQty=0
			SoldQty=rs("LotQty")			
			BoughtAmount=0
			SoldAmount=rs("NetAmount")
		 else
		 	BoughtQty=rs("LotQty")			
			SoldQty=0
			BoughtAmount=rs("NetAmount")
			SoldAmount=0
		 end if
		 BalanceQty=BookQty		 
      
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
	BalanceAmount=BookValue		 
	
	PreviousBookPrice=BookPrice
	PreviousBookQty=BookQty
	PreviousBookValue=Bookvalue
		
	PreviousBookaValue=BookValue
	PreviousBuy=BookPrice
	
	PreviousSecurity=rs("SecurityName")
	PreviousCode=rs("SecurityCode")
			
	SecurityBoughtQty=SecurityBoughtQty+BoughtQty	
	SecurityBoughtAmount=SecurityBoughtAmount+BoughtAmount
	SecuritySoldQty=SecuritySoldQty+soldQty
	SecuritySoldAmount=SecuritySoldAmount+SoldAmount
	
	PreviousBoughtQty=BoughtQty
	PreviousSoldQty=SoldQty
	PreviousBalanceQty=BalanceQty
	PreviousBoughtAmount=BoughtAmount
	PreviousSoldAmount=SoldAmount
	PreviousBalanceAmount=BalanceAmount
	
	Security2=Security1
	first=0
	Rs.MoveNext
	
	Loop
	TotalBookValue=TotalBookValue+PreviousBookaValue
	TotalBoughtAmount=TotalBoughtAmount+SecurityBoughtAmount
	TotalSoldAmount=TotalSoldAmount+SecuritySoldAmount
	TotalBalanceAmount=TotalBalanceAmount+PreviousBalanceAmount
	
	TotalBoughtQty=TotalBoughtQty+SecurityBoughtQty
	TotalSoldQty=TotalSoldQty+SecuritySoldQty
	TotalBalanceQty=TotalBalanceQty+PreviousBalanceQty
		
	%>	
	<tr>      
		<td><font face="Arial Narrow" size="2"><%=PreviousCode%></font></td>      
		<td><font face="Arial Narrow" size="2"><%=PreviousSecurity%></font></td>						
		<td align="Right"><font face="Arial Narrow" size="2"><%=FormatNumEx(SecurityBoughtQty,0)%></font></td>      
		<td align="Right"><font face="Arial Narrow" size="2"><%=FormatNumEx(SecuritySoldQty,0)%></font></td>
		<td align="Right"><font face="Arial Narrow" size="2"><%=FormatNumEx(PreviousBalanceQty,0)%></font></td>      
		<td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(SecurityBoughtAmount)%></font></td>      
		<td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(SecuritySoldAmount)%></font></td>
		<td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(PreviousBalanceAmount)%></font></td>      		
	</tr>
	<tr>      
		<td align="right" colspan="8"><font face="Arial Narrow" size="3">&nbsp;</font></td>      				
	</tr>
	<tr>      
		<td align="right" colspan="2"><font face="Arial Narrow" size="2"><b>PORTFOLIO&nbsp;TOTAL</b></font></td>      		
		<td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b><%=FormatNumEx(TotalBoughtQty,0)%></b></font></td>
		<td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b><%=FormatNumEx(TotalSoldQty,0)%></b></font></td>      
		<td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b><%=FormatNumEx(TotalBalanceQty,0)%></b></font></td>            
		<td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b><%=FormatNum(TotalBoughtAmount)%></b></font></td>
		<td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b><%=FormatNum(TotalSoldAmount)%></b></font></td>      
		<td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2"><b><%=FormatNum(TotalBalanceAmount)%></b></font></td>            
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

