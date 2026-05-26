<!--#include file="../libroutines.asp"-->

<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Portfolio BS</title>  
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
FirstDay=DateSerial(Year(Date), Month(Date)-3 + iOffset, 1)
'FirstDay=Date-90
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
	<form method="POST" action="PortfolioBs.asp" Name="frmMain" id="frmMain">
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
	sqlStr = "Select * From [DB_Portfolios] WHERE Client_DPA_ = " & selectedClient & " And TransDate between '" & CDate(selectedFromDate) -1 & "' and '" & CDate(selectedToDate)+1 & "' and OrderTypeSale=0"
	
	'Response.Write sqlstr
	'Response.end
	
	
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
	  <td bgcolor="#000000" width="80%" nowrap align="left"><font color="#FFFFFF" face="Impact" size="2">PORTFOLIO&nbsp;&nbsp;BALANCE SHEET&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></td>
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
  
</table>
<table border="0" cellspacing=2 cellpadding=2 width="900">	
	<tr>
      <td style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2">Code&nbsp;&nbsp;&nbsp;&nbsp;</font></td>
      <td style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2">Share&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></td>
      <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Q</font></td>
      <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Buy</font></td>
      <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2">LAST&nbsp;SPOT&nbsp;PRICE</font></td>
      <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2">Book&nbsp;Value</font></td>
      <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2">Current&nbsp;Value</font></td>
      <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;P/L</font></td>
      <td align="Right" style="border: 1px solid #000000" valign="top"><font face="Arial Narrow" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;%P/L</font></td>      
    </tr>
    <%
Dim Quantity
Dim Price
Dim AVPrice
Dim BookValue
Dim CurrentValue
Dim QuantityPrice

Dim Security1
Dim Security2

Security1=0
Security2=0
SecurityName=""
SecurityCode=""

first=1
    
Do Until Rs.EOF
Security1=rs("Security_DPA_")
	if(Security1<>Security2) and first<>1 then
	AVPrice=QuantityPrice/Quantity
	BookValue=Quantity*AVPrice
	ClientValue=Quantity*Price
	PL=ClientValue-BookValue
	PL_=(PL/BookValue)*100
	%>
	<tr>
      <td><font face="Arial Narrow" size="2"><%=SecurityCode%></font></td>
      <td><font face="Arial Narrow" size="2"><%=SecurityName%></font></td>
      <td align="Right"><font face="Arial Narrow" size="2"><%=formatNumEx(Quantity,0)%></font></td>
      <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(AVPrice)%></font></td>
      <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(Price)%></font></td>
      <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(BookValue)%></font></td>
      <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(ClientValue)%></font></td>
      <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(PL)%></font></td>
      <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNumEx(PL_,2)%>%</font></td>      
    </tr>
	<%
	Equityvalue=Equityvalue+bookValue
	Equityvalue1=Equityvalue1+ ClientValue
	
	Quantity=0
	QuantityPrice=0
	end if
	Quantity=Quantity + rs("LotQty")
	QuantityPrice =QuantityPrice + rs("LotQty")*rs("LotPrice")
		
    CashOnHand=rs("CurrentBal")    
    
    SecurityName=rs("SecurityName")
    Securitycode=rs("SecurityCode")
    Price=rs("SecurityMktPrice")
    
    Security2=Security1
    
    first=0
    	
	Rs.MoveNext
	
	Loop	
	
	AVPrice=QuantityPrice/Quantity
	BookValue=Quantity*AVPrice
	ClientValue=Quantity*Price
	PL=ClientValue-BookValue
	PL_=(PL/BookValue)*100
	
	if(PL_<0.01) then
	PL_=0
	end if
	%>
	<tr>
      <td><font face="Arial Narrow" size="2"><%=SecurityCode%></font></td>
      <td><font face="Arial Narrow" size="2"><%=SecurityName%></font></td>
      <td align="Right"><font face="Arial Narrow" size="2"><%=formatNumEx(Quantity,0)%></font></td>
      <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(AVPrice)%></font></td>
      <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(Price)%></font></td>
      <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(BookValue)%></font></td>
      <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(ClientValue)%></font></td>
      <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNum(PL)%></font></td>
      <td align="Right"><font face="Arial Narrow" size="2"><%=FormatNumEx(PL_,2)%>%</font></td>      
    </tr>
	<%
	Equityvalue=Equityvalue+bookValue
	Equityvalue1=Equityvalue1+ ClientValue
	
	Total=Equityvalue1 + CashOnHand
	
	Set rs = Nothing
	Set Conn = Nothing
	
	TotalPL	=((EquityValue1-EquityValue)/EquityValue)*100
	
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
	<!--
	<tr>
	<td colspan="9">&nbsp;</td>
	</tr>
	-->
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