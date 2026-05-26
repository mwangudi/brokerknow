<!--#include file="../libroutinesTEST.asp"-->

<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">

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

<body Class="Reports" leftmargin="25">
<%
FirstDay=Now()
genReport = Request.Form("genReport")
selectedClient = Request.Form("cboClient")
selectedFromDate = Request.Form("txtFromDate")
If genReport <> "1" Then%>
		<Script Language="JavaScript">
		report_SetBodyClass();
		function validateForm(frm){			
			
			frm.target = '_self';			
			frm.submit();
		}
		
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "txtFromDate","cmdDate","<%= FormatDate(FirstDay) %>",1);
	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000; width:185; height:10"></div>
	<form method="POST" action="HoldingsValuationSimple.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">	
		<table width="567">			
			<tr>
				<td width="81">Client: </td>
				<td width="470"><input type = 'text' name ='txtClientCode' id = 'txtClientCode' size="10" onBlur="txtval = this.value; selectItem(cboClient);"><span lang="en-gb">&nbsp;
                </span><select name = 'cboClient' id = "cboClient" size="1" 
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

					    </select></td>
				
			</tr>
			<tr>
				<td width="81">Select date:</td>
				<td width="470">
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>
			</tr>
			
			<tr>
				<td colspan=2 width="555" ><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... " id=Button1 name=Button1>&nbsp;&nbsp;</td>
			</tr>
		</table>
		
	</form>
	
	
	<%
	Response.End
End If

%>

<% DrawPageFunctions True, True, True, True %>

<%

	If selectedClient="" Then
		%>
		<Script Language="JavaScript">
			alert("Please select The client")
			window.history.go(-1);
		</Script>
		<%
		Response.End
	End If

	'=============================================================================
	' Generate Client Holdings Valuation 
	'=============================================================================

	Set conn = GetActiveConnection("KBroker")
	Set Rs = CreateObject("ADODB.Recordset")						        
	sqlStr = "Select * From [DB_Portfolios] WHERE Client_DPA_ = " & selectedClient & " And Cast(floor(Cast(TransDate as Float)) as datetime) <= '" & formatdate(CDate(selectedFromDate)+1) & "' order by Security_DPA_,TransDate"

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
	
	sqlStr = "SELECT Client.ClientAddr,Client.ClientContact,'W' + ClientOfficeTel + '/' + Client.ClientEmail as Contacts,ClientName FROM Client " & _			
			" WHERE Client_DPA_ = " & selectedClient
	
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

<p id="toPDFOrient" name="toPDFOrient" value="L" style="display:none;">L
<p id="toPDF" name="toPDF">

<table border=0 align=center>
	<tr><td height=50>&nbsp;</td></tr>
</table>

<table border="0" cellspacing=2 cellpadding=2 width="971" height="600" align="center"> 
	<tr>
		<td style="border: 2px solid #000000" align="center" height="70" width="970">
			<table border =0 width="734" height="100%">
			<tr><td align="center" height="500" width="728"><font size="7"><b><%= ClientName%></b></font></td></tr>
			<tr><td valign="bottom" align="center" width="728"><Img Src="../data/photos/aaprintlogofooter.jpg"></td></tr>
			</table>
		</td>
	</tr>
</table>

<!--<table border=0 align=center>
	<tr><td height=200>&nbsp;</td></tr>
</table>  -->

<BR class="newpage">

<table border=0 align=center>
	<tr><td height=115>&nbsp;</td></tr>
</table>

<table border="0" cellspacing=0 cellpadding=2 width="971" style="border-collapse: collapse" bordercolor="#111111"  align="center"> 
	<tr>
		<td style="border: 2px solid #000000" align="center" height="100" width="907"><font size="5"><b>Portfolio Valuation Statement</b></font><br><font size="2"><b>
			<%
			tarehe = split(formatdatefull(selectedFromDate),",")
			response.write replace(tarehe(1),"-"," ")
			%>
		</b></font></td>
	</tr>
</table>

<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" width="760"  align="center"> 
	<tr><td align="right" height="8" width="801">&nbsp;</td></tr>  
</table>

</td></tr>
</table>
	
<table border="0" cellspacing=0 cellpadding=2 width="969" style="border-collapse: collapse" bordercolor="#111111" height="243"  align="center">	
	<tr>
		<td align="right" height="18" width="398" style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;"><font face="Arial Narrow" size="2"><b>Instrument</b></font></td>		
		<td align="right" height="18" width="87" style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;"><font face="Arial Narrow" size="2"><b>Holding</b></font></td>      
		<td align="right" height="18" width="91" style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;"><font face="Arial Narrow" size="2"><b>Book&nbsp;Price</b></font></td>
		<td align="right" height="18" width="83" style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;"><font face="Arial Narrow" size="2"><b>Book&nbsp;Value</b></font></td>      
		<td align="right" height="18" width="98" style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;"><font face="Arial Narrow" size="2"><b>Market&nbsp;Price</b></font></td>      
		<td align="right" height="18" width="129" style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;"><font face="Arial Narrow" size="2"><b>Market&nbsp;Value</b></font></td>
		<td align="right" height="18" width="122" style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;"><font face="Arial Narrow" size="2"><b>Profit/(Loss)</b></font></td>
	</tr>

	<tr>	
		<td align="Left" style="border: 0px solid #000000" valign="top" colspan="6" width="873" height="1"></td>                        
	</tr>

	<tr>	
		<td align="Left" style="border: 0px solid #000000" valign="top" height="16" colspan="6" width="873"><font face="Arial Narrow" size="2" ><b>Equity&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b></font></td>                        
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
MarketPrice = 0
PreviousMarketPrice = 0

PreviousSecurity=""
PreviousCode=""

PreviousBookaValue=0
PreviousBuy=0

TotalBookValue=0	
first = 1
    
Do Until Rs.EOF
Security1=rs("Security_DPA_")	
	if(Security1<>Security2) then
		if(first=0) then
			if(PreviousBookQty<>0) then
				ClientValue=PreviousBookQty * Price
				
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
					<td align="Right" height="1" width="398"><font face="Arial Narrow" size="2"><%=PreviousSecurity%></font></td>		
					<td align="Right" height="1" width="87"><font face="Arial Narrow" size="2"><%=FormatNumEx(PreviousBookQty,0)%></font></td>      
					<td align="Right" height="1" width="91"><font face="Arial Narrow" size="2"><%=FormatNum(PreviousBuy)%></font></td>
					<td align="Right" height="1" width="83"><font face="Arial Narrow" size="2"><%=FormatNum(PreviousBookaValue)%></font></td>      
					<td align="Right" height="1" width="98"><font face="Arial Narrow" size="2"><%=FormatNum(Price)%></font></td>      
					<td align="Right" height="1" width="129"><font face="Arial Narrow" size="2"><%=FormatNum(ClientValue)%></font></td>
					<td align="Right" height="1" width="122"><font face="Arial Narrow" size="2"><%=FormatNum(PL)%></font></td>
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
			'Price = rs("SecurityMktPrice")
			
			Price = GetMarketPrice(rs("Security_DPA_"),FormatDate(CDate(selectedFromDate)))
			
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
				<td align="Right" height="1" width="398"><font face="Arial Narrow" size="2"><%=PreviousSecurity%></font></td>		
				<td align="Right" height="1" width="87"><font face="Arial Narrow" size="2"><%=FormatNumEx(PreviousBookQty,0)%></font></td>      
				<td align="Right" height="1" width="91"><font face="Arial Narrow" size="2"><%=FormatNum(PreviousBuy)%></font></td>
				<td align="Right" height="1" width="83"><font face="Arial Narrow" size="2"><%=FormatNum(PreviousBookaValue)%></font></td>
				<td align="Right" height="1" width="98"><%=FormatNum(Price)%></td>
				<td align="Right" height="1" width="129"><font face="Arial Narrow" size="2"><%=FormatNum(ClientValue)%></font></td>
				<td align="Right" height="1" width="122"><font face="Arial Narrow" size="2"><%=FormatNum(PL)%></font></td>           
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
	
		if(Abs(TotalPL)<0.01) then
		TotalPL=0
		end if
		%>
	<tr>
		<td colspan="7" height="14" width="941">&nbsp;&nbsp;</td>
	</tr>

	<tr height="25">
		<td style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="18" width="398"><font face="Arial Narrow" size="2" ><b><i>Total Equity Value</i></b></font></td>
		<td style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="18" width="87"><font face="Arial Narrow" size="2" >&nbsp;</font></td>
		<td style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="18" width="91"><font face="Arial Narrow" size="2">&nbsp;</font></td>
		<td align="Right" style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="18" width="83"><font face="Arial Narrow" size="2"><%=FormatNum(EquityValue)%></font></td>
		<td align="Right" style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="18" width="98">&nbsp;</td>
		<td align="Right" style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="18" width="129"><font face="Arial Narrow" size="2"><%=FormatNum(EquityValue1)%></font></td>
		<td align="Right" style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="18" width="122"><font face="Arial Narrow" size="2"><%=FormatNum(EquityValue1-EquityValue)%></font></td>
	</tr>
	   
	<tr>
		<td colspan="7" height="14" width="941">&nbsp;&nbsp;</td>
	</tr>
		
	<tr>
		<td colspan="7" height="14" width="941">&nbsp;&nbsp;<b><i>Money Market</i></b></td>
	</tr>
		
	<tr>
		<td height="16" width="398" ><font face="Arial Narrow" size="2">&nbsp;&nbsp;&nbsp;&nbsp;Cash&nbsp;In&nbsp;Hand</font></td>
		<td height="16" width="87" ><font face="Arial Narrow" size="2">&nbsp;</font></td>
		<td height="16" width="91"><font face="Arial Narrow" size="2">&nbsp;</font></td>
		<td height="16" width="83"><font face="Arial Narrow" size="2">&nbsp;</font></td>
		<td height="16" width="98"></td>
		<td align="Right" height="16" width="129"><font face="Arial Narrow" size="2"><%=FormatNum(CashOnHand)%></font></td>
		<td height="16" width="122"><font face="Arial Narrow" size="2">&nbsp;</font></td>      
	</tr>
	    
	<tr>
		<td colspan="7" height="14" width="941">&nbsp;</td>
	</tr>
		
	<tr>
		<td style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="18" width="398"><font face="Arial Narrow" size="2">&nbsp;&nbsp;<b><i>Total Money Market Value</i></b></font></td>
		<td style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="18" width="87"><font face="Arial Narrow" size="2">&nbsp;</font></td>
		<td style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="18" width="91"><font face="Arial Narrow" size="2">&nbsp;</font></td>
		<td style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="18" width="83"><font face="Arial Narrow" size="2">&nbsp;</font></td>
		<td style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="18" width="98">&nbsp;</td>
		<td align="Right" style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="18" width="129"><font face="Arial Narrow" size="2"><%=FormatNum(Total)%></font></td>
		<td style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="18" width="122"><font face="Arial Narrow" size="2">&nbsp;</font></td>      
	</tr>

	<tr>
		<td colspan="7" height="14" width="941">&nbsp;&nbsp;</td>
	</tr>

	<tr>
		<td style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="18" width="398"><font face="Arial Narrow" size="2">&nbsp;&nbsp;<b><i>Total Portifolio Value</i></b></font></td>
		<td style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="18" width="87"><font face="Arial Narrow" size="2">&nbsp;</font></td>
		<td style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="18" width="91"><font face="Arial Narrow" size="2">&nbsp;</font></td>
		<td style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="18" width="83"><font face="Arial Narrow" size="2">&nbsp;</font></td>
		<td style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="18" width="98">&nbsp;</td>
		<td align="Right" style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="18" width="129"><font face="Arial Narrow" size="2"><%=FormatNum(Total)%></font></td>
		<td style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="18" width="122"><font face="Arial Narrow" size="2">&nbsp;</font></td>    
	</tr>	
</table>	

</body>

</html>
<%
Function GetMarketPrice(Security_DPA_,MktDate)
	If Len(Security_DPA_)>0 And Len(MktDate)>0 Then
		SQL = "SELECT datastream_Market.MktClose AS Price FROM Security INNER JOIN" & _
			" datastream_Securities ON Security.Security_DPA_ = datastream_Securities.SecKnow_DPA INNER JOIN" & _
			" datastream_Market ON Security.SecurityCode = datastream_Market.MktCode" & _
			" WHERE (Security.Security_DPA_ = "& Security_DPA_ &") AND (datastream_Market.MktDate = '"& MktDate &"')"
		Set RST = Conn.Execute(SQL)
		
		If Not (RST.EOF And RST.BOF) Then
			GetMarketPrice = RST("Price")
		Else
			SQL = "SELECT datastream_Market.MktClose AS Price" & _
				" FROM Security INNER JOIN" & _
				" datastream_Securities ON Security.Security_DPA_ = datastream_Securities.SecKnow_DPA INNER JOIN" & _
				" datastream_Market ON Security.SecurityCode = datastream_Market.MktCode" & _
				" WHERE (Security.Security_DPA_ = "& Security_DPA_ &") AND (datastream_Market.MktDate < '"& MktDate &"')" & _
				" ORDER BY datastream_Market.MktDate DESC"
			Set RST = Conn.Execute(SQL)
			
			If Not (RST.EOF And RST.BOF) Then
				GetMarketPrice = RST("Price")
			Else
				GetMarketPrice = 0
			End If	
		End If		
	Else
		GetMarketPrice = 0
	End If
End Function
%>