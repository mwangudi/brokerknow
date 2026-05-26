<html>

<head>

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
				margin-left: 0cm;
				margin-right: 0cm;
				margin-top: 0cm;    
				margin-bottom: 0cm;
				writing-mode: tb-rl;
				br.newpage{
					page-break-before:always;
				}		
			}		 
	</style>
</head>

<body Class="Reports">

<!--#include file="../libroutinesVB.asp"-->
<form name=frm1 id=frm1>
	<%
	selectedClient = Request.QueryString("a")
	selectedFromDate = Request.QueryString("c")
	
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
	
	If rs.EOF Or rs.BOF Then
		Set Rs = Nothing
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
	
<table border="0" width="100%" align="center"> 
<tr>
  <td style="border: 2px solid #000000" align="center" height="70" width="970">
  <table border =0 width="734" height="100%">
	<tr><td align="center" height="400" width="728"><font size="7"><b><% response.write ClientName
	%></b></font></td></tr>
	<tr><td valign="bottom" align="center" width="728"><Img Src="xx-xx-xx">		</td></tr>
</table>
</td></tr>
</table>

  <BR class="newpage">

<table border=0 align=center>
	<tr><td height=50>&nbsp;</td></tr>
</table>    

<br>
<table border="0" cellspacing=0 cellpadding=2 width="100%" style="border-collapse: collapse" bordercolor="#111111"  align="center"> 
	<tr>
      <td style="border: 2px solid #000000" align="center" height="100" width="907"><font size="5"><b>
      Portfolio Valuation Statement</b></font>
	<br><font size="2"><b><% tarehe = split(formatdatefull(selectedFromDate),",")
		response.write replace(tarehe(1),"-"," ")
	%></b></font></td></tr>
</table>

<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" width="100%"  align="center"> 
	<tr>
		<td align="right" height="8" width="801">
            &nbsp;	
		</td>		
	</tr>  
</table>


	
	<table border="0" cellspacing=0 cellpadding=2 width="100%" style="border-collapse: collapse" bordercolor="#111111" height="243"  align="center">	
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
<td colspan="7" height="14" width="100%">&nbsp;&nbsp;</td>
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
		<td colspan="7" height="14" width="100%">&nbsp;&nbsp;</td>
	</tr>
    <tr>
		<td colspan="7" height="14" width="100%">&nbsp;&nbsp;<b><i>Money Market</i></b></td>
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
		<td colspan="7" height="14" width="100%">&nbsp;</td>
	</tr>
   	<tr>
       <td style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="18" width="398"><font face="Arial Narrow" size="2">&nbsp;&nbsp;<b><i>Total Money Market Value</i></b></font></td>
      <td style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="18" width="87"><font face="Arial Narrow" size="2">&nbsp;</font></td>
      <td style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="18" width="91"><font face="Arial Narrow" size="2">&nbsp;</font></td>
      <td style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="18" width="83"><font face="Arial Narrow" size="2">&nbsp;</font></td>
      <td style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="18" width="98">
     &nbsp;</td>
      <td align="Right" style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="18" width="129"><font face="Arial Narrow" size="2"><%=FormatNum(Total)%></font></td>
      <td style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="18" width="122"><font face="Arial Narrow" size="2">&nbsp;</font></td>      
    </tr>	
	<tr>
		<td colspan="7" height="14" width="100%">&nbsp;&nbsp;</td>
	</tr>
   	<tr>
       <td style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="18" width="398"><font face="Arial Narrow" size="2">&nbsp;&nbsp;<b><i>Total Portifolio Value</i></b></font></td>
      <td style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="18" width="87"><font face="Arial Narrow" size="2">&nbsp;</font></td>
      <td style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="18" width="91"><font face="Arial Narrow" size="2">&nbsp;</font></td>
      <td style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="18" width="83"><font face="Arial Narrow" size="2">&nbsp;</font></td>
      <td style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="18" width="98">
      &nbsp;</td>
      <td align="Right" style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="18" width="129"><font face="Arial Narrow" size="2"><%=FormatNum(Total)%></font></td>
      <td style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="18" width="122"><font face="Arial Narrow" size="2">&nbsp;</font></td>    
    </tr>	
</table>	
  
<BR class="newpage">

 <table border=0 align=center width="100%">
	<tr><td height=50>&nbsp;</td></tr>
</table>
 
<%  
increment = 100
 
    '=======================================================================================
    'Generate client statement for the last one month
	'=======================================================================================

	Set conn = GetActiveConnection("KBroker")
	Set Rs = CreateObject("ADODB.Recordset")
	Rs.CursorLocation = adUseClient		
						        
  
    selectedFromDate = dateadd("m",-1,selectedFromDate)
    
	selectedFromDate = 1 & "-" & DatePart("m",selectedFromDate) & "-" & DatePart("yyyy",selectedFromDate)
	
	selectedFromDate  = formatdate(selectedFromDate)
  
	sqlStr = "SELECT * FROM StatementList WHERE Client_DPA_ = " & selectedClient & " AND TransDate >= '" & FormatDate(selectedFromDate) & "'"
	
	Rs.CursorLocation = adUseClient	
	Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
	
	Set rsClient = Conn.Execute ("SELECT * FROM Client WHERE Client_DPA_ = " & selectedClient)
	
	If Not (rsClient.EOF Or rsClient.BOF) Then
		accountDesc = rsClient.Fields("ClientName").Value & "&nbsp;&nbsp;&nbsp;" & "[" & rsClient.Fields("Client_DPA_").Value & "]"
		accountAddress = rsClient.Fields("ClientAddr").Value
	End If
	
	Set rsClient = Nothing
	
	Rs.PageSize = 25

	Rs.CacheSize = Rs.PageSize
		intPageCount = Rs.PageCount 
		intRecordCount = Rs.RecordCount 
	
		first=0
		
	' Now you must double check to make sure that you are not before the start
	' or beyond end of the recordset.  If you are beyond the end, set 
	' the current page equal to the last page of the recordset.  If you are
	' before the start, set the current page equal to the start of the recordset.	

	Rs.CacheSize = Rs.PageSize

	if Not (rs.eof or rs.bof) then
	intPageCount = Rs.PageCount 
	intRecordCount = Rs.RecordCount 
	else
	intPageCount=1
	intRecordCount=0
	end if	
		
        PageNumber1=PageNumber1 + 1
        
        intPage=0
        
         intPageCount=Cint(intPageCount)
        
      t=0      
	
	dim Valuewrite
	dim runningBal
	totalDebits = 0
    totalCredits = 0
    runningBal=0
    
	if Not (rs.eof or rs.bof) then
		if (rs("IsOpeningBalance")<>1) then
		sqlStr = "SELECT SUM(ISNULL(dbo.StatementList.Credit - dbo.StatementList.Debit, 0)) + dbo.Client.ClientOpeningBal AS CurrentBal, dbo.StatementList.Client_DPA_" & _
					 " FROM  dbo.StatementList INNER JOIN dbo.Client ON dbo.StatementList.Client_DPA_ = dbo.Client.Client_DPA_" & _
					 " WHERE (dbo.Client.Deleted = 0) AND (dbo.StatementList.TransDate < '" & FormatDate(selectedFromDate) & "') GROUP BY dbo.StatementList.Client_DPA_, dbo.Client.ClientOpeningBal having dbo.StatementList.Client_DPA_=" & selectedClient 

		Set rsClient = Conn.Execute(sqlStr)
	
			If Not (rsClient.EOF Or rsClient.BOF) Then
				OpeningBalance=rsClient("CurrentBal")
			end if
		end if
	else
		sqlStr = "SELECT SUM(ISNULL(dbo.StatementList.Credit - dbo.StatementList.Debit, 0)) + dbo.Client.ClientOpeningBal AS CurrentBal, dbo.StatementList.Client_DPA_" & _
					 " FROM  dbo.StatementList INNER JOIN dbo.Client ON dbo.StatementList.Client_DPA_ = dbo.Client.Client_DPA_" & _
					 " WHERE (dbo.Client.Deleted = 0) AND (dbo.StatementList.TransDate < '" & FormatDate(selectedFromDate) & "') GROUP BY dbo.StatementList.Client_DPA_, dbo.Client.ClientOpeningBal having dbo.StatementList.Client_DPA_=" & selectedClient 
	

	
	Set rsClient = Conn.Execute(sqlStr)
	
		If Not (rsClient.EOF Or rsClient.BOF) Then
			OpeningBalance=rsClient("CurrentBal")
		end if
	end if	
	
	kwanza=0

	Do while Cint(intPage) < intPageCount	
	intPage=intPage + 1
	'Response.write(intpage)

	if(Cint(first)=1) then
	%>
             <BR class="newpage">
    <%
	end if

	first=1

	If CInt(intPage) > CInt(intPageCount) Then intPage = intPageCount
	If CInt(intPage) <= 0 Then intPage = 1
	
	 'Make sure that the recordset is not empty.  If it is not, then set the 
	 'AbsolutePage property and populate the intStart and the intFinish variables.
	
	'if Not(Rs.eof and Rs.bof) Then

	If intRecordCount > 0 Then 'and Not(Rs.eof and Rs.bof)
		
		Rs.AbsolutePage = intPage
		intStart = Rs.AbsolutePosition
		'Response.write(intStart)
		
		If CInt(intPage) = CInt(intPageCount) Then
			intFinish = intRecordCount
		Else
			intFinish = intStart + (Rs.PageSize - 1)
		End if
	End If	  
	
	if Not (rs.eof or rs.bof) then
	isOpeningBalance = CBool(Rs.Fields("IsOpeningBalance").Value)
	else
	isOpeningBalance=false
	end if
	
	If Not IsOpeningBalance Then		
		OpeningBalance = OpeningBalance		
			Else
		OpeningBalance =  Rs.Fields("Balance").Value		
		
	End If	

%>
<table border=0 align=center width="100%">
	<tr><td height=<%=increment+20%>>&nbsp;</td></tr>
</table>

<table border="0" cellspacing=0 cellpadding=2 width="100%" style="border-collapse: collapse" bordercolor="#111111"  align="center"> 
	<tr>
      <td style="border: 2px solid #000000" align="center" height="100" width="100%">
      <font size="5"><b>Transactions Summary</b></font><br><font size="2"><b><% 'tarehe = split(formatdatefull(selectedFromDate),",")
		response.write replace(tarehe(1),"-"," ")
	%></b><br><br>
	Page&nbsp;<%=intPage%>&nbsp;of&nbsp;<%=intPageCount%>
	</font></td></tr>
</table>


  <table border="0" cellspacing="0" cellpadding="0" style="font-family: Arial Narrow; LEFT-MARGIN:100PX; border-collapse:collapse"   align="center" width="100%" bordercolor="#111111">
    <tr>
      <td colspan="6" align="right" height="12%" width="100%">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
    </tr>  
    <tr>      
      <td width="80" style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="12%">
      <p align="left"><b><font face="Arial Narrow" size="2">Date:</font></b></td>
      <td style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" width="120" height="12%">
      <p align="left"><b><font face="Arial Narrow" size="2">Ref:</font></b></td>
      <td style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" width="451" height="12%">
      <p align="left"><b><font face="Arial Narrow" size="2">Particulars:</font></b></td>
      <td align="right" style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" width="120" height="12%">
      <p><b><font face="Arial Narrow" size="2">Debit:</font></b></td>
      <td align="right" style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" width="120" height="12%">
      <p><b><font face="Arial Narrow" size="2">Credit:</font></b></td>
      <td align="right" style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" width="120" height="12%">
      <p><b><font face="Arial Narrow" size="2">Balance:</font></b></td>
    </tr> 
      
    <%    
    'Do Until Rs.EOF       
 if(Cint(kwanza=0)) then   
	If Not IsOpeningBalance Then

	if(abs(FormatNum(OpeningBalance))=0) then
	OpeningBalance=0
	end if
	%>
	<tr>      
      <td width="80" height="15"><b><font size="1"><%= Day(selectedFromDate) & " " & MonthName(Month(selectedFromDate), True) & " " & Right(Year(selectedFromDate),2) %></font></b></td>
      <td height="15" width="80"><b><font size="1">&nbsp;</font></b></td>
      <td height="15" width="451"><b><font size="1">Opening Balance</font></b></td>
      <td align="right" height="15" width="120"><b><font size="1">&nbsp;</font></b></td>
      <td align="right" height="15" width="120"><b><font size="1">&nbsp;</font></b></td>
      <td align="right" height="15" width="120"><b><font size="1"><%=CreditDebitValue(FormatNum(OpeningBalance))%></font></b></td>
    </tr> 
	<%

	runningBal = runningBal + OpeningBalance
	end if
end if
kwanza=1

if Not (rs.eof or rs.bof) then
    For intRecord = 1 to Rs.PageSize
    
		totalDebits = totalDebits + Rs.Fields("Debit").Value 
		totalCredits = totalCredits + Rs.Fields("Credit").Value
		
		if intRecord = 1 then
			If intPage <> 1 Then
			%>
			<tr>      
		      <td width="80" height="15"><b><font size="1"><%= Day(lastDate) & " " & MonthName(Month(lastDate), True) & " " & Right(Year(lastDate),2) %></font></b></td>
		      <td height="15" width="80"><b><font size="1">&nbsp;</font></b></td>
		      <td height="15" width="451"><b><font size="1">BALANCE B/F</font></b></td>
		      <td align="right" height="15" width="120"><b><font size="1">&nbsp;</font></b></td>
		      <td align="right" height="15" width="120"><b><font size="1">&nbsp;</font></b></td>
		      <td align="right" height="15" width="120"><b><font size="1"><%=FormatNum(BalOnNextPage)%></font></b></td>
		    </tr>  
			<%
			End If
		end if
		%>
		<tr>		
		  <td height="12%" width="80" ><font size="1"><%= Day(rs.Fields("TransDate")) & " " & MonthName(Month(rs.Fields("TransDate")), True) & " " & Right(Year(rs.Fields("TransDate")),2) %></font></td>		  
		  <td height="12%" nowrap width="120" ><font size="1"><%= Rs.Fields("PaymentReceiptNo").Value %></font></td>
		  <%
		  if(Rs("receipttype")=1) then
		  %>
		  <td height="12%" width="451" ><font size="1">RECEIPT:<%= Mid(Ucase(Rs.Fields("Particulars").Value),1,47) %></font></td>
		  <%
		  else
			if(Rs("receipttype")=3) then
			%>
			<td width="451" height="1"><font size="1">PAYMENT:<%= Mid(Ucase(Rs.Fields("Particulars").Value),1,47) %></font></td>
			<%			
				else
					if(rs("IsOpeningBalance")=1) then				
						If Not IsOpeningBalance Then
						else
						%>
						<td height="12%" width="451"><font size="1"><%= Mid(Ucase(Rs.Fields("Particulars").Value),1,55) %></font></td>
						<%
						end if
					else
					%>
					<td height="12%" width="451"><font size="1"><%= Mid(Ucase(Rs.Fields("Particulars").Value),1,55) %></font></td>
					<%					
				end if
			end if
		  end if
		  %>
		  <td align="right" height="12%" width="120"><font size="1"><% If Rs.Fields("Debit").Value <> "0" Then 
									Response.Write FormatNum(Rs.Fields("Debit").Value) 
							   End If			%>
            </font>
		  </td>
		  <td align="right" height="12%" width="120"><font size="1"><% If Rs.Fields("Credit").Value <> "0" Then
									Response.Write FormatNum(Rs.Fields("Credit").Value) 
							   End If			
							%>
            </font>
		  </td>
		  <td align="right" height="12%" width="120">
            <font size="1">
			<% 					

			If Not IsOpeningBalance Then
					Valuewrite=(Rs.Fields("Credit").Value - Rs.Fields("Debit").Value)										
					Response.write(FormatNum(CreditDebitValue(runningBal+ValueWrite)))											
					runningBal = runningBal + (Rs.Fields("Credit").Value - Rs.Fields("Debit").Value)										
				Else											
					if(rs("IsOpeningBalance")=1) then						
						Response.Write FormatNum(CreditDebitValue(Rs.Fields("Balance").Value))						
						runningBal = rs("CreditBal")
					else
					    
						Valuewrite=(Cdbl(Rs.Fields("Credit").Value) - CDbl(Rs.Fields("Debit").Value))
						'Response.write("yeah")
						if(abs(FormatNum(runningBal+Cdbl(ValueWrite)))=0) then
						Response.write(CreditDebitValue(abs(FormatNum(runningBal+Cdbl(ValueWrite)))))
						else
						Response.write(CreditDebitValue(FormatNum(runningBal+Cdbl(ValueWrite))))
						end if
						runningBal = runningBal + Valuewrite
						
						'Response.Write FormatNum(Rs.Fields("Balance").Value)
						
						'Response.write(runningBal)
						  if(abs(FormatNum(runningBal))=0) then
							runningBal=abs(FormatNum(runningBal))
						  end if
						  'Response.write(runningBal)

					end if
				End If 
				lastDate = Rs.Fields("TransDate").Value%>
            </font>
		  </td>
		</tr>
	
	<%	'Rs.MoveNext	
	rs.MoveNext
        
		If Rs.EOF Then Exit for

        Next
  else
  lastDate=Date
  end if      
	if Cint(intPage) <> intPageCount then 	
  	%>
	<tr>
		<td colspan="6" align="right" height="15" width="971">&nbsp;</td>
	</tr>
	<%
	if(abs(FormatNum(OpeningBalance))=0) then
	OpeningBalance=0
	end if
	
	%>
    
	<%
	if(abs(FormatNum(runningBal))=0) then
		runningBal=abs(FormatNum(runningBal))
	end if
	%>

    <tr>
      <td style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="12%" width="80">&nbsp;</td>
      <td style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="12%" width="120">&nbsp;</td>
      <td colspan="3" style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="12%" width="691"><b><font size="1">Running&nbsp;Balance:</font></b></td>
      <td NOWRAP align="right" style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="12%" width="120"><b><font size="1"><%= CreditDebitValue(FormatNum(runningBal)) %></font></b></td>
    </tr>

   	<%
   	BalOnNextPage = CreditDebitValue(runningbal)
   	end if
   	%>	
	</table>
	<%	
	loop
	%>
	<table border="0" cellspacing="0" cellpadding="0"  align="center" width="100%" style="border-collapse: collapse" bordercolor="#111111">

<tr>		  
	<td colspan="6" height="12%">&nbsp;</td>
</tr>


<tr>		  
	<td width="80" style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="12%"><font size="1"><%= Day(lastDate) & " " & MonthName(Month(lastDate), True) & " " & Right(Year(lastDate),2) %></font></td>
	<td width="120" style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="12%">&nbsp;</td>
	<td width="451" align="left" style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="12%">&nbsp;<font size="1">CLOSING BALANCE</font></td>
	<td width="120" style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="12%">&nbsp;</td>
	<td width="120" style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="12%">&nbsp;</td>
	<td width="120" align="right" style="BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid;" height="12%"><font size="1"><%= CreditDebitValue(FormatNum(CDbl(runningBal))) %></font></td>
</tr>

    <%
	Dim BeforeBal
	Dim LaterBal
		
	BeforeBal =	ClientBalance(30)
	
	if(abs(FormatNum(BeforeBal))=0) then
	BeforeBal=0
	end if

	LaterBal = ClientBalance(60)

	if(abs(FormatNum(LaterBal))=0) then
	LaterBal=0
	end if

	%>
  
  </table>
   
<%
'Set Rs = Nothing

function ClientBalance(Days)
sqlStr="SELECT SUM(ISNULL(dbo.StatementList.Credit - dbo.StatementList.Debit, 0)) + dbo.Client.ClientOpeningBal AS CurrentBal," & _ 
       " dbo.StatementList.Client_DPA_ FROM  dbo.StatementList INNER JOIN dbo.Client ON dbo.StatementList.Client_DPA_ = dbo.Client.Client_DPA_" & _
	   " WHERE (dbo.StatementList.TransDate < GETDATE() - " & Days & ") and (StatementList.Client_DPA_=" & selectedClient & ")" & _
	   " GROUP BY dbo.StatementList.Client_DPA_, dbo.Client.ClientOpeningBal"

Set Rs = Conn.Execute(sqlStr)
 if not(Rs.eof and Rs.Bof) then
 ClientBalance=Rs("CurrentBal")
 else
 ClientBalance=0
 end if
end function 

Set Rs = Nothing
Set Conn = Nothing

function convertSign(balance)
	dim sign
	sign = sgn(balance) '1 indicates positive, -1 indicates negative
	if sign = 1 then
		convertSign = "-" & balance	'// make this positive number a negative number
	elseif sign = -1 then
		convertSign = Abs(balance)	'// make this negative number a positive number. Abs removes the - element.
	end if
end function

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

<input type="hidden" value="" id="hidData" name="hidData">
<input type="hidden" value="<%=selectedClient%>" id="hidClient" name="hidClient">
<input type="hidden" value="HoldingsValuation" id="hidCategory" name="hidCategory">
	
</form>

</body>

</html>
<% Response.End %>
<script language="vbscript">
	theData = frm1.innerHTML 
	
	frm1.hidData.value = theData
	frm1.method = "post"
	frm1.action = "PDFMail.asp"
	frm1.submit 
</script>