<html>
<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Bonds Evaluation</title>
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
	 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
	 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
	<!--CALENDAR -->
	<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
	<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
	<SCRIPT language=VBScript src="../scripts/reports.vbs"></SCRIPT>
	<SCRIPT language=Javascript src="../scripts/reports.js"></SCRIPT>

	<style media="print">
	
		@page {
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

<!--#include file="../libroutines.asp"-->

<%

genReport = Request.Form("genReport")

If genReport <> "1"  Then%>
	<Script Language="JavaScript">
		report_SetBodyClass();
		
		function validateForm(frm){			
			if (frm.client.value==''){
				alert("Select a Client");
				frm.client.focus();
				return
			}
			if(frm.txtDate.value==''){
				alert("Specify the Valuation Date");
				frm.txtDate.focus();
				return;
			
			}
			
			frm.target = '_self';			
			frm.submit();
		}
		
		function UpdateClientCode(){
		 var item = document.frmMain
		 item.code.value = item.client[item.client.selectedIndex].value;
		 
		}
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "txtDate","cmdDate","<%= FormatDate(Date) %>",1);
	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="BondsEvaluation.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">	
		<table>
		<tr>
	<td>Client Code: </td>
	<td><input type="text" name="code" id="code"  size="10" onBlur="txtval = this.value; selectItem(client);"></td>
	</tr>
			<tr>
				<td>Client: </td>
				<td>
				<select size="1" name="client" id="client" onchange="UpdateClientCode();"
    onfocus="txtval = '';inputIsItemCode = 1;"
    onblur="txtval = '';inputIsItemCode = 1;" readonly>
    <option value="" SearchCode = '' SearchText = '' ></option>
    <%
		Set conn = GetActiveConnection("KBroker")
        sqlStr = "SELECT * FROM [Clientlist] Order By ClientName"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF
						
				 %><option SearchCode = '<%=rs.Fields("Client_DPA_")%>' SearchText = '<%=rs.Fields("ClientName")%>' value = '<%=rs.Fields("Client_DPA_")%>'><%=mid(rs.Fields("ClientName"),1,30)%></option><%  
                rs.MoveNext
                Loop
        End If
        rs.close
%>
    </select>
				</td>
			</tr>
	<tr>
	<td>Date of Valuation: </td>
	<td><SCRIPT language="JavaScript">cal.writeControl();</SCRIPT></td>
	</tr>		
			<tr>
				<td colspan=2><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... ">&nbsp;&nbsp;</td>
			</tr>
		</table>
		
	</form>
	
	<%
	Set rs = Nothing
	Set Conn = Nothing
	Response.End
End If

%>


<% 
DrawPageFunctions True, True, True

Client = Request.Form("client")
SettlementDate = Request.Form("txtDate")

Set conn = GetActiveConnection("KBroker")
	sqlStr = "SELECT * FROM BondsEvaluationList WHERE Client_DPA_ = " & Client & _
	         " AND Included = 1 AND #" & SettlementDate & "# >= IssueDate AND #" & SettlementDate & _
	         "# <= MaturityDate Order by BondIssue ASC"
			        
	Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
    If rs.EOF Or rs.BOF Then%>
	   <script language = 'vbscript'>
	    window.self.ShowMessage "There are no records for the selected Client."             		
	   </script>
		<% response.end
	else
	 rs.movefirst
	End If

%>

<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
     
     <tr>
		<td colspan="13" align="center" class="Title"><%=ucase(rs.Fields("ClientName"))%></td>
	</tr>
	<tr>
		<td colspan="13" align="center" class="Title"><u>TREASURY BOND VALUATION REPORT</td>
	</tr>
	<tr>
		<td colspan="13" align="center" class="Title">AS AT <%=ucase(formatDate(SettlementDate))%></td>
	</tr>	
    <br>
</table>				
<p>
    <table border="1" width="100%" cellPadding="0" cellSpacing="0" bordercolor="#000000">
    
    <tr>		
	  <td nowrap class="normal">&nbsp;<b>Issue&nbsp;</b></td>
      <td nowrap class="normal">&nbsp;<b>Date&nbsp;</b></td>
      <td nowrap class="normal">&nbsp;<b>Date&nbsp;</b></td>
      <td nowrap class="normal">&nbsp;<b>Number of&nbsp;&nbsp;</b></td>
      <td nowrap class="normal">&nbsp;<b>Coupons&nbsp;</b></td>
      <td nowrap class="normal">&nbsp;<b>Buying Price&nbsp;</b></td>
      <td nowrap class="normal">&nbsp;<b>Face value&nbsp;</b></td>
      <td nowrap class="normal">&nbsp;<b>Cost&nbsp;</b></td>
      <td nowrap class="normal">&nbsp;<b>Offer Price&nbsp;</b></td>
      <td nowrap class="normal">&nbsp;<b>Current&nbsp;</b></td>
      <td nowrap class="normal">&nbsp;<b>Offer amount&nbsp;</b></td>
      <td nowrap class="normal">&nbsp;<b>Capital&nbsp;</b></td>
      <td nowrap class="normal">&nbsp;<b>Selling&nbsp;</b></td>
      
      </tr>
      <tr>		
	  <td nowrap class="normal">&nbsp;</td>
      <td nowrap class="normal">&nbsp;<b>of Purchase&nbsp;&nbsp;</b></td>
      <td nowrap class="normal">&nbsp;<b>of valuation&nbsp;</b></td>
      <td nowrap class="normal">&nbsp;<b>days held&nbsp;</b></td>
      <td nowrap class="normal">&nbsp;<b>Paid&nbsp;</b></td>
      <td nowrap class="normal">&nbsp;</td>
      <td nowrap class="normal">&nbsp;</td>
      <td nowrap class="normal">&nbsp;</td>
      <td nowrap class="normal">&nbsp;</td>
      <td nowrap class="normal">&nbsp;<b>Yields&nbsp;</b></td>
      <td nowrap class="normal">&nbsp;</td>
      <td nowrap class="normal">&nbsp;<b>Gain (Ksh)&nbsp;</b></td>
      <td nowrap class="normal">&nbsp;<b>Yield&nbsp;</b></td>
      
      </TR>
    <%
    Dim TotalFValue 
    Dim TotalCost
    Dim TotalOAmount
    Dim TotalCGain
    Dim lastCouponDate
    Dim CouponsPaid
    Dim Cost
    Dim OfferPrice
    Dim TSellingYield
    Dim TBasis
    Dim TofferPrice
    Dim TDaysHeld
    Dim TBuyingPrice
    
    TotalFValue = 0
    TotalCost = 0
    TotalOAmount = 0
    TotalCGain = 0
    TSellingYield = 0
    TBasis = 0
    TofferPrice = 0
    TDaysHeld = 0
    TBuyingPrice = 0
    
    Function NumFormat(No,decimals)
    Dim tempNo
     if isnull(decimals) then decimals = 0 'Default value
     if No = "" or not isnumeric(No)then
      NumFormat = No
      exit function 
     end if
    
     if Cdbl(No) < 0 then
      tempNo = formatnumber(abs(No),decimals)
      NumFormat = Cdbl("-" & tempNo)
     else
      NumFormat = formatnumber(No,decimals)
     end if
    End Function
    
    
	Function Calculator(fRate, DaysRemaining, CouponAmount, NoCouponPaymentsLeft)
	   
		dim LeftCoupons
		'Initialise arrays
		Dim I(300)
		Dim K(300)
		Dim O(300)
		Dim Q(300)
		Dim H(300)
	    
	   if isnumeric(NoCouponPaymentsLeft) then
	   
		for a = 1 to  NoCouponPaymentsLeft
		    H(a-1) =((1/((1+(fRate/2)))^((a-1)+(DaysRemaining/DaysInCouponPeriod))))
		    I(a-1) = CouponAmount * H(a-1)
		    if a = 1 then K(a-1) = I(a-1) else  K(a-1) = K(a-2) + I(a-1)
		    O(a-1) = FaceValue * H(a-1)
		    Q(a-1) = O(a-1) + K(a-1)
		next 
		
		if (NoCouponPaymentsLeft-1) < 0 then
			'LeftCoupons=0
			Calculator = 0
		else
			LeftCoupons=NoCouponPaymentsLeft-1
			Calculator = Q(LeftCoupons)
		end if
	    
	   end if
	End function

    function BondDirtyPrice(Rate,DaysRemaining,CouponAmount,CouponsLeft,FaceValue)
       BondDirtyPrice = (Calculator(Rate,DaysRemaining,CouponAmount,CouponsLeft)*(100/FaceValue))
    end function

     Do until rs.EOF
     
      IssueDate = formatDate(rs.Fields("IssueDate"))
      MDate = Cdate(rs.Fields("MaturityDate"))
      DaysInCouponPeriod = cdbl(iif(isnull(rs.Fields("DaysInCoupon")),0,rs.Fields("DaysInCoupon")))
      Basis =  cdbl(iif(isnull(rs.Fields("Basis")),0,rs.Fields("Basis")))
      'The Basis And Days In Coupon Date must have been Specified.
      
      if DaysInCouponPeriod = 0 or Basis= 0 then
       %>
                <script language = 'vbscript'>
                ShowMessage "Error! Days In Coupon or Basis not Specified. Please Specify to Continue"
                window.self.close
                </script>
                <% response.end
      end if
      
      FaceValue  = formatnumber(cdbl(rs.Fields("FaceValue")),2)
      CouponRate = cdbl(rs.Fields("CouponRate"))
      ForwardRate = formatnumber(cdbl(iif(isnull(rs.Fields("ForwardRate")),0,rs.Fields("ForwardRate"))),2)
      BuyingPrice =  formatnumber(rs.Fields("Bprice"),4) 
      
      lastCouponDate = Dateadd("d",(int(DateDiff("d",IssueDate,SettlementDate)/DaysInCouponPeriod)*DaysInCouponPeriod),IssueDate)
     
      CouponsPaid = int(Datediff("d",IssueDate,lastCouponDate)/DaysInCouponPeriod)
      Cost = formatnumber(( BuyingPrice * FaceValue)/100,2)
      
      MonthsInCoupon = ((12 * DaysInCouponPeriod)/Basis)
      
      PeriodicIRate = (((CouponRate/100) * MonthsInCoupon)/12) * 100
      CAmount = (PeriodicIRate * FaceValue)/100
      
      TotalCoupons = (12 * datediff("d",IssueDate,MDate)) / (Basis * MonthsInCoupon)
      CouponsLeft = TotalCoupons - CouponsPaid
      
      NextCouponDate = dateadd("d",DaysInCouponPeriod,lastCouponDate)
      DaysRemaining =  datediff("d",SettlementDate,NextCouponDate)
      OfferPrice = formatnumber(BondDirtyPrice(ForwardRate/100,DaysRemaining,CAmount,CouponsLeft,FaceValue),4)
      OfferAmount = formatnumber((OfferPrice * FaceValue)/100,2)
      Div = (100*12/MonthsInCoupon)
    
      CGain =  formatnumber((OfferAmount - Cost) + (CouponsPaid*(CouponRate/Div)*FaceValue),2)
      DaysHeld = datediff("d",rs.Fields("IssueDate"),SettlementDate)
      
     '' SellingYield = formatnumber(((((OfferPrice/100)*FaceValue)-((BuyingPrice/100)*FaceValue))*((BuyingPrice/100)/FaceValue))*(Basis/DaysHeld)*100,2)
     
      SellingYield = NumFormat(((CGain/Cost) * (Basis/DaysHeld)*100),2)
      
      %>
      <tr>		
	  <td nowrap  class="normal">&nbsp;<%=rs.Fields("BondIssue")%>&nbsp;</td>
	  <td nowrap class="normal">&nbsp;<%=formatDate(rs.Fields("IssueDate"))%>&nbsp;</td>
	  <td nowrap class="normal">&nbsp;<%=formatDate(SettlementDate)%>&nbsp;</td>
	  <td nowrap class="normal">&nbsp;<%=DaysHeld%>&nbsp;</td>
	  <td nowrap class="normal">&nbsp;<%=CouponsPaid%>&nbsp;</td>
	  <td nowrap class="normal">&nbsp;<%=BuyingPrice%>%&nbsp;</td>
	  <td nowrap class="normal">&nbsp;<%=FaceValue%>&nbsp;</td>
	  <td nowrap class="normal">&nbsp;<%=Cost%>&nbsp;</td>
	  <td nowrap class="normal">&nbsp;<%=OfferPrice%>%&nbsp;</td>
	  <td nowrap class="normal">&nbsp;<%=ForwardRate%>&nbsp;</td>
	  <td nowrap class="normal">&nbsp;<%=OfferAmount%>&nbsp;</td>
	  <td nowrap class="normal">&nbsp;<%=CGain%>&nbsp;</td>
	  <td nowrap class="normal">&nbsp;<%=SellingYield%>%&nbsp;</td>
      </tr>
      <tr>		
	  <td nowrap>&nbsp;</td>
	  <td nowrap>&nbsp;</td>
	  <td nowrap>&nbsp;</td>
	  <td nowrap>&nbsp;</td>
	  <td nowrap>&nbsp;</td>
	  <td nowrap>&nbsp;</td>
	  <td nowrap>&nbsp;</td>
	  <td nowrap>&nbsp;</td>
	  <td nowrap>&nbsp;</td>
	  <td nowrap>&nbsp;</td>
	  <td nowrap>&nbsp;</td>
	  <td nowrap>&nbsp;</td>
	  <td nowrap>&nbsp;</td>
      </tr>
      <%
      
      TotalFValue = TotalFValue + cdbl(rs.Fields("FaceValue"))
      TotalCost = TotalCost + cdbl(Cost) 
      TotalOAmount = TotalOAmount + cdbl(OfferAmount)
      TotalCGain = TotalCGain + cdbl(CGain)
      TBasis = TBasis + cdbl(Basis)
      TofferPrice = TofferPrice + cdbl(OfferPrice)
      TDaysHeld = TDaysHeld + cdbl(DaysHeld)
      TBuyingPrice = TBuyingPrice + cdbl(BuyingPrice)
      
     rs.movenext
     loop
    %>  
      <tr>		
	  <td nowrap>&nbsp;</td>
	  <td nowrap>&nbsp;</td>
	  <td nowrap>&nbsp;</td>
	  <td nowrap>&nbsp;</td>
	  <td nowrap>&nbsp;</td>
	  <td nowrap>&nbsp;</td>
	  <td nowrap>&nbsp;</td>
	  <td nowrap>&nbsp;</td>
	  <td nowrap>&nbsp;</td>
	  <td nowrap>&nbsp;</td>
	  <td nowrap>&nbsp;</td>
	  <td nowrap>&nbsp;</td>
	  <td nowrap>&nbsp;</td>
      </tr>
      <tr>		
	  <td nowrap>&nbsp;</td>
	  <td nowrap>&nbsp;</td>
	  <td nowrap>&nbsp;</td>
	  <td nowrap>&nbsp;</td>
	  <td nowrap>&nbsp;</td>
	  <td nowrap>&nbsp;</td>
	  <td nowrap  class="normal">&nbsp;<b><u><%=formatnumber(TotalFValue,2)%></u></b>&nbsp;</td>
	  <td nowrap class="normal">&nbsp;<b><u><%=formatnumber(TotalCost,2)%></u></b>&nbsp;</td>
	  <td nowrap>&nbsp;</td>
	   <td nowrap>&nbsp;</td>
	  <td nowrap class="normal">&nbsp;<b><u><%=formatnumber(TotalOAmount,2)%></u></b>&nbsp;</td>
	  <td nowrap class="normal">&nbsp;<b><u><%=formatnumber(TotalCGain,2)%></u></b>&nbsp;</td>
	  <%
	''  TSellingYield = formatnumber(((((TofferPrice/100)*TotalFValue)-((TBuyingPrice/100)*TotalFValue))*((TBuyingPrice/100)/TotalFValue))*(TBasis/TDaysHeld)*100,2)
	  TSellingYield =  (TotalCGain/TotalCost) * (TBasis/TDaysHeld)*100
	  %>
	  <td nowrap class="normal">&nbsp;<b><u><%=NumFormat(TSellingYield,2)%>%</u></b>&nbsp;</td>
	 
      
      </tr>
  </table>
  <p>
  <table border="0" cellspacing="0" cellpadding="2" width="100%">
     
     <tr>
		<td colspan="13" align="left" class="FootNote">This portfolio is based on the findings of our bonds department and are subject to change without notice</td>
	</tr>
	<tr>
		<td colspan="13" align="left" class="FootNote">It is for private use to the person to whom it is provided without any liability whatsoever on the part of Dyer and Blair Limited</td>
	</tr>
		
    
</table>
</body>

</html>