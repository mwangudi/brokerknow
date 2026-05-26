<!--#include file="../libroutines.asp"-->
<%

	ID = Request.QueryString("ID")
	action=Request.QueryString("action")
	ID1 = ID

' DB Connection
Dim Conn 
    Set Conn = CreateObject("ADODB.Connection")
    theDBName = "KBroker" 
    Conn.ConnectionString =  "FILE NAME=" & GetUDLPath(theDBName) 
    Conn.Open
    
   Function GetUDLPath(theDBName) 
    Dim tmpStr
    
    tmpStr = StrReverse(Request.ServerVariables("APPL_PHYSICAL_PATH"))
    
    tmpStr = Mid(tmpStr, InStr(1, tmpStr, "\") + 1)
    
    tmpStr = StrReverse(tmpStr)
    
    GetUDLPath = tmpStr & "\UDL\" & Trim(theDBName) & ".UDL"

End Function
'Bond Parameters

		Dim IssueNo
		Dim FaceValue
		Dim IssueDate
		Dim MaturityDate
		Dim SettlementDate
		Dim CouponRate
		Dim ForwardRate
		Dim Client
		Dim BondType
		Dim TradeType
		Dim CouponPayment
		Dim ValidityDate
		Dim SecurityCode
		Dim ProposalDate
		Dim CounterParty
		Dim AccManager
		Dim Commission
		Dim CommissionRate
		Dim CommissionRate1
		Dim Salutation
		SecurityCode=0
		BondType=0
		Dim Consideration
		Dim NetAmt
		Dim txtCounterparty
		Dim AlternatePrice
		Dim Hold
		Dim Compound
		Dim OrderDate
		Dim InterBank
		Dim guid
		Dim SqlStr

	 'Fetch Input Data from DB
	set rst = CreateObject("ADODB.Recordset")
	
	'SQLStr = "Select * From BondsInputs"
	SQLStr = "Select * from EditBondProposals where Proposal_DPA_= "& Clng(ID)	
	set rst = Conn.Execute (SQLStr)

	If rst.EOF Or rst.BOF Then%>
	   <script language = 'javascript'>
	   alert('Process Stopped! There are no Bonds Input Values.')  		
	   </script>
	     <% 
	        ''   response.end
	else
	
		Do Until rst.EOF
			ProposalNo =trim(rst("Proposal_DPA_"))
		    IssueNo = trim(rst("BondIssue"))
			FaceValue = trim(rst("FaceValue"))
			IssueDate = cdate(trim(rst("BondIDate")))
			MaturityDate = cdate(trim(rst("BondMDate")))
			SettlementDate = cdate(trim(rst("SettlementDate")))
			CouponRate = trim(rst("CouponRate"))
			ForwardRate = trim(rst("ForwardRate"))
			DaysInCouponPeriod = trim(rst("CouponPeriodDays"))
			Basis = trim(rst("Basis"))
			ChangeInYield = trim(rst("ChangeInYield"))
			InputID = trim(rst("Proposal_DPA_"))							        		
			Client = trim(rst("Client_DPA_"))
			ClientName =trim(rst("clientName"))
			BondType = trim(rst("Security_DPA_"))
			TradeType = trim(rst("TradeType"))
			CouponPayment = trim(rst("BondPayment"))
			ValidityDate = trim(rst("Validity"))
			BondDirtyPrice1 = trim(rst("BondDirtyPrice"))
			BondCleanPrice1 = trim(rst("BondCleanPrice"))
			Duration1 = trim(rst("Duration"))
			Convexity1 = trim(rst("Convexity"))
			AccruedInterest1 = trim(rst("AccruedInterest"))
			ProposalDate = trim(rst("ProposalDate"))
   			SecurityCode1  = trim(rst("SecurityCode"))
   			SecurityCode  = trim(rst("Bond_DPA_"))
			CommissionRate = trim(rst("Comm_DPA"))
			CommissionRate1=trim(rst("CommDesc"))
			Commission = trim(rst("Commission"))
			CounterParty = trim(rst("OwnerName"))
			AccManager = trim(rst("OwnerName"))
			Salutation = trim(rst("Salutation"))
			AlternatePrice = trim(rst("AlternatePrice"))
			BondDirtyPrice1 = trim(rst("BondDirtyPrice"))
			BondCleanPrice1 = trim(rst("BondCleanPrice"))
			Duration1 = trim(rst("Duration"))
			Convexity1 = trim(rst("Convexity"))
			AccruedInterest1 = trim(rst("AccruedInterest"))
			Consideration =trim(rst("Consideration"))
			NetAmt =trim(rst("NetAmount"))
		    rst.MoveNext
		Loop
	End If

	'save the records in tborder 
	 buttonAction = Ucase(Request.Form("buttonAction")) 
	 if buttonAction ="SAVE" then

		OrdDate = cdate(request.form("OrderDate"))
		if trim(request.form("Hold"))="" then Hold =0 else Hold=request.form("Hold")
		if trim(request.form("Compound"))="" then Compound =0 else Compound=request.form("Compound")
		if trim(request.form("Interbank"))="" then Interbank =0 else Interbank =request.form("Interbank")
		 set guid = server.createobject("NDUtils.CGUID")
		 guidStr = guid.GenerateGUID
	
	Dim rsMax
	Dim MaxNo
	set rsMax = server.CreateObject("Adodb.recordset")
	set rsMax=Conn.Execute("Select Max(Order_DPA_)+1 as MaxNo from tbOrder")
	if (isnull(rsMax("MaxNo")) or trim(rsMax("MaxNo"))="") then
		MaxNo =1
	else
		MaxNo= rsMax("MaxNo")
	end if
	rsMax.close
	
	conn.beginTrans
	SqlStr = "INSERT INTO tbOrder(Branch_DPA_, Order_DPA_, Client_DPA_, Order_EIT_, OrderSecType_DPA_, OrderType_DPA_, OrderDate, Proposal_DPA_,OrderRef, " &_
     		 " OrderHold, OrderCompounded, InterBank,changedby) values ( 1,"& MaxNo & "," &_
			 " " & Client & ",'" & guidStr & "',1,"& TradeType &",'" & FormatDate(cdate(OrdDate)) & "',"&_
			 ""& ProposalNo &",' ',"& Hold &"," &_
			 "" & Compound & ", " & InterBank & ","& session("UserID") &")" 
		
		' Insert into the details table
		conn.execute(SqlStr)

		set rsMax=Conn.Execute("Select Max(OrdDetail_DPA_)+1 as MaxNo from OrdDetail")
		if (isnull(rsMax("MaxNo")) or trim(rsMax("MaxNo"))="") then
			MaxNo1 =1
		else
			MaxNo1= rsMax("MaxNo")
		end if
		rsMax.close
		
		'
		set rsDPA = server.createObject("Adodb.recordset")
		set rsDPA =Conn.Execute("Select Order_DPA_ from tbOrder where Order_EIT_='"& guidStr &"' ")
		
		
		if (rsDPA.eof or rsDPA.bof) then
			OrderDPA =""
			
		else
			rsDPA.movefirst
			OrderDPA= rsDPA("Order_DPA_")
		end if
		'Insert related records in the orderdetail table
		SqlStr = "Insert into OrdDetail(OrdDetail_DPA_,Order_DPA_, OrdDetailPrice,Security_DPA_,OrdDetailQty, BondDescription,Bond_DPA_)" &_
		"values ("& MaxNo1 &","& OrderDPA &","& BondCleanPrice1 &","& BondType &","& FaceValue &",'"& IssueNo &"'," &_
				"" & SecurityCode & " )"
		
		conn.execute(SqlStr)
		conn.Execute ("Update BondProposals set Accepted=1 where Proposal_DPA_="& ProposalNo )
		conn.commitTrans
		
		%>
		<Script Language="JavaScript">
		try{
			window.parent.dialogArguments.opener.location.reload();
			window.self.close()
			}
			catch(e){window.self.close()}
    </Script>
		<%
		response.end
	 set rsMax=nothing
	 end if
%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Bond Orders</title>
 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<!--CALENDAR -->
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
 <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css"> 
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT>
<script language="JavaScript" src="CALENDAR/calendar.js"></script>
<!--END CALENDAR -->
<!--#include file="../pageLoadingTop.asp"-->
<script language='javascript'>
	
		var validNavigate = false;
		function ReleaseRecord()
		{
			if(!validNavigate)
			{
 				event.returnValue = "Please use the cancel button to close the dialog"
 			}
		}
		
		function AllowedNavigation()
		{
			validNavigate = true;
		}
		
		
</script>
</head>

<body leftmargin="20"  Class="Dialog">
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<!--#include file="pageLoadingBody.asp"-->
<%
'Defalt Dates
  
 if isnull(MaturityDate) or trim(MaturityDate)="" then MatDate = "" else MatDate = formatdate(MaturityDate)
 if isnull(IssueDate) or trim(IssueDate)="" then IssDate = "" else IssDate = formatdate(IssueDate )
 if isnull(SettlementDate) or trim(SettlementDate)=""  then SettDate = formatdate(Date()) else SettDate = formatdate(SettlementDate)
 if isnull(ProposalDate) or trim(ProposalDate)=""  then PropDate = FormatDate(Date()) else PropDate = formatdate(ProposalDate)
 if isnull(ValidityDate) or trim(ValidityDate)="" then valDate = "" else valDate = formatdate(ValidityDate)
 if isnull(OrderDate) or trim(OrderDate)="" then OrderDate = formatdate(now()) else OrderDate = formatdate(OrderDate)


%>
<SCRIPT language="JavaScript">
	var cal=new ctlSpiffyCalendarBox("cal", "frmBondOrder", "OrderDate","cmdOrderDate","<%=OrderDate %>",1);
</SCRIPT>
<form name="frmBondOrder" method = 'post' action = 'AddRegisterOrders2.asp?ID=<%=ID1%>' id = "frmBondOrder" >
<p><b>&nbsp;</b></p>

<p>
<input type="Submit" value="SAVE" name="buttonAction">
<input type="button" value="CLOSE" name="buttonAction" onclick="javascript: self.close();">
<input type="hidden" value="<%=InputID%>" name="InputID">
<input type="hidden" value="<%=SecurityCode%>" name="SecurityCode" id="SecurityCode">
<input type="hidden" value="<%=FormatDate(now)%>" name="ValidDate">
<input type="hidden" value="CALCULATE" name="action">
<input type="Hidden" name="HidIssueNo" id="HidIssueNo" size="1" Value="<%=IssueNo%>"></p>

<table border="0" width="90%" bordercolor="#000000" cellspacing="0" cellpadding="0">
 <tr>
    <td width="30%" style="border-left-style: solid; border-top-style: solid">&nbsp;Bond Clean Price</td>
    <td width="70%" align="right" style="border-right-style: solid; border-top-style: solid">
    <% 
		 Response.Write formatnumber(BondCleanPrice1,4)
   
    %></td>
  </tr>
  <tr>
    <td width="30%" style="border-left-style: solid">&nbsp;Accrued Interest</td>
    <td width="70%" align="right" style="border-right-style: solid">
    <% 
		 Response.Write formatnumber(AccruedInterest1,2)   
    %>
	</td>
  </tr>
  <tr>
    <td width="30%" style="border-left-style: solid">&nbsp;Bond Dirty Price</td>
    <td width="70%" align="right" style="border-right-style: solid">
    <% 
       	 Response.Write formatnumber(BondDirtyPrice1,4)
    
     %>&nbsp;</td>
  </tr>
  <tr>
    <td width="30%" style="border-left-style: solid">&nbsp;Duration</td>
    <td width="70%" align="right" style="border-right-style: solid">
    <% 
         Response.Write formatnumber(Duration1,2)
      %></td>
  </tr>
  <tr>
    <td width="30%" style="border-left-style: solid; border-bottom-style: solid">&nbsp;Convexity</td>
    <td width="70%" align="right" style="border-right-style: solid; border-bottom-style: solid">
    <% 
       Response.Write formatnumber(Convexity1,2)
     
    %></td>
  </tr>

</table>
<p><b>Order Parameters</b></p>

<table border="1" width="90%" cellspacing="0" cellpadding="0" bordercolor="#000000">
 <tr>
    <td width="30%" style="border-left-style: solid; border-top-style: solid">&nbsp;Order Date</td>
    <td width="70%" style="border-right-style: solid; border-top-style: solid">
	<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>
    </td>
  </tr>
  <tr>
    <td width="30%" style="border-left-style: solid">&nbsp;Hold</td>
    <td width="70%" style="border-right-style: solid">
	<input type="checkBox" name="Hold" value="1">
	</td>
  </tr>
  <tr>
    <td width="30%" style="border-left-style: solid">&nbsp;Compound</td>
    <td width="70%"  style="border-right-style: solid">
	<input type="checkBox" name="Compound" value="1">
    </td>
  </tr>
  <tr>
    <td width="30%" style="border-left-style: solid">&nbsp;Interbank</td>
    <td width="70%" style="border-right-style: solid">
	<input type="checkBox" name="Interbank" value="1">
    </td>
  </tr>
 </table>



<p><b>Proposal Parameters</b></p>
<table border="1" width="90%" cellspacing="0" cellpadding="0" bordercolor="#000000">
  <tr>
    <td width="30%">&nbsp;Proposal No.</td>
    <td width="70%">
    <input readonly = 'true' style="width:256" class=readonly type ="text" Name ="ProposalNo" ID ="ProposalNo" value="<%=ProposalNo%>" size="20">
     </td>
  </tr>
    <tr>
    <td width="30%">&nbsp;Proposal Date</td>
    <td width="70%"><input readonly class="readonly" style="width:256" type="txt" name="cboOwner" ID ="cboOwner" value="<%=formatdate(PropDate)%>" size="20" >&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;Client</td>
    
    <td>
       <input type = 'hidden' name ='cboClient' id = 'cboClient' size="10" onBlur="txtval = this.value; selectItem(cboClient);" value="<%=client%>">
         <input type = 'text' name ='txtClient' style="width:256"  id = 'txtClient' size="10" readonly class=readonly style="width:256" value="<%=clientName%>">
    </td>
  </tr>
  <tr>
    <td width="30%">&nbsp;Trade Type
    
     </td>
    <td width="70%">
    <input type="hidden"  readonly class=readonly name="TradeType" value="<%=TradeType%>">
    <% if TradeType=2 then
			 s="Sale"
		elseif TradeType=1 then
			  s="Purchase"
		end if
	%>
     <input type="text"  readonly class=readonly style="width:256" name="TradeType1" value="<%=s%>" size="20">
    <div style="display:none;">
    <select name="cboTradeType" disabled>
		<option></option>
		<%
		set rs = Conn.Execute("select * from OrderType") 
		if TradeType="" then TradeType=2
		while not rs.eof 
		
		%>
		
		<option value="<%=rs("OrderType_DPA_")%>" <%if(rs("OrderType_DPA_")=cint(TradeType)) then Response.Write "selected"%>><%=rs("OrderTypeDescription")%></option>
		<%
			rs.movenext
			wend
		%>
	</select>
	</div>
	</td>
  </tr>
 <tr>
    <td width="30%">&nbsp;Bond Type</td>
    <td width="70%">
    <input type="hidden" class=readonly name="BondType" value="<%=BondType%>">
    <input type="text" style="width:256" name="SecurityCode1" id="SecurityCode1" readonly class="readonly" value="<%=SecurityCode1%>" size="20">
   
	</div>
	</td>
  </tr>
  
   <tr>
    <td>&nbsp;Bond Issue</td>
    <td>
    <input type="text"  readonly style="width:256" class=readonly name="IssueNo" value="<%=IssueNo%>" size="20">
   </td>
  </tr>
  <tr>
    <td width="30%">&nbsp;Issue Date</td>
    <td width="70%">
    <input type = "text" name="IssueDate" style="width:256"  value="<%=formatdate(IssueDate)%>" readonly class="readonly" size="20"></td>
  </tr>
   <tr>
    <td width="30%">&nbsp;Settlement Date</td>
    <td width="70%"><input readonly class="readonly" style="width:256" type="txt" name="cboOwner" ID ="cboOwner" value="<%=MatDate%>" size="20" ></SCRIPT>&nbsp;</td>
  </tr>
    <td width="30%">&nbsp;Maturity Date</td>
    <td width="70%">
    <input type = "text" name="MaturityDate" value="<%=formatdate(MaturityDate)%>" readonly class="readonly" style="width:256" size="20" ></td>
  </tr>
   <tr>
     <td width="30%">&nbsp;Coupon Payment</td>
    <td width="70%">
    <input type = "text" name="CouponPayment" value="<%=CouponPayment%>" readonly class="readonly" style="width:256" size="20" >
    </td>
  </tr>
   <tr>
    <td width="30%">&nbsp;Coupon Rate (%)</td>
    <td width="70%"><input type="text" readonly class="readonly" name="CouponRate"  id="CouponRate" size="39" Value="<%=CouponRate%>" style="width:256" ></td>
  </tr>
   </tr>
    <td width="30%">&nbsp;Face Value</td>
    <td width="70%"><input type="text"  readonly class="readonly" style="width:256" name="FaceValue" size="39" Value="<%=formatnumber(FaceValue,2)%>"  onchange="JavaScript: formatnumber(this)"></td>
  </tr>
   
  <tr>
    <td width="30%">&nbsp;Forward Rate (%)</td>
    <td width="70%"><input type="text" name="ForwardRate" size="39"  style="width:256" readonly class="readonly" Value="<%=ForwardRate%>"></td>
  </tr>
  
  <tr>
    <td width="30%">Validity &nbsp;</td>
    <td width="70%"> <input readonly class="readonly" style="width:256" type="txt" name="cboOwner" ID ="cboOwner" value="<%=formatdate(ValidityDate)%>" size="20" ></td>
  </tr>
  <tr>
    <td width="30%">&nbsp;Account Manager &nbsp;</td>
    <td width="70%">
    <input readonly class="readonly" style="width:256" type="txt" name="cboOwner" ID ="cboOwner" value="<%=AccManager%>" size="20" >
   </td>
    
   
  </tr>
  <tr>
		<td width="16%">Commission Rate</td>
		<td width="30%">
		<input type="hidden" readonly class="readonly" name = 'cboCommission' id = 'cboCommission' style= "width:256" value="<%=CommissionRate%>" width="39">
		<input type="text" readonly class="readonly" name = 'cboCommission1' id = 'cboCommission1' style= "width:256" value="<%=CommissionRate1%>" width="39" size="20">
		</td>
	</tr>

  <tr>
    <td width="30%">&nbsp;Commission</td>
    <%if Isnumeric(Commission) then Commission=formatnumber(replace(formatnumber((RoundPoint05(Commission)),2),",",""),2) else Commission ="0.00"%>
    <td width="70%">
    <input type="text" readonly class="readonly" name="Commission" style="width:256" Value="<%=Commission%>" onchange="JavaScript: formatnumber(this)" size="20">
    
    </td>
  </tr>	
  <tr>
    <td width="30%">&nbsp;Salutation</td>
    <td width="70%"><input type="text" name="Salutation" readonly class="readonly" size="39" style="width:256" Value="<%=Salutation%>"></td>
  </tr>	
 
  <tr>
    <td>&nbsp;Counter Party</td>
    <td>
       <input type="text" readonly class="readonly" name="OwnerName" style="width:256" Value="<%=CounterParty%>" onchange="JavaScript: formatnumber(this)" size="20">
   </td>
  </tr>
   <tr>
    <td width="30%">&nbsp;Consideration</td>
	<%if isnumeric(Consideration) then Consideration =formatnumber((replace(formatnumber((RoundPoint05(Consideration)),2),",","")),2) %>
    <td width="70%"><input type="text" name="Consideration" size="39" Value="<%=formatnumber(Consideration,2)%>"reaonly class=readonly style="width:256" > </td>
  </tr>	
   <tr>
	<%if isnumeric(NetAmt) then NetAmt =formatnumber((replace(formatnumber((RoundPoint05(NetAmt)),2),",","")),2) %>
    <td width="30%">&nbsp;Net Amount</td>
    <td width="70%"><input type="text" name="NetAmt" size="39" Value="<%=NetAmt%>" readonly class=readonly style="width:256" ></td>
  </tr>	
  <tr>
    <td width="30%">&nbsp;Alternate Price</td>
    <td width="70%">
	<%if isnumeric(AlternatePrice) then AlternatePrice = formatnumber((replace(formatnumber(((AlternatePrice)),2),",","")),2)%>
    <input type="text" name="AlternatePrice" size="39" Value="<%=AlternatePrice%>" readonly class="readonly" style="width:256"></td>
  </tr>
</table>

</form>
  <!--#include file="../pageLoadingBottom.asp"-->
</body>

</html>