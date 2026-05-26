
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

Function SQLServerFormatWithCustomMax(sqlStr)
        Const startStr = "iif(isnull(max("
        Const midStr1 = ")),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = '"
        Const midStr2 = "'),max("
        Const endStr = ") + 1)"
        Const SQLstartStr = "ISNULL(MAX("
        Const SQLmidStr = ") + 1,(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = '"
        Const SQLendStr = "'))"
        Dim startStrPos 
        Dim midStrPos1
        Dim midStrPos2 
        Dim endStrPos 
        Dim fldName 
        Dim fldNameStartPos 
        Dim tblName
        Dim tblNameStartPos
        Dim oldStr 
        Dim newStr
                
        sqlStr = Replace(sqlStr, "[", "")
        sqlStr = Replace(sqlStr, "]", "")
        sqlStr = Replace(sqlStr, "#", "'")
        sqlStr = Replace(sqlStr, vbCrLf, "' + CHAR(13) + CHAR(10) + '")
        
        'sqlStr = lcase(sqlStr)
        do while 1 < 2
				startStrPos = InStr(1, sqlStr, startStr, vbTextCompare)
				If startStrPos < 1 Then
				       ' UpdateAuditTrail sqlStr
				        SQLServerFormatWithCustomMax = sqlStr
				        Exit Function
				End If
        
				midStrPos1 = InStr(startStrPos, sqlStr, midStr1, vbTextCompare)
				fldNameStartPos = startStrPos + Len(startStr)
				fldName = Mid(sqlStr, fldNameStartPos, midStrPos1 - fldNameStartPos)
				midStrPos2 = InStr(midStrPos1, sqlStr, midStr2, vbTextCompare)
				tblNameStartPos = midStrPos1 + Len(midStr1)
				tblName = Mid(sqlStr, tblNameStartPos, midStrPos2 - tblNameStartPos)
        
				newStr = SQLstartStr & fldName & SQLmidStr & tblName & SQLendStr
        
				endStrPos = InStr(midStrPos2, sqlStr, endStr, vbTextCompare)
				oldStr = Mid(sqlStr, startStrPos, (endStrPos + Len(endStr) - startStrPos))
				sqlStr = Replace(sqlStr, oldStr, newStr)  
		loop      

End Function

Function HandleQuote(sqlStr)
        Dim RE 
        Dim matches 
        Dim match 
        Dim i 
        Dim subStr 
        Dim strArray 
        Set RE = server.createobject("VBScript.RegExp")
        
        RE.Pattern = "([^,\s,',)]'[^,\s,',)])"
        Do
            Set matches = RE.Execute(sqlStr)
            For Each match In matches
                    subStr = match.Value
                    strArray = Split(subStr, "'")
                    subStr = Join(strArray, "''")
                    sqlStr = Replace(sqlStr, match.Value, subStr, 1, 1)
            Next
        Loop Until matches.Count = 0
        
        HandleQuote = sqlStr
End Function

Dim action
 
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

'Period Parameters
Dim DaysInCouponPeriod
Dim Basis
Dim ChangeInYield

'Validation Parameters
Dim inValid1  ' Mod((MaturityDate-IssueDate)/DaysInCouponPeriod) = 0
Dim inValid2  ' Mod((DaysInCouponPeriod-Basis)/91) = 0
Dim inValid3  ' IssueDate < MaturityDate
Dim inValid4  ' SettlementDate > IssueDate
Dim inValid5  ' SettlementDate < MaturityDate
Dim inValid6  ' CouponRate > 0
Dim inValid7  ' CouponRate < 30%
Dim inValid8  ' Forwardrate > 0
Dim inValid9  ' ForwardRate < 30%
Dim inValid10 ' NoPaymentsLeft < 50

action= ucase(Request.Form("action"))


if action = "CALCULATE" then
   ProposalNo = trim(Request.Form("ProposalNo"))	
   SecurityCode = trim(Request.Form("SecurityCode"))
   SecurityCode1 = trim(Request.Form("SecurityCode1"))
   IssueNo = trim(Request.Form("IssueNo"))
   FaceValue = trim(Request.Form("FaceValue"))
   IssueDate = cdate(trim(Request.Form("IssueDate")))
   MaturityDate = cdate(trim(Request.Form("MaturityDate")))
   SettlementDate = Cdate(trim(Request.Form("SettlementDate")))
   CouponRate = trim(Request.Form("CouponRate"))
   ForwardRate = trim(Request.Form("ForwardRate"))
   
   DaysInCouponPeriod = trim(Request.Form("DaysInCouponPeriod"))
   Basis = trim(Request.Form("Basis"))
   ChangeInYield = trim(Request.Form("ChangeInYield"))
   InputID = trim(Request.Form("InputID"))
   Client = trim(Request.Form("cboClient"))
   clientName = trim(Request.Form("txtClient"))
   BondType = trim(Request.Form("BondType"))
   TradeType = trim(Request.Form("TradeType"))
   CouponPayment = trim(Request.Form("CouponPayment"))
   ValidityDate = trim(Request.Form("ValidityDate"))
   ProposalDate = cdate(trim(Request.Form("ProposalDate")))
   CommissionRate1 = trim(Request.Form("cboCommission1"))
   CommissionRate = trim(Request.Form("cboCommission"))
   Commission = trim(Request.Form("Commission"))
   CounterParty = trim(Request.Form("cboCounterParty"))
   AccManager = trim(Request.Form("cboOwner"))	
   AlternatePrice= trim(Request.Form("AlternatePrice"))	
   
    CommissionRate = trim(Request.Form("cboCommission"))
   CommissionRate1 = trim(Request.Form("cboCommission1"))
   Commission = trim(Request.Form("Commission"))
   CounterParty = trim(Request.Form("cboCounterParty"))
   txtCounterparty = trim(Request.Form("cboCounterParty"))
   AccManager = trim(Request.Form("cboOwner"))	
   Salutation = trim(Request.Form("Salutation"))
   AlternatePrice= trim(Request.Form("AlternatePrice"))

	FaceValue = cdbl(replace(formatnumber((RoundPoint05(FaceValue)),2),",",""))
	ForwardRate = cdbl(replace(formatnumber(RoundPoint05(ForwardRate),2),",",""))
	Commission = cdbl(replace(formatnumber(RoundPoint05(Commission),2),",",""))
	Consideration = cdbl(replace(formatnumber((RoundPoint05(Consideration)),2),",",""))
	CommissionRate = cdbl(replace(formatnumber(((CommissionRate)),2),",",""))
	NetAmt = formatnumber(replace(formatnumber((RoundPoint05(NetAmt)),2),",",""),2)
	if isnull(AlternatePrice) or trim(AlternatePrice)="" then AlternatePrice=0 
	AlternatePrice = cdbl(replace(formatnumber(RoundPoint05(AlternatePrice),4),",",""))
   Salutation = trim(Request.Form("Salutation"))	
  buttonAction = Ucase(Request.Form("buttonAction")) 
       
  if buttonAction = "GENERATE" then
       Call Validation
       call CommissionCalc 
  elseif buttonAction = "SAVE" then
	call Validation
	call CommissionCalc  
    Conn.BeginTrans 
		'Update Bondproposal Table
if isnumeric(Convexity) and isnumeric(BondDirtyPrice(ForwardRate/100)) and isnumeric (BondCleanPrice) and isnumeric(AccruedInterest) then
	Dim sqlStr1
	if not isnumeric(alternateprice)then alternateprice=0
	FaceValue = cdbl(replace(formatnumber((RoundPoint05(FaceValue)),2),",",""))
	ForwardRate = cdbl(replace(formatnumber(RoundPoint05(ForwardRate),2),",",""))
	Commission = cdbl(replace(formatnumber(RoundPoint05(Commission),2),",",""))
	Consideration1 = cdbl(replace(formatnumber((RoundPoint05(Consideration)),2),",",""))
	CommissionRate = cdbl(replace(formatnumber(((CommissionRate)),2),",",""))
	Duration1 = cdbl(replace(formatnumber((Duration),2),",",""))
	Convexity2 = cdbl(replace(formatnumber((Convexity),2),",",""))
	AccruedInterest1 = cdbl(replace(formatnumber((AccruedInterest),2),",",""))
	NetAmount1 = cdbl(replace(formatnumber((RoundPoint05(NetAmt)),2),",",""))
	BondDirtyPrice1= cdbl(replace(formatnumber((BondDirtyPrice(ForwardRate/100)),4),",",""))
	BondCleanPrice1= cdbl(replace(formatnumber((BondCleanPrice),4),",",""))
	AlternatePrice1 = cdbl(replace(formatnumber((AlternatePrice),4),",",""))
	
	SqlStr =" Update BondProposals set ProposalDate='"& cdate(ProposalDate) &"',"&_
				" SettlementDate ='"& cdate(SettlementDate) &"'," &_ 
				" ForwardRate="& ForwardRate &",Consideration="& Consideration1 &" ," &_
				" BondCleanPrice=" & BondCleanPrice1 & ",AccruedInterest=" & AccruedInterest1 & "," &_ 
				" BondDirtyPrice=" & BondDirtyPrice1 & ", Validity='"& ValidityDate &"', " &_
				" CouponPeriodDays='"& DaysInCouponPeriod &"', Basis='"& Basis &"'," &_ 
				" RemainingDaysToCoupon='"& DaysToNextCoupon&"', PreviousCouponPayments='"& NoCouponPaymentsLeft&"'," &_ 
				" ChangeInYield='"& ChangeInYield &"', CouponRate="& CouponRate &"," &_ 
				" FaceValue= '" & ccur(FaceValue) & "',Duration='" & Duration1 & "'," &_ 
				" Convexity ='" & Convexity2 & "',Commission='" & Commission & "',AlternatePrice=" & AlternatePrice1 & ", NetAmount= "& NetAmount1 &"," &_
				" CounterParty='" & CounterParty & "',Salutation='" & Salutation & "' ,ModifiedBy='" & session("UserID") & "'  ,DateModified='" & now() & "' where Proposal_DPA_=" & clng(ProposalNo)
				
	 Conn.Execute SqlStr
	Conn.Execute "update Bond set BondRate ="& CouponRate &" where Bond_DPA_ =" & SecurityCode
	 'response.write " bond dpa "& SecurityCode
	Conn.CommitTrans 

	%>
	<Script Language="JavaScript">
		try{
			window.parent.dialogArguments.opener.location.reload();
			}
			catch(e){window.self.close()}
    </Script>
 <%
  end if
  end if
  
else
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
	           response.end
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
			CounterParty = trim(rst("CounterParty"))
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
	
end if

'*************************** Start Output Functions ***************************************************************
function Convexity()
 a = AskCleanPrice
 b = BidCleanPrice
 c = BondCleanPrice
 
 if not isnumeric(a) or not isnumeric(b) or not isnumeric(c) then
   Convexity = "#INVALID ENTRY!!"
 else
   Convexity = (a+b-2*c)/(2*c*((ChangeInYield/10000)^2))
 end if
end function

function AskCleanPrice()
 if Validation = 0 then
   AskCleanPrice = "#INVALID ENTRY!!"
 else
   AskCleanPrice = (Calculator(AskForwardRate) -  AccruedInterest)/1000000
 end if
end function

function CommissionCalc()
	Dim rsComm 
	Dim LowerRate, UpperRate
	
	Dim GrossAmt
	Dim Price
	
	set rsComm = server.CreateObject("Adodb.recordset")
	set rsComm = Conn.Execute("select * from Commission where Commission_DPA_=" & CommissionRate) 
	if not rsComm.EOF or not rsComm.BOF then
		 LowerRate = rsComm("BondCommission")
		 UpperRate = rsComm("UpperBondCommission")
		 GrossAmt =rsComm("BondBoundary")
		 
    	 if isnumeric(AlternatePrice)then Price =AlternatePrice
		 	if  Price <> 0 then
			 	Price = AlternatePrice
			 else
			 	Price = BondDirtyPrice(ForwardRate/100)
		 	end if
    	
		 if isnumeric(price) then
		 Consideration =((Price * FaceValue)/100)
		if Consideration < GrossAmt then
		Commission = (LowerRate/100)*Consideration
		elseif Consideration > GrossAmt then
			commission1 = (LowerRate/100)* GrossAmt
			commission2 = (UpperRate/100) * (Consideration - GrossAmt)
			Commission = commission1 + commission2
		elseif Commission <  rsComm("MinimumBondCommission") then
			Commission = LowerRate
		end if
		end if
		if TradeType =1 then
			NetAmt=Consideration + Commission 
		elseif TradeType=2 then
			NetAmt=Consideration - Commission 
		end if
		NetAmt = formatnumber(NetAmt,2)
		else 
		
		end if	
		NetAmt = formatnumber(replace(formatnumber((RoundPoint05(NetAmt)),2),",",""),2)
		CommissionCalc=  formatnumber((replace(formatnumber((RoundPoint05(Commission)),2),",","")),2)
	
end function



function BidCleanPrice()
 if Validation = 0 then
   BidCleanPrice = "#INVALID ENTRY!!"
 else
   BidCleanPrice = (Calculator(BidForwardRate) -  AccruedInterest)/1000000
 end if
end function

function Calculator(fRate)
	Dim DaysRemaining
	Dim CouponAmount
	'Initialise arrays
	Dim I(50)
	Dim K(50)
	Dim O(50)
	Dim Q(50)
	Dim H(50)
    dim LeftCoupons
	DaysRemaining = DaysToNextCoupon
	CouponAmount = CouponAmountPayment
   if isnumeric(NoCouponPaymentsLeft) then
	for a = 1 to  NoCouponPaymentsLeft
	    'H(a-1) = 1/((1+(fRate/2))^((DaysRemaining/DaysInCouponPeriod)+(a-1)))
	    H(a-1) =((1/((1+(fRate/2)))^((a-1)+(DaysRemaining/DaysInCouponPeriod))))
	    I(a-1) = CouponAmount * H(a-1)
	    if a = 1 then K(a-1) = I(a-1) else  K(a-1) = K(a-2) + I(a-1)
	    O(a-1) = FaceValue * H(a-1)
	    Q(a-1) = O(a-1) + K(a-1)
	next 
	
	if (NoCouponPaymentsLeft-1)<0 then
		LeftCoupons=0
	else
		LeftCoupons=NoCouponPaymentsLeft-1
	end if
    Calculator = Q(LeftCoupons)
   end if
end function

function Duration()
 a = AskCleanPrice
 b = BidCleanPrice
 c = BondCleanPrice
 
 if not isnumeric(a) or not isnumeric(b) or not isnumeric(c) then
   Duration = "#INVALID ENTRY!!"
 else
	Duration = (b - a)/(c*(AskForwardRate - BidForwardRate))	
 end if
 
end function

function BondDirtyPrice(Rate)
 if Validation = 0 then
   BondDirtyPrice = "#INVALID ENTRY!!"
 else
    'BondDirtyPrice = (Calculator(Rate)/1000000)
    BondDirtyPrice = (Calculator(Rate)/1000000)*(100000000/FaceValue)
 end if
end function

function BondCleanPrice()
 F3 = Calculator(ForwardRate/100)
 
 if Validation = 0 or not isnumeric(F3) then
   BondCleanPrice = "#INVALID ENTRY!!"
 else
   'BondCleanPrice =  ((F3 - AccruedInterest)/1000000)
   BondCleanPrice =  ((F3 - AccruedInterest)/1000000)*(100000000/FaceValue)
 end if
end Function 

function AccruedInterest()
if Validation = 0 then
   AccruedInterest = "#INVALID ENTRY!!"
 else
   AccruedInterest = (FaceValue * (InterestDaysAccrued/Basis) * (CouponRate/100)) 
 end if
 
end function


'*********************************** END Output Functions*******************************************************


'**************************** START DERIVED FUNCTIONS ***************************************************
function BidForwardRate()
  BidForwardRate = (ForwardRate/100) - (ChangeInYield/10000)
end function

function AskForwardRate()
 AskForwardRate = (ForwardRate/100) + (ChangeInYield/10000)
end function

Function deciChangeInYield()
 deciChangeInYield = (AskForwardRate - BidForwardRate)/2
end function

function DaysToNextCoupon()
 DaysToNextCoupon = Datediff("d",SettlementDate,NextCouponDate)
end function

function InterestDaysAccrued()
 InterestDaysAccrued = Datediff("d",lastCouponDate,SettlementDate)
end function

function CouponAmountPayment()
  CouponAmountPayment = (PeriodicInterestRate*FaceValue)/100
end function

function PeriodicInterestRate()
  PeriodicInterestRate = (((CouponRate/100) *MonthsInCouponPeriod)/12)*100 
end function

function BondMaturity()
   BondMaturity = Datediff("d",NextCouponDate,MaturityDate)/Basis
end function

function BondLife()
   BondLife = DateDiff("d",IssueDate,MaturityDate)/Basis
end function

function NextCouponDate()
  NextCouponDate = Dateadd("d",DaysInCouponPeriod,lastCouponDate)
end function

function CouponPayments()
 Months = MonthsInCouponPeriod
  if Months = "" then 
    CouponPayments = "NOT OK"
  elseif Months = 1 then 
    CouponPayments = "Monthly"
  elseif Months = 3 then
    CouponPayments = "Quarterly"
  elseif Months = 6 then
    CouponPayments = "Semi - Annually"
  elseif Months = 12 then 
    CouponPayments = "Annually"
  end if
end function
 
function lastCouponDate()
 lastCouponDate = Dateadd("d",(int(DateDiff("d",IssueDate,SettlementDate)/DaysInCouponPeriod)*DaysInCouponPeriod),IssueDate)
end function

function noPreviousCouponPeriod()
 noPreviousCouponPeriod = int(Datediff("d",IssueDate,lastCouponDate)/DaysInCouponPeriod)
end function

function NoCouponPaymentsLeft()
 if Not isnumeric(TotalCouponPayments) or not isnumeric(noPreviousCouponPeriod)then
   NoCouponPaymentsLeft = "NOT OK"
 else
   NoCouponPaymentsLeft = TotalCouponPayments - noPreviousCouponPeriod
 end if
end function

function TotalCouponPayments()
 if couponPaymentOK = 1 then
  TotalCouponPayments = (12 * Datediff("d",IssueDate,MaturityDate))/(364*MonthsInCouponPeriod)
 else
  TotalCouponPayments = "NOT OK"
 end if
end function

Function MonthsInCouponPeriod()
   MonthsInCouponPeriod = ((12*DaysInCouponPeriod)/Basis)
End function

function couponPaymentOK ()
 i = MonthsInCouponPeriod
  if PeriodDaysOK = 1 then
    if i = 1 or i = 3 or i = 6 or i = 12 then couponPaymentOK = 1 else couponPaymentOK = 0
  else 
   couponPaymentOK = 0
  end if
end function

function PeriodDaysOK()
  if Modulus(Datediff("d",IssueDate,MaturityDate),91) = 0 then PeriodDaysOK = 1 else PeriodDaysOK = 0
end function

'********************************** END DERIVED FUNCTIONS **********************************************************************

Function Validation ()
  Validation = 1 'Default: None of the validation rules has been violated.
  
  '**** Default: No validation rules have been violation ***********
  inValid1 = 0  ' Mod((MaturityDate-IssueDate)/DaysInCouponPeriod) = 0
  inValid2 = 0  ' Mod((DaysInCouponPeriod-Basis)/91) = 0
  inValid3 = 0  ' IssueDate < MaturityDate
  inValid4 = 0  ' SettlementDate > IssueDate
  inValid5 = 0  ' SettlementDate < MaturityDate
  inValid6 = 0  ' CouponRate > 0
  inValid7 = 0  ' CouponRate < 30%
  inValid8 = 0  ' Forwardrate > 0
  inValid9 = 0  ' ForwardRate < 30%
  inValid10 = 0 ' NoPaymentsLeft < 50
   
   if Modulus(Datediff("d",IssueDate,MaturityDate),DaysInCouponPeriod) = 0 then inValid1 = 0 else inValid1 = 1
   if Modulus(DaysInCouponPeriod-Basis,91) = 0 then inValid2 = 0 else inValid2 = 1
   if IssueDate < MaturityDate then inValid3 = 0  else inValid3 = 1
   if SettlementDate >= IssueDate then inValid4 = 0  else inValid4 = 1
   if SettlementDate < MaturityDate then inValid5 = 0  else inValid5 = 1
   if CDbl(CouponRate) >= 0 then  inValid6 = 0 else inValid6 = 1
   if CDbl(CouponRate)  < 30 then  inValid7 = 0 else inValid7 = 1
   if CDbl(ForwardRate) > 0 then  inValid8 = 0 else inValid8 = 1
   if CDbl(ForwardRate)  < 30 then  inValid9 = 0 else inValid9 = 1
   
   if not isnumeric(NoCouponPaymentsLeft)then 
       inValid10 = 1 
   else 
    if NoCouponPaymentsLeft < 50 then inValid10 = 0 else inValid10 = 1
   end if
  
   if ((inValid1+inValid2+inValid3+inValid4+inValid5+inValid6+inValid7+inValid8+inValid9+inValid10) > 0) then Validation = 0
End function

Function Modulus(number,divisor)
     
	if not isnumeric(number) or not isnumeric(divisor) then 
	 Modulus = ""
	 exit function
	end if
	
  Modulus = ((number/divisor) - int(number/divisor))*divisor
End function

Function FormatDate(theDate)
	On Error Resume Next
	FormatDate = Day(theDate) & "-" & MonthName(Month(theDate), True) & "-" & Year(theDate)
	If Err.Number > 0 Then
		FormatDate = theDate
	End If
End Function

dim currentEntityType
		
action = ucase(Request.Form("action"))
ID = Request.Form("ID")
currentEntityType = 1
currentEntityType=cint(id)

%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>BOND CALCULATOR</title>
 
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
		
		function IssueOnChange(){
			var item=document.frmBondCalculator
		
			item.IssueDate.value=item.IssueNo[item.IssueNo.selectedIndex].cboIDate;
			item.MaturityDate.value=item.IssueNo[item.IssueNo.selectedIndex].cboMDate;
			item.SecurityCode.value=item.IssueNo[item.IssueNo.selectedIndex].SecCode;
			//item.CouponPayment.options[item.IssueNo[item.IssueNo.selectedIndex].cboCPayment].selected;
			
			if (item.IssueNo[item.IssueNo.selectedIndex].cboCPayments=="Monthly"){
				item.CouponPayment.selectedIndex=0;
			}
			if (item.IssueNo[item.IssueNo.selectedIndex].cboCPayments=="Quaterly"){
				item.CouponPayment.selectedIndex=1;
			}
			if (item.IssueNo[item.IssueNo.selectedIndex].cboCPayments=="Annually"){
				item.CouponPayment.selectedIndex=2;
				
			}
			if (item.IssueNo[item.IssueNo.selectedIndex].cboCPayments.toLowerCase()=="semi-annually"){
				item.CouponPayment.selectedIndex=3;
				//document.frmBondCalculator.CouponPayment.selectedIndex=3
			}
			
			if (item.IssueNo[item.IssueNo.selectedIndex].cboCPayments.toLowerCase()=="biannually"){
				item.CouponPayment.selectedIndex=3;
				}
			//alert(item.IssueNo[item.IssueNo.selectedIndex].cboCPayments.toLowerCase());
			item.CouponRate.value=item.IssueNo[item.IssueNo.selectedIndex].cboRate;
			item.FaceValue.value=item.IssueNo[item.IssueNo.selectedIndex].cboFaceValue;
			if (item.IssueNo[item.IssueNo.selectedIndex].cboCPayments.toLowerCase()==""){
				item.CouponPayment.selectedIndex=-1;
				}
					
		}	
		function chkDate(){
			var item=document.frmBondCalculator 
			if(item.ValidityCheck.checked==true){
				item.ValidityDate.value=item.ValidDate.value;
			}
			else
			{
				item.ValidityDate.value='';
			}
		}
		
		function formatnumber(theTxt)
	{	
	var theprice = theTxt.value;
		theTxt.value=formatNum(theprice);
	}
	
function format_number(p,d) 
	{
  	var r;
  	if(p<0)
  		{
  		p=-p;
  		r=format_number2(p,d);
  		r="-"+r;
  		}
  	else
  		{
  		r=format_number2(p,d);
  		}
  return r;
	}

function format_number2(pnumber,decimals) 
	{
  	var strNumber = new String(pnumber);
  	var arrParts = strNumber.split('.');
  	var intWholePart = parseInt(arrParts[0],10);
  	var strResult = '';
  	if (isNaN(intWholePart))
    intWholePart = '0';
  	if(arrParts.length > 1)
  		{
    	var decDecimalPart = new String(arrParts[1]);
    	var i = 0;
    	var intZeroCount = 0;
     	while ( i < String(arrParts[1]).length )
     		{
       		if( parseInt(String(arrParts[1]).charAt(i),10) == 0 )
       			{
         		intZeroCount += 1;
         		i += 1;
       			}
       		else
         	break;
    		}
    	decDecimalPart = parseInt(decDecimalPart,10)/Math.pow(10,parseInt(decDecimalPart.length-decimals-1)); 
    	Math.round(decDecimalPart); 
    	decDecimalPart = parseInt(decDecimalPart)/10; 
    	decDecimalPart = Math.round(decDecimalPart); 

    	//If the number was rounded up from 9 to 10, and it was for 1 'decimal' 
    	//then we need to add 1 to the 'intWholePart' and set the decDecimalPart to 0. 

    	if(decDecimalPart==Math.pow(10, parseInt(decimals)))
    		{ 
      		intWholePart+=1; 
      		decDecimalPart="0"; 
    		} 
    	var stringOfZeros = new String('');
    	i=0;
    	if( decDecimalPart > 0 )
    		{
      		while( i < intZeroCount)
      			{
        		stringOfZeros += '0';
        		i += 1;
      			}
    		}
    	decDecimalPart = String(intWholePart) + "." + stringOfZeros + String(decDecimalPart); 
    	var dot = decDecimalPart.indexOf('.');
    	if(dot == -1)
    		{
      		decDecimalPart += '.'; 
      		dot = decDecimalPart.indexOf('.'); 
    		} 
    	var l=parseInt(dot)+parseInt(decimals); 
    	while(decDecimalPart.length <= l) 
    		{
      		decDecimalPart += '0'; 
    		}
    	strResult = decDecimalPart;
  		}
  	else
  		{
    	var dot; 
    	var decDecimalPart = new String(intWholePart); 

    	decDecimalPart += '.'; 
    	dot = decDecimalPart.indexOf('.'); 
    	var l=parseInt(dot)+parseInt(decimals); 
    	while(decDecimalPart.length <= l) 
    		{
      		decDecimalPart += '0'; 
    		}
    	strResult = decDecimalPart;
  		}
  	return strResult;
	}

	
</script>
<script language="Javascript">
		function validate(){
			var item=document.frmBondCalculator
			
			if(item.SettlementDate.value==''){
				alert('Please select the settlement date');
				item.SettlementDate.focus;
				return false;			
			}
			if(item.TradeType.selectedIndex==''){
				alert('Please select a Trade Type');
				item.TradeType.focus;				
				return false;			
			}
			if(item.BondType.selectedIndex==''){
				alert('Please select a Bond Type');
				item.BondType.focus;				
				return false;			
			}
			if(item.IssueNo.selectedIndex==''){
				alert('Please select a Bond Issue');
				item.IssueNo.focus;				
				return false;			
			}
			if(item.CouponPayment.selectedIndex==''){
				alert('Please select a Coupon Payment');
				item.CouponPayment.focus;				
				return false;			
			}
			if(item.ForwardRate.value==''){
				alert('Please select a Forward Rate');
				item.ForwardRate.focus;				
				return false;			
			}
			
			return true;
		}
		
function FetchAccounts(theList)
		{
			var i = 0;
			var entity = theList.value;
			var toList = document.frmBondCalculator.IssueNo;
			var issueno = document.frmBondCalculator.HidIssueNo.value;								
			
			frm = document.frmBondCalculator;				
			xmlhttp = createXMLHTTPObj();
			
			url="GetList.asp?ID="+entity+"&action=GetBondList&issueno="+issueno;
			xmlhttp.open("GET",url,true);
			xmlhttp.onreadystatechange=function() {
				if (xmlhttp.readyState==4) {
				returnStr = xmlhttp.responseText;
				returnStr = getBodyHTML(returnStr);
			
				var secList = "<select name = '" + toList.name + "' id = '" + toList.name + "' size='1'' onchange='IssueOnChange();' ";
				secList += "OnClick='event.cancelBubble=true;' " ;
				secList += "onChange='event.cancelBubble=true;' " ;
				secList += "onKeypress='return (dodefaultaction()==\"\"); ' "  ;
				secList += "onKeydown='return (dodefaultaction()==\"\");' " ; 
				secList += "onKeyup='return (change(" + toList.name + "));' " ; 
				secList += "onfocus='txtval = \"\";inputIsItemCode = 1;' "  ;
				secList += "onblur='txtval = \"\";inputIsItemCode = 1;'>" ;
				secList += returnStr ;
				secList += "</select>";
				
				toList.outerHTML = secList;														
				}
				}
			xmlhttp.setRequestHeader('Accept','message/x-jl-formresult');
			xmlhttp.send(); 
		
		
		
		}
	
</script>
</head>
<!--#include file="../libroutines.asp"-->
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
   


%>
<SCRIPT language="JavaScript">

	var cal=new ctlSpiffyCalendarBox("cal", "frmBondCalculator", "IssueDate","cmdIssueDate","<%=IssDate %>",'');
	var cal1=new ctlSpiffyCalendarBox("cal1", "frmBondCalculator", "MaturityDate","cmdMaturityDate","<%=MatDate %>",'');
	var cal2=new ctlSpiffyCalendarBox("cal2", "frmBondCalculator", "SettlementDate","cmdSettlementDate","<%=SettDate %>",1);
	var cal3=new ctlSpiffyCalendarBox("cal3", "frmBondCalculator", "ValidityDate","cmdValidityDate","<%=valDate%>",1);
	var cal4=new ctlSpiffyCalendarBox("cal4", "frmBondCalculator", "ProposalDate","cmdProposalDate","<%=PropDate %>",1);
</SCRIPT>
<form name="frmBondCalculator" method = 'post' action = 'EditBondProposal2.asp?ID=<%=ID1%>' id = "frmBondCalculator" onsubmit="return validate();">
<p><b>BOND CALCULATOR</b></p>

<p><input type="Submit" value="GENERATE" name="buttonAction">
<input type="Submit" value="SAVE" name="buttonAction">
<input type="Submit" value="CLOSE" name="buttonAction" onclick="javascript: self.close();">
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
   
   
     if isnumeric(BondCleanPrice)  then Response.Write formatnumber(replace(formatnumber(((BondCleanPrice)),4),",",""),4) else Response.Write BondCleanPrice 
   
    %>&nbsp;</td>
  </tr>
  <tr>
    <td width="30%" style="border-left-style: solid">&nbsp;Accrued Interest</td>
    <td width="70%" align="right" style="border-right-style: solid">
    <% 
    
     if isnumeric(AccruedInterest) then Response.Write formatnumber(replace(formatnumber(((AccruedInterest)),2),",",""),2) else Response.Write AccruedInterest
   
    %>&nbsp;</td>
  </tr>
  <tr>
    <td width="30%" style="border-left-style: solid">&nbsp;Bond Dirty Price</td>
    <td width="70%" align="right" style="border-right-style: solid">
    <% 
       		if isnumeric(BondDirtyPrice(ForwardRate/100)) then Response.Write formatnumber(replace(formatnumber(((BondDirtyPrice(ForwardRate/100))),4),",",""),4) else Response.Write BondDirtyPrice(ForwardRate/100)
    
     %>&nbsp;</td>
  </tr>
  <tr>
    <td width="30%" style="border-left-style: solid">&nbsp;Duration</td>
    <td width="70%" align="right" style="border-right-style: solid">
    <% 
        if isnumeric(Duration) then Response.Write formatnumber(replace(formatnumber(((Duration)),2),",",""),2) else Response.Write Duration
      %>&nbsp;</td>
  </tr>
  <tr>
    <td width="30%" style="border-left-style: solid; border-bottom-style: solid">&nbsp;Convexity</td>
    <td width="70%" align="right" style="border-right-style: solid; border-bottom-style: solid">
    <% 
         if isnumeric(Convexity) then Response.Write formatnumber(replace(formatnumber(((Convexity)),2),",",""),2) else Response.Write Convexity
     
    %>&nbsp;</td>
  </tr>
  <%
	if action = "CALCULATE" then 
         call CommissionCalc
       end if 
  %>
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
    <td width="70%"><SCRIPT language="JavaScript">cal4.writeControl();</SCRIPT>&nbsp;</td>
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
    <div style="display:none;">
    <select name = 'IssueNo1' id = 'IssueNo1' size="1" >
		<option selected SecCode = '0' SearchCode = "0" SearchText = ""  value = '' cboIDate='' cboMDate=''
		cboCPayments='' cboRate='' cboFaceValue='' disabled></option>
	<%
		sqlStr = "SELECT  distinct IssueList.* From [IssueList] WHERE Security_DPA_ =" & bondType & " Order By SecurityCode"        
    dim rsAccount
	set rsAccount= server.CreateObject("Adodb.recordset")
    Set rsAccount = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
    If Not (rsAccount.EOF Or rsAccount.BOF) Then
            rsAccount.MoveFirst
            Do Until rsAccount.EOF
                                   
               %> <option selected SearchCode ='<% = rsAccount.Fields("SecurityCode")%>'
										 SearchText = '<% = rsAccount.Fields("BondIssue")%>'
										value = '<% = rsAccount.Fields("BondIssue")%>'
										cboIDate='<% =  rsAccount.Fields("BondIDate")%>'
										cboMDate='<% =  rsAccount.Fields("BondMDate") %>'
										cboCPayments='<% = rsAccount.Fields("BondPayment")%>'
										cboRate='<% =  rsAccount.Fields("BondRate") %>'
										cboFaceValue='<% =  rsAccount.Fields("FaceValue")%>'>
										<%= rsAccount.Fields("BondIssue")%> </option>
                   <% rsAccount.MoveNext
            Loop
    else
		
   
    End If
    	
	%>
    </select></div></td>
  </tr>
  <tr>
    <td width="30%">&nbsp;Issue Date</td>
    <td width="70%">
    <input type = "text" name="IssueDate" style="width:256"  value="<%=formatdate(IssueDate)%>" readonly class="readonly" size="20"></td>
  </tr>
   <tr>
    <td width="30%">&nbsp;Settlement Date</td>
    <td width="70%"><SCRIPT language="JavaScript">cal2.writeControl();</SCRIPT>&nbsp;</td>
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
    <td width="70%"><input type="text" name="CouponRate"  id="CouponRate" size="39" Value="<%=CouponRate%>" style="width:256" ></td>
  </tr>
   </tr>
    <td width="30%">&nbsp;Face Value</td>
    <td width="70%"><input type="text" name="FaceValue" size="39" Value="<%=formatnumber(FaceValue,2)%>"  onchange="JavaScript: formatnumber(this)"></td>
  </tr>
   
  <tr>
    <td width="30%">&nbsp;Forward Rate (%)</td>
    <td width="70%"><input type="text" name="ForwardRate" size="39" Value="<%=ForwardRate%>"></td>
  </tr>
  
  <tr>
    <td width="30%">Validity &nbsp;</td>
    <td width="70%">&nbsp;<input type="checkbox" name="ValidityCheck" onclick="chkDate();" value="ON">&nbsp;<SCRIPT language="JavaScript">cal3.writeControl();</SCRIPT></td>
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
    <td width="70%"><input type="text" name="Salutation" size="39" Value="<%=Salutation%>"></td>
  </tr>	
 
  <tr>
    <td>&nbsp;Counter Party</td>
    <td>
       <input type = 'text' name ='txtCounterParty' id = 'txtCounterParty' size="10" onBlur="txtval = this.value; selectItem(cboCounterParty);" value="<%=client%>">
    <select name = 'cboCounterParty' id = 'cboCounterParty' size="1" 
			onKeypress="return (dodefaultaction()==''); " 
			onKeydown="return (dodefaultaction()==''); " 
			onKeyup="return (UpdateCode(change(cboCounterParty,0),cboCounterParty,txtCounterParty));" 
			onChange="UpdateCode(true,cboCounterParty,txtCounterParty);"
			onfocus="txtval = '';inputIsItemCode = 1;" 
			onblur="txtval = '';inputIsItemCode = 1;" readonly>
    	<!--option selected SearchCode = "" SearchText = "" value = ''></option-->
    	<option selected SearchCode = "" SearchText = "" value = ''></option>
		<%
		
		set rs = server.CreateObject("Adodb.recordset")
        sqlStr = "SELECT * FROM [ClientList]"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF					                
                
                ClientName=rs.Fields("ClientName")
			    NameClient=rs.Fields("Client_DPA_") & " " & Mid(ClientName,1,20)
			    if isnull(CounterParty) or trim(CounterParty)="" then CounterParty=0
					    %>                    
                        <option SearchCode = "<%=rs.Fields("Client_DPA_")%>" SearchText = "<%=rs.Fields("ClientName")%>" value = '<%=rs.Fields("Client_DPA_")%>' 
                        <%if (rs.Fields("Client_DPA_")=clng(CounterParty))then
							Response.Write "selected"
                         end if%>><%=NameClient%></option>
                        <%rs.MoveNext
                Loop
        End If
              
		%>
    </select></td>
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
    <input type="text" name="AlternatePrice" size="39" Value="<%=AlternatePrice%>"></td>
  </tr>
</table>
<p><b>Period Parameters</b></p>
<table border="1" width="90%" cellspacing="0" cellpadding="0" bordercolor="#000000">
  <tr>
    <td width="50%">&nbsp;No of Days in Coupon Period</td>
    <td width="50%"><input type="text" name="DaysInCouponPeriod" size="39" Value="<%=DaysInCouponPeriod%>"></td>
  </tr>
  <tr>
    <td width="50%">&nbsp;Basis</td>
    <td width="50%"><input type="text" name="Basis" size="39" Value="<%=Basis%>"></td>
  </tr>
  <tr>
    <td width="50%">&nbsp;Change in&nbsp; Yield (Basis Points)</td>
    <td width="50%"><input type="text" name="ChangeInYield" size="39" Value="<%=ChangeInYield%>"></td>
  </tr>
</table>
<p><b>Derived Parameters</b></p>
<table border="1" width="90%" bordercolor="#000000" cellspacing="0" cellpadding="0">
  <tr>
    <td width="51%">&nbsp;Coupon Payment</td>
    <td width="49%"><% Response.Write CouponPayments%>&nbsp;</td>
  </tr>
  <tr>
    <td width="51%">&nbsp;Last Coupon Date</td>
    <td width="49%"><% Response.Write FormatDate(lastCouponDate)%>&nbsp;</td>
  </tr>
  <tr>
    <td width="51%">&nbsp;Next Coupon Date</td>
    <td width="49%"><%  Response.Write FormatDate(NextCouponDate)%>&nbsp;</td>
  </tr>
  <tr>
    <td width="51%">&nbsp;Bond life (Clean, Year)</td>
    <td width="49%"><% Response.Write BondLife%>&nbsp;</td>
  </tr>
  <tr>
    <td width="51%">&nbsp;Bond Maturity (Clean, Year)</td>
    <td width="49%"><% Response.Write BondMaturity%>&nbsp;</td>
  </tr>
  <tr>
    <td width="51%">&nbsp;Periodic Interest rate</td>
    <td width="49%"><% Response.Write PeriodicInterestRate & "%"%>&nbsp;</td>
  </tr>
  <tr>
    <td width="51%">&nbsp;S/A Coupon Amount/ Payment (KES)&nbsp;</td>
    <td width="49%"><% Response.Write Formatnumber(CouponAmountPayment,2)%>&nbsp;</td>
  </tr>
  <tr>
    <td width="51%">&nbsp;Interest Days Accrued</td>
    <td width="49%"><% Response.Write InterestDaysAccrued%>&nbsp;</td>
  </tr>
  <tr>
    <td width="51%">&nbsp;Days remaining to Next Coupon</td>
    <td width="49%"><% Response.Write DaysToNextCoupon%>&nbsp;</td>
  </tr>
  <tr>
    <td width="51%">&nbsp;No of Previous Coupon Payments</td>
    <td width="49%"><% Response.Write noPreviousCouponPeriod%>&nbsp;</td>
  </tr>
  <tr>
    <td width="51%">&nbsp;No of Coupon Payments left</td>
    <td width="49%"><% Response.Write NoCouponPaymentsLeft%>&nbsp;</td>
  </tr>
  <tr>
    <td width="51%">&nbsp;Change in yield (Decimal)</td>
    <td width="49%"><% Response.Write deciChangeInYield%>&nbsp;</td>
  </tr>
</table>
</form>
  <!--#include file="../pageLoadingBottom.asp"-->
</body>

</html>