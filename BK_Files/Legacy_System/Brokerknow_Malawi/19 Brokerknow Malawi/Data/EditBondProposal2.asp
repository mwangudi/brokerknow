
<%

ID =Request.QueryString ("ID")

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

if isnull(action) then action=Request.Form("action")
 
if action = "CALCULATE" then 

    'Get Input Values
    ProposalNo = trim(Request.Form("ProposalNo"))
   SecurityCode = trim(Request.Form("SecurityCode"))
   IssueNo = trim(Request.Form("IssueNo"))
   FaceValue = trim(Request.Form("FaceValue"))
   IssueDate = Cdate(trim(Request.Form("IssueDate")))
   MaturityDate = Cdate(trim(Request.Form("MaturityDate"))) 
   SettlementDate = Cdate(trim(Request.Form("SettlementDate")))
   ForwardRate = trim(Request.Form("ForwardRate"))
    
   DaysInCouponPeriod = trim(Request.Form("DaysInCouponPeriod"))
   Basis = trim(Request.Form("Basis"))
   CouponRate = trim(Request.Form("CouponRate"))
   ChangeInYield = trim(Request.Form("ChangeInYield"))
   Client = trim(Request.Form("cboClient"))
   BondType = trim(Request.Form("BondType"))
   TradeType = trim(Request.Form("TradeType"))
   CouponPayment = trim(Request.Form("CouponPayment"))
   ValidityDate = trim(Request.Form("ValidityDate"))
   ProposalDate = cdate(trim(Request.Form("ProposalDate")))
   CommissionRate = trim(Request.Form("CommissionRate"))
   CommissionID = trim(Request.Form("CommissionID"))
   Commission = trim(Request.Form("Commission"))
   CounterParty = trim(Request.Form("cboCounterParty"))
   txtCounterparty = trim(Request.Form("cboCounterParty"))
   AccManager = trim(Request.Form("cboOwner"))	
   Salutation = trim(Request.Form("Salutation"))
   AlternatePrice= trim(Request.Form("AlternatePrice"))
   NetAmt = trim(Request.Form("NetAmt"))
  
   'Input Validations
   
   If trim(FaceValue) = "" Then%>
	   <script language = 'javascript'>
	   alert('Please Specify the Face Value.') 
	   window.history.back ()		
	   </script>
	     <% response.end
   End If
           
   If not isnumeric(trim(FaceValue)) Then%>
	   <script language = 'javascript'>
	   alert('Invalid Face Value.')  
	   window.history.back ()		
	   </script>
	     <% response.end
   End If
   
	If trim(ForwardRate) = "" Then%>
	   <script language = 'javascript'>
	   alert('Please Specify the Forward Rate.')
	   window.history.back ()  		
	   </script>
	     <% response.end
   End If
           
   If not isnumeric(ForwardRate) Then%>
	   <script language = 'javascript'>
	   alert('Invalid Forward Rate.')  
	   window.history.back ()		
	   </script>
	     <% response.end
   End If
	
	If trim(DaysInCouponPeriod) = "" Then%>
	   <script language = 'javascript'>
	   alert('Please Specify the Days in Coupon Period.')  	
	   window.history.back ()	
	   </script>
	     <% response.end
   End If
           
   If not isnumeric(trim(DaysInCouponPeriod)) Then%>
	   <script language = 'javascript'>
	   alert('Invalid Days in Coupon Period.')  
	   window.history.back ()		
	   </script>
	     <% response.end
   End If 
	
	If trim(Basis) = "" Then%>
	   <script language = 'javascript'>
	   alert('Please Specify the Basis.')  	
	   window.history.back ()	
	   </script>
	     <% response.end
   End If
           
   If not isnumeric(trim(Basis)) Then%>
	   <script language = 'javascript'>
	   alert('Invalid Basis.')  
	   window.history.back ()		
	   </script>
	     <% response.end
   End If 
	
   If trim(ChangeInYield) = "" Then%>
	   <script language = 'javascript'>
	   alert('Please Specify the Change in Yield.')  
	   window.history.back ()		
	   </script>
	     <% response.end
   End If
           
   If not isnumeric(trim(ChangeInYield)) Then%>
	   <script language = 'javascript'>
	   alert('Invalid Change in Yield') 
	   window.history.back () 		
	   </script>
	     <% response.end
   End If 
   	
   'Manipulate n Format values
   if isnull(CommissionRate) or trim(CommissionRate)="" then CommissionRate=0 
   if isnull(CommissionID) or trim(CommissionID)="" then CommissionID=0 
   if isnull(Commission) or trim(Commission)="" then Commission=0 
   if isnull(AlternatePrice) or trim(AlternatePrice)="" then AlternatePrice=0 
   if isnull(NetAmt) or trim(NetAmt)="" then NetAmt=0 
   
	FaceValue = formatnumber(RoundPoint05(cdbl(FaceValue)),2)
	ForwardRate = formatnumber(RoundPoint05(cdbl(ForwardRate)),2)
	CommissionRate = formatnumber(CommissionRate,2)
	Commission = formatnumber(RoundPoint05(Commission),2)
	Consideration = formatnumber(RoundPoint05(Consideration))
	NetAmt = formatnumber((RoundPoint05(cdbl(NetAmt))),2)
	AlternatePrice = formatnumber(RoundPoint05(AlternatePrice),4)

  buttonAction = Ucase(Request.Form("buttonAction")) 
       
  if buttonAction = "GENERATE" then
       Call Validation
       call CommissionCalc 
  elseif buttonAction = "SAVE" then
    
    call Validation
    call CommissionCalc 
    Conn.BeginTrans 
    
	'Save the bond and period parameters
	
	if not isnull(ValidityDate) or trim(ValidityDate)= "" then
		ValidityDate=cdate(date())
	else
		ValidityDate=cdate(ValidityDate)
	end if
	
	FaceValue = cdbl(FaceValue)
	ForwardRate = cdbl(ForwardRate)
	CommissionRate = Cdbl(CommissionRate)
	Commission = Cdbl(Commission)
	Consideration = Cdbl(Consideration)
	NetAmt = cdbl(NetAmt)
	AlternatePrice = cdbl(AlternatePrice)
	
	if isnumeric(Convexity) and isnumeric(BondDirtyPrice(ForwardRate/100)) and isnumeric (BondCleanPrice) and isnumeric(AccruedInterest) then
	Dim sqlStr1
	
	FaceValue = cdbl(replace(formatnumber((RoundPoint05(FaceValue)),2),",",""))
	ForwardRate = cdbl(replace(formatnumber(RoundPoint05(ForwardRate),2),",",""))
	Commission = cdbl(replace(formatnumber(RoundPoint05(Commission),2),",",""))
	Consideration = cdbl(replace(formatnumber((Consideration),2),",",""))
	CommissionRate = cdbl(replace(formatnumber(((CommissionRate)),2),",",""))
	Duration1 = cdbl(replace(formatnumber((Duration),2),",",""))
	Convexity2 = cdbl(replace(formatnumber(Convexity,2),",",""))
	AccruedInterest1 = cdbl(replace(formatnumber(AccruedInterest,2),",",""))
	NetAmount1 = cdbl(replace(formatnumber((RoundPoint05(NetAmt)),2),",",""))
	BondDirtyPrice1= cdbl(replace(formatnumber(RoundPoint05(BondDirtyPrice(ForwardRate/100)),4),",",""))
	BondCleanPrice1= cdbl(replace(formatnumber((BondCleanPrice),4),",",""))
	AlternatePrice1 = cdbl(replace(formatnumber((AlternatePrice),4),",",""))
	
    SqlStr =" Update BondProposals set ProposalDate=#"& FormatDate(cdate(ProposalDate)) &"#,"&_
				" SettlementDate =#"& FormatDate(cdate(SettlementDate)) &"#," &_ 
				" ForwardRate="& ForwardRate &",Consideration="& Consideration &" ," &_
				" BondCleanPrice=" & BondCleanPrice1 & ",AccruedInterest=" & AccruedInterest1 & "," &_ 
				" BondDirtyPrice=" & BondDirtyPrice1 & ", Validity=#"& FormatDate(ValidityDate) &"#, " &_
				" CouponPeriodDays="& DaysInCouponPeriod &", Basis= "& Basis &" ," &_ 
				" RemainingDaysToCoupon= "& DaysToNextCoupon&" , PreviousCouponPayments= "& NoCouponPaymentsLeft&"," &_ 
				" ChangeInYield="& ChangeInYield &"," &_ 
				" FaceValue= " & ccur(FaceValue) & ",Duration=" & Duration1 & "," &_ 
				" Convexity =" & Convexity2 & ",Commission=" & Commission & ",AlternatePrice=" & AlternatePrice1 & ", NetAmount= "& NetAmount1 &"," &_
				" CounterParty='" & CounterParty & "',Salutation='" & Salutation & "' ,ModifiedBy=" & session("UserID") & "  ,DateModified=#" & FormatDate(now()) & "# where Proposal_DPA_=" & clng(ProposalNo)
	
	
	sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
				
	Conn.Execute SqlStr
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
  else
  %>
	<script language = "javascript">
		 alert ('Output not saved');
	</script>
   <% 
  end if
else
   
   if action = "CALCULATE" then ID = ProposalNo
 'Fetch Input Data from DB
	set rst = CreateObject("ADODB.Recordset")

	'SQLStr = "Select * From BondsInputs"
	SQLStr = "Select * from BondProposalList where Proposal_DPA_= "& Clng(ID)	
	set rst = Conn.Execute (SQLStr)
	If rst.EOF Or rst.BOF Then%>
	   <script language = 'javascript'>
	   alert('Process Stopped! There are no Bond Proposal Values.')  		
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
			Client = trim(rst("Client_DPA_"))
			ClientName =trim(rst("clientName"))
			BondType = trim(rst("Security_DPA_"))
			TradeType = trim(rst("TradeType"))
			CouponPayment = trim(rst("BondPayment"))
			ValidityDate = trim(rst("Validity"))
			Duration1 = trim(rst("Duration"))
			Convexity1 = trim(rst("Convexity"))
			AccruedInterest1 = trim(rst("AccruedInterest"))
			ProposalDate = trim(rst("ProposalDate"))
   			SecurityCode1  = trim(rst("SecurityCode"))
   			SecurityCode  = trim(rst("Bond_DPA_"))
			CommissionID = trim(rst("Comm_DPA"))
			CommissionRate=trim(rst("CommissionRate"))
			Commission = trim(rst("Commission"))
			CounterParty = trim(rst("CounterParty"))
			AccManager = trim(rst("OwnerName"))
			Salutation = trim(rst("Salutation"))
			AlternatePrice = trim(rst("AlternatePrice"))
			BondDirtyPrice1 = trim(rst("BondDirtyPrice"))
			BondCleanPrice1 = trim(rst("BondCleanPrice"))
			Consideration =trim(rst("Consideration"))
			NetAmt =trim(rst("NetAmount"))
			
		    rst.MoveNext
		Loop
	End If
	
end if

'*************************** Start Output Functions ***************************************************************
function Convexity()
 a = Cdbl(AskCleanPrice)
 b = cdbl(BidCleanPrice)
 c = cdbl(BondCleanPrice)
 
 if not isnumeric(a) or not isnumeric(b) or not isnumeric(c) then
   Convexity = "#INVALID ENTRY!!"
 else
   Convexity = (a+b-2*c)/(2*c*((ChangeInYield/10000)^2))
 end if
end function

function CommissionCalc()
	Dim rsComm 
	Dim LowerRate, UpperRate
	
	Dim GrossAmt
	Dim Price
	
	set rsComm = server.CreateObject("Adodb.recordset")
	
	set rsComm = Conn.Execute("select * from Commission where Commission_DPA_ = " & CommissionID) 
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
		
		'Factor in Agent Returnable Commission 
		 set rsAgent = server.CreateObject("Adodb.recordset")
		 Sql = "select * from AgentRCommissionList where Client_DPA_ = " & Client & " " & _
		       "AND AgentStatus = 1"
	     set rsAgent = Conn.Execute(Sql)
	     
		 if not rsAgent.EOF or not rsAgent.BOF then
		   Commission = Commission - (cdbl(rsAgent.Fields("Commission"))/100*Commission)
		 end if
		 
		end if
		if TradeType =1 then
			NetAmt=Consideration + Commission 
		elseif TradeType=2 then
			NetAmt=Consideration - Commission 
		end if
		
		else 
		
		end if	
		NetAmt = formatnumber(RoundPoint05(NetAmt),2)
		CommissionCalc=  formatnumber(RoundPoint05(Commission),2)
end function

function AskCleanPrice()
 if Validation = 0 then
   AskCleanPrice = "#INVALID ENTRY!!"
 else
   dirtyPrice2 = formatnumber((formatnumber(Calculator(AskForwardRate),4)) * (100 / FaceValue),4)
   AskCleanPrice = formatnumber(Cdbl(dirtyPrice2) - (AccruedInterest*100/FaceValue),4)
 end if
end function

function BidCleanPrice()
 if Validation = 0 then
   BidCleanPrice = "#INVALID ENTRY!!"
 else
   dirtyPrice2 = formatnumber((formatnumber(Calculator(BidForwardRate),4)) * (100 / FaceValue),4)
   BidCleanPrice = formatnumber(dirtyPrice2 - (AccruedInterest*100/FaceValue),4)
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
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/accountList.js"></SCRIPT>
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT>
<script language="JavaScript" src="CALENDAR/calendar.js"></script>
<!--CALENDAR -->

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
 <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css"> 
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">

 
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
			item.CouponPayment.value=item.IssueNo[item.IssueNo.selectedIndex].cboCPayments;
			//item.ForwardRate.value=item.IssueNo[item.IssueNo.selectedIndex].ForwardRate;
			
			if (item.IssueNo[item.IssueNo.selectedIndex].cboCPayments.toLowerCase()=="biannually"){
				item.CouponPayment.selectedIndex=3;
				}
			
			item.CouponRate.value=item.IssueNo[item.IssueNo.selectedIndex].cboRate;
			item.FaceValue.value=item.IssueNo[item.IssueNo.selectedIndex].cboFaceValue;
			if (item.IssueNo[item.IssueNo.selectedIndex].cboCPayments.toLowerCase()==""){
				item.CouponPayment.selectedIndex=-1;
				}
					
		}	
		function setCommision(){
			var item=document.frmBondCalculator
			item.Commission.value=item.cboCommission[item.cboCommission.selectedIndex].comm;
		
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
		function Format2Decimals(theItem)
	{
		var item = document.frmBondCalculator;
			theItem.value=format2NumberCommasOnly(theItem.value)
			theItem.value= formatNum(theItem.value);
	 //document.all.item("txtTotal").value= formatNum(document.all.item("txtTotal").value);
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
	
	
function FetchAccManager()
		{
			var item = document.frmBondCalculator
			
			item.cboOwner.value = item.cboClient[item.cboClient.selectedIndex].AccManager;
			item.CommissionID.value = item.cboClient[item.cboClient.selectedIndex].Comm_DPA;
			item.CommissionRate.value = item.cboClient[item.cboClient.selectedIndex].CommRate;
			
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

	var cal=new ctlSpiffyCalendarBox("cal", "frmBondCalculator", "IssueDate","cmdIssueDate","<%=IssDate %>",1);
	var cal1=new ctlSpiffyCalendarBox("cal1", "frmBondCalculator", "MaturityDate","cmdMaturityDate","<%=MatDate %>",1);
	var cal2=new ctlSpiffyCalendarBox("cal2", "frmBondCalculator", "SettlementDate","cmdSettlementDate","<%=SettDate %>",1);
	var cal3=new ctlSpiffyCalendarBox("cal3", "frmBondCalculator", "ValidityDate","cmdValidityDate","<%=valDate%>",1);
	var cal4=new ctlSpiffyCalendarBox("cal4", "frmBondCalculator", "ProposalDate","cmdProposalDate","<%=PropDate %>",1);
</SCRIPT>
<form name="frmBondCalculator" method = 'post' action = 'EditBondProposal2.asp' id = "frmBondCalculator">
<p><b>BOND CALCULATOR</b></p>

<p>
<input type="Submit" value="GENERATE" name="buttonAction">
<input type="Submit" value="SAVE" name="buttonAction">
<input type="Submit" value="CLOSE" name="buttonAction" onclick="javascript: self.close();">
<input type="hidden" value="<%=SecurityCode%>" name="SecurityCode" id="SecurityCode">
<input type="hidden" value="<%=FormatDate(now)%>" name="ValidDate">
<input type="hidden" value="CALCULATE" name="action">
<input type="Hidden" name="HidIssueNo" id="HidIssueNo" size="1" Value="<%=IssueNo%>"></p>

<table border="0" width="90%" bordercolor="#000000" cellspacing="0" cellpadding="0">
  <tr>
    <td width="30%" style="border-left-style: solid; border-top-style: solid">&nbsp;Bond Clean Price</td>
    <td width="70%" align="right" style="border-right-style: solid; border-top-style: solid">
    <% 
   
    if action = "CALCULATE" then 
     if isnumeric(BondCleanPrice)  then Response.Write formatnumber(replace(formatnumber(((BondCleanPrice)),4),",",""),4) else Response.Write BondCleanPrice 
    else
     Response.Write BondCleanPrice1
    end if
    %>&nbsp;</td>
  </tr>
  <tr>
    <td width="30%" style="border-left-style: solid">&nbsp;Accrued Interest</td>
    <td width="70%" align="right" style="border-right-style: solid">
    <% 
    if action = "CALCULATE" then 
     if isnumeric(AccruedInterest) then Response.Write formatnumber(replace(formatnumber(((AccruedInterest)),2),",",""),2) else Response.Write AccruedInterest
    else
     Response.Write AccruedInterest1
    end if
    %>&nbsp;</td>
  </tr>
  <tr>
    <td width="30%" style="border-left-style: solid">&nbsp;Bond Dirty Price</td>
    <td width="70%" align="right" style="border-right-style: solid">
    <% if action = "CALCULATE" then 
       		if isnumeric(BondDirtyPrice(ForwardRate/100)) then Response.Write formatnumber(replace(formatnumber(((BondDirtyPrice(ForwardRate/100))),4),",",""),4) else Response.Write BondDirtyPrice(ForwardRate/100)
     else
        Response.Write BondDirtyPrice1
     end if
     %>&nbsp;</td>
  </tr>
  <tr>
    <td width="30%" style="border-left-style: solid">&nbsp;Duration</td>
    <td width="70%" align="right" style="border-right-style: solid">
    <% if action = "CALCULATE" then 
        if isnumeric(Duration) then Response.Write formatnumber(replace(formatnumber(((Duration)),2),",",""),2) else Response.Write Duration
       else
        Response.Write Duration1
       end if %>&nbsp;</td>
  </tr>
  <tr>
    <td width="30%" style="border-left-style: solid; border-bottom-style: solid">&nbsp;Convexity</td>
    <td width="70%" align="right" style="border-right-style: solid; border-bottom-style: solid">
    <% if action = "CALCULATE" then 
         if isnumeric(Convexity) then Response.Write formatnumber(replace(formatnumber(((Convexity)),2),",",""),2) else Response.Write Convexity
       else
        Response.Write Convexity1
       end if 
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
    <input type ="text" Name ="ProposalNo" ID ="ProposalNo" value="<%=ProposalNo%>" readonly class=readonly style= "width:256" size="20"></td>
  </tr>
    <tr>
    <td width="30%">&nbsp;Proposal Date</td>
    <td width="70%"><SCRIPT language="JavaScript">cal4.writeControl();</SCRIPT>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;Client</td>
    <td>
    <input type = 'text' name ='txtClientCode' id = 'txtClientCode' size="10" onBlur="txtval = this.value; selectItem(cboClient); FetchAccManager();" value="<%=client%>">
    <select name = 'cboClient' id = 'cboClient' size="1" 
			onKeypress="return (dodefaultaction()==''); " 
			onKeydown="return (dodefaultaction()==''); " 
			onKeyup="return (UpdateCode(change(cboClient,0),cboClient,txtClientCode));" 
			onChange="UpdateCode(true,cboClient,txtClientCode);FetchAccManager();"
			onfocus="txtval = '';inputIsItemCode = 1;" 
			onblur="txtval = '';inputIsItemCode = 1;" readonly >
    	<option selected SearchCode = '' SearchText = '' value = '' Comm_DPA='' CommRate='' AccManager=''></option>
		<%
		
		dim ClientName
		dim NameClient
		set rs = server.CreateObject("Adodb.recordset")
       
	'sqlStr =" SELECT dbo.BondClientList.*, dbo.Commission.BondCommission AS CommissionRate, " &_
	'		" dbo.OwnerList.OwnerName AS Owner FROM dbo.BondClientList INNER JOIN " &_
	'		" dbo.Commission ON dbo.BondClientList.Commission_DPA_ = dbo.Commission.Commission_DPA_ " &_
	'		" LEFT OUTER JOIN dbo.OwnerList ON dbo.BondClientList.Owner_DPA_ = dbo.OwnerList.Owner_DPA_"
	
	sqlStr = " SELECT  dbo.Client.ClientName, dbo.Client.Commission_DPA_,   " & _
            " dbo.Commission.BondCommission AS CommissionRate, dbo.OwnerList.OwnerName AS Owner, dbo.Client.Client_DPA_ " & _
            " FROM  dbo.Client LEFT OUTER JOIN " & _
            " dbo.OwnerList ON dbo.Client.Owner_DPA_ = dbo.OwnerList.Owner_DPA_ LEFT OUTER JOIN " & _
            " dbo.Commission ON dbo.Client.Commission_DPA_ = dbo.Commission.Commission_DPA_"

        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF					                
                
                ClientName=rs.Fields("ClientName")
					    %>                    
                        <option CommRate="<%=rs.Fields("CommissionRate")%>" Comm_DPA="<%=rs.Fields("Commission_DPA_")%>" AccManager = "<%=rs.Fields("Owner")%>" SearchCode = "<%=rs.Fields("Client_DPA_")%>" SearchText = "<%=rs.Fields("ClientName")%>" value = '<%=rs.Fields("Client_DPA_")%>' 
                        <%if (rs.Fields("Client_DPA_")=clng(client))then
							Response.Write "selected"
                         end if%> ><%=mid(ClientName,1,30)%></option>
                        <%rs.MoveNext
                Loop
        End If
              
		%>
    </select></td>
  </tr>
  <tr>
    <td width="30%">&nbsp;Trade Type</td>
    <td width="70%"><select name="TradeType">
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
	</select></td>
  </tr>
  
  <tr>
    <td width="30%">&nbsp;Bond Type</td>
    <td width="70%">
    <select name="BondType" id="BondType" size ='1' onchange='FetchAccounts(this);'>
    <%
		dim rst
		set rst = server.CreateObject("Adodb.recordset")
	    sqlStr = "SELECT * FROM [Security] where OrderSecType_DPA_ = 1 order by Security_DPA_"

        Set rst = Conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		
    %>
		<option></option>
		<% while not rst.eof %>
		<option  value="<%=rst("Security_DPA_")%>" <%if(rst("Security_DPA_")=clng(BondType)) then Response.Write "Selected"%>><%=rst("SecurityCode")%></option>
	<%
		rst.movenext
		wend
		rst.close
		set rst= nothing	
	%>
	</select></td>
  </tr>
   <tr>
    <td>&nbsp;Bond Issue</td>
    <td>
    <select name = 'IssueNo' id = 'IssueNo' size="1" onchange="IssueOnChange();">
		<option SecCode = '0' SearchCode = '0' SearchText = ''  value = '' cboIDate='' cboMDate=''
		cboCPayments='' cboRate='' cboFaceValue='' ForwardRate=''>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>
	<%
		sqlStr = "SELECT  * From [IssueList] WHERE Security_DPA_ =" & BondType  & "  and  bondMdate> getdate()Order By SecurityCode"        
    	dim rsAccount
		set rsAccount= server.CreateObject("Adodb.recordset")
    	Set rsAccount = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
    If Not (rsAccount.EOF Or rsAccount.BOF) Then
            rsAccount.MoveFirst
            Do Until rsAccount.EOF
                                   
               %> <option ForwardRate = '<%= rsAccount.Fields("ForwardRate")%>'	SecCode = '<%= rsAccount.Fields("Bond_DPA_")%>'
								SearchText = '<% = rsAccount.Fields("BondIssue")%>'
								value = '<% = rsAccount.Fields("BondIssue")%>'
								cboIDate='<% =  rsAccount.Fields("BondIDate")%>'
								cboMDate='<% =  rsAccount.Fields("BondMDate") %>'
								cboCPayments='<% = rsAccount.Fields("BondPayment")%>'
								cboRate='<% =  rsAccount.Fields("Rate") %>'
								cboFaceValue='<% =  formatnumber(rsAccount.Fields("FaceValue"),2)%>'
								 
								<% if(rsAccount.Fields("Bond_DPA_") = cint(SecurityCode)) then response.write "selected"%>
								>
								<%= rsAccount.Fields("BondIssue")%> </option>
                   <% rsAccount.MoveNext
            Loop
    else
		%>
		<option selected SecCode = '0' SearchCode = "0" SearchText = ""  value = '' cboIDate='' cboMDate=''
		cboCPayments='' cboRate='' cboFaceValue=''   ForwardRate=''>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>
		<%
   
    End If
    	
	%>
    </select></td>
  </tr>

  <tr>
    <td width="30%">&nbsp;Issue Date</td>
    <td width="70%">
    <input type="text" name="IssueDate" class="readonly" readonly value="<%=formatdate(IssueDate)%>" style= "width:256" size="20"></td>
  </tr>
   <tr>
    <td width="30%">&nbsp;Settlement Date</td>
    <td width="70%"><SCRIPT language="JavaScript">cal2.writeControl();</SCRIPT>&nbsp;</td>
  </tr>
    <td width="30%">&nbsp;Maturity Date</td>
    <td width="70%">
    <input type="text" name="MaturityDate" class="readonly" readonly value="<%=formatdate(MaturityDate)%>" style= "width:256" size="20"></td>
  </tr>
  <tr>
    <td width="30%">&nbsp;Coupon Payment</td>
    <td width="70%">
	<input type ="text" class="readonly" readonly style="width:256" name ="CouponPayment" value="<%=CouponPayment%>">
	</td>
  </tr>
  <tr>
    <td width="30%">&nbsp;Coupon Rate (%)</td>
    <td width="70%">
    <input type="text" name="CouponRate"  id="CouponRate" style="width:256"  Value="<%=CouponRate%>" size="20"></td>
  </tr>
    <td width="30%">&nbsp;Face Value</td>
    <td width="70%"><input type="text" name="FaceValue" size="39" Value="<%=formatnumber(FaceValue,2)%>"  onchange="JavaScript: formatnumber(this)"></td>
  </tr>
   
  <tr>
    <td width="30%">&nbsp;Forward Rate (%)</td>
    <td width="70%"><input type="text" name="ForwardRate" id="ForwardRate" size="39" Value="<%=ForwardRate%>"></td>
  </tr>
  
  <tr>
    <td width="30%">Validity &nbsp;</td>
    <td width="70%">&nbsp;<input type="checkbox" name="ValidityCheck" onclick="chkDate();" value="ON">&nbsp;<SCRIPT language="JavaScript">cal3.writeControl();</SCRIPT></td>
  </tr>
  <tr>
    <td width="30%">&nbsp;Account Manager &nbsp;</td>
    <td width="70%">
    <input type="text" readonly class= readonly name="cboOwner" ID ="cboOwner" value="<%=AccManager%>" style= "width:256" size="20">
    
    </td>
   
  </tr>
  <tr>
		<td width="16%">Commission Rate</td>
		<td width="30%">
		<input type="hidden" class='readonly' readonly name = 'CommissionID' id = 'CommissionID' style= "width:256" value="<%=CommissionID%>" width="39">
		<input type="text" readonly class="readonly" name = 'CommissionRate' id = 'CommissionRate' style= "width:256" value="<%=CommissionRate%>" width="39" size="20">
		</td>
	</tr>
 <tr>
    <td width="30%">&nbsp;Commission</td>
    <%if Isnumeric(Commission) then Commission= formatnumber(RoundPoint05(Commission),2) else Commission ="0.00"%>
    <td width="70%"><input readonly style= "width:256" class= "readonly" type="text" id="Commission" name="Commission" size="39" Value="<%=Commission%>" onchange="JavaScript: formatnumber(this)"></td>
  </tr>	
  <tr>
    <td width="30%">&nbsp;Salutation</td>
    <td width="70%"><input type="text" name="Salutation" size="39" Value="<%=Salutation%>"></td>
  </tr>	
 
  <tr>
    <td>&nbsp;Counter Party</td>
    <td>
       <input type = 'text' name ='txtCounterParty' id = 'txtCounterParty' size="10" onBlur="txtval = this.value; selectItem(cboCounterParty);" value="<%=txtCounterparty%>">
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
    <td width="70%">
	<%if isnumeric(Consideration) then Consideration =formatnumber((replace(formatnumber((RoundPoint05(Consideration)),2),",","")),2) %>
    <input type="text" name="Consideration" style= "width:256" Value="<%=Consideration%>" readonly class= readonly style= "width:256" size="20"></td>
  </tr>	
   <tr>
    <td width="30%">&nbsp;Net Amount</td>
    <td width="70%"><input type="text" name="NetAmt" size="39" Value="<%=NetAmt%>" readonly class="readonly" style= "width:256"></td>
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
    <td width="49%"><% if action = "CALCULATE" then Response.Write CouponPayments%>&nbsp;</td>
  </tr>
  <tr>
    <td width="51%">&nbsp;Last Coupon Date</td>
    <td width="49%"><% if action = "CALCULATE" then Response.Write FormatDate(lastCouponDate)%>&nbsp;</td>
  </tr>
  <tr>
    <td width="51%">&nbsp;Next Coupon Date</td>
    <td width="49%"><% if action = "CALCULATE" then Response.Write FormatDate(NextCouponDate)%>&nbsp;</td>
  </tr>
  <tr>
    <td width="51%">&nbsp;Bond life (Clean, Year)</td>
    <td width="49%"><% if action = "CALCULATE" then Response.Write BondLife%>&nbsp;</td>
  </tr>
  <tr>
    <td width="51%">&nbsp;Bond Maturity (Clean, Year)</td>
    <td width="49%"><% if action = "CALCULATE" then Response.Write BondMaturity%>&nbsp;</td>
  </tr>
  <tr>
    <td width="51%">&nbsp;Periodic Interest rate</td>
    <td width="49%"><% if action = "CALCULATE" then Response.Write PeriodicInterestRate & "%"%>&nbsp;</td>
  </tr>
  <tr>
    <td width="51%">&nbsp;S/A Coupon Amount/ Payment (KES)&nbsp;</td>
    <td width="49%"><% if action = "CALCULATE" then Response.Write Formatnumber(CouponAmountPayment,2)%>&nbsp;</td>
  </tr>
  <tr>
    <td width="51%">&nbsp;Interest Days Accrued</td>
    <td width="49%"><% if action = "CALCULATE" then Response.Write InterestDaysAccrued%>&nbsp;</td>
  </tr>
  <tr>
    <td width="51%">&nbsp;Days remaining to Next Coupon</td>
    <td width="49%"><% if action = "CALCULATE" then Response.Write DaysToNextCoupon%>&nbsp;</td>
  </tr>
  <tr>
    <td width="51%">&nbsp;No of Previous Coupon Payments</td>
    <td width="49%"><% if action = "CALCULATE" then Response.Write noPreviousCouponPeriod%>&nbsp;</td>
  </tr>
  <tr>
    <td width="51%">&nbsp;No of Coupon Payments left</td>
    <td width="49%"><% if action = "CALCULATE" then Response.Write NoCouponPaymentsLeft%>&nbsp;</td>
  </tr>
  <tr>
    <td width="51%">&nbsp;Change in yield (Decimal)</td>
    <td width="49%"><% if action = "CALCULATE" then Response.Write deciChangeInYield%>&nbsp;</td>
  </tr>
</table>
<p><b>Validations</b></p>
<table border="1" width="90%" bordercolor="#000000" cellspacing="0" cellpadding="0">
  <tr>
    <td width="81%">&nbsp;Check : Mod ((Maturity Date-Issue Date) / Coupon Period) = 0</td>
    <td width="19%" align="right"><% if inValid1 = 0  and not inValid1 = "" then 
                              Response.Write "OK"
                             elseif inValid1 = 1 then
                              Response.Write "NOT OK" 
                             End if
                   %>&nbsp;</td>
  </tr>
  <tr>
    <td width="81%">&nbsp;Check : Mod (Basis/Coupon Period) = 0</td>
    <td width="19%" align="right"><% if inValid2 = 0  and not inValid2 = "" then 
                              Response.Write "OK"
                             elseif inValid2 = 1 then
                              Response.Write "NOT OK" 
                             End if
                   %>&nbsp;</td>
  </tr>
  <tr>
    <td width="81%">&nbsp;Check : Issue Date &lt; Maturity Date</td>
    <td width="19%" align="right"><% if inValid3 = 0  and not inValid3 = "" then 
                              Response.Write "OK"
                             elseif inValid3 = 1 then
                              Response.Write "NOT OK" 
                             End if
                   %>&nbsp;</td>
  </tr>
  <tr>
    <td width="81%">&nbsp;Check : Settlement Date >= Issue Date</td>
    <td width="19%" align="right"><% if inValid4 = 0  and not inValid4 = "" then 
                              Response.Write "OK"
                             elseif inValid4 = 1 then
                              Response.Write "NOT OK" 
                             End if
                   %>&nbsp;</td>
  </tr>
  <tr>
    <td width="81%">&nbsp;Check : Settlement Date &lt; Maturity Date</td>
    <td width="19%" align="right"><% if inValid5 = 0  and not inValid5 = "" then 
                              Response.Write "OK"
                             elseif inValid5 = 1 then
                              Response.Write "NOT OK" 
                             End if
                   %>&nbsp;</td>
  </tr>
  <tr>
    <td width="81%">&nbsp;Check : Coupon Rate >= 0</td>
    <td width="19%" align="right"><% if inValid6 = 0  and not inValid6 = "" then 
                              Response.Write "OK"
                             elseif inValid6 = 1 then
                              Response.Write "NOT OK" 
                             End if
                   %>&nbsp;</td>
  </tr>
  <tr>
    <td width="81%">&nbsp;Check : Coupon Rate &lt; 30%</td>
    <td width="19%" align="right"><% if inValid7 = 0  and not inValid7 = "" then 
                              Response.Write "OK"
                             elseif inValid7 = 1 then
                              Response.Write "NOT OK" 
                             End if
                   %>&nbsp;</td>
  </tr>
  <tr>
    <td width="81%">&nbsp;Check : Forward Rate > 0</td>
    <td width="19%" align="right"><% if inValid8 = 0  and not inValid8 = "" then 
                              Response.Write "OK"
                             elseif inValid8 = 1 then
                              Response.Write "NOT OK" 
                             End if
                   %>&nbsp;</td>
  </tr>
  <tr>
    <td width="81%">&nbsp;Check : Forward Rate &lt; 30%</td>
    <td width="19%" align="right"><% if inValid9 = 0  and not inValid9 = "" then 
                              Response.Write "OK"
                             elseif inValid9 = 1 then
                              Response.Write "NOT OK" 
                             End if
                   %>&nbsp;</td>
  </tr>
  <tr>
    <td width="81%">&nbsp;Check : No Of Coupon Payments Left &lt; 50</td>
    <td width="19%" align="right"><% if inValid10 = 0  and not inValid10 = "" then 
                              Response.Write "OK"
                             elseif inValid10 = 1 then
                              Response.Write "NOT OK" 
                             End if
                   %>&nbsp;</td>
  </tr>
  
</table>
</form>
  <!--#include file="../pageLoadingBottom.asp"-->
</body>

</html>