<!--#include file="../libroutinesTEST.asp"-->
<%
ID = Request.QueryString("ID")

'ID = 2

Set Conn = GetActiveConnection("KBroker")

sqlStr = "SELECT * FROM PaymentRequestList WHERE PaymentRequest_DPA_=" & ID
Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))

if (Rs.bof or Rs.eof) then
	%>
	<script language = 'vbscript'>
   		ShowMessage "Invalid data sent. Please try again"
	</script>
	<%
	Response.End
end if

function stripFormatting(theVal, delimiter)
	Dim i
	Dim tempVal, char
	  
	for i = 1 to len(theVal)
		char = mid(theVal,i,1)
		if isnumeric(char) OR  char = delimiter then tempVal = tempVal & char
	next 
	  
	stripFormatting = tempVal
end function
 
function ConvertTowords(theVal, delimiter)
	Dim tempVal, strVal,mainpart, decimalpart
	Dim i,d
	theVal = trim(theVal)
	 
	'******************* Validate parameter **************
	 
	if not isnumeric(theVal) or theVal = "" then
	 ConvertTowords = theVal
	 exit function
	end if  
	 
	if delimiter = "" then delimiter = "." 'default delimiter: decimal
	 
	'There should only be one delimiter 
	d=0
	for i = 1 to len(theVal)
	 char = mid(theVal,i,1)
	 if char = delimiter then d = d + 1
	next
	 
	if d > 1 then
	 ConvertTowords = theVal
	 exit function
	end if 
	 
	theVal = stripFormatting(theVal, delimiter)
  
	pos = 0
	pos = instr(1,theVal, delimiter,1)
	 
	if pos <> 0 then ' Separate main and decimal parts of the number
	 mainpart = left(theVal,pos-1)
	 decimalpart = right(theVal,len(theVal)- pos)
	else
	 mainpart = theVal
	 decimalpart = ""
	end if
	 
	strVal = translate(mainpart,0)
	  
	if strVal <> "" then  
	  if trim(strVal) = "ONE" then 
	   strVal = strVal & " SHILLING "
	  else
	   strVal = strVal & " SHILLINGS "
	  end if
	end if
	 
	decimalpart = left(decimalpart,2) 'Default decimal Places: 2
	 
	if decimalpart <> "" then
	 strVal2 = translate(decimalpart,1)
 
	 	if strVal2 <> "" AND strVal <> "" then 
	 	  strVal = strval & " AND " & strVal2 & " CENTS "
	 	elseif strVal2 <> "" then 
	 	  strVal = strval & " " & strVal2 & " CENTS "
	 	end if
	end if
	 
	ConvertTowords = strVal
end function
 
function translate(tempNum, deci)
	 Dim k,convnum,strNum
	 Dim noElements, noCategory
	  
	if deci = "" then deci = 0 'default
	strNum = "" 'initialise empty string to hold amount in words
	convnum  = int(tempNum) 'Remove leading zeros
	noElements = len(tempNum)
	 
	extra = 0
	if noElements > 3 then
	 if (Abs(noElements / 3) - int(noElements / 3)) > 0 then extra = 1
	 noCategory = int(noElements/3) + extra
	else
	 noCategory=1
	end if
	 
	 'Translate given number
	    
	    for k = 1 to  noCategory
	    
	      if len(convnum) >= 3 then
	       num = right(convnum,3) 'fetch number to be translated 
	       convnum = left(convnum,len(convnum)-3) 'strip number from original  
	      else
	       num = convnum
	       convnum = ""
	      end if
	 
	      strEquivalent = wordEquivalent(num,deci)'Get word equivalent
	       
	      if strNum <> "" then comma = "," 'Format output
				
	      if k = 1 then
	       strNum = strEquivalent & strNum
	      elseif k = 2 AND strEquivalent <> "" then
	          if strNum <> "" then
	  		 if mid(tempNum,len(tempNum)-2,1) = 0 then 
	  		  comma = " AND " 'Format output
	  		 else
	  		  comma = "," 'Format output
	  		 end if
	  	    end if
	       strNum = strEquivalent & " THOUSAND" & comma & " " & strNum 
	      elseif k = 3 AND strEquivalent <> "" then
	       strNum = strEquivalent & " MILLION" & comma & " " & strNum
	      elseif k = 4 AND strEquivalent <> "" then
	       strNum = strEquivalent & " BILLION" & comma & " " & strNum
	      elseif k = 5 AND strEquivalent <> "" then
	       strNum = strEquivalent & " TRILLION" & comma & " " & strNum
	      end if  
 
	   next
	 
	translate = strNum
end function
 
function wordEquivalent(Num,deci)
	Dim tempNum, strNum, lenNum
	 
	if deci = "" then deci =0
	 
	'**** Validate *****
	if Num = "" then
	 wordEquivalent = ""
	 exit function
	end if
	 
	lenNum = len(Num)
	 
	'Force leading zeros as appropriate
	if deci = 0 then     ' Main number part  
	  if lenNum = 1 then
	   num = "00" & num
	  elseif lenNum = 2 then
	   num = "0" & num
	  end if
	elseif deci =1 then ' Decimal part
	  if lenNum = 1 then
	   num = "0" & num & "0"
	  elseif lenNum = 2 then
	   num = "0" & num
	  end if

	end if
	 
	'wordEquivalent = num
	'exit function        
	strNum = "" 'initialise empty string to hold amount in words
	isNonZero = left(Num,1) <> 0 'first number not a zero
	isNotOne = mid(Num,2,1) <> 1 'Second number not a one
 
	 for i = 1 to 3
	  
	   tempNum = mid(Num, i, 1)
	    
	    select case tempNum
	       case "0"
	          if i = 2  OR i = 3 then strNum = strNum & ""
	       case 1
	          if i = 1 then
	           strNum = " ONE HUNDRED "
	          elseif i = 2  then
	             if isNonZero then strNum = strNum & " AND "
	             select case right(Num,1)
	  				case 0 
	  				 strNum = strNum & " TEN "
	  				case 1
	  				  strNum = strNum & " ELEVEN "
	  				case 2
	  				  strNum = strNum & " TWELVE "
	  				case 3
	  				  strNum = strNum & " THIRTEEN "
	  				case 4
	  				  strNum = strNum & " FOURTEEN "
	  				case 5
	  				  strNum = strNum & " FIFTEEN "
	  				case 6
	  				  strNum = strNum & " SIXTEEN "
	  				case 7
	  				  strNum = strNum & " SEVENTEEN "
	  				case 8
	  				  strNum = strNum & " EIGHTEEN "
	  				case 9
	  				  strNum = strNum & " NINETEEN "
	             end select 
	          elseif i = 3 AND isNotOne then
	           if mid(Num,2,1) = 0 AND isNonZero then  strNum = strNum & " AND "
	           strNum = strNum &  " ONE "
	          end if 
	       case 2
	          if i = 1 then
	           strNum = " TWO HUNDRED "
	          elseif i = 2  then
	           if isNonZero then strNum = strNum & " AND "
	           strNum = strNum  & " TWENTY "
	          elseif i = 3 AND isNotOne then
	           if mid(Num,2,1) = 0 AND isNonZero then  strNum = strNum & " AND "
	           strNum = strNum & " TWO "
	          end if 
	       case 3
	          if i = 1 then
	           strNum = " THREE HUNDRED "
	          elseif i = 2  then
	           if isNonZero then strNum = strNum & " AND "
	           strNum = strNum  & " THIRTY "
	          elseif i = 3 AND isNotOne then
	           if mid(Num,2,1) = 0 AND isNonZero then  strNum = strNum & " AND "
	           strNum = strNum & " THREE "
	          end if 
	       case 4
	          if i = 1 then
	           strNum = " FOUR HUNDRED "
	          elseif i = 2  then
	           if isNonZero then strNum = strNum & " AND "
	           strNum = strNum  & " FOURTY "
	          elseif i = 3 AND isNotOne then
	           if mid(Num,2,1) = 0 AND isNonZero then  strNum = strNum & " AND "
	           strNum = strNum & " FOUR "
	          end if 
	       case 5
	          if i = 1 then
	           strNum = " FIVE HUNDRED "
	          elseif i = 2  then
	           if isNonZero then strNum = strNum & " AND "
	           strNum = strNum  & " FIFTY "
	          elseif i = 3 AND isNotOne then
	           if mid(Num,2,1) = 0 AND isNonZero then  strNum = strNum & " AND "
	           strNum = strNum & " FIVE "
	          end if 
	       case 6
	          if i = 1 then
	           strNum = " SIX HUNDRED "
	          elseif i = 2  then
	           if isNonZero then strNum = strNum & " AND "
	           strNum = strNum  & " SIXTY "
	          elseif i = 3 AND isNotOne then
	           if mid(Num,2,1) = 0 AND isNonZero then  strNum = strNum & " AND "
	           strNum = strNum & " SIX "
	          end if 
	       case 7
	          if i = 1 then
	           strNum = " SEVEN HUNDRED "
	          elseif i = 2  then
	           if isNonZero then strNum = strNum & " AND "
	           strNum = strNum  & " SEVENTY "
	          elseif i = 3 AND isNotOne then
	           if mid(Num,2,1) = 0 AND isNonZero then  strNum = strNum & " AND "
	           strNum = strNum & " SEVEN "
	          end if
	       case 8
	          if i = 1 then
	           strNum = " EIGHT HUNDRED "
	          elseif i = 2  then
	           if isNonZero then strNum = strNum & " AND "
	           strNum = strNum  & " EIGHTY "
	          elseif i = 3 AND isNotOne then
	           if mid(Num,2,1) = 0 AND isNonZero then  strNum = strNum & " AND "
	           strNum = strNum & " EIGHT "
	          end if
	       case 9
	          if i = 1 then
	           strNum = " NINE HUNDRED "
	          elseif i = 2  then
	           if isNonZero then strNum = strNum & " AND "
	           strNum = strNum  & " NINETY "
	          elseif i = 3 AND isNotOne then
	           if mid(Num,2,1) = 0 AND isNonZero then  strNum = strNum & " AND "
	           strNum = strNum & " NINE "
	          end if
	    end select
	   
	 next
	 wordEquivalent = strNum
End function
%>
<html>

<head>
<title>Payment Request Form</title>
<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<SCRIPT language=VBScript src="../scripts/reports.vbs"></SCRIPT>
 <SCRIPT language=Javascript src="../scripts/reports.js"></SCRIPT>
<link rel="stylesheet" type="text/css" href="../Operations/CALENDAR/calendar.css">
<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">

<style media="print">
	@page 
		{
		@top{font-family: Helvetica, Arial, sans-serif;
			font-size: 150%;
			font-weight: bolder;
			text-align: left;
			content: "<%= FormatDate(Date) %>";			
		}
			
		margin-left: 0cm;
		margin-right: 0cm;
		margin-top: 0cm;    
		margin-bottom: 0cm;
		size: portrait;
			
		}
	}
</style>
<style type="text/css">
	td {font-family: Trebuchet MS;}
</style>
</head>

<body>
<%
DrawPageFunctions2 true, true, true, "EditPaymentRequest.asp?ID=" & ID, True
%>
<p id="toPDFOrient" name="toPDFOrient" value="P" style="display:none;">P
<p id="toPDF" name="toPDF">

<TABLE cellSpacing=3 cellPadding=2 border=0 width="80%" align="center">
	<TR>	  
	    <TD colspan=5 width="100%" align="center"><img src="../data/photos/aaprintlogo.jpg" border="0"><TD>
    </TR>
	
	<tr><td colspan=5>&nbsp;</td></tr>
	
	<TR>	  
		<td style="border:1 solid black;" colspan=5 align="center" height=30><b>TRADE SETTLEMENT PAYMENT REQUISITION</b></td>
    </TR>
	
	<tr><td colspan=5>&nbsp;</td></tr>
	
	<TR>
		<td nowrap width="5%" align=left><b>DATE:</b></td>
		<td nowrap style="border: 1px solid #000000;" width="35%" height=25><%=formatDate(rs("PaymentPDate"))%></td>
		<td nowrap width="20%" height=25>&nbsp;</td>
		<td nowrap width="5%" align=right><b>OUR REFERENCE:</b></td>
		<td nowrap style="border: 1px solid #000000;" width="35%" height=25>&nbsp;</td>
	</TR>
	
	<tr><td colspan=5 height=10>&nbsp;</td></tr>
	
	<tr>
		<td nowrap width="5%"><b>ISSUE CHEQUE IN FAVOUR OF:</b>&nbsp;&nbsp;</td>
		<td align=center nowrap width="95%" colspan="4" style="border: 1 solid black;" height="30"><%=rs.Fields("ClientName")%>&nbsp;</td>
	<tr>
	
	<tr><td colspan=5 height=10>&nbsp;</td></tr>
	
	<tr>
		<td nowrap width="5%"><b>AMOUNT IN WORDS:</b>&nbsp;&nbsp;</td>
		<td align=center nowrap width="95%" colspan="4" style="border: 1 solid black;" height="80"><%=ConvertTowords(rs("PaymentAmount"),"")%></td>
	<tr>
	
	<tr><td colspan=5 height=10>&nbsp;</td></tr>
	
	<tr>
		<td nowrap width="5%"><b>AMOUNT IN FIGURES:</b>&nbsp;&nbsp;</td>
		<td nowrap width="60%" colspan="3">&nbsp;</td>
		<td align=center nowrap style="border: 1px solid #000000;" width="35%" height=25><%=FormatNum(rs("PaymentAmount"))%></td>
	<tr>
	
	<tr><td colspan=5 height=10>&nbsp;</td></tr>
	
	<tr>
		<td nowrap width="5%"><b>CURRENCY:</b>&nbsp;&nbsp;</td>
		<td nowrap width="60%" colspan="3">&nbsp;</td>
		<td align=center nowrap style="border: 1px solid #000000;" width="35%" height=25><b>USHS</b></td>
	<tr>
	
	<tr><td colspan=5 height=10>&nbsp;</td></tr>
	
	<tr>
		<td nowrap width="5%"><b>CHEQUE NUMBER:</b>&nbsp;&nbsp;</td>
		<td nowrap width="60%" colspan="3">&nbsp;</td>
		<td align=center nowrap style="border: 1px solid #000000;" width="35%" height=25><%=rs("PaymentReference")%></td>
	<tr>
	
	<tr><td colspan=5 height=10>&nbsp;</td></tr>
	
	<tr>
		<td nowrap width="5%"><b>DEBIT A/C NUMBER:</b>&nbsp;&nbsp;</td>
		<td nowrap width="60%" colspan="3">&nbsp;</td>
		<%
		If Len(rs("BankAccountName"))>0 Then
			bnk = split(rs("BankAccountName")," ")
			BankAccountName = bnk(Ubound(bnk))
		End If
		%>
		<td align=center nowrap style="border: 1px solid #000000;" width="35%" height=25><%=BankAccountName%></td>
	<tr>
	
	<tr><td colspan=5 height=10>&nbsp;</td></tr>
	
	<tr>
		<td nowrap width="5%"><b>PREPARED BY:</b>&nbsp;&nbsp;</td>
		<td align=center nowrap width="95%" colspan="4" style="border: 1 solid black;" height="25"><%=rs("ChangedBy")%></td>
	<tr>
	
	<tr><td colspan=5>&nbsp;</td></tr>
	
	<tr>
		<td colspan=5>
			<table cellpadding="1" cellspacing="1" border="0" width="100%">
				<tr>
				<td width="10%" nowrap><b>APPROVED BY:</b>&nbsp;&nbsp;</td>
				<td align=right width="36%" nowrap style="border: 1 solid black;" height=25>&nbsp;</td>
				<td width="8%" nowrap>&nbsp;</td>
				<td width="10%" nowrap><b>APPROVED BY:</b>&nbsp;&nbsp;</td>
				<td width="36%" nowrap style="border: 1 solid black;" height=25>&nbsp;</td>
				</tr>
			</table>
		</td>
	</tr>
	
	<tr><td colspan=5>&nbsp;</td></tr>
	
	<tr>
		<td colspan=5>
			<table cellpadding="1" cellspacing="1" border="0" width="100%">
				<tr>
				<td width="10%" nowrap><b>SIGNATURE:</b>&nbsp;&nbsp;</td>
				<td align=right width="36%" nowrap style="border: 1 solid black;" height=50>&nbsp;</td>
				<td width="8%" nowrap>&nbsp;</td>
				<td width="10%" nowrap><b>SIGNATURE:</b>&nbsp;&nbsp;</td>
				<td width="36%" nowrap style="border: 1 solid black;" height=50>&nbsp;</td>
				</tr>
			</table>
		</td>
	</tr>
	
	<tr><td colspan=5>&nbsp;</td></tr>
	
	<tr>
		<td colspan=5>
			<table cellpadding="1" cellspacing="1" border="0" width="100%">
				<tr>
				<td width="10%" nowrap><b>DATE:</b>&nbsp;&nbsp;</td>
				<td align=right width="36%" nowrap style="border: 1 solid black;" height=25>&nbsp;</td>
				<td width="8%" nowrap>&nbsp;</td>
				<td width="10%" nowrap><b>DATE:</b>&nbsp;&nbsp;</td>
				<td width="36%" nowrap style="border: 1 solid black;" height=25>&nbsp;</td>
				</tr>
			</table>
		</td>
	</tr>
</TABLE>  
</body>

</html>