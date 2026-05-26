<!--#include file="../libroutines.asp"-->
<%

Const report_ViewName = "ReceiptForm"
Const reportPage = "ReceiptForm.asp"
Const headerColCount = 3
Const groupingHeaderCol = 0
Const reportTitle = "Receipt Form"

ReceiptID = Request.QueryString("ID")
Set Conn = GetActiveConnection("KBroker")

sqlStr = "SELECT * FROM ReceiptList WHERE (Payment_DPA_ = " & ReceiptID & ")"
' WHERE (Payment_DPA_ = " & ReceiptID & ")"

Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))

if (Rs.bof or Rs.eof) then
	%>
	<script language = 'vbscript'>
    		ShowMessage "Invalid data sent. Please try again"
	</script>
	<%
	Response.End

end if

'Build In Payment Of field information
 if (rs("EntityType_DPA_") = 3) then
    InPaymentOf = " Broker Receipt "
 elseif (rs("EntityType_DPA_") = 1 AND isnull(rs("Order_DPA_"))) then
       InPaymentOf = "Client Receipt"
 elseif (rs("EntityType_DPA_") = 1 AND trim(rs("Order_DPA_")) <> "") then
       InPaymentOf = "Payment for Order No  " & rs("Order_DPA_")
       
       if (rs("OrderSecType_DPA_")=2) then
        InPaymentOf2 = "SHARES"
       elseif (rs("OrderSecType_DPA_")=1) then
        InPaymentOf2 = "BONDS"
       end if
       
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
<meta http-equiv="Content-Language" content="en-us">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Reciept Form</title>
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

</head>
<%
'br.newpage{page-break-before:always;
%>
<body>
<%
DrawPageFunctions2 true, true, true, "EditReceipt.asp?ID=" &  ReceiptID
%>
<div align="right">
<TABLE cellSpacing=0 cellPadding=0 border=0 width="100%" bordercolor="#111111"  style="border-collapse: collapse">
  <TBODY 
  style="border-left-color: black; border-bottom-color: black; border-top-color: black; border-right-color: black">
	<TR>	  
	    <TD colspan="4" width="100%" align="center" valign="top">
			<!--#include file="../Reports/Header.asp"-->
		<TD>
    </TR>
	<tr><td colspan="4">&nbsp;</td></tr>
	<TR>	  
		<td colspan="4" align="center"><b><font size ="4">RECEIPT</font></b></td>
    </TR>

	<TR>
		<td colspan="4">
			<table cellpadding="0" cellspacing="0" border="0" width="35%" bordercolor="#111111"  style="border-collapse: collapse" align="right">
			<tr>
			<td align="left" style="border: 1px solid #000000" valign="top" width="50%">&nbsp;Receipt Number:</td>
			<td align="right" style="border: 1px solid #000000; border-right-color: 1px solid #000000;" valign="top" width="50%" ><%=rs("PaymentReceiptNo")%>&nbsp;</td></tr>
			<tr><td colspan="2">&nbsp;&nbsp;</td></tr>
			<tr>
			<td align="left" style="border: 1px solid #000000" valign="top" width="50%">&nbsp;Date:</td>
			<td align="right" style="border: 1px solid #000000; border-right-color: 1px solid #000000;" valign="top" width="50%"><%=formatDate(rs("PaymentPDate"))%>&nbsp;</td></tr>
			</table>
			</td>
	  </TR>
	<TR>	
	<%
		set temprs= server.createobject("Adodb.recordset")
		sqlStr = "SELECT Client.ClientAddr,Client.ClientContact,'W' + ClientOfficeTel + '/' + Client.ClientEmail as Contacts,ClientName FROM Client " & _			
			" WHERE Client_DPA_ = " & rs("EntityCode")
		'Response.Write sqlstr
		'Response.end
	
	Set temprs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
	
	Dim clientID
	If Not (temprs.EOF Or temprs.BOF) Then		
		'clientAddress = Replace(temprs.Fields("ClientAddr").Value, Chr(13), ",")
		clientAddress = temprs.Fields("ClientAddr").Value
		clientID = rs("EntityCode")
		ClientContact=temprs("ClientContact")
		Contacts=temprs("Contacts")
		ClientName=temprs("ClientName")
	End If
	
	Set temprs = Nothing
	%>
	<tr><td colspan="4">&nbsp;</td></tr>
	<tr><td colspan="4">
		<table cellpadding="0" cellspacing="0" border="1" width="100%" bordercolor="#111111"  style="border-collapse: collapse" align="left">
		<tr><td>
		<table cellpadding="0" cellspacing="0" border="0" width="100%" bordercolor="#111111"  style="border-collapse: collapse" align="left">
		<tr><td  width="5%">&nbsp;</td><td>&nbsp;</td></tr>
		<tr><td>&nbsp;</td><td>&nbsp;[<%=clientID%>]&nbsp;<%=ClientName%></td></tr>
		<tr><td width="5%">&nbsp;</td><td>&nbsp;<%=clientAddress%></td></tr>	
		<tr><td>&nbsp;</td><td>&nbsp;</td></tr>
		</table>
		</td></tr>
		</table>
	</td><tr>
	<tr><td colspan="4">&nbsp;</td></tr>
	<tr><td colspan="4">
		<table cellpadding="0" cellspacing="0" border="1" width="100%" bordercolor="#111111"  style="border-collapse: collapse" align="left">
			<tr><td width="15%">&nbsp;Mode:</td><td>&nbsp;<%= rs("PaymentType")%></td></tr>	
		</table>
	</td></tr>
	<tr><td colspan="4">&nbsp;</td></tr>

	<tr><td colspan="4">
		<table cellpadding="0" cellspacing="0" border="0" width="100%" bordercolor="#111111"  style="border-collapse: collapse" align="left">
		<tr><td colspan="4">
		
		<table cellpadding="0" cellspacing="0" border="1" width="100%" bordercolor="#111111"  style="border-collapse: collapse" align="left" >
		<tr><td  width="20%" valign="bottom">&nbsp;<br>&nbsp;<b>Cheque Number</b></td>
		<td width="60%" valign="bottom" align="center"> &nbsp;<b>Description</b></td>
		<td  width="20%" valign="bottom" align="center">&nbsp;<b>Detailed Amount</b></td>
		</tr>

		<tr height="200"><td  width="20%" valign="top">&nbsp;<br>&nbsp;<%=rs("PaymentReference")%></td>
		<td width="60%" valign="top" align="left">&nbsp;<br>&nbsp;<%=rs("PaymentNarrative")%></td>
		<td  width="20%" valign="top" align="right">&nbsp;<br><%=formatNumber(rs("PaymentAmount"))%>&nbsp;</td>
		</tr>

		
		</table>
		</td></tr>
		
		</table>
	</td><tr>
	<tr><td colspan="4">
		
	<table cellpadding="0" cellspacing="0" border="0" width="100%" bordercolor="#111111"  style="border-collapse: collapse" align="left" >
		<tr height="20" style="border: 0px solid #FFFFFF;" ><td  width="20%" >&nbsp;</td>
		<td width="60%"  align="right"><b>Total&nbsp;</b></td>
		<td align=right style="border-style: solid; border-color: #000000; border-width: 1" height="30px"><%= FormatNum(rs("PaymentAmount")) %></td></td>
		</tr>
	</table>
	</td></tr>

	<tr><td colspan="4">&nbsp;</td></tr>
	<tr><td colspan="4">&nbsp;</td></tr>
	<tr><td colspan="4">&nbsp;</td></tr>
	<tr><td colspan="4">&nbsp;</td></tr>
	<tr><td colspan="4">&nbsp;</td></tr>
	<tr><td colspan="4">&nbsp;&nbsp;For AAMS signed by:______________________________________________</td></tr>
	
</TBODY></TABLE>  
  
</div>
</body>

</html>