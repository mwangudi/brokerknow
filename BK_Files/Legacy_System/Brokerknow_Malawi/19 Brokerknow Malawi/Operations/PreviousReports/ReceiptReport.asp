<!--#include file="../libroutines.asp"-->
<%

Const report_ViewName = "ReceiptForm"
Const reportPage = "ReceiptReport.asp"
Const headerColCount = 3
Const groupingHeaderCol = 0
Const reportTitle = "Receipts"

'ReceiptID = Request.QueryString("ID")
 
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
   
   strVal = translate(mainpart)
    
   if strVal <> "" then  
     if trim(strVal) = "ONE" then 
      strVal = strVal & " SHILLING "
     else
      strVal = strVal & " SHILLINGS "
     end if
   end if
   
   decimalpart = left(decimalpart,2) 'Default decimal Places: 2
   
   if decimalpart <> "" then
    strVal2 = translate(decimalpart)
 
		if strVal2 <> "" AND strVal <> "" then 
		  strVal = strval & " AND " & strVal2 & " CENTS "
		elseif strVal2 <> "" then 
		  strVal = strval & " " & strVal2 & " CENTS "
		end if
   end if
   
   ConvertTowords = strVal
 end function
 
 function translate(tempNum)
   Dim k,convnum,strNum
   Dim noElements, noCategory
  
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
  
        strEquivalent = wordEquivalent(num)'Get word equivalent
        
        if strNum <> "" then
			 if mid(tempNum,len(tempNum)-K,1) = 0 then 
			  comma = " AND " 'Format output
			 else
			  comma = "," 'Format output
			 end if
		end if
			
        if k = 1 then
         strNum = strEquivalent & strNum
        elseif k = 2 AND strEquivalent <> "" then
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
 
 function wordEquivalent(Num)
  Dim tempNum, strNum, lenNum
  
  '**** Validate *****
  if Num = "" then
   wordEquivalent = ""
   exit function
  end if
  
  lenNum = len(Num)
  
  'Force leading zeros as appropriate       
  if lenNum = 1 then
   num = "00" & num
  elseif lenNum = 2 then
   num = "0" & num
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
<title>Receipt Report</title>
<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<SCRIPT language=VBScript src="../scripts/reports.vbs"></SCRIPT>
 <SCRIPT language=Javascript src="../scripts/reports.js"></SCRIPT>
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
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

<body Class="Reports">
<%

genReport = Request.Form("genReport")
selectedContractDate = Request.Form("txtDate")

If genReport <> "1" Or selectedContractDate = "" Then%>
		<Script Language="JavaScript">
		report_SetBodyClass();
		function validateForm(frm){			
			if (frm.txtDate.value==''){
				alert("Select a date");
				frm.txtDate.focus();
				return;
			}
			
			frm.target = '_self';			
			frm.submit();
		}
		
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "txtDate","cmdDate","<%= FormatDate(Date) %>",1);
	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="ReceiptReport.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">	
		<table>
			<tr>
				<td>Select date:</td>
				<td>
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>
			</tr>
			
			<tr>
				<td colspan=2><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... " id=Button1 name=Button1>&nbsp;&nbsp;</td>
			</tr>
		</table>
		
	</form>
	
	<%
	Response.End
End If

%>

<%
DrawPageFunctions True, True, True
%>
<%
Set Conn = GetActiveConnection("KBroker")

sqlStr = "SELECT * FROM ReceiptList WHERE (DAY(PaymentPDate) = DAY('" & selectedContractDate & "'))"

'Response.Write(sqlStr)
'Response.End 

Set rs = conn.Execute(sqlStr)

if (Rs.bof or Rs.eof) then
	%>
	<script language = 'vbscript'>
    		ShowMessage "Invalid data sent. Please try again"
	</script>
	<%
	Response.End

end if

Dim pageNumber	
pageNumber = 0

Do Until Rs.EOF
'Build In Payment Of field information
 if (rs("EntityType_DPA_") = 3) then
    InPaymentOf = "Broker Receipt"
 elseif (rs("EntityType_DPA_") = 1 AND rs("Order_DPA_") = "") then
    
       InPaymentOf = "Client Receipt"
 elseif (rs("EntityType_DPA_") = 1 AND rs("Order_DPA_") <> "") then
       InPaymentOf = "Payment for Order No  " & rs("Order_DPA_")
       
       if (rs("OrderSecType_DPA_")=2) then
        InPaymentOf2 = "SHARES"
       elseif (rs("OrderSecType_DPA_")=1) then
        InPaymentOf2 = "BONDS"
       end if
       
 end if
 pageNumber = pageNumber + 1
%>
<TABLE cellSpacing=0 cellPadding=0 width="100%" border=0 bordercolor="#111111">
<THEAD>	
	<!--
	<tr>
		<td align="right" height="200" style="border-bottom: 2px inset #000000">
			<Img Src="../data/photos/image002.gif">			
		</td>		
	</tr>
	<IFRAME FRAMEBORDER=0 marginwidth="0" marginheight="0" NAME="detail" SCROLLING=yes SRC="<%=DataSource%>Item.asp?ID=<%=ID%>&action=<%=action%>" width="870px" height="330px"></IFRAME>
	-->
	<tr class="pageNumbering">
		<td align="left" height="18">
			<FONT FACE=ARIAL SIZE=2><B>Page <%=pageNumber%>	</B></FONT>	
		</td>		
	</tr>
	<tr>
		<td align="right" height="200">
			<!--<Img Src="../data/photos/image002.gif">-->
			<!--#Include file="Final_Logo.htm"-->			
		</td>		
	</tr>
	<THEAD>
	<TR>
	<td>
	<table cellSpacing="0" cellPadding="0" border="0" width="100%">
	<TR>
    <TD colspan="3" style="border: 1px solid #000000" valign="top">
      <P>&nbsp;Received From
      <br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<STRONG><%=rs("EntityCode")%> &nbsp;&nbsp;
      <%=rs("EntityName")%></STRONG>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br></P></TD>
      <td width="20%">&nbsp;</td>
    </TR>
  <TR>
    <TD colspan="3" style="border: 1px solid #000000" valign="top">
      <P>&nbsp;The sum of shillings <br><STRONG>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br> 
      <%=ConvertTowords(rs("PaymentAmount")   ,".")%></STRONG></P>
      <P>&nbsp;</P></TD>
    <TD valign="bottom">
      <P><STRONG style="TEXT-ALIGN: center" 
      >&nbsp;<FONT size="3">&nbsp;&nbsp;&nbsp;&nbsp;Receipt</font><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
      <%=rs("PaymentReceiptNo")%></STRONG></P></TD></TR>
  <TR>
  <TR>
  <td colspan="4" height="0">&nbsp;</td>
  </TR>
    <TD colspan="3" style="border: 1px solid #000000" valign="top">
      <P>&nbsp;In Payment of&nbsp;&nbsp;<br><STRONG>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      <%=InPaymentOf%>&nbsp;  
      &nbsp;<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
      <%=InPaymentOf2%><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      <%=rs("PaymentNarrative")%></STRONG>   </P>
      <P>&nbsp;</P></TD>
    <TD rowspan="2">
    <table cellSpacing="10" cellPadding="0" border="0">
    <tr>
    <td>
    &nbsp;    
    </td>
    </tr>
    <tr>
    <td style="border: 1px solid #000000" valign="bottom">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br>
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br>
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br>
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br>
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;    
    </tr></td></table></TD></TR>
  <TR>
          <TD width="25%" style="border: 1px solid #000000" valign="top">&nbsp;Shs&nbsp;&nbsp; 
            <STRONG><%=FormatNum(rs("PaymentAmount"))%></STRONG></TD>
          <TD width="25%" style="border: 1px solid #000000" valign="top">&nbsp;Reference<STRONG>&nbsp;&nbsp;&nbsp;&nbsp; 
            <%=rs("PaymentReference")%></STRONG></TD>
          <TD width="25%" style="border: 1px solid #000000" valign="top">&nbsp;Date&nbsp; 
            <STRONG><%
             if rs("PaymentPDate") <> "" then
              Response.Write FormatDate(rs("PaymentPDate"))
             End if
            %></STRONG></TD></TD>
    </TR>
	</table>
	</td>
    </TR>
    </Table>
    <%
    Rs.moveNext
    If Not Rs.EOF Then %>
			<BR class="newpage">
	<%	End If
		
    loop
    %>
</body>

</html>