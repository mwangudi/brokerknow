<html>
<!--#include file="../libroutines.asp"-->
<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Summary Batch Download Batch Applications</title>
	<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
	<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
	<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
	<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
	<SCRIPT language=VBScript src="../scripts/reports.vbs"></SCRIPT>
	<SCRIPT language=Javascript src="../scripts/reports.js"></SCRIPT>

	<script language="javascript">
	 function isNumeric(val){return(parseFloat(val,10)==(val*1));}

	 function validateNumeric(ctrl){
	  var value = ctrl.value;
	  
	  if (value!=''){
		if (isNumeric(value)){}else{ctrl.value=''}
	  }
	 }
	</script>

	<style media="print">
		@page {
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
			
			br.newpage{
				page-break-before:always;
			}		
			
		}

	</style>
</head>

<body Class="Reports">

<% DrawPageFunctions True, True, True %>
<Script Language="JavaScript">
		report_SetBodyClass();
		function validateForm(frm){			
			if (frm.FromNo.value==''){
			    alert("Please specify the from batch number.");
				frm.FromNo.focus();
				return;
			}
			
			frm.target = '_self';			
			frm.submit();
		}
		
	</Script>
	<form method="POST" action="DownloadBatchApplication.asp" Name="frmMain" id="frmMain">
	
    <table border="0" cellspacing="0" cellpadding="0" style="font-family: Arial Narrow" width="100%">
	<tr>
		<td width="100%" align="middle"><img src="../data/photos/aaprintlogo.jpg" border="0" width="482" height="178"></td>
	</tr>
	<br>
	<tr>
		<td width="100%" nowrap align="left"><font face="Impact" size="4">BATCHED APPLICATIONS</font></td>      
	</tr>
	<tr>
		<td width="100%" nowrap height="0" align="right">&nbsp;</td>      
	</tr>
</table>
<br>

	<%
	 Dim sqlstr, rs, rst, conn

     Set Conn = Server.CreateObject("ADODB.Connection")
     Set rs = Server.CreateObject("ADODB.Recordset")
	 Set rst = Server.CreateObject("ADODB.Recordset")
     Set Conn = GetActiveConnection("KBroker")

	 sqlstr = "SELECT Batch_No as BatchNo, SUM(Alloted_Rights) AS TotalQty,  " & _
		 "     SUM(ISNULL(Alloted_Rights, 0) * ISNULL(Offering_Price, 0)) AS TotalAmt, COUNT(Offering_DPA_)  " & _
		 "     AS AppCount " & _
		 " FROM  OfferingsList " & _
		 " where (NOT (Batch_No IS NULL)) AND LastDownloaded = 1 "  & _
		 " GROUP BY Batch_No "
    
     set rs = conn.execute(sqlstr)

	 intrscount =  rs.recordcount

	 if intrscount <= 0 then
		 %>
		 <script language="javascript">
		  alert("No batch application has been downloaded.")
		   document.location.href = 'DownloadBatchApplication.asp';
		 </script>
		 <%
		  set rs = nothing
		  set conn = nothing
		  response.end
	 else
       rs.movefirst
       batchData =  rs.getrows
	 end if

   '==========================================================================================
   'Prepare cover page
    '=========================================================================================
	%>
	<br>
	<table border="0" cellspacing="2" cellpadding="2" style="font-family: Arial Narrow" width="100%">
		 <tr>
			 <td style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;"><b>BATCH</b></td>
			 <td style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;"><b>PAYMENT</b></td>
			 <td align="right" style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;"><b>APPLICATIONS</b></td>
			 <td align="right" style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;"><b>TOTAL SHARES</b></td>
			 <td align="right" style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;"><b>TOTAL AMOUNT</b></td>
		 </tr>
	
	<%
	TotalAppNo = 0
	TotalShares = 0
    TotalAmount = 0

   for intcount=0 to intrscount-1
     BatchID = trim(batchData(0,intCount))
     AppNo = trim(batchData(3,intCount))
     TotalQty = trim(batchData(1,intCount))
     TotalAmt = trim(batchData(2,intCount))
     
	 %>
	 <tr>
		  <td ><%=BatchID%></td>
		  <td >Agent Cheque</td>
		  <td align="right">&nbsp;<%=AppNo%></td>
		  <td align="right">&nbsp;<%=FormatNumCommasOnly(TotalQty)%></td>
		  <td align="right">&nbsp;<%=FormatNum(TotalAmt)%></td>
	 </tr>
	 <%
     TotalAppNo = TotalAppNo + AppNo
	 TotalShares = TotalShares + TotalQty
     TotalAmount = TotalAmount + TotalAmt
   next
   %>
  
	<tr>
		  <td colspan="2"><b>TOTALS</b></td>
		  <td align="right">&nbsp;<b><%=TotalAppNo%></b></td>
		  <td align="right">&nbsp;<b><%=FormatNumCommasOnly(TotalShares)%></b></td>
		  <td align="right">&nbsp;<b><%=FormatNum(TotalAmount)%></<b></td>
	 </tr>
   </TABLE>
   <br><br>
  <table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
	 <tr>
		 <td width="5%" height = "100">&nbsp;</td>
		 <td height = "100" style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;">&nbsp;</td>
		 <td width="5%">&nbsp;</td>
		 <td height = "100" style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;">&nbsp;</td>
		 <td height = "100" width="5%">&nbsp;</td>
	 </tr>
	 <tr>
		 <td width="10%">&nbsp;</td>
		 <td width="35%" align="center">&nbsp;Agent stamp and signature</td>
		 <td width="10%">&nbsp;</td>
		 <td width="35%" align="center">&nbsp;Receiving bank and signature</td>
		 <td width="10%">&nbsp;</td>
	 </tr>
  </table>
   <BR class="newpage">
   <%
    for intcount = 0 to intrscount-1
     BatchID = trim(batchData(0,intCount))
     
	  sqlStr = "SELECT Offerings_Date, Pal_No,Client_DPA_, ClientName, Alloted_Rights, Offering_Price,SecurityName,BatchFileName FROM OfferingsList where Batch_No = " & BatchID & _
	         " order by cast(floor(cast(Offerings_Date as float)) as datetime), Rtrim(ltrim(ClientName)) "
      
	   set rst = conn.execute(sqlstr)

	   intrstcount = rst.recordcount

	   if intrstcount > 0 then
	   %>
	   <table border="0" cellspacing="0" cellpadding="0" style="font-family: Arial Narrow" width="100%">
	<tr>
		<td width="100%" align="middle"><img src="../data/photos/aaprintlogo.jpg" border="0" width="482" height="178"></td>
	</tr>
	<br>
	<tr>
		<td width="100%" nowrap align="left"><font face="Impact" size="4">BATCHED FORWARDS</font></td>      
	</tr>
	<tr>
		<td width="100%" nowrap height="0" align="right">&nbsp;</td>      
	</tr>
</table>
<br>
	<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
		<tr>
		  <td width="10%"><b>DATE:</b></td>
		  <td width="90%"><%= FormatDate(Date) %></td>
		</tr>   
		<tr>
		  <td width="10%"><b>OFFER</b></td>
		  <td width="90%"><font face="Arial Narrow" size="3"><%= ucase(trim(rst("SecurityName"))) %></font></td>
		</tr>
		<tr>
		  <td width="10%"><font face="Arial Narrow" size="2"><b>PAYMENT</b></font></td>
		  <td width="90%"><font face="Arial Narrow" size="2">TO BE PAID BY AGENT</font></td>
		</tr>
		<tr>
		  <td width="10%"><font face="Arial Narrow" size="2"><b>BATCH NO</b></font></td>
		  <td width="90%"><font face="Arial Narrow" size="2"><%=BatchID%></font></td>
		</tr>
		<tr>
		  <td width="10%"><font face="Arial Narrow" size="2"><b>FILE NAME</b></font></td>
		  <td width="90%"><font face="Arial Narrow" size="2"><%=trim(rst("BatchFileName"))%></font></td>
		</tr>
	</table>
	<BR>

	 <table border="0" cellspacing="3" cellpadding="0" style="font-family: Arial Narrow; LEFT-MARGIN:100PX"  width="100%">
    <tr>
	  <td><b><font face="Arial Narrow" size="2">#</font></b></td>
      <td style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;"><b><font face="Arial Narrow" size="2">Date</font></b></td>
      <td style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;"  nowrap><b><font face="Arial Narrow" size="2">Serial No</font></b></td>
      <td style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;"><b><font face="Arial Narrow" size="2">Code</font></b></td>
      <td style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;"  nowrap><b><font face="Arial Narrow" size="2">Client Name</font></b></td>
      <td align="right" style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;"  align=right nowrap><b><font face="Arial Narrow" size="2">Quantity</font></b></td>
	  <td align="right" style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;"  nowrap><b><font face="Arial Narrow" size="2">Price</font></b></td>
	  <td align="right" style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;"  nowrap><b><font face="Arial Narrow" size="2">Payable</font></b></td>
    </tr>
	   <%
         astrRecs = rst.getrows()
          
		  totalquantity = 0
		  totalpayable = 0

		 for rstcount = 0 to intrstcount- 1 
		        Price = trim(astrRecs(5,rstcount))
				Qty = trim(astrRecs(4,rstcount))
				%>
				<tr>
					<td><b><%=rstcount + 1%></b></td>
					<td nowrap><%=FormatDate(trim(astrRecs(0,rstcount)))%></td>
					<td><%=trim(astrRecs(1,rstcount))%></td>
					<td><%=trim(astrRecs(2,rstcount))%></td>
					<td nowrap><%=Mid(trim(astrRecs(3,rstcount)),1,30)%></td>		
					<td align="right"><%=FormatNumCommasOnly(Qty)%></td>
					<td align="right"><%=FormatNum(Price)%></td>
					<td align="right"><%=FormatNum(Price*Qty)%></td>
				</tr>
				<%
				totalquantity = totalquantity + Qty
				totalpayable = totalpayable + (Price*Qty)
		 next

	   end if
      %>
	    <tr>
			<td colspan="5" align="center"><b>Totals</b></td>
			<td align="right"><b><%=FormatNumCommasOnly(totalquantity)%></b></td>
			<td align="right">&nbsp;</td>
			<td align="right"><b><%=FormatNum(totalpayable)%></b></td>
        </tr>
      </table>
	  <br><br>
	  <table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
	 <tr>
		 <td width="5%" height = "100">&nbsp;</td>
		 <td height = "100" style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;">&nbsp;</td>
		 <td width="5%">&nbsp;</td>
		 <td height = "100" style="border-right-style: solid; border-right-width: 1; border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1;border-color: black;">&nbsp;</td>
		 <td height = "100" width="5%">&nbsp;</td>
	 </tr>
	 <tr>
		 <td width="10%">&nbsp;</td>
		 <td width="35%" align="center">&nbsp;Agent stamp and signature</td>
		 <td width="10%">&nbsp;</td>
		 <td width="35%" align="center">&nbsp;Receiving bank and signature</td>
		 <td width="10%">&nbsp;</td>
	 </tr>
  </table> 
	  <%
	next
	%>
	</form>
	
</body>
</html>
