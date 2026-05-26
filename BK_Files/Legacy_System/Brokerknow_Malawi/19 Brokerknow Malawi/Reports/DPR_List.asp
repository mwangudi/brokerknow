<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Daily Processing Reconciliation</title>
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
	 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
	 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
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

<body >

<Script Language="JavaScript">
	
	function GoBack() {
		window.location.replace("Import.asp");
	}
	
	function ShowUnmatchedTrades() {
		window.location.replace("UnmatchedTradesBody.asp");
	}
	
</Script>

<!--#include file="../libroutinesTEST.asp"-->


<table border=0 cellspacing=5 cellpadding=5 id="StatusRow" style="display:none;">
	<tr>	
		<TH noWrap align=left width="20%" bgColor=khaki><b>&nbsp;Processing next batch ...&nbsp;</b></TH>
		<td>
			&nbsp;
		</td>
	</tr>
</table>



<%

'Begin Listing Page
DrawPageFunctions True, False, False, False
 %>
<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
     <tr>
		<td nowrap>
			<b><font face="Arial Narrow" size="4">
			DAILY PROCESSING RECONCILIATION
			</font></b>
		</td>
		<td nowrap>
			&nbsp;
		</td>
		<td nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>
	</tr>	
       <tr>
		  <td COLSPAN=2><font face="Arial" size="2">&nbsp;</font></td>
	</tr>
</table>			



    <table border="0" width="100%" cellPadding="2" cellSpacing=0>
    <tr bgColor="#000000">
			
	
  </table>
  
	



<SCRIPT Language="JavaScript">
	//Account Statement
	function ShowAccountStatement()	{
		document.getElementById("cmdAccountStatement").style.display="none";
		window.location.replace("DPR.asp?pageTitle=Account%20Statement&page=AccountsStatement.asp");
	}
	//Contract Schedule
	function ShowContractSchedule()	{
		document.getElementById("cmdContractSchedule").style.display="none";
		window.location.replace("DPR.asp?page=DownloadContractschedule.asp");
	}
	//Commission
	function ShowCommission()	{
		document.getElementById("cmdCommission").style.display="none";
		window.location.replace("DPR.asp?page=TradedLeviesCommission.asp");
	}
	//Fine Settlement
	function ShowFinalSettlement()	{
		document.getElementById("cmdFinalSettlement").style.display="none";
		window.location.replace("DPR.asp?page=FinalEquitySettlement.asp");
	}
	//Journal Entries
	function ShowJournalEntries()	{
		document.getElementById("cmdJournalEntries").style.display="none";
		window.location.replace("DPR.asp?page=JournalEntriesListing.asp");
	}
	//Levies Non Custodian
	function ShowLeviesNonCustodian()	{
		document.getElementById("cmdLeviesNonCustodian").style.display="none";
		window.location.replace("DPR.asp?page=TradedLeviesNonCust.asp");
	}
	//Levies Custodian
	function ShowLeviesCustodian()	{
		document.getElementById("cmdLeviesCustodian").style.display="none";
		window.location.replace("DPR.asp?page=TradedLeviesCust.asp");
	}
	//Payment Schedule
	function ShowPaymentSchedule()	{
		document.getElementById("cmdPaymentSchedule").style.display="none";
		window.location.replace("DPR.asp?page=PaymentSchedule.asp");
	}
	//Receipt Schedule
	function ShowReceiptSchedule()	{
		document.getElementById("cmdReceiptSchedule").style.display="none";
		window.location.replace("DPR.asp?page=ReceiptSchedule.asp");
	}
	
</SCRIPT>

<table border=0 cellspacing=5 cellpadding=5>
	<tr>	
		<td>
			<INPUT type=Button  value="Account Statement" name="cmdAccountStatement" ID="cmdAccountStatement" OnClick="JavaScript: ShowAccountStatement();">
		</td>
	</tr>
	<tr>	
		<td>
			<INPUT type=Button  value="Commission" name="cmdCommission" ID="cmdCommission" OnClick="JavaScript: ShowCommission();">
		</td>
	</tr>
	<tr>	
		<td>
			<INPUT type=Button  value="Contract Schedule" name="cmdContractSchedule" ID="cmdContractSchedule" OnClick="JavaScript: ShowContractSchedule();">
		</td>
	</tr>
	<tr>	
		<td>
			<INPUT type=Button  value="Final Settlement (CDS)" name="cmdFinalSettlement" ID="cmdFinalSettlement" OnClick="JavaScript: ShowFinalSettlement();">
		</td>
	</tr>
	<tr>	
		<td>
			<INPUT type=Button width="200"  value="Journal Entries" name="cmdJournalEntries" ID="cmdJournalEntries" OnClick="JavaScript: ShowJournalEntries();">
		</td>
	</tr>
	<tr>	
		<td>
			<INPUT type=Button  value="Levies Non Custodian" name="cmdLeviesNonCustodian" ID="cmdLeviesNonCustodian" OnClick="JavaScript: ShowLeviesNonCustodian();">
		</td>
	</tr>
	<tr>	
		<td>
			<INPUT type=Button  value="Levies Custodian" name="cmdLeviesCustodian" ID="cmdLeviesCustodian" OnClick="JavaScript: ShowLeviesCustodian();">
		</td>
	</tr>
	<tr>	
		<td>
			<INPUT type=Button  value="Payment Schedule" name="cmdPaymentSchedule" ID="cmdPaymentSchedule" OnClick="JavaScript: ShowPaymentSchedule();">
		</td>
	</tr>
	<tr>	
		<td>
			<INPUT type=Button  value="Receipt Schedule" name="cmdReceiptSchedule" ID="cmdReceiptSchedule" OnClick="JavaScript: ShowReceiptSchedule();">
		</td>
	</tr>
	
	
	
	
</table>



</body>

</html>
