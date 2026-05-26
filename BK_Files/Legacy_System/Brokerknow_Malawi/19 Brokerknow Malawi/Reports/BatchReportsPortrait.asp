<html>

<head>
<title>Batch Report [Portrait]</title>
  
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
<SCRIPT language=VBScript src="../scripts/reports.vbs"></SCRIPT>
<SCRIPT language=Javascript src="../scripts/reports.js"></SCRIPT>
	<style media="print">
		@page {
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
					page-break-before: always;
				}

				tr.pageNumbering{
					display:none;
				}
		      }
	</style>

	<script language="vbscript">
		Function DoSendMail()
			theData = document.repToPDF0.innerHTML
			frm1.hidData0.value = theData
		
			theData = document.repToPDF1.innerHTML
			frm1.hidData1.value = theData
		
			theData = document.repToPDF2.innerHTML
			frm1.hidData2.value = theData
		
			theData = document.repToPDF3.innerHTML
			frm1.hidData3.value = theData
		
			theData = document.repToPDF4.innerHTML
			frm1.hidData4.value = theData
		
			theData = document.repToPDF5.innerHTML
			frm1.hidData5.value = theData
		
			frm1.method = "post"
			frm1.action = "../Mailer/EmailBatchReports.asp"
			frm1.submit
		End Function
	</script>
</head>

<body Class="Reports">

<!--#include file="../libroutinesTEST.asp"-->
<!--#include file="BatchReportsFunctions.asp"-->

<%
genReport = Request.QueryString("genReport")
selectedContractDate = Request.QueryString("selectedContractDate")
	
If genReport <> "1" Or selectedContractDate = "" Then
	%>
	<Script Language="JavaScript">
		report_SetBodyClass();
		
		function validateForm(frm)
		{			
			if (frm.txtDate.value=='')
			{
				alert("Select a date");
				frm.txtDate.focus();
				return;
			}
						
			frm.action = 'BatchReportsPortrait.asp?genReport=1&selectedContractDate='+frm.txtDate.value;
			frm.target = '_self';			
			frm.submit();
		}
				
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "txtDate","cmdDate","<%= FormatDate(Date) %>",1);
	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	
	<form method="POST" action="BatchReportsPortrait.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">	
		<table>
			<tr>
				<td>Select date:</td>
				<td>
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>
			</tr>
					
			<tr>
				<td colspan=2><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... ">&nbsp;&nbsp;</td>
			</tr>
		</table>
	</form>
	<%
	Response.End
End If
%>

<% DrawPageFunctions True, False, True, False %>

<form id="frm1" name="frm1">
	<table border="0" width="100%">
		<tr>
			<td width="100%" align="left"><font face="Tahoma" size="2" color="#000080"><b>To:</b></font>&nbsp;<input type="text" name="To" size="40">&nbsp;<input type="button" class="Buttons" value=" Send " name="B1" OnClick="JavaScript: DoSendMail()"></td>
		</tr>
	</table>
	
	<input type="hidden" name="hidDate" id="hidDate" value="<%=selectedContractDate%>">
	
	<input type="hidden" name="hidData0" id="hidData0">
	<input type="hidden" name="repToPDFOrient0" id="repToPDFOrient0" value="L">
	
	<input type="hidden" name="hidData1" id="hidData1">
	<input type="hidden" name="repToPDFOrient1" id="repToPDFOrient1" value="P">
	
	<input type="hidden" name="hidData2" id="hidData2">
	<input type="hidden" name="repToPDFOrient2" id="repToPDFOrient2" value="P">
	
	<input type="hidden" name="hidData3" id="hidData3">
	<input type="hidden" name="repToPDFOrient3" id="repToPDFOrient3" value="P">
	
	<input type="hidden" name="hidData4" id="hidData4">
	<input type="hidden" name="repToPDFOrient4" id="repToPDFOrient4" value="P">
	
	<input type="hidden" name="hidData5" id="hidData5">
	<input type="hidden" name="repToPDFOrient5" id="repToPDFOrient5" value="P">
</form>

<p id="toPDFOrient" name="toPDFOrient" value="P" style="display:none;">P
<p id="toPDF" name="toPDF">
<%

ContractScheduleGenerate

ContractsGenerate

ContractsCompoundedGenerate

AgentContractsGenerate

AgentContractsCompoundedGenerate

SettlementGenerate

%>
</body>
</html>



