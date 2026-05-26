<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE>Client Documents - Doc No</TITLE>
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT>
</head>
<!--#include virtual="libroutines.asp"-->

<SCRIPT Language="JavaScript">
		
	function validateCContr(frm){			
		if (frm.DocNum.value==''){
			alert("You must enter a contract number.");
			frm.DocNum.focus();
		return;
	}
	frm.target= '_self';
	frm.action = '../Reports/SingleClientContract.asp?genReport=1&DocNo=' + frm.DocNum.value;			
	frm.submit();
	}

	function validateCContrCom(frm){			
		if (frm.DocNum.value==''){
			alert("You must enter a contract number.");
			frm.DocNum.focus();
		return;
	}
	frm.target= '_self';
	frm.action = '../Reports/SingleClientCompounded.asp?genReport=1&DocNo=' + frm.DocNum.value;		
	frm.submit();
	}
	
	function validateAContr(frm){			
		if (frm.DocNum.value==''){
			alert("You must enter a contract number.");
			frm.DocNum.focus();
		return;
	}
	frm.target= '_self';
	frm.action = '../Reports/SingleAgentContract.asp?genReport=1&DocNo=' + frm.DocNum.value;				
	frm.submit();
	}
	
	function validateAContrCom(frm){			
		if (frm.DocNum.value==''){
			alert("You must enter a contract number.");
			frm.DocNum.focus();
		return;
	}
	frm.target= '_self';
	frm.action = '../Reports/SingleAgentContractCompounded.asp?genReport=1&DocNo=' + frm.DocNum.value;	
	frm.submit();
	}
				
</SCRIPT>
<body class="Reports" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
	<form method="post" enctype="multipart/form-data"  action="" name="frmDocs">
	<input type="hidden" value="1" name="genReport">
	<input type=hidden name=delPhoto value=0>
	<input type=hidden name=delPhotoPath value=0>
	<table border=0 cellspacing=5 cellpadding=5 width=50%>
		<tr>
			<td colspan=2> GENERATE DOCUMENT BY:</td>
		</tr>
		<tr>
			<td width=1>&nbsp;</td>
			<td width=90%>
				<input type="radio" class="BorderLess" value="1" id="AgNo" name="AgNo" onclick ="javascript: window.location.replace('AgentDocs.asp');" >&nbsp;&nbsp;<label for="AgNo" style="cursor: hand">Agent</label>
				<br>
				<input type="radio" class="BorderLess" value="1" id="DocNo" name="DocNo" onclick ="javascript: window.location.replace('ClientDocs.asp');">&nbsp;&nbsp;<label for="DocNo" style="cursor: hand">Client</label>
				<br>
				<input type="radio" checked class="BorderLess" value="1" id="OrdNo" name="OrdNo" >&nbsp;&nbsp;<label for="OrdNo" style="cursor: hand" >Document Number<label>
			</td>
		</tr>
	</table>
	<br>
	<table border=0>
		<tr>	
			<td nowrap>&nbsp;&nbsp; Enter the Document Number</td>
			<td>
				<INPUT class="cal-textboxBlue" value="" name="DocNum" ID="DocNum" type="text" size="30" TITLE="Enter the document number">
			</td>
		</tr>
		<tr>	
			<td colspan=2>
				&nbsp;
			</td>
		</tr>
		<tr>	
			<td colspan=2>
				<table align="center" style="border:1 solid gray;background-color:gainsboro;" border="0" width="100%">
					<tr>
						<td width="100%" align="left">
						Client
						</td>
						
						<td width="100%" align="left">
						Agent
						</td>
					</tr>
					
					<tr>
						<td width="100%" align="left">
						<INPUT style="width:300px;" type=Button  value="Generate Client Contract" name="genCContr" ID="genCContr" TITLE="Generate the Client Contract" OnClick="JavaScript: validateCContr(document.all.item('frmDocs'));">
						</td>
						
						<td width="100%" align="left">
						<INPUT style="width:300px;" type=Button  value="Generate Agent Contract" name="genAContr" ID="genAContr" TITLE="Generate the Agent Contract" OnClick="JavaScript: validateAContr(document.all.item('frmDocs'));">
						</td>
					</tr>
					
					<tr>
						<td width="100%" align="left">
						<INPUT style="width:300px;" type=Button  value="Generate Compounded Client Contract" name="genCContrCom" ID="genCContrCom" TITLE="Generate the Client Contract (Compounded)" OnClick="JavaScript: validateCContrCom(document.all.item('frmDocs'));">
						</td>
						
						<td width="100%" align="left">
						<INPUT style="width:300px;" type=Button  value="Generate Compounded Agent Contract" name="genAContrCom" ID="genAContrCom" TITLE="Generate the Agent Contract (Compounded)" OnClick="JavaScript: validateAContrCom(document.all.item('frmDocs'));">
						</td>
					</tr>
				</table>			</td>
		</tr>
	</table>
	</form>
</body>

</html>