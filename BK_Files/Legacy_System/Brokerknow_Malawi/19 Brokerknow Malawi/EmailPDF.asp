<html>

<head>
	<meta http-equiv="Content-Language" content="en-uk">
	<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
	<title>Email PDF</title>
	<LINK href="STYLE/default.css" type=TEXT/CSS rel=STYLESHEET> 
	<LINK href="STYLE/webparts.css" type=TEXT/CSS rel=STYLESHEET>
	<SCRIPT language=Javascript src="scripts/common.js"></SCRIPT>
</head>

<!--#include file="libroutinesTEST.asp"-->

<Script Language="JavaScript">
	function DoSendMail(){
		var to = document.all.item('to');
		var cc = document.all.item('cc');
		var bcc = document.all.item('bcc');
		var subject = document.all.item('subject');
		var bodytext = document.all.item('bodytext');
		var attachment = document.all.item('attachment');
		
		if (to.value=='') {
			alert('Ensure that the recipient address is entered in the "To" field.')
			to.focus();
			return;
		}
		
		if (subject.value=='') {
			alert('Ensure that the mail subject is entered.');
			subject.focus();			
			return;
		}
		
		window.parent.dialogArguments.document.all.item('to2').value = to.value;
		window.parent.dialogArguments.document.all.item('cc2').value = cc.value;
		window.parent.dialogArguments.document.all.item('bcc2').value = bcc.value;
		window.parent.dialogArguments.document.all.item('subject2').value = subject.value;
		window.parent.dialogArguments.document.all.item('bodytext2').value = bodytext.value;	
		window.parent.dialogArguments.document.all.item('attachment').value = attachment.value;	
				
		window.parent.dialogArguments.emailPDFDoc();
		window.parent.self.close(); 
	}
</Script>

<body Class="Dialog">

<form id=form1 name=form1>

	<input type=hidden name="sendPointer" value=1>
	<input type=hidden name="filePath" value="<%= pathTo %>">
		
	<table border="0" width="100%">
		<tr>
			<td width="21%"><font face="Tahoma" size="2" color="#000080"><b>To:</b></font></td>
			<td width="79%"><input type="text" name="To" size="40"></td>
		</tr>
		<tr>
			<td width="21%"><font face="Tahoma" size="2" color="#000080"><b>cc:</b></font></td>
			<td width="79%"><input type="text" name="cc" size="40"></td>
		</tr>
		<tr>
			<td width="21%"><font face="Tahoma" size="2" color="#000080"><b>bcc:</b></font></td>
			<td width="79%"><input type="text" name="bcc"  size="40"></td>
		</tr>
		<tr>
			<td width="21%"><font face="Tahoma" size="2" color="#000080"><b>Subject</b></font></td>
			<td width="79%"><input type="text" name="subject" size	="40"></td>
		</tr>
		<tr style="visibility:hidden;display:none;">
			<td width="21%"><font face="Tahoma" size="2" color="#000080"><b>Attachment</b></font></td>
			<td width="79%"><input disabled type="text" name="attachment"  id ="attachment" size="40" value="PDF FILE NAME"></td>
		</tr>
		<tr>
			<td width="21%"><font face="Tahoma" size="2" color="#000080"><b>Body Text(Optional)</b></font></td>
			<td width="79%"><textarea rows="7" name="bodyText" cols="40"></textarea></td>
		</tr>
	</table>
	
	<p align="right"><input type="button" class="Buttons" value=" Send " name="B1" OnClick="JavaScript: DoSendMail()">&nbsp;&nbsp;&nbsp;<input type="button" class="Buttons" value=" Close " name="B2" OnClick="JavaScript: window.self.close();"></p>

</form>

</body>

</html>
