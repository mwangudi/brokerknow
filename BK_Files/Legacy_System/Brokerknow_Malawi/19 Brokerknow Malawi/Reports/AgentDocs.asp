<!--#include file="../libroutines.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE>Agent Documents</TITLE>
	<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
	<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
	<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
	<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
	<SCRIPT language=VBScript src="../scripts/reports.vbs"></SCRIPT>
	<SCRIPT language=Javascript src="../scripts/reports.js"></SCRIPT>
	<SCRIPT language=Javascript src="../scripts/accountList.js"></SCRIPT>



<SCRIPT Language="JavaScript">
	
	function validateAContr(frm){			
		if (frm.txtAgentCode.value==''){
			alert("Please specify the Agent.");
			frm.cboAgent.focus();
		return;
	}
	
	frm.target= '_self';
	frm.action = '../Reports/SingleAgentContract.asp?genReport=1&Agent= ' + frm.txtAgentCode.value + '&transFromDate=' + frm.transFromDate.value + '&transToDate=' + frm.txtToDate.value + '';				
	frm.submit();
	}
	
	function validateAContrCom(frm){			
		if (frm.txtAgentCode.value==''){
			alert("Please specify the Agent.");
			frm.cboAgent.focus();
		return;
	}
	frm.target= '_self';
	frm.action = '../Reports/SingleAgentContractCompounded.asp?genReport=1&Agent= ' + frm.txtAgentCode.value + '&transFromDate=' + frm.transFromDate.value + '&transToDate=' + frm.txtToDate.value + '';					
	frm.submit();
	}
		
</SCRIPT>
<body class="Reports" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="post" action="" name="frmDocs" id="frmDocs">
	<table border=0 cellspacing=5 cellpadding=5 width=50%>
		<tr>
			<td colspan=2> GENERATE DOCUMENT BY:</td>
		</tr>
		<tr>
			<td width=1>&nbsp;</td>
			<td width=90%>
				<input type="radio" checked class="BorderLess" value="1" id="DocNo" name="DocNo" >&nbsp;&nbsp;<label for="DocNo" style="cursor: hand">Agent</label>				<br>
				<input type="radio" class="BorderLess" value="1" id="AgNo" name="AgNo" onclick ="javascript: window.location.replace('ClientDocs.asp');">&nbsp;&nbsp;<label for="AgNo" style="cursor: hand">Client</label>
				<br>
				<input type="radio" class="BorderLess" value="1" id="OrdNo" name="OrdNo" onclick ="javascript: window.location.replace('ClientDocsOrd.asp');" >&nbsp;&nbsp;<label for="OrdNo" style="cursor: hand" >Document Number<label>
			</td>
		</tr>
	</table>
		<br>
	<table border=0>
		<tr>
			<td style="padding-left=10px"> Specify Agent</td>
		</tr>
		<tr>	
			<td width="30%" height="10" style="padding-left=50px">
			<input type = 'text' name ='txtAgentCode' id = 'txtAgentCode' size="10" onBlur="txtval = this.value; selectItem(cboAgent);UpdateCodes(true,cboAgent,txtCdsNo);">
			<input type = 'hidden' name ='txtCdsNo' id = 'txtCdsNo' size="16" onBlur="txtval = this.value; selectItems(cboAgent);UpdateCode(true,cboAgent,txtAgentCode);">
			<select name = 'cboAgent' id = 'cboAgent' size="1" 
					onKeypress="return (dodefaultaction()==''); " 
					onKeydown="return (dodefaultaction()==''); " 
					onKeyup="return (UpdateCode(change(cboAgent,0),cboAgent,txtAgentCode));" 
					onChange="UpdateCode(true,cboAgent,txtAgentCode);UpdateCodes(true,cboAgent,txtCdsNo);"
					onfocus="txtval = '';inputIsItemCode = 1;" 
					onblur="txtval = '';inputIsItemCode = 1;" readonly>
				    	
					<%
					dim AgentName
					dim NameAgent 
					dim conn
					dim rs
					dim sqlStr
					
					Set Conn = GetActiveConnection("KBroker")
					
					Set rs = Server.CreateObject("ADODB.Recordset")

					sqlStr = "SELECT * FROM Agent WHERE (Deleted = 0) ORDER BY AgentName ASC"
					Set rs = conn.Execute(sqlStr)
					If Not (rs.EOF Or rs.BOF) Then
						    rs.MoveFirst
					        Do Until rs.EOF					                
								        
					        AgentName=rs.Fields("AgentName")
						    NameAgent=Mid(AgentName,1,30)
								    %>                    
					                <option SearchCode = "<%=rs.Fields("Agent_DPA_")%>" SearchText = "<%=rs.Fields("AgentName")%>" value = '<%=rs.Fields("Agent_DPA_")%>'><%=NameAgent%></option>
					                <%rs.MoveNext
					        Loop
					End If
					
					Set rs = Nothing
					Set Conn = Nothing
					
					%>
			</select>		
			</td>
		</tr>
	</table>
	<%FirstDay = (Date-30)%>
	<table>			<tr>	
			<td colspan=2>
				&nbsp;
			</td>
		</tr>		
		<Script Language="JavaScript">
		
			var cal=new ctlSpiffyCalendarBox("cal", "frmDocs", "transFromDate","cmdDate","<%= FormatDate(FirstDay) %>",1);
			var cal1=new ctlSpiffyCalendarBox("cal1", "frmDocs", "txtToDate","cmdDate2","<%= FormatDate(Date) %>",1);

		</Script>
		
			<tr>
				<td colspan="2">Select date from:</td>
				<td>
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>				
			</tr>
			<tr>
				<td colspan="2">To date:</td>
				<td>
					<SCRIPT language="JavaScript">cal1.writeControl();</SCRIPT>	
				</td>
				
			</tr>			
	</table>
	<table>
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
						Agent
						</td>
					</tr>
					
					<tr>
						<td width="100%" align="left">
						<INPUT style="width:300px;" type=Button  value="Generate Agent Contract" name="genAContr" ID="genAContr" TITLE="Generate the Agent Contract" OnClick="JavaScript: validateAContr(document.all.item('frmDocs'));">
						</td>
					</tr>
					
					<tr>
						<td width="100%" align="left">
						<INPUT style="width:300px;" type=Button  value="Generate Compounded Agent Contract" name="genAContrCom" ID="genAContrCom" TITLE="Generate the Agent Contract (Compounded)" OnClick="JavaScript: validateAContrCom(document.all.item('frmDocs'));">
						</td>
					</tr>
				</table>			</td>
		</tr>	</table>
	</form>
</body>

</html>