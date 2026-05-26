<!--#include file="../libroutines.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE>Client Documents</TITLE>
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
		function validateStatement(frm){			
		if (frm.txtClientCode.value=='' || frm.txtClientCode.value.toLowerCase()=='code' || frm.txtClientCode.value=='0'){
			alert("Please specify the client.");
			frm.cboClient.focus();
		return;
	}
	frm.target= '_self';
	frm.action = '../Reports/ClientStatement.asp?genReport=1&cboClient=' + frm.txtClientCode.value + '&transFromDate=' + frm.transFromDate.value +'&transToDate=' + frm.txtToDate.value + '';
	frm.submit();
	}

	function validateCContr(frm){			
		if (frm.txtClientCode.value=='' || frm.txtClientCode.value.toLowerCase()=='code' || frm.txtClientCode.value=='0'){
			alert("Please specify the client.");
			frm.cboClient.focus();
		return;
	}
	frm.target= '_self';
	frm.action = '../Reports/SingleClientContract.asp?genReport=1&Client= ' + frm.txtClientCode.value + '&transFromDate=' + frm.transFromDate.value + '&transToDate=' + frm.txtToDate.value + '';				
	frm.submit();
	}
		function validateCContrCom(frm){			
		if (frm.txtClientCode.value=='' || frm.txtClientCode.value.toLowerCase()=='code' || frm.txtClientCode.value=='0'){
			alert("Please specify the client.");
			frm.cboClient.focus();
		return;
	}
	frm.target= '_self';
	frm.action = '../Reports/SingleClientCompounded.asp?genReport=1&Client= ' + frm.txtClientCode.value + '&transFromDate=' + frm.transFromDate.value + '&transToDate=' + frm.txtToDate.value + '';				
	frm.submit();
	}
		function validateHV(frm){			
		if (frm.txtClientCode.value=='' || frm.txtClientCode.value.toLowerCase()=='code' || frm.txtClientCode.value=='0'){
			alert("Please specify the client.");
			frm.cboClient.focus();
		return;
	}
	frm.target= '_self';
	frm.action = '../Reports/HoldingsValuation.asp?genReport=1&Client= ' + frm.txtClientCode.value + '&txtFromDate=' + frm.transFromDate.value + '&txtToDate=' + frm.txtToDate.value + '';				
	frm.submit();
	}

	
		function ClearFields(element)
		{
		   if (element == 'txtClientCode')
		   {
			document.frmDocs.elements("txtClientCode").value = '';
			document.frmDocs.elements("txtcdsno").value = 'CDS No.';
			document.frmDocs.elements("txtclientname").value = 'Client Name';
			return;
		   }
		   if (element == 'txtcdsno')
		   {
			document.frmDocs.elements("txtClientCode").value = 'Code';
			document.frmDocs.elements("txtcdsno").value = '';
			document.frmDocs.elements("txtclientname").value = 'Client Name';
			return;
		   }
		   if (element == 'txtclientname')
		   {
		    document.frmDocs.elements("txtclientname").value = '';
			document.frmDocs.elements("txtClientCode").value = 'Code';
			document.frmDocs.elements("txtcdsno").value = 'CDS No.';
			return;
		   }		
		   
		}

		function updatefields(selectedclient)
		{
		 document.frmDocs.elements("txtclientname").value = '';
		 document.frmDocs.elements("txtcdsno").value = '';
		 document.frmDocs.elements("txtClientCode").value = selectedclient.value;
		 //alert(selectedclient.value);
			LoadMyClient();
		}

		function LoadMyClient()
		{
			var clientcode = document.frmDocs.elements("txtClientCode").value
			var clientcds = document.frmDocs.elements("txtCdsNo").value
			var clientcobo = document.getElementById("cboAccount");		 
			var guidstr = Math.random();
			
			xmlhttp = createXMLHTTPObj();
				
			url="GetList.asp?clientcode="+clientcode+"&cdsno="+clientcds+"&clientname=&action=SLoadClient&guidstr="+guidstr;
			
			xmlhttp.open("GET",url,true);

			xmlhttp.onreadystatechange=function() 
			{
				if (xmlhttp.readyState==4) 
				{
					returnStr = xmlhttp.responseText;
					returnStr = getBodyHTML(returnStr);
									
					myArray = returnStr.split("<->");
					//alert(myArray);
									
					document.frmDocs.elements("txtclientname").value = myArray[7];
					document.frmDocs.elements("txtClientCode").value = myArray[5];
					document.frmDocs.elements("txtCdsNo").value = myArray[9]; 
				}
		   }
				 
		xmlhttp.setRequestHeader('Accept','message/x-jl-formresult');
		xmlhttp.send();
		}

		function LoadClient(accountno, element)
		{
		 var clientcode = document.frmDocs.elements("txtClientCode").value;
		 var clientcds = document.frmDocs.elements("txtcdsno").value;
		 var clientname = document.frmDocs.elements("txtclientname").value;
		 var clientcobo = document.getElementById("cboAccount");
		 
	     var x_clientname;
		
		
		 if (element == 'txtClientCode')
		 {
			clientcds = '';
			clientname = '';
			
			if (clientcode == '')
			{
			document.frmDocs.elements("txtClientCode").value = 'Code'
			clientcobo.length = 1;
			clientcobo[0].text = 'Load Account';
			clientcobo[0].value = '';
			loadIframe(0);
			return;
			}
			
		 }
		/* else if (element == 'txtcdsno')
		 {
			clientcode = '';
			clientname = '';

			if (clientcds == '')
			{
			document.frmDocs.elements("txtcdsno").value = 'CDS No.'
			clientcobo.length = 1;
			clientcobo[0].text = 'Load Account';
			clientcobo[0].value = '';
			loadIframe(0);
			return;
			}
						
		 }*/
		 else if (element == 'txtclientname')
		 {
			clientcode = '';
			clientcds = '';

			if (clientname == '')
			{
			document.frmDocs.elements("txtclientname").value = 'Client Name';
			clientcobo.length = 1;
			clientcobo[0].text = 'Load Account';
			clientcobo[0].value = '';
			loadIframe(0);
			return;
			}
			
		 }

				xmlhttp = createXMLHTTPObj();
				
				url="GetList.asp?clientcode="+clientcode+"&cdsno="+clientcds+"&clientname="+clientname+"&action=SLoadClient";
				
				//alert(url);

				xmlhttp.open("GET",url,true);

				xmlhttp.onreadystatechange=function() 
				  {
							if (xmlhttp.readyState==4) 
							{
								returnStr = xmlhttp.responseText;
								returnStr = getBodyHTML(returnStr);
								
								//alert(returnStr);

								myArray = returnStr.split("<->");
								
								x_clientname = myArray[7];

								if (x_clientname.length > 12) 
								{
									x_clientname = x_clientname.substring(0,16)  + '...';
								}
								
								//document.getElementById("cboAccount").options.length = 0;
								clientcobo.length = 1;
								if (element != 'txtclientname')
								{
									
									document.frmDocs.elements("txtclientname").value = x_clientname;
									document.frmDocs.elements("txtClientCode").value = myArray[5];
									//document.frmDocs.elements("txtCdsNo").value = myArray[9]; 
									
									clientcobo[0].Credit = myArray[0];
									clientcobo[0].CurrentBal = myArray[1];
									clientcobo[0].Agent = myArray[2];
									clientcobo[0].Owner = myArray[3];
									clientcobo[0].AgentID = myArray[4];
									clientcobo[0].SearchCode = myArray[5];
									clientcobo[0].OrderContact = myArray[6];
									clientcobo[0].SearchText = myArray[7];
									clientcobo[0].OwnerID = myArray[8];
									clientcobo[0].SearchCDS = myArray[9];
									clientcobo[0].IsCustodian = myArray[10];
									
									clientcobo[0].text = myArray[7];
									clientcobo[0].value = myArray[5];

								}
								else
								{
									var myArrayx;
									var myArrayz;
									
									//alert(returnStr);
									myArrayx = returnStr.split("|");
									myArrayxsize = myArrayx.length - 1;
									
									//alert(myArrayxsize);

									for (i=myArrayxsize; i>=0; i--)
									{
										
										myArrayz = myArrayx[i].split("<->");

										//alert(myArrayz)
										
										document.frmDocs.elements("txtClientCode").value = '';
										//document.frmDocs.elements("txtcdsno").value = '';
										document.frmDocs.elements("txtclientname").value = '';
										document.frmDocs.elements("txtClientCode").value = myArray[5];
										//document.frmDocs.elements("txtCdsNo").value = myArray[9]; 
										document.frmDocs.elements("txtClientname").value = myArray[7];
										
										document.getElementById("cboAccount").options[i] = new Option(myArrayz[7],myArrayz[5],myArrayz[6],myArrayz[10],myArrayz[4],myArrayz[8],myArrayz[3],myArrayz[3],myArrayz[0],myArrayz[1],myArrayz[5],myArrayz[7],myArrayz[9]);
																				
									}
									
								}
							}
					}
				 
				 xmlhttp.setRequestHeader('Accept','message/x-jl-formresult');
				 xmlhttp.send();
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
				<input type="radio" class="BorderLess" value="1" id="AgNo" name="AgNo" onclick ="javascript: window.location.replace('AgentDocs.asp');">&nbsp;&nbsp;<label for="AgNo" style="cursor: hand">Agent</label>
				<br>
				<input type="radio" checked class="BorderLess" value="1" id="DocNo" name="DocNo" >&nbsp;&nbsp;<label for="DocNo" style="cursor: hand">Client</label>
				<br>
				<input type="radio"  class="BorderLess" value="1" id="OrdNo" name="OrdNo" onclick ="javascript: window.location.replace('ClientDocsOrd.asp');" >&nbsp;&nbsp;<label for="OrdNo" style="cursor: hand" >Document Number<label>
			</td>
		</tr>
	</table>
		<br>
		<table border=0>
		<tr>
			<td style="padding-left=10px"> Specify Client</td>
		</tr>
		<tr>	
			<td width="30%" height="10" style="padding-left=50px">
			<input type = 'text' name ='txtClientCode' id = 'txtClientCode' size="5" value = 'Code' onBlur="txtval = this.value;LoadClient(cboAccount, this.name);"  onClick  = "ClearFields(this.name)">
	&nbsp;
	<input style="display:none" type = 'text' name ='txtcdsno' id = 'txtcdsno' size="10" value = 'CDS NO.' onBlur="txtval = this.value;LoadClient(cboAccount, this.name);" onClick  = "ClearFields(this.name)">

	<input type = 'text' name ='txtclientname' id = 'txtclientname' size="15" value = 'Client Name' onBlur="txtval = this.value;LoadClient(cboAccount, this.name);" onClick  = "ClearFields(this.name)">

	<select name = 'cboAccount' id = "cboAccount" size="1" onChange ="updatefields(this)">		
		<option SearchCode = "" SearchText = "" value = '' >Load Account</option>
	</select> </td>
	
			</td>
		</tr>
	</table>	
	<%FirstDay = (Date-30)%>	
	<table>
				<tr>	
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
						Client
						</td>
					</tr>
					
					<tr>
						<td width="100%" align="left">
						<INPUT style="width:300px;" type=Button style="1000px" value="Generate Client Statement" name="genOrder" ID="genOrder" TITLE="Generate Client Statement" OnClick="JavaScript: validateStatement(document.all.item('frmDocs'));">
						</td>
					</tr>
					
					<tr>
						<td width="100%" align="left">
						<INPUT style="width:300px;" type=Button  value="Generate Client Contract" name="genCContr" ID="genCContr" TITLE="Generate the Client Contract" OnClick="JavaScript: validateCContr(document.all.item('frmDocs'));">
						</td>
					</tr>
					
					<tr>
						<td width="100%" align="left">
						<INPUT style="width:300px;" type=Button  value="Generate Compounded Client Contract" name="genCContrCom" ID="genCContrCom" TITLE="Generate the Client Contract (Compounded)" OnClick="JavaScript: validateCContrCom(document.all.item('frmDocs'));">
						</td>
					</tr>
					
				</table>			</td>
		</tr>	</table>
	</form>
</body>

</html>