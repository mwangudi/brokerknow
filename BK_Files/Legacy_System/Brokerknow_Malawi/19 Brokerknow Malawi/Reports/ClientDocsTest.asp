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
	<style> 
	.hide 
	{
	display:none; 
	}
	</style>


<SCRIPT Language="JavaScript">
	
	function validateStatement(frm){			
		if (frm.txtClientCode.value==''){
			alert("Please specify the client.");
			frm.cboClient.focus();
		return;
	}
	frm.target= '_self';
	frm.action = '../Reports/ClientStatement.asp?genReport=1&cboClient=' + frm.txtClientCode.value + '&transFromDate=' + frm.transFromDate.value +'&transToDate=' + frm.txtToDate.value + '';
	//prompt('','ClientStatement.asp?genReport=1&cboClient=' + frm.txtClientCode.value + '&transFromDate=' + frm.transFromDate.value +'&transToDate=' + frm.txtToDate.value + '');
	frm.submit();
	}

	function validateCContr(frm){			
		if (frm.txtClientCode.value==''){
			alert("Please specify the client.");
			frm.cboClient.focus();
		return;
	}
	frm.target= '_self';
	frm.action = '../Reports/SingleClientContract.asp?genReport=1&Client= ' + frm.txtClientCode.value + '&transFromDate=' + frm.transFromDate.value + '&transToDate=' + frm.txtToDate.value + '';				
	frm.submit();
	}
	
	function validateCContrCom(frm){			
		if (frm.txtClientCode.value==''){
			alert("Please specify the client.");
			frm.cboClient.focus();
		return;
	}
	frm.target= '_self';
	frm.action = '../Reports/SingleClientCompounded.asp?genReport=1&Client= ' + frm.txtClientCode.value + '&transFromDate=' + frm.transFromDate.value + '&transToDate=' + frm.txtToDate.value + '';				
	frm.submit();
	}
	
	function validateHV(frm){			
		if (frm.txtClientCode.value==''){
			alert("Please specify the client.");
			frm.cboClient.focus();
		return;
	}
	frm.target= '_self';
	frm.action = '../Reports/HoldingsValuation.asp?genReport=1&Client= ' + frm.txtClientCode.value + '&txtFromDate=' + frm.transFromDate.value + '&txtToDate=' + frm.txtToDate.value + '';				
	frm.submit();
	}
	
	function changecursor (cursortype)
	{
		
		if (cursortype == 0)
		{
			document.body.style.cursor = 'default';
			
			if(document.getElementById("div1"))
			{
				document.getElementById("div1").style.display='none'; 
				document.images["jsbutton"].src= "images/spacer.gif";
			}
		}
		
		if (cursortype == 1)
		{
			document.body.style.cursor = 'wait';
			document.getElementById("div1").style.display='inline'; 
			document.images["jsbutton"].src= "images/roller.gif"; 

		}	
		
	}
	
function LoadClient(accountno, element)
		{
		 var clientcode = document.frmDocs.elements("txtClientCode").value
		 var clientcds = document.frmDocs.elements("txtcdsno").value
		 var clientname = document.frmDocs.elements("txtclientname").value
		 var clientcobo = document.getElementById("cboClient");		 
		 
	     var x_clientname;
			
		 if (element == 'txtClientCode')
		 {
			clientcds = ''
			clientname = ''
			
			if (clientcode == '')
			{
			document.frmDocs.elements("txtClientCode").value = 'Code'
			clientcobo.length = 1;
			clientcobo[0].text = 'Load Account';
			clientcobo[0].value = '';
			return;
			}
			
		 }
		 else if (element == 'txtcdsno')
		 {
			clientcode = ''
			clientname = ''

			if (clientcds == '')
			{
			document.frmDocs.elements("txtcdsno").value = 'CDS No.'
			clientcobo.length = 1;
			clientcobo[0].text = 'Load Account';
			clientcobo[0].value = '';
			return;
			}
						
		 }
		 else if (element == 'txtclientname')
		 {
			clientcode = ''
			clientcds = ''

			if (clientname == '')
			{
			document.frmDocs.elements("txtclientname").value = 'Client Name';
			clientcobo.length = 1;
			clientcobo[0].text = 'Load Account';
			clientcobo[0].value = '';
			return;
			}
			
		 }

		xmlhttp = createXMLHTTPObj();
		
		url="GetList.asp?clientcode="+clientcode+"&cdsno="+clientcds+"&clientname="+clientname+"&action=SLoadClient";
		
		//alert(url);
		//changecursor(1)

		xmlhttp.open("GET",url,true);

		xmlhttp.onreadystatechange=function() 
		  {
					if (xmlhttp.readyState==4) 
					{
						returnStr = xmlhttp.responseText;
						returnStr = getBodyHTML(returnStr);
						
						//alert(returnStr);
						
						myArray = returnStr.split("<->");
						alert(myArray[2]);
						x_clientname = myArray[2]

						if (x_clientname.length > 12) 
						{
							x_clientname = x_clientname.substring(0,16)  + '...';
						}
						
						//document.getElementById("cboAccount").options.length = 0;
						clientcobo.length = 1;
						if (element != 'txtclientname')
						{
							document.frmDocs.elements("txtClientCode").value = myArray[8];
							document.frmDocs.elements("txtcdsno").value =myArray[9];
							document.frmDocs.elements("txtclientname").value = x_clientname;
							
							clientcobo[0].SearchCode = myArray[5];
							clientcobo[0].SearchText = myArray[7];
							
							clientcobo[0].text = myArray[2];
							clientcobo[0].value = myArray[3];
							
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
								
								myArrayz = myArrayx[i].split(";");

								alert(myArrayz)
								
								document.frmDocs.elements("txtClientCode").value = '';
								document.frmDocs.elements("txtcdsno").value = '';
								document.frmDocs.elements("txtclientname").value = '';
								
								document.getElementById("cboClient").options[i] = new Option(myArrayz[2],myArrayz[0]);
																		
							}
							
						}
					}
			}
		 
		 xmlhttp.setRequestHeader('Accept','message/x-jl-formresult');
		 xmlhttp.send();
}


function updatebalances()
		{
			var url;

				xmlhttp = createXMLHTTPObj();
				
				url="GetList.asp?action=UpdateBalances";
				
				xmlhttp.open("GET",url,true);

				xmlhttp.onreadystatechange=function() 
				  {
							if (xmlhttp.readyState==4) 
							{
								returnStr = xmlhttp.responseText;
								returnStr = getBodyHTML(returnStr);
								
								if (returnStr = "1")
								{
									alert("All balances have been updated succesfully.");
									changecursor(0);
								}
								
							}
					}
				 
				 xmlhttp.setRequestHeader('Accept','message/x-jl-formresult');
				 xmlhttp.send();
		}

		
		function ClearFields(element)
		{
		   if (element == 'txtClientCode')
		   {
			document.frmDocs.elements("txtClientCode").value = ''
			document.frmDocs.elements("txtcdsno").value = 'CDS No.'
			document.frmDocs.elements("txtclientname").value = 'Client Name'
			return;
		   }
		   if (element == 'txtcdsno')
		   {
			document.frmDocs.elements("txtClientCode").value = 'Code'
			document.frmDocs.elements("txtcdsno").value = ''
			document.frmDocs.elements("txtclientname").value = 'Client Name'
			return;
		   }
		   if (element == 'txtclientname')
		   {
		    document.frmDocs.elements("txtclientname").value = ''
			document.frmDocs.elements("txtClientCode").value = 'Code'
			document.frmDocs.elements("txtcdsno").value = 'CDS No.'
			return;
		   }		
		   
		}

		function updatefields(selectedclient)
		{
		 document.frmDocs.elements("txtclientname").value = '';
		 document.frmDocs.elements("txtcdsno").value = '';
		 document.frmDocs.elements("txtClientCode").value = selectedclient.value;
		}

		function Toggle() 
		{
			if( document.getElementById("div1").style.display == 'none') 
			{
				document.getElementById("div1").style.display='inline'; 
				document.images["jsbutton"].src= "spacer.gif"; 
			}
		
			else 
			{
				document.getElementById("div1").style.display='none'; 
				document.images["jsbutton"].src= "roller.gif"; 
			}
		}
</SCRIPT>
<body class="Reports" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<SCRIPT Language="JavaScript">
	changecursor(0)
</script>
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
			<td nowrap>
				&nbsp;
				<input type = 'text' name ='txtClientCode' id = 'txtClientCode' size="5" value = 'Code' onBlur="LoadClient(cboClient, this.name);"  onClick  = "ClearFields(this.name)"> 
				&nbsp; <input type = 'text' name ='txtcdsno' id = 'txtcdsno' size="10" value = 'CDS NO.' onBlur="LoadClient(cboClient, this.name);" onClick  = "ClearFields(this.name)"> 
				<input type = 'text' name ='txtclientname' id = 'txtclientname' size="15" value = 'Client Name' onBlur="LoadClient(cboClient, this.name);" onClick  = "ClearFields(this.name)"> 
				<select name = 'cboClient' id = "cboClient" size="1" onChange ="updatefields(this);">
				<option SearchCode = "" SearchText = "" value = '' >Load Account</option>
				</select>
			</td>	
		</tr>
  
	</table>
	
	<%FirstDay = (Date-30)%>
	
	<table>
		
		<tr>	
			<td colspan=2>&nbsp;
				
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
			<td colspan=2>&nbsp;
				
			</td>
		</tr>
		
		<tr>	
			<td colspan=2>
				<table align="center" style="border:1 solid gray;background-color:gainsboro;" border="0" width="94%">
          <tr> 
            <td colspan="2" align="left"> Client </td>
          </tr>
          <tr> 
            <td colspan="2" align="left"> <INPUT style="width:300px;" type=Button style="1000px" value="Generate Client Statement" name="genOrder" ID="genOrder" TITLE="Generate Client Statement" OnClick="JavaScript: validateStatement(document.all.item('frmDocs'));"> 
            </td>
          </tr>
          <tr> 
            <td align="left"> <INPUT style="width:300px;" type=Button  value="Generate Client Contract" name="genCContr" ID="genCContr" TITLE="Generate the Client Contract" OnClick="JavaScript: validateCContr(document.all.item('frmDocs'));"> 
            </td>
            <td width="75%" rowspan="4" align="left">
			
			  </td>
          </tr>
          <tr> 
            <td align="left"> <INPUT style="width:300px;" type=Button  value="Generate Compounded Client Contract" name="genCContrCom" ID="genCContrCom" TITLE="Generate the Client Contract (Compounded)" OnClick="JavaScript: validateCContrCom(document.all.item('frmDocs'));"> 
            </td>
          </tr>
          <tr> 
            <td align="left"><input style="width:300px;" type=Button  value="Generate Holdings Valuation" name="genHV" id="genHV" title="Generate Holdings Valuation" onClick="JavaScript: validateHV(document.all.item('frmDocs'));"></td>
          </tr>
         <!--<tr> 
          <td width="25%" align="left"><input style="width:300px;" type=Button  value="Update All Client Balances" name="updblnce" id="updblnce" title="Generate Holdings Valuation" onClick="JavaScript: updatebalances();changecursor(1)"> 
            </td>
          </tr>-->
        </table>
			</td>
			<td>
			  <table width="75%" border="0">
    <tr>
      <td height="53" align="center">
	  <div class="hide" id="div1">
			<img name="jsbutton" id="jsbutton" src="images/roller.gif" width="32" height="32">
              <table width="41%" border="0">
                <tr>
                  
              <td nowrap><font color="#333333" face="Verdana, Arial, Helvetica, sans-serif">Please 
                Wait while Client Balances Updates</font></td>
                </tr>
              </table>
			  </div>
	  </td>
    </tr>
  </table>
</td>
		</tr>
	</table>
</form>
</body>
<SCRIPT Language="JavaScript">
	changecursor(0)
</script>
</html>