<!--#include file="../libroutines.asp"-->
<HTML>
<HEAD>
<TITLE> SMS </TITLE>
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css">
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT>
<SCRIPT language="javascript">

function SendSms(sMobileNo,hidID){
        
		var sUrl;
        var sAPI_ID, sPassword, sUsername;
		var xmlhttp, sPostData, sResult;
		var max_credits, escalate, from;
		var sText = document.all.item(hidID).value
		
		sUrl = "https://api.clickatell.com/http/sendmsg" 
		sAPI_ID = 1528651
		sPassword = "alliance1"
		sUsername = "africanalliance"
		max_credits = 3
		escalate =3
		from = "AAKS"
        
		sPostData = "api_id=" +  sAPI_ID;
		sPostData = sPostData + "&user=" + sUsername;
		sPostData = sPostData + "&password=" + sPassword;
		sPostData = sPostData + "&to=" + sMobileNo;
		sPostData = sPostData + "&text=" + sText;
		sPostData = sPostData + "&max_credits=" + max_credits;
		sPostData = sPostData + "&escalate=" + escalate;
		sPostData = sPostData + "&from=" + from;
        
        //alert (hidID);
        //alert (sText);
        //alert (sPostData);
        
        xmlhttp = createXMLHTTPObj() ;

		try
		{
			xmlhttp.open("POST", sUrl, false)
		}
		catch(e)
		{
			return ( e.number + ': ' + e.description)
		}
        
		xmlhttp.setRequestHeader("Content-Type","application/x-www-form-urlencoded");
		xmlhttp.send(sPostData) ;
        
		sResult = xmlhttp.responseText;
        return (sResult) 
		
	}

	function ValidateSMS(){
	var recordcount = document.frmMain.NoSms.value;
	var i,j, m;
	var data;
	var elem
  
		 if (recordcount!=''){
		      m=0;
			   for (i = 0; i < recordcount; i++){
			    j = i + 1
				  elemName = "Debit" + j
				  
				  elem = document.getElementById(elemName)
				  
					 if (elem.checked==true){
					 
					    if (isNaN(parseInt(elem.value))){
						 }else{
							  if (m==0){
								data = elem.value;
							  }else {
								data = data + ';' + elem.value;
							  }
							  m = 1;
						 }
						  
					 }
			   }
		  }


		 if (data!=''){
          	window.location.replace("DebitSMS.asp?gen=1&Debits=1&CanSendSMS=1&data="+data);	 
		 }
	} 

</SCRIPT>
<SCRIPT Language="JavaScript">
	function Generate()
	{
		var stockWatch
		var Contracts
		var Debits
		
		if (FrmSMS.stockWatch.checked == true)
			{
				stockWatch = 1;
			}
		else
			{
				stockWatch = 0;
			}
		
		if (FrmSMS.Contracts.checked == true)
			{
				Contracts = 1;
			}
		else
			{
				Contracts = 0;
			}
				
		if (FrmSMS.Debits.checked == true)
			{
				Debits = 1;
			}
		else
			{
				Debits = 0;
			}
					
		window.location.replace("DebitSMS.asp?gen=1&stockWatch="+stockWatch+"&Contracts="+Contracts+"&Debits="+Debits);
	}
	
	function Download(){
		window.location.replace("downloadSMS.asp");
	}
</SCRIPT>
</HEAD>

<BODY>
	
<%
gen = trim(Request.QueryString("gen"))

If gen = "" Then gen=0
	
If gen <> 1 Then
	%>
	<form method="post" action="DebitSMS.asp" name="FrmSMS">
		<table width="50%">
			<tr>
				<td>&nbsp;<BR></td> <td></td>
			</tr>
			<tr>
				<td colspan="2"><U><B>Select the SMS files you want to Generate</B></U></td>
			</tr>
			<tr>
				<td><input type="checkbox" disabled name="stockWatch" value="1" class="borderless">&nbsp;Stock Watch</td><td></td>
			</tr>
			<tr>
				<td><input type="checkbox" disabled name="Contracts" value="1" class="borderless">&nbsp;Contracts</td><td></td>
			</tr>
			<tr>
				<td><input type="checkbox" checked name="Debits" value="1" class="borderless">&nbsp;Debits</td><td></td>
			</tr>
			<tr>
				<td colspan="2"><input type="button" value="Generate" id="btnGenerate" name="btnGenerate" title="Generate SMS Files" onclick="JavaScript: Generate();"></td>
			</tr>
			<!--<tr>
				<td colspan="2"><input type="button" value="Download" id="btnDownload" name="btnDownload" title="Download SMS Files" onclick="JavaScript: Download();"></td>
			</tr>-->
		</table>
	</form>
	<%
Else
%>
<form method="post" action="" name="frmMain" id="frmMain">
<%
  '=========================================================================================
  ' Generate sms
  '=========================================================================================
	stockWatch = trim(Request.QueryString("stockWatch"))
	If stockWatch="" Then stockWatch=0
		
	Contracts = trim(Request.QueryString("Contracts"))
	If Contracts ="" Then Contracts=0
	
	Debits = trim(Request.QueryString("Debits"))
	If Debits ="" Then Debits =0
	
	CanSendSMS = trim(Request.QueryString("CanSendSMS"))
	If CanSendSMS ="" Then CanSendSMS =0

	''create the server objects here 
	Set rs=server.createobject("Adodb.recordset")
	Set rs1=server.createobject("Adodb.recordset")

	Set Conn = GetActiveConnection("KBroker")	
	
	'Check for debits
	If Debits =1 Then
	     
		'get all the list of clients with mobile numbers
		Set rs=server.createobject("Adodb.recordset")
		Set rs1=server.createobject("Adodb.recordset")

		Set Conn = GetActiveConnection("KBroker")
		
		if CanSendSMS = 0 then

			sqlStr= "SELECT * FROM smsDebtors WHERE (ClientCellTel IS Not NULL) AND (ClientCellTel <> N'') AND (updateOnDebt=1)"
			
			rs.open sqlStr, conn,0,1
			
			If rs.EOF Or rs.BOF Then
				%>
				<SCRIPT LANGUAGE="JavaScript">
					alert("No Debtor Information To SMS");
					window.location.href='DebitSMS.asp'
				</SCRIPT>
				<%
				'Response.End
			Else
				smsStr = ""
				rs.movefirst

				Response.Write "<br>&nbsp;<br><b>Debit List</b><br>"
						%>
						<br><br>

						<table border=1 cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111" width="94%">	
						<tr>
								<td width="5%">&nbsp;<b>Code</b></td>
								<td width="26%">&nbsp;<b>Name</b></td>
								<td width="12%">&nbsp;<b>Cell No.</b></td>
								<td width="43%">&nbsp;<b>Message</b></td>
								<td width="15%">&nbsp;<b>To SMS</b></td>
						</tr>
						<%
				NoSms = 0

				While Not rs.EOF 
						'create a string for the sms here
						' format the celltell number
						
						if len(rs("ClientCellTel")) > 0 then
						cellNo = replace(rs("ClientCellTel")," ","")
						cellNo = replace(rs("ClientCellTel"),"-","")
						else
						cellNo = ""
						end if
								
						

						if cellNo <> "" then
						     'Validate Cell no
							 If mid(cellNo,1,1) = 0 Then cellNo = "" 
							 If len(cellNo) < 9  Then cellNo = "" 
							
							If Not isnumeric(cellNo)  Or trim(cellNo)="" Then
								cellNo = cellNo & " *"
								ValidCellNo = false
							Else
								ValidCellNo = true
							End If
						Else
							ValidCellNo = false
						End if

						smsStr = "Please note that you have an outstanding balance of KES " & formatnumber(0-rs("Balance"),2) & " overdue. Please settle."& vbCrLf

						'tabulate the Debit list
									%>
									<tr>
										<td width="5%">&nbsp;<%=rs.fields("Client_DPA_")%></td>
										<td width="26%">&nbsp;<%=mid(rs.fields("ClientName"),1,30)%></td>
										<td width="12%">&nbsp;<%=cellNo%></td>
										<td width="43%" nowrap>&nbsp;<%=smsStr%></td>
										<td width="15%">&nbsp;
										<%
											 If Len(cellNo) > 0 AND ValidCellNo = true Then
											   NoSms = NoSms + 1
											 %> 
											   <input type="checkbox" name="Debit<%=NoSms%>" ID="Debit<%=NoSms%>" value="<%=rs.fields("Client_DPA_")%>" class="borderless" checked>
											 <%
											Else	
												   response.write "Invalid mobile number."
											End If
										%>
										</td>
								   </tr>
								  <%
								  smsStr  = ""
					rs.movenext
				Wend
				%>
					</table>
					<%		 
				rs.close
				conn.close
						 
			End If 'rs.EOF
        Else
		' Send SMS

		 data = trim(Request.QueryString("data"))

		if data <> "" then
		 con = split(data,";")
		 conbound = ubound(con)
         
		 for i = 0 to conbound

          if i = 0 then
		   if isnumeric(trim(con(i)))AND trim(con(i)) <> "" then
		      ClientID = con(i)
		    end if
		  else
           if isnumeric(trim(con(i)))AND trim(con(i)) <> "" then
		      if ClientID = "" then
		       ClientID =  con(i) 
		      else
		       ClientID = ClientID & "," & con(i) 
		      end if 
		    end if
		  end if

		 next
         
         if ClientID = "" then ClientID = 0
         ClientSQL = " AND Client_DPA_ IN  (" & ClientID & ")  "
         
         sqlStr= "SELECT * FROM smsDebtors WHERE (ClientCellTel IS Not NULL) AND (ClientCellTel <> N'') AND (updateOnDebt=1) " & ClientSQL
			
			rs.open sqlStr, conn,0,1
			
			If rs.EOF Or rs.BOF Then
				%>
				<SCRIPT LANGUAGE="JavaScript">
					alert("No Debtor Information To SMS");
					window.location.href='DebitSMS.asp'
				</SCRIPT>
				<%
				'Response.End
			Else
				smsStr = ""
				rs.movefirst
               smscount = 0
				While Not rs.EOF 
						'create a string for the sms here
						' format the celltell number
						
						
						if len(rs("ClientCellTel")) > 0 then
						cellNo = replace(rs("ClientCellTel")," ","")
						cellNo = replace(rs("ClientCellTel"),"-","")
						else
						cellNo= ""
						end if
								
						if cellNo <> "" then
							 'Validate Cell no
							 If mid(cellNo,1,1) = 0 Then cellNo = "" 
							 If len(cellNo) < 9  Then cellNo = "" 
							
							If Not isnumeric(cellNo)  Or trim(cellNo)="" Then
								cellNo = cellNo & " *"
								ValidCellNo = false
							Else
								ValidCellNo = true
							End If
						Else
							ValidCellNo = false
						End if

						smsStr = "Please note that you have an outstanding balance of KES " & formatnumber(0-rs("Balance"),2) & " overdue. Please settle."& vbCrLf

						'Send SMS if valid Cell Number
						if  ValidCellNo = true then
								  
									'Store sms message in form element
								   %>
								   <input type="hidden" name="sMessage<%=rs("Client_DPA_")%>" id="sMessage<%=rs("Client_DPA_")%>" value="<%=smsStr%>">
								   <%
										 'Send Sms here
										 'cellNo = 254721613909 'Hard Code
											 %>
												 <script language="javascript">
												 var msgStatus, replyText 
												   msgStatus = SendSms(<%=cellNo%>,'sMessage<%=rs("Client_DPA_")%>')
                                                   
												   //Handle any errors encountered
                                                   replyText = msgStatus.substr(0,2)
			                                       
												   if (replyText.toUpperCase() == 'ID'){
												      //message sent successfully
													  
												   }else{
												    //errors encountered
													alert('Unexpected error occurred.\n ' + msgStatus);
													window.location.href='stockWatchSMS.asp'
												   }

												 </script>
											 <%
											 smsStr = ""
									smscount = smscount + 1
						  end if
					rs.movenext
				Wend
				 
				rs.close
				conn.close
						 
			End If 'rs.EOF
			 
			End if 'data <> ""
			
%>
<script language="javascript">
	alert('<%=smscount%> SMS Text Message(s) sent successfully.')
	window.location.href='DebitSMS.asp'
</script>
<%
		End If 'if CanSendSMS = 0 then
	End If 'Debits = 1

	%>
	<BR>
	<input type ="button" name ="back" value="<< Back" onclick ="javascript: window.location.href='DebitSMS.asp'">
	<input type="hidden" name="NoSms" id="NoSms" value="<%=NoSms%>">
	<input type ="button" name ="SMSData" id ="SMSData" value="Send SMS" onclick="javascript: ValidateSMS()">
	
	<%
End If
%>
</form>
</BODY>
</HTML>
