<!--#include file="../libroutines.asp"-->
<HTML>
<HEAD>
<TITLE> SMS </TITLE>
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css">
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT>

<SCRIPT Language="JavaScript">
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
		window.location.replace("stockWatchSMS.asp?gen=1&stockWatch="+stockWatch+"&Contracts="+Contracts+"&Debits="+Debits);
	}
	
	function Download(){
		window.location.replace("downloadSMS.asp");
	}
	
	function ValidateSMS(){
	var recordcount = document.frmMain.NoSms.value;
	var i,j, m;
	var data = '';
	var elem
	var val 
		
		 if (recordcount!=''){
		         m=0;
			   for (i = 0; i < recordcount; i++){
			    j = i + 1
				  elemName = "StockWatch" + j
				  
				  elem = document.getElementById(elemName)
				  
					 if (elem.checked==true){
					 val  =  elem.value
					 
					    if (isNaN(parseInt(val))){
						 }else{
							  if (m==0){
								data = elem.value;
							  }else {
								data = data + ';' + elem.value;
							  }
							  m =1;
						 }  
					 }
			   }
		  }
		 
		 //alert(data);
		
			data = ''; 
			for(i=0; i<document.frmMain.elements.length; i++)
			{
				if (document.frmMain.elements[i].type=="checkbox")
				{
					if (document.frmMain.elements[i].checked==true)
					{
						if (document.frmMain.elements[i].value!='')
						{
						data = data + ';' + document.frmMain.elements[i].value;
						}
					}
				}
			}
		
		 //alert(data);
		 
		 window.location.replace("stockWatchSMS.asp?gen=1&stockWatch=1&CanSendSMS=1&owners=1&data="+data);
	} 
</SCRIPT>

</HEAD>

<BODY>
	
<%
gen = trim(Request.QueryString("gen"))

If gen = "" Then gen=0
	
If gen <> 1 Then
	%>
	<form method="post" action="stockWatchSMS.asp" name="FrmSMS">
		<input type="hidden" name="hidData" id="hidData">
		<table width="50%">
			<tr>
				<td>&nbsp;<BR></td> <td></td>
			</tr>
			<tr>
				<td colspan="2"><U><B>Select the SMS files you want to Generate</B></U></td>
			</tr>
			<tr>
				<td><input type="checkbox" checked name="stockWatch" value="1" class="borderless">&nbsp;Stock Watch</td><td></td>
			</tr>
			<tr>
				<td><input type="checkbox" disabled name="Contracts" value="1" class="borderless">&nbsp;Contracts</td><td></td>
			</tr>
			<tr>
				<td><input type="checkbox" disabled name="Debits" value="1" class="borderless">&nbsp;Debits</td><td></td>
			</tr>
			<tr>
				<td colspan="2"><input type="button" value="Generate" id="btnGenerate" name="btnGenerate" title="Generate SMS Files" onclick="JavaScript: Generate();"></td>
			</tr>
			<!--
			<tr>
				<td colspan="2"><input type="button" value="Download" id="btnDownload" name="btnDownload" title="Download SMS Files" onclick="JavaScript: Download();"></td>
			</tr>-->
		</table>
	</form>
	<%
Else
%>
<form method="post" action="" name="frmMain" id="frmMain">
<%
'==================================================================================
'Generate stock watch list or sms stock watch list
'==================================================================================

	stockWatch = trim(Request.QueryString("stockWatch"))
	If stockWatch="" Then stockWatch=0
		
	Contracts = trim(Request.QueryString("Contracts"))
	If Contracts ="" Then Contracts=0
	
	Debits = trim(Request.QueryString("Debits"))
	If Debits ="" Then Debits =0

	CanSendSMS = trim(Request.QueryString("CanSendSMS"))
	If CanSendSMS ="" Then CanSendSMS =0
	
	'create the server objects here 
	Set rs=server.createobject("Adodb.recordset")
	Set rs1=server.createobject("Adodb.recordset")

	Set Conn = GetActiveConnection("KBroker")	

	'generate all urs stuff here, first check the files one wants to generate
	If stockWatch = 1 Then

      if CanSendSMS = 0 then
			 'get all the list of clients with mobile numbers
			 sqlStr ="SELECT DISTINCT TOP 100 PERCENT Client_DPA_,ClientCellTel, ClientName FROM StockWatchList order by client_DPA_"
			
			 rs.open sqlStr, conn,0,1
							
			 If rs.EOF Or rs.BOF Then
					%>
					<SCRIPT LANGUAGE="JavaScript">
						alert("No Stock Watches To SMS");
						window.location.href='StockWatchSMS.asp'
					</SCRIPT>
					<%
				Else
					'check the last time the stocks were updated 
					Set equityRs = server.createObject("Adodb.recordset")

					equityRs.open "SELECT MAX(MktDate) AS LastImportDate FROM datastream_Market ", conn, 0,1

					If Not rs.EOF Or Not rs.BOF Then
						lastModifiedDate = formatdate(equityRs("LastImportDate"))
					End If

					equityRs.Close
					Set equityRs= Nothing

					' give an alert on the date the last import happened
					%>
					<SCRIPT LANGUAGE="JavaScript">
						alert('The last importation of the Price List was on <%=lastModifiedDate%>');
					</SCRIPT>
					<%
				
					rs.movefirst
					
					Response.Write "<br>&nbsp;<br><b>Stock Watch List</b><br>"
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
					DefaultMsgExists = false
					While Not rs.EOF 
							'create a string for the sms here
							clientNo = rs("Client_DPA_")
															
							sqlstr1 ="select * from StockWatchList where client_DPA_ =" & rs("Client_DPA_")

							set rs1 = conn.execute(sqlstr1) 
											
							i=0
							secStr =""
							While Not rs1.EOF 
									'get all the securities for the selected client
									If  trim(secStr) = "" Then
										secStr = rs1("SecurityCode") & " " & formatnumber(rs1("Price"),2)	
									Else
										'check If the characters are greater than 150

									  secStr = secStr & " : " & rs1("SecurityCode")	& " " & formatnumber(rs1("Price"),2)
									End If

								rs1.movenext
							Wend
							  
							rs1.close
							set rs1 = nothing
							
							  ' format the celltell number
								if len(rs("ClientCellTel")) > 0 then
								cellNo = replace(rs("ClientCellTel")," ","")
								else
								cellNo= ""
								end if
								'cellNo = replace(rs("ClientCellTel"),"-","")

								if cellNo <> "" then
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

								smsStr = secStr
                                DefaultMsgExists = true
								 'tabulate the stock watch list
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
										   <input type="checkbox" name="StockWatch<%=NoSms%>" ID="StockWatch<%=NoSms%>" value="<%=rs.fields("Client_DPA_")%>" class="borderless" checked>
										 <%
									    Else	
											   response.write "Invalid mobile number."
										End If
									%>
									</td>
									
							   </tr>
							  <%
					smsStr = ""
						rs.movenext
					Wend
					set rs  = nothing
                    
					'SMS Account Managers
					SQLStr = " Select * from ownerlist where sendSMS = 1 AND MobileNo IS NOT NULL  "
                    
					set rs = conn.execute(SQLStr)

					if not (rs.eof or rs.bof) then
					 rs.movefirst
                    Owners = "" 'Mark that there are owners to be smsed
					NoOwners = 0
					 do while not rs.EOF
					    Owners = 1
                        ' format the celltell number
								
								
								if len(rs("MobileNo")) > 0 then
								cellNo = replace(rs("MobileNo")," ","")
								cellNo = replace(cellNo,"-","")
								else
								cellNo= ""
								end if
						
								if cellNo <> "" then
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
								
							'Determine default message
							if DefaultMsgExists then
							 ownersms = smsStr
							else
                                  sqlstr1 ="select  TOP 1 * from StockWatchList "

								  set rs1= conn.execute(sqlstr1) 
													
									i=0
									secStr =""
									While Not rs1.EOF 
											'get all the securities for the selected client
											If  trim(secStr) = "" Then
												secStr = rs1("SecurityCode") & " " & formatnumber(rs1("Price"),2)	
											Else
												'check If the characters are greater than 150

											  secStr = secStr & " : " & rs1("SecurityCode")	& " " & formatnumber(rs1("Price"),2)
											End If

										rs1.movenext
									Wend
									  
									rs1.close
									set  rs1 = nothing
									ownersms = secStr
							end if
					 %>
					  <tr>
									<td width="5%">&nbsp;Default</td>
									<td width="26%">&nbsp;<%=mid(rs.fields("OwnerName"),1,30)%></td>
									<td width="12%">&nbsp;<%=cellNo%></td>
									<td width="43%" nowrap>&nbsp;<%=ownersms%></td>
									<td width="15%">&nbsp;
									<%
										 If Len(cellNo) > 0 AND ValidCellNo = true Then
										   NoOwners = NoOwners + 1
										 %> 
										   <input type="checkbox" disabled name="Stock<%=NoOwners%>" ID="Stock<%=NoOwners%>"value="" class="borderless" checked>
										 <%
									    Else	
											   response.write "Invalid mobile number."
										End If
									%>
									</td>
									
							   </tr>
					 <%
					 ownersms = ""
					 rs.movenext
					 Loop
					end if
					
					%>
					</table>
					<%
					Set rs=Nothing
					conn.close
				End If 'rs.EOF

		Else

		'Send SMS
		data = trim(Request.QueryString("data"))
		owners = trim(Request.QueryString("owners")) 'Check if there are any account managers to be smsed
        If owners ="" Then owners =0
		DefaultMsgExists = false
		smscount = 0
		
		data= mid(data,2)
		
		if data <> "" then
			ClientID = ""
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
         ClientSQL = " Client_DPA_ IN  (" & ClientID & ")  "
         
		 'get all the list of clients with mobile numbers
			 sqlStr ="SELECT DISTINCT TOP 100 PERCENT Client_DPA_,ClientCellTel, ClientName FROM StockWatchList Where " & ClientSQL & " order by client_DPA_"
				
				'Response.Write sqlStr
				'Response.End 
				
			 rs.open sqlStr, conn,0,1
					smscount = 0		
				If rs.EOF Or rs.BOF Then
					%>
					<SCRIPT LANGUAGE="JavaScript">
						alert("No Stock Watches To SMS");
						window.location.href='StockWatchSMS.asp'
					</SCRIPT>
					<%
				Else
				 'Process sms
				 rs.movefirst
					
					While Not rs.EOF 
							'create a string for the sms here
							clientNo = rs("Client_DPA_")
															
							sqlstr1 ="select * from StockWatchList where client_DPA_ =" & rs("Client_DPA_")

							rs1.open sqlstr1, conn, 0, 1
											
							i=0
							secStr =""
							While Not rs1.EOF 
									'get all the securities for the selected client
									If  trim(secStr) = "" Then
										secStr = rs1("SecurityCode") & " " & formatnumber(rs1("Price"),2)	
									Else
										'check If the characters are greater than 150

									  secStr = secStr & " : " & rs1("SecurityCode")	& " " & formatnumber(rs1("Price"),2)
									End If

								rs1.movenext
							Wend
							  
							rs1.close
							
							 ' format the celltell number
								
								if len(rs("ClientCellTel")) > 0 then
								cellNo = replace(rs("ClientCellTel")," ","")
								cellNo = replace(rs("ClientCellTel"),"-","")
								else
								cellNo= ""
								end if
																
								if cellNo <> "" then
								     'Validate cell no
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

								smsStr = secStr
								DefaultMsgExists = true
                                
								if  ValidCellNo = true then
								  
								  'Response.Write rs("Client_DPA_") & "<br>"
								  'Response.End 
									'Store sms message in form element
								   %>
								   <input type="hidden" name="sMessage<%=rs("Client_DPA_")%>" id="sMessage<%=rs("Client_DPA_")%>" value="<%=smsStr%>">
								   <%
										 'Send Sms here
										 
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

					Set rs=Nothing
					'conn.close
				End If 'rs.EOF
          
          End if ' data <> ""
			
		  if owners  =  1 then
               'SMS Account Managers
					SQLStr = " Select * from ownerlist where sendSMS = 1 AND MobileNo IS NOT NULL  "
                    
					set rs = conn.execute(SQLStr)

					if not (rs.eof or rs.bof) then
					 rs.movefirst
                    
					 do while not rs.EOF
                        ' format the celltell number
								if cellNo <> "" then
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
							
							DefaultMsgExists = false
				
							'Determine default message
							if DefaultMsgExists then
							 ownersms = smsStr
							else
                                  sqlstr1 ="select  TOP 1 * from StockWatchList "

								  set rs1= conn.execute(sqlstr1) 
													
									i=0
									secStr =""
									While Not rs1.EOF 
											'get all the securities for the selected client
											If  trim(secStr) = "" Then
												secStr = rs1("SecurityCode") & " " & formatnumber(rs1("Price"),2)	
											Else
												'check If the characters are greater than 150

											  secStr = secStr & " : " & rs1("SecurityCode")	& " " & formatnumber(rs1("Price"),2)
											End If

										rs1.movenext
									Wend
									  
									rs1.close
									set  rs1 = nothing
									ownersms = secStr
							end if
							'ready to sms
                            %>
								   <input type="hidden" name="sMessage<%=rs("Owner_DPA_")%>" id="sMessage<%=rs("Owner_DPA_")%>" value="<%=ownersms%>">
								   <%

							If Len(cellNo) > 0 AND ValidCellNo = true Then
                              
                                 %>
									 <script language="javascript">
										 var msgStatus, replyText 
										   msgStatus = SendSms(<%=cellNo%>, 'sMessage<%=rs("Owner_DPA_")%>')
										   
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
								ownersms =""
								smscount = smscount + 1
							end if
					 
					 rs.movenext
					 Loop
					end if
		  end if

%>
<script language="javascript">
	alert('<%=smscount%> SMS Text Message(s) sent successfully.')
	window.location.href='stockWatchSMS.asp'
</script>
<%
        End if 'Can Send SMS = 0

		 
	End If

	%>
	<BR>
	<input type="hidden" name="NoSms" id="NoSms" value="<%=NoSms%>">
	<input type="hidden" name="txtOwners" id="txtOwners" value="<%=Owners%>">
	<input type ="button" name ="back" value="<< Back" onclick ="javascript: window.location.href='stockWatchSMS.asp'">
	<input type ="button" name ="SMSData" id ="SMSData" value="Send SMS" onclick="javascript: ValidateSMS()">
	</form>
	<%
End If
%>

</BODY>
</HTML>