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
		escalate = 3
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
			return (e.number + ': ' + e.description)
		}
        
		xmlhttp.setRequestHeader("Content-Type","application/x-www-form-urlencoded");
		xmlhttp.send(sPostData) ;
        
		sResult = xmlhttp.responseText;
        return (sResult) 
        
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
					
		window.location.replace("contractSMS.asp?gen=1&stockWatch="+stockWatch+"&Contracts="+Contracts+"&Debits="+Debits);
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
				  elemName = "Contract" + j
				  
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
		   
         window.location.replace("contractSMS.asp?gen=1&Contracts=1&CanSendSMS=1&owners=1&data="+data);
         	 
	} 

</SCRIPT>
</HEAD>

<BODY>
	
<%
gen = trim(Request.QueryString("gen"))

If gen = "" Then gen=0
	
If gen <> 1 Then
	%>
	<form method="post" action="contractSMS.asp" name="FrmSMS">
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
				<td><input type="checkbox" checked name="Contracts" value="1" class="borderless">&nbsp;Contracts</td><td></td>
			</tr>
			<tr>
				<td><input type="checkbox" disabled name="Debits" value="1" class="borderless">&nbsp;Debits</td><td></td>
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
	
	'create the server objects here 
	Set rs=server.createobject("Adodb.recordset")
	Set rs1=server.createobject("Adodb.recordset")

	Set Conn = GetActiveConnection("KBroker")	

	'Check whether the Contracts has been selected
	If Contracts = 1 Then	
		'get all the list of clients with mobile numbers
		Set rs=server.createobject("Adodb.recordset")
		Set rs1=server.createobject("Adodb.recordset")

		Set Conn = GetActiveConnection("KBroker")

		if CanSendSMS = 0 then
           
		   sqlStr= "SELECT * FROM ShowSMSAllContractsCompounded"
		   
			rs.open sqlStr, conn,0,1
				
			If rs.EOF Or rs.BOF Then
				%>
				<SCRIPT LANGUAGE="JavaScript">
					alert("No Contracts To SMS");
					window.location.href='ContractSMS.asp';
				</SCRIPT>
				<%
			Else
					rs.movefirst

					Response.Write "<br>&nbsp;<br><b>Contracts List</b><br>"
					%>
					<br><br>
					<table border=1 cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">	
					<tr>
							<td width="5%">&nbsp;<b>Code</b></td>
							<td width="20%">&nbsp;<b>Name</b></td>
							<td width="5%">&nbsp;<b>Cell No.</b></td>
							<td width="60%">&nbsp;<b>Message</b></td>
							<td width="10%">&nbsp;<b>To SMS</b></td>
					</tr>
					<%
					NoSms = 0

					While Not rs.EOF 
							'Tabulate the sms's here

							' format the celltell number
								IF LEN(rs("ClientCellTel")) > 0 THEN
								cellNo = replace(rs("ClientCellTel")," ","")
								cellNo = replace(rs("ClientCellTel"),"-","")
								ELSE
								cellNo = ""
								END IF
								'Validate Cell no
								
								If Not isnumeric(cellNo)  Or trim(cellNo)= ""  Then
									cellNo = cellNo & " *"
									ValidCellNo = false
									msg = "Invalid mobile number"
								Else
								  if mid(cellNo,1,1) = 0 Or len(cellNo) < 9 then
								    cellNo = cellNo & " *"
									ValidCellNo = false
									msg = "Invalid mobile number"
								  Else
									ValidCellNo = true
									msg = ""
								  End if
								End If
							
							If Len(cellNo) > 0   Then
								
									'If it is a purchase contract Then
									If rs("OrderType_DPA_")=1 Then
										smsStr = formatnumber(rs("LotQty"),0) & " " & rs("SecurityCode")& " SHARES PURCHASED ON " & formatdate(rs("lotTdate")) & " @ KES " & formatnumber(rs("LotPrice"),2) &" PER SHARE. SETTLEMENT OF KES " & formatnumber(rs("contractAmount"),2) & " DUE " & formatdate(rs("contractSettlementDate"))& ". " & vbCrLf
									Else
										smsStr = formatnumber(rs("LotQty"),0) & " " & rs("SecurityCode")& " SHARES SOLD ON "& formatdate(rs("lotTdate")) &" @ KES " & formatnumber(rs("LotPrice"),2) & " PER SHARE. " & vbCrLf
									End If
									
								DefaultMsgExists = true
								
							If trim(rs("updateOnContract")) Then
							 updateOnContract = true
							else
							 updateOnContract = false
							end if 
							
		 
							%>
							<tr>
							<td width="5%">&nbsp;<%=rs.fields("Client_DPA_")%>&nbsp;</td>
							<td width="20%">&nbsp;<%=mid(rs.fields("ClientName"),1,30)%>&nbsp;</td>
							<td width="5%">&nbsp;<%=cellNo%>&nbsp;</td>
							<td width="60%" >&nbsp;<%=smsStr%>&nbsp;</td>
							<td width="10%">
                             <%
                             
                               if updateOnContract then
                           
                                 if ValidCellNo  then
									 NoSms = NoSms + 1
									 %> 
									<input type="checkbox" name="Contract<%=NoSms%>" ID="Contract<%=NoSms%>" value="<%=rs.fields("Contract_DPA_")%>" class="borderless" checked>
									 <%
								 else
									response.write msg
									 
							      End if 
                              
                               end if
                             
							   
							 %>
							</td>
							</tr>
							<%
						  
							End If
							smsStr = ""
						rs.movenext
						
					Wend
				
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
								cellNo = replace(rs("MobileNo")," ","")
								cellNo = replace(cellNo,"-","")

								if cellNo <> "" then 
									
									If Not isnumeric(cellNo)  Or trim(cellNo)= ""  Then
										cellNo = cellNo & " *"
										ValidCellNo = false
										msg = "Invalid mobile number"
									Else
										if mid(cellNo,1,1) = 0 Or len(cellNo) < 9 then
										  cellNo = cellNo & " *"
											ValidCellNo = false
											msg = "Invalid mobile number"
										Else
											ValidCellNo = true
											msg = ""
										End if
									End If
									
								Else
									ValidCellNo = false
								End if
								
							'Determine default message
							if DefaultMsgExists then
							 ownersms = smsStr
							else
                                  sqlstr1 = " select  TOP 1 * from ShowSMSAllContractsCompounded "

								  set rs1= conn.execute(sqlstr1) 
													
									i=0
									secStr =""
									While Not rs1.EOF 
											
											'If it is a purchase contract Then
											If rs1("OrderType_DPA_")=1 Then
												secStr = formatnumber(rs1("LotQty"),0) & " " & rs1("SecurityCode")& " SHARES PURCHASED ON " & formatdate(rs1("lotTdate")) & " @ KES " & formatnumber(rs1("LotPrice"),2) &" PER SHARE. SETTLEMENT OF KES " & formatnumber(rs1("contractAmount"),2) & " DUE " & formatdate(rs1("contractSettlementDate"))& ". " & vbCrLf
											Else
												secStr = formatnumber(rs1("LotQty"),0) & " " & rs1("SecurityCode")& " SHARES SOLD ON "& formatdate(rs1("lotTdate")) &" @ KES " & formatnumber(rs1("LotPrice"),2) & " PER SHARE. " & vbCrLf
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
									<td width="43%">&nbsp;<%=ownersms%></td>
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
            End if
        Else
        

''SENDING THE SMS
data = trim(Request.QueryString("data"))
owners = trim(Request.QueryString("owners")) 'Check if there are any account managers to be smsed

If owners = "" Then owners =0
        
DefaultMsgExists = false
smscount = 0
ContractID = ""

''ORDINARY
if data <> "" then
	con = split(data,";")
	conbound = ubound(con)
	         
	for i = 0 to conbound
		if i = 0 then 
			if isnumeric(trim(con(i)))AND trim(con(i)) <> "" then
				ContractID = con(i)
			end if
		else
			if isnumeric(trim(con(i)))AND trim(con(i)) <> "" then
				if ContractID = "" then
					ContractID = con(i)
				else
					ContractID = ContractID & "," & con(i)
				end if
			end if
		end if
	next
     
	if ContractID <> "" then
		ContractSQL = " Contract_DPA_ IN (" & ContractID & ")"

		sqlStr= "SELECT * FROM ShowSMSAllContractsCompounded WHERE  " & ContractSQL
		rs.open sqlStr, conn,0,1
						
		If rs.EOF Or rs.BOF Then
			%>
			<SCRIPT LANGUAGE="JavaScript">
				alert("No Contracts To SMS");
				window.location.href='ContractSMS.asp';
			</SCRIPT>
			<%
		Else
			rs.movefirst
								    
			smscount = 0
			
			While Not rs.EOF 
					'FORMAT THE CELL PHONE NUMBER
					cellNo = replace(rs("ClientCellTel")," ","")
					cellNo = replace(rs("ClientCellTel"),"-","")
												
					'VALIDATE CELL NO
					If Not isnumeric(cellNo)  Or trim(cellNo)=""  Then
						cellNo = cellNo & " *"
						CorrectCellNo = false
					Else
						if mid(cellNo,1,1) = 0 Or len(cellNo) < 9 then
							cellNo = cellNo & " *"
							CorrectCellNo = false
						else
							CorrectCellNo = true
						end if
					End If
					                            
					If trim(rs("updateOnContract")) Then
						updateOnContract = true
					else
						updateOnContract = false
					end if 
											
					If Len(cellNo) > 0 AND CorrectCellNo  AND updateOnContract Then
						NoSms = NoSms + 1
						cellNo = Replace(cellNo," ","")
															
						'If it is a purchase contract Then
						If rs("OrderType_DPA_")=1 Then
							smsStr =smsStr & formatnumber(rs("LotQty"),0) & " " & rs("SecurityCode")& " shares purchased on " & formatdate(rs("lotTdate")) & " @ KES " & formatnumber(rs("LotPrice"),2) &" per share. Settlement of KES " & formatnumber(rs("contractAmount"),2) & " due " & formatdate(rs("contractSettlementDate"))& ". " & vbCrLf
							OwnerSMSstr = OwnerSMSstr & formatnumber(rs("LotQty"),0) & " " & rs("SecurityCode")& " purchased @KES " & formatnumber(rs("LotPrice"),2) &"." & vbCrLf
						Else
							smsStr =smsStr & formatnumber(rs("LotQty"),0) & " " & rs("SecurityCode")& " shares sold on "& formatdate(rs("lotTdate")) &" @ KES " & formatnumber(rs("LotPrice"),2) & " per share. " & vbCrLf
							OwnerSMSstr = OwnerSMSstr & formatnumber(rs("LotQty"),0) & " " & rs("SecurityCode")& " sold @KES " & formatnumber(rs("LotPrice"),2) &"." & vbCrLf
						End If
					                                
						'Response.write  smsStr & "<br>"
						
						If Len(smsStr) > 0 Then
							%>
							<input type="hidden" name="sMessage<%=Rs("Client_DPA_")&Rs("Contract_DPA_")%>" id="sMessage<%=Rs("Client_DPA_")&Rs("Contract_DPA_")%>" value="<%=smsStr%>">
							<%

							%>
							<script language="javascript">
								var msgStatus, replyText 
								msgStatus = SendSms(<%=cellNo%>, 'sMessage<%=Rs("Client_DPA_")&Rs("Contract_DPA_")%>')
								                                                   
								//Handle any errors encountered
								replyText = msgStatus.substr(0,2)
											                                       
								if (replyText.toUpperCase() == 'ID')
								{
								//message sent successfully
								}
								else
								{
								//errors encountered
								alert('Unexpected error occurred.\n ' + msgStatus);
								window.location.href='contractSMS.asp'
								}
							</script>
							<%
						End If		 		 
						smsStr = ""
						smscount = smscount + 1
					End If
				rs.movenext
			Wend
		End if
	end if
End if 

''OWNERS		
if owners  =  1 then
	SQLStr = " Select * from ownerlist where sendSMS = 1 AND MobileNo IS NOT NULL  "
	set rs = conn.execute(SQLStr)

	if not (rs.eof or rs.bof) then
		rs.movefirst
		do while not rs.EOF
				'FORMAT THE CELLPHONE NO
				cellNo = replace(rs("MobileNo")," ","")
				cellNo = replace(cellNo,"-","")
		
				if cellNo <> "" then 
					If Not isnumeric(cellNo)  Or trim(cellNo)= ""  Then
						cellNo = cellNo & " *"
						ValidCellNo = false
						msg = "Invalid mobile number"
					Else
						if mid(cellNo,1,1) = 0 Or len(cellNo) < 9 then
							cellNo = cellNo & " *"
							ValidCellNo = false
							msg = "Invalid mobile number"
						Else
							ValidCellNo = true
							msg = ""
						End if
					End If
				Else
					ValidCellNo = false
				End if
				
				DefaultMsgExists = false
				smsStr = Replace(mid(OwnerSMSstr,1,160),",","")
				
				'Determine default message
				if DefaultMsgExists then
					ownersms = smsStr
				else
					sqlstr1 = " select  TOP 1 * from ShowSMSAllContractsCompounded "
					set rs1= conn.execute(sqlstr1) 
																		
					i=0
					secStr =""
					
					While Not rs1.EOF 
							'If it is a purchase contract Then
							If rs1("OrderType_DPA_")=1 Then
								secStr = formatnumber(rs1("LotQty"),0) & " " & rs1("SecurityCode")& " shares purchased on " & formatdate(rs1("lotTdate")) & " @ KES " & formatnumber(rs1("LotPrice"),2) &" per share. Settlement of KES " & formatnumber(rs1("contractAmount"),2) & " due " & formatdate(rs1("contractSettlementDate"))& ". " & vbCrLf
							Else
								secStr = formatnumber(rs1("LotQty"),0) & " " & rs1("SecurityCode")& " shares sold on "& formatdate(rs1("lotTdate")) &" @ KES " & formatnumber(rs1("LotPrice"),2) & " per share. " & vbCrLf
							End If

						rs1.movenext
					Wend
														  
					rs1.close
					set  rs1 = nothing
					
					ownersms = secStr
				end if
				
				'Response.write  smsStr & "<br>"
				
				If Len(ownersms) > 0 Then
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
											                                       
							if (replyText.toUpperCase() == 'ID')
							{
							//message sent successfully
							}
							else
							{
							//errors encountered
							alert('Unexpected error occurred.\n ' + msgStatus);
							window.location.href='contractSMS.asp'
							}
						</script>
						<%
						ownersms = ""
						
						smscount = smscount + 1
					End If
				End If
			rs.movenext
		Loop
	end if
end if
		  
%>
<script language="javascript">
	alert('<%=smscount%> SMS Text Message(s) sent successfully.')
	window.location.href='contractSMS.asp'
</script>
<%
             
End if

End If
	%>
	<BR>
	<input type ="button" name ="back" value="<< Back" onclick ="javascript: window.location.href='contractSMS.asp'">
	<input type="hidden" name="NoSms" id="NoSms" value="<%=NoSms%>">
	<input type ="button" name ="SMSData" id ="SMSData" value="Send SMS" onclick="javascript: ValidateSMS()">
	</form>
	<%
End If
%>

</BODY>
</HTML>