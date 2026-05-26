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

		try{
		   xmlhttp.open("POST", sUrl, false)
		}catch(e){
		   return ( e.number + ': ' + e.description)
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
					
		window.location.replace("SMS.asp?gen=1&stockWatch="+stockWatch+"&Contracts="+Contracts+"&Debits="+Debits);
	}
	
	function Download(){
		window.location.replace("downloadSMS.asp");
	}
	
	function ValidateSMS(){
	
	var debitcount = document.frmMain.DebitNoSms.value;
	var stockcount = document.frmMain.StockNoSms.value;
	var contractcount = document.frmMain.ContractNoSms.value;
	var i,j, m;
	var datadebit, datacontract, datastock;
	var elem;
	var GenDebits = document.frmMain.GenDebits.value; 
	var GenContract = document.frmMain.GenContract.value; 
	var Genstock = document.frmMain.Genstock.value;
      
		if (GenDebits==1) {
		 if (debitcount!=''){
		      m = 0 ;
				   for (i = 0; i < debitcount; i++){
				    j = i + 1
					  elemName = "Debit" + j
					  
					  elem = document.getElementById(elemName)
					  
						 if (elem.checked==true){
						
						    if (isNaN(parseInt(elem.value))){
							 }else{
								  if (m==0){
									datadebit = elem.value;
								  }else {
									datadebit = datadebit + ';' + elem.value;
								  }
								 m = 1 
							 }
							  
						 }
						 
				   }
			  }
		  }
		
		 if (Genstock==1){
		  if (stockcount!=''){
		     m = 0 ;
				for (i = 0; i < stockcount; i++){
				 j = i + 1
					  elemName = "StockWatch" + j
					  
					  elem = document.getElementById(elemName)
					  
						 if (elem.checked==true){
						  
						    if (isNaN(parseInt(elem.value))){
							 }else{
								  if (m==0){
									datastock = elem.value;
								  }else {
									datastock = datastock + ';' + elem.value;
								  }
								  m = 1
							 }
							  
						 }
					
				}
		    }
		 }
		  
		  if (GenContract==1){
		   if (contractcount!=''){
		    m =1;
			   for (i = 0; i < contractcount; i++){
			    j = i + 1
				  elemName = "Contract" + j
				  
				  elem = document.getElementById(elemName)
				  
					 if (elem.checked==true){
					
					    if (isNaN(parseInt(elem.value))){
						 }else{
							  if (m==0){
								datacontract = elem.value;
							  }else {
								datacontract = datacontract + ';' + elem.value;
							  }
							  m=1;
						 }
						  
					 }
			   }
		  }
		  }
		  
		 if (datadebit != '' || datacontract !='' || datastock !=''){
          	window.location.replace("SMS.asp?gen=1&Debits=" + GenDebits + "&stockWatch=" + Genstock + "&Contracts=" + GenContract + "&CanSendSMS=1&owners=1&datacontract="+ datacontract + "&datastock=" + datastock + "&datadebit=" + datadebit);	 
		 }
	} 
</SCRIPT>
</HEAD>

<BODY>
	
<%
gen = trim(Request.QueryString("gen"))

If gen = "" Then gen=0
	
If gen <> 1 Then
	%>
	<form method="post" action="sms.asp" name="FrmSMS">
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
				<td><input type="checkbox" checked name="Contracts" value="1" class="borderless">&nbsp;Contracts</td><td></td>
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
	
	smscount = 0 'count the number of smses sent

	Set Conn = GetActiveConnection("KBroker")	
	 
 
	'Check whether the Contracts has been selected
	If Contracts = 1 Then	
		'get all the list of clients with mobile numbers
		Set rs=server.createobject("Adodb.recordset")
		Set rs1=server.createobject("Adodb.recordset")

		if CanSendSMS = 0 then
           
		   sqlStr= "SELECT * FROM ShowSMSAllContractsCompounded"
		   
			set rs = conn.execute(sqlStr) 
				
			If rs.EOF Or rs.BOF Then
				
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
					ContractNoSms = 0

					While Not rs.EOF 
							'Tabulate the sms's here

							' format the celltell number
								cellNo = replace(rs("ClientCellTel")," ","")
								cellNo = replace(rs("ClientCellTel"),"-","")
								
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
									'If rs("OrderType_DPA_")=1 Then
									'	smsStr = formatnumber(rs("LotQty"),0) & " " & rs("SecurityCode")& " SHARES PURCHASED ON " & formatdate(rs("lotTdate")) & " @ KES " & formatnumber(rs("LotPrice"),2) &" PER SHARE. SETTLEMENT OF KES " & formatnumber(rs("contractAmount"),2) & " DUE " & formatdate(rs("contractSettlementDate"))& ". " & vbCrLf
									'Else
									'	smsStr = formatnumber(rs("LotQty"),0) & " " & rs("SecurityCode")& " SHARES SOLD ON "& formatdate(rs("lotTdate")) &" @ KES " & formatnumber(rs("LotPrice"),2) & " PER SHARE. " & vbCrLf
									'End If
									
									If rs("OrderType_DPA_")=1 Then
										smsStr =smsStr & formatnumber(rs("LotQty"),0) & " " & rs("SecurityCode")& " shares purchased on " & formatdate(rs("lotTdate")) & " @ KES " & formatnumber(rs("LotPrice"),2) &" per share. Settlement of KES " & formatnumber(rs("contractAmount"),2) & " due " & formatdate(rs("contractSettlementDate"))& ". " & vbCrLf
										OwnerSMSstr = OwnerSMSstr & formatnumber(rs("LotQty"),0) & " " & rs("SecurityCode")& " purchased @KES " & formatnumber(rs("LotPrice"),2) &"." & vbCrLf
									Else
										smsStr =smsStr & formatnumber(rs("LotQty"),0) & " " & rs("SecurityCode")& " shares sold on "& formatdate(rs("lotTdate")) &" @ KES " & formatnumber(rs("LotPrice"),2) & " per share. " & vbCrLf
										OwnerSMSstr = OwnerSMSstr & formatnumber(rs("LotQty"),0) & " " & rs("SecurityCode")& " sold @KES " & formatnumber(rs("LotPrice"),2) &"." & vbCrLf
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
									 ContractNoSms = ContractNoSms + 1
									 %> 
									<input type="checkbox" name="Contract<%=ContractNoSms%>" ID="Contract<%=ContractNoSms%>" value="<%=rs.fields("Contract_DPA_")%>" class="borderless" checked>
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
								
								DefaultMsgExists = false
								
							'Determine default message
							if DefaultMsgExists then
							 ownersms = smsStr
							else
                                  sqlstr1 = " select  TOP 1 * from ShowSMSAllContractsCompounded "

								  set rs1= conn.execute(sqlstr1) 
													
									i=0
									secStr =""
									While Not rs1.EOF 
											
										If rs1("OrderType_DPA_")=1 Then
											secStr = formatnumber(rs1("LotQty"),0) & " " & rs1("SecurityCode")& " shares purchased on " & formatdate(rs1("lotTdate")) & " @ KES " & formatnumber(rs1("LotPrice"),2) &" per share. Settlement of KES " & formatnumber(rs1("contractAmount"),2) & " due " & formatdate(rs1("contractSettlementDate"))& ". " & vbCrLf
										Else
											secStr = formatnumber(rs1("LotQty"),0) & " " & rs1("SecurityCode")& " shares sold on "& formatdate(rs1("lotTdate")) &" @ KES " & formatnumber(rs1("LotPrice"),2) & " per share. " & vbCrLf
										End If

										rs1.movenext
									Wend
									  
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
         
		'Send SMS
		data = trim(Request.QueryString("datacontract"))
		
		owners = trim(Request.QueryString("owners")) 'Check if there are any account managers to be smsed
        If owners = "" Then owners =0
        
		DefaultMsgExists = false
		
        ContractID = ""
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
		 
		 set rs=conn.execute(sqlStr)
				
				   If rs.EOF Or rs.BOF Then
						
					Else
					 rs.movefirst
					    
						 While Not rs.EOF 
								'Send the sms's here

								'format the celltell number
								cellNo = replace(rs("ClientCellTel")," ","")
								cellNo = replace(rs("ClientCellTel"),"-","")
								
								'Validate Cell no
								
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
									ContractNoSms = ContractNoSms + 1
								cellNo = Replace(cellNo," ","")
										
								'If it is a purchase contract Then
								If rs("OrderType_DPA_")=1 Then
									smsStr =smsStr & formatnumber(rs("LotQty"),0) & " " & rs("SecurityCode")& " shares purchased on " & formatdate(rs("lotTdate")) & " @ KES " & formatnumber(rs("LotPrice"),2) &" per share. Settlement of KES " & formatnumber(rs("contractAmount"),2) & " due " & formatdate(rs("contractSettlementDate"))& ". " & vbCrLf
								Else
									smsStr =smsStr & formatnumber(rs("LotQty"),0) & " " & rs("SecurityCode")& " shares sold on "& formatdate(rs("lotTdate")) &" @ KES " & formatnumber(rs("LotPrice"),2) & " per share. " & vbCrLf
								End If
                                
                                'DefaultMsgExists = true
                                
									'Store sms message in form element
								   %>
								   <input type="hidden" name="sMessage<%=rs("Client_DPA_")&rs("Contract_DPA_")%>" id="sMessage<%=rs("Client_DPA_")&rs("Contract_DPA_")%>" value="<%=smsStr%>">
								   <%
										 'Send Sms here
										 'cellNo = 254721613909 'Hard Code
											 %>
												 <script language="javascript">
												 var msgStatus, replyText 
												   msgStatus = SendSms(<%=cellNo%>,'sMessage<%=rs("Client_DPA_")&rs("Contract_DPA_")%>')
                                                   
												   //Handle any errors encountered
                                                   replyText = msgStatus.substr(0,2)
			                                       
												   if (replyText.toUpperCase() == 'ID'){
												      //message sent successfully
													  
												   }else{
												    //errors encountered
													alert('Unexpected error occurred.\n ' + msgStatus);
													window.location.href='SMS.asp'
												   }
												  
   
												 </script>
											 <%
											 smsStr = ""
									smscount = smscount + 1
							
								End If
							rs.movenext
							
						 Wend
		   
					End if
                     
		       end if 'ContractID <> ""

		End if ' data <> ""
		
		if owners  =  1 then
		 
                     'SMS Account Managers
					 SQLStr = " Select * from ownerlist where sendSMS = 1 AND MobileNo IS NOT NULL  "
                    
					set rs = conn.execute(SQLStr)

					if not (rs.eof or rs.bof) then
					 rs.movefirst
                    
					 do while not rs.EOF
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
							
							DefaultMsgExists = false
							
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
											
											 Contract_DPA_ = rs1("Contract_DPA_")		
										rs1.movenext
									Wend
									  
									set  rs1 = nothing
									ownersms = secStr
							end if
							'ready to sms
                            %>
								   <input type="hidden" name="sMessage<%=rs("Owner_DPA_")&Contract_DPA_%>" id="sMessage<%=rs("Owner_DPA_")&Contract_DPA_%>" value="<%=ownersms%>">
								   <%

							If Len(cellNo) > 0 AND ValidCellNo = true Then
                              
                                 %>
									 <script language="javascript">
										 var msgStatus, replyText 
										   msgStatus = SendSms(<%=cellNo%>, 'sMessage<%=rs("Owner_DPA_")&Contract_DPA_%>')
										   
										   //Handle any errors encountered
										   replyText = msgStatus.substr(0,2)
										   
										   if (replyText.toUpperCase() == 'ID'){
											  //message sent successfully
											  
										   }else{
											//errors encountered
											alert('Unexpected error occurred.\n ' + msgStatus);
											window.location.href='SMS.asp'
										   }
                                         
									 </script>
								<%
								ownersms = ""
								smscount = smscount + 1
							end if
					 
					 rs.movenext
					 Loop
					end if

		  end if
		  
		End if 'CanSendSMS = 0
            set rs = nothing
			
	End If
	
	'generate all urs stuff here, first check the files one wants to generate
	If stockWatch = 1 Then
        Set rs=server.createobject("Adodb.recordset")
		Set rs1=server.createobject("Adodb.recordset")
		
      if CanSendSMS = 0 then
			 'get all the list of clients with mobile numbers
			 sqlStr ="SELECT DISTINCT TOP 100 PERCENT Client_DPA_,ClientCellTel, ClientName FROM StockWatchList order by client_DPA_"
			
			 set rs = conn.execute(sqlStr)
							
			 If rs.EOF Or rs.BOF Then
					
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
						//alert('The last importation of the Price List was on <%=lastModifiedDate%>');
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
					StockNoSms = 0
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
							  If len(rs("ClientCellTel")) > 0 then
								cellNo = replace(rs("ClientCellTel")," ","")
								cellNo = replace(rs("ClientCellTel"),"-","")
							else
								cellNo = ""
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
										   StockNoSms = StockNoSms + 1
										 %> 
										   <input type="checkbox" name="StockWatch<%=StockNoSms%>" ID="StockWatch<%=StockNoSms%>" value="<%=rs.fields("Client_DPA_")%>" class="borderless" checked>
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
								cellNo = replace(rs("MobileNo")," ","")
								cellNo = replace(cellNo,"-","")

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
				End If 'rs.EOF

		Else
        
		'Send SMS
		data = trim(Request.QueryString("datastock"))
		
		owners = trim(Request.QueryString("owners")) 'Check if there are any account managers to be smsed
        If owners ="" Then owners =0
		DefaultMsgExists = false
		
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
       ClientSQL = " Client_DPA_ IN  (" & ClientID & ")  "
       
		 'get all the list of clients with mobile numbers
			 sqlStr ="SELECT DISTINCT TOP 100 PERCENT Client_DPA_,ClientCellTel, ClientName FROM StockWatchList Where " & ClientSQL & " order by client_DPA_"
			'' Response.Write sqlStr
			'' Response.end
			 set rs = conn.execute(sqlStr) 
							
				If rs.EOF Or rs.BOF Then
					
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
								cellNo = replace(rs("ClientCellTel")," ","")
								cellNo = replace(rs("ClientCellTel"),"-","")

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
													window.location.href='SMS.asp'
												   }

												 </script>
												 
											 <%
											 smsStr = ""
									smscount = smscount + 1
								 end if
					
						rs.movenext
					Wend

					Set rs=Nothing
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
								cellNo = replace(rs("MobileNo")," ","")
								cellNo = replace(cellNo,"-","")

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
											
											StockWatch_DPA_ = rs1("StockWatch_DPA_")
										rs1.movenext
									Wend
									  
									rs1.close
									set  rs1 = nothing
									ownersms = secStr
							end if
							'ready to sms
                            %>
								   <input type="hidden" name="sMessage<%=rs("Owner_DPA_")&StockWatch_DPA_%>" id="sMessage<%=rs("Owner_DPA_")&StockWatch_DPA_%>" value="<%=ownersms%>">
								   <%

							If Len(cellNo) > 0 AND ValidCellNo = true Then
                              
                                 %>
									 <script language="javascript">
										 var msgStatus, replyText 
										   msgStatus = SendSms(<%=cellNo%>, 'sMessage<%=rs("Owner_DPA_")&StockWatch_DPA_%>')
										   
										   //Handle any errors encountered
										   replyText = msgStatus.substr(0,2)
										   
										   if (replyText.toUpperCase() == 'ID'){
											  //message sent successfully
											  
										   }else{
											//errors encountered
											alert('Unexpected error occurred.\n ' + msgStatus);
											window.location.href='SMS.asp'
										   }
										   

									 </script>
								<%
								ownersms=""
								smscount = smscount + 1
							end if
					 
					 rs.movenext
					 Loop
					end if
		  end if

        End if 'Can Send SMS = 0

		 set rs = nothing
	End If 'StockList = 1
	
	'Check for debits
	If Debits =1 Then
	     
		'get all the list of clients with mobile numbers
		Set rs=server.createobject("Adodb.recordset")
		Set rs1=server.createobject("Adodb.recordset")

		Set Conn = GetActiveConnection("KBroker")
		
		if CanSendSMS = 0 then

			sqlStr= "SELECT * FROM smsDebtors WHERE (ClientCellTel IS Not NULL) AND (ClientCellTel <> N'') AND (updateOnDebt=1)"
			
			set rs= conn.execute(sqlStr)
			
			If rs.EOF Or rs.BOF Then
				
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
				DebitNoSms = 0

				While Not rs.EOF 
						'create a string for the sms here
						' format the celltell number
						cellNo = replace(rs("ClientCellTel")," ","")
						cellNo = replace(rs("ClientCellTel"),"-","")

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
											   DebitNoSms = DebitNoSms + 1
											 %> 
											   <input type="checkbox" name="Debit<%=DebitNoSms%>" ID="Debit<%=DebitNoSms%>" value="<%=rs.fields("Client_DPA_")%>" class="borderless" checked>
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
				%>
					</table>
					<%		 
						 
			End If 'rs.EOF
        Else
		' Send SMS

		 data = trim(Request.QueryString("datadebit"))
         
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
			
			set rs = conn.execute(sqlStr)
			
			If rs.EOF Or rs.BOF Then
				
			Else
				smsStr = ""
				rs.movefirst
               
				While Not rs.EOF 
						'create a string for the sms here
						' format the celltell number
						cellNo = replace(rs("ClientCellTel")," ","")
						cellNo = replace(rs("ClientCellTel"),"-","")

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
								   <input type="hidden" name="sMessage<%=rs("Client_DPA_")&rs("ClientCDSNo")%>" id="sMessage<%=rs("Client_DPA_")&rs("ClientCDSNo")%>" value="<%=smsStr%>">
								   <%
										 'Send Sms here
										 'cellNo = 254721613909 'Hard Code
											 %>
												 <script language="javascript">
												 var msgStatus, replyText 
												   msgStatus = SendSms(<%=cellNo%>,'sMessage<%=rs("Client_DPA_")&rs("ClientCDSNo")%>')
                                                   
												   //Handle any errors encountered
                                                   replyText = msgStatus.substr(0,2)
			                                       
												   if (replyText.toUpperCase() == 'ID'){
												      //message sent successfully
													  
												   }else{
												    //errors encountered
													alert('Unexpected error occurred.\n ' + msgStatus);
													window.location.href='SMS.asp'
												   }
												  
												 </script>
											 <%
											 smsStr= ""
									smscount = smscount + 1
						  end if

						
					rs.movenext
				Wend
						 
			End If 'rs.EOF
			
			End if 'data <> ""
		End If 'if CanSendSMS = 0 then
	End If 'Debits = 1
	
	if CanSendSMS = 1 then
		%>
		<script language="javascript">
			alert('<%=smscount%> SMS Text Message(s) sent successfully.')
			window.location.href='SMS.asp'
		</script>
		<%
	end if
	%>
	<BR>
	<input type ="button" name ="back" value="<< Back" onclick ="javascript: window.location.href='SMS.asp'">
	<input type="hidden" name="StockNoSms" id="StockNoSms" value="<%=StockNoSms%>">
	<input type="hidden" name="DebitNoSms" id="DebitNoSms" value="<%=DebitNoSms%>">
	<input type="hidden" name="ContractNoSms" id="ContractNoSms" value="<%=ContractNoSms%>">
	<input type="hidden" name="GenContract" id="GenContract" value="<%=Contracts%>">
	<input type="hidden" name="Genstock" id="Genstock" value="<%=stockWatch%>">
	<input type="hidden" name="GenDebits" id="GenDebits" value="<%=Debits%>">
	<input type ="button" name ="SMSData" id ="SMSData" value="Send SMS" onclick="javascript: ValidateSMS()">
	</form> 
	<%
	set conn = nothing
End If
%>

</BODY>
</HTML>
