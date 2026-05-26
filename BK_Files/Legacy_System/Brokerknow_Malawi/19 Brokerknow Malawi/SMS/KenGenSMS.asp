<!--#include file="../libroutines.asp"-->
<HTML>
<HEAD>
<TITLE> KenGen SMS </TITLE>
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css">
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT>

<style type="text/css">
	td {border: 1 solid gray;}
</style>

<SCRIPT Language="JavaScript">
function SendSms(sMobileNo)
	{
		var sUrl;
		var sAPI_ID, sPassword, sUsername;
		var xmlhttp, sPostData, sResult;
		var max_credits, escalate, from;
		var sText = 'KenGen Shares available.20th March to 11th April.@11.90.Minimum 500 shares.AAKS email:gathigic@africanalliance.co.ke.Tel:2735138.';
				
		sUrl = "https://api.clickatell.com/http/sendmsg" 
		sAPI_ID = 1528651
		sPassword = "alliance1"
		sUsername = "africanalliance"
		max_credits = 3
		escalate =3
		from = "AfricanAlliance"
		        
		sPostData = "api_id=" +  sAPI_ID;
		sPostData = sPostData + "&user=" + sUsername;
		sPostData = sPostData + "&password=" + sPassword;
		sPostData = sPostData + "&to=" + sMobileNo;
		sPostData = sPostData + "&text=" + sText;
		sPostData = sPostData + "&max_credits=" + max_credits;
		sPostData = sPostData + "&escalate=" + escalate;
		sPostData = sPostData + "&from=" + from;
		        
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
	var i;
	var data = '';
		
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
	
	document.frmMain.hidData.value = data;
	
	frmMain.method = 'POST';
	frmMain.action = 'KenGenSMS.asp?SendSMS=1';
	frmMain.submit();
	} 
</SCRIPT>

</HEAD>

<BODY>
<%
SendSMS = trim(Request.QueryString("SendSMS"))

If SendSMS = "" Then SendSMS = 0
	
If SendSMS <> 1 Then
	''DISPLAY CLIENTS TO SEND
	sqlStr ="SELECT Client_DPA_, ClientName, ClientCellTel FROM Client WHERE (ClientCellTel <> '254') AND (ClientCellTel <> '') AND (ClientCellTel IS NOT NULL) AND (DELETED = 0) ORDER by ClientName"
	
	Set RS = Server.CreateObject("ADODB.RecordSet")
	Set Conn = GetActiveConnection("KBroker")	
	
	RS.Open sqlStr, Conn, 0, 1
								
	If RS.EOF Or RS.BOF Then
		%>
		<SCRIPT LANGUAGE="JavaScript">
			alert("No SMS to send.");
			window.location.href='KenGenSMS.asp'
		</SCRIPT>
		<%
	Else
		%>
		<br><br>
		
		<form method="post" action="" name="frmMain" id="frmMain">
		
		<table align="center" border=0 cellspacing="0" cellpadding="1" style="border: 0 solid gray;" bordercolor="#111111" width="80%">	
			<tr>
				<td colspan="5" width="100%" style="border: 0;">&nbsp;<b>KenGen IPO SMS</b></td>
			</tr>
			
			<tr>
				<td height="20" colspan="5" width="100%" style="border: 0;">&nbsp;</td>
			</tr>
	
			<tr>
				<td width="5%">&nbsp;<b>Code</b></td>
				<td width="35%">&nbsp;<b>Name</b></td>
				<td width="20%">&nbsp;<b>Cell No</b></td>
				<td width="20%">&nbsp;<b>To SMS</b></td>
			</tr>
			<%
			While Not RS.EOF 
				if Len(RS("ClientCellTel")) = 12 And Instr(1,RS("ClientCellTel"),"+") = 0 then
					CellNo = replace(RS("ClientCellTel")," ","")
				elseif Len(RS("ClientCellTel")) >= 10 And Instr(1,RS("ClientCellTel"),"+") = 0 then
					If Instr(1,RS("ClientCellTel"),";") > 0 Then
						CellNo = mid(replace(RS("ClientCellTel")," ",""),1,12)
					End If
					
					If Instr(1,RS("ClientCellTel"),"254") = 0 Then
						CellNo = replace(RS("ClientCellTel")," ","")
					End If
				else
					CellNo = ""
				end if
				%>
				<tr>
					<td width="5%">&nbsp;<%=RS("Client_DPA_")%></td>
					<td width="35%">&nbsp;<%=RS("ClientName")%></td>
					
					<%
					If Len(CellNo) > 0 Then
						%> 
						<td width="20%">&nbsp;<input type="text" name="txtC<%=RS("Client_DPA_")%>" ID="txtC<%=RS("Client_DPA_")%>" value="<%=cellNo%>" style="border:0;background-color:white;color:black;font-size:8pt;font-family:tahoma;">
						</td>
						<td width="20%">&nbsp;
							<input type="checkbox" name="C<%=RS("Client_DPA_")%>" ID="C<%=RS("Client_DPA_")%>" value="C<%=RS("Client_DPA_")%>" class="borderless" checked>
						<%
					Else
						%> 
						<td width="20%">&nbsp;<%=RS("ClientCellTel")%></td>
						<td width="20%">&nbsp;
						<%	
						response.write "-INVALID-"
					End If
					%>
					</td>
				</tr>
				<%
				RS.movenext
			Wend
			%>
		<!--</table>-->
		<%
	End If    

	''SMS ACCOUNT MANAGERS
	sqlStr = "SELECT * FROM ownerlist WHERE sendSMS = 1 AND MobileNo IS NOT NULL"
	set RS = Conn.Execute(sqlStr)

	if not (RS.eof or RS.bof) then
		%>
		<tr>
			<td height="20" colspan="5" width="100%" style="border: 0;">&nbsp;</td>
		</tr>
		<tr>
			<td colspan="5" width="100%" style="border: 0;">&nbsp;<b>Account Managers</b></td>
		</tr>
		
		<%
		do while not RS.EOF
				if Len(RS("MobileNo")) = 12 And Instr(1,RS("MobileNo"),"+") = 0 then
					CellNo = replace(RS("MobileNo")," ","")
				elseif Len(RS("MobileNo")) >= 10 And Instr(1,RS("MobileNo"),"+") = 0 then
					If Instr(1,RS("MobileNo"),";") > 0 Then
						CellNo = mid(replace(RS("MobileNo")," ",""),1,12)
					End If
					
					If Instr(1,RS("MobileNo"),"254") = 0 Then
						CellNo = replace(RS("MobileNo")," ","")
					End If
				else
					CellNo = ""
				end if
				%>
				<tr>
					<td width="5%">&nbsp;</td>
					<td width="35%">&nbsp;<%=RS("OwnerName")%></td>
					<%
					If Len(CellNo) > 0 Then
						%> 
						<td width="20%">&nbsp;<input type="text" name="txtO<%=RS("Owner_DPA_")%>" ID="txtO<%=RS("Owner_DPA_")%>" value="<%=RS("MobileNo")%>" style="border:0;background-color:white;color:black;font-size:8pt;font-family:tahoma;"></td>
						<td width="20%">&nbsp;
							<input type="checkbox" disabled name="O<%=RS("Owner_DPA_")%>" ID="O<%=RS("Owner_DPA_")%>" value="O<%=RS("Owner_DPA_")%>" class="borderless" checked>
						<%
					Else
						%> 
						<td width="20%">&nbsp;<%=RS("Owner_DPA_")%></td>
						<td width="20%">&nbsp;
						<%	
						response.write "-INVALID-"
					End If
					%>
					</td>
				</tr>
				<%
			RS.movenext
		Loop
	end if
	%>
	
	<tr>
		<td height="20" colspan="5" width="100%" style="border: 0;">&nbsp;</td>
	</tr>
	
	<tr>
		<td align="center" height="20" colspan="5" width="100%" style="border: 0;">
			<input type ="hidden" name ="hidData" id ="hidData" value="">
			<input type ="button" name ="SMSData" id ="SMSData" value="Send SMS" onclick="javascript: ValidateSMS()">
		</td>
	</tr>
	
	</table>
	
	</form>
	<%
Else
	''SEND SMS
	data = Request("hidData")
	
	ExtraCellNos = ";254721689695"
	
	data = Mid(data,2) & ExtraCellNos
	
	EachItem = Split(data,";")
	
	Dim i
	
	For i = 0 To UBound(EachItem)
		If Instr(1,EachItem(i),"C") > 0 Then
			''CLIENT
			CellNo = Request.Form("txt"&EachItem(i))
			
			'Response.Write EachItem(i) & "=" & CellNo & "<br>"
		End If
		
		If Instr(1,EachItem(i),"O") > 0 Then
			''ACCOUNT MANAGER
			CellNo = Request.Form("txt"&EachItem(i))
			
			'Response.Write EachItem(i) & "=" & CellNo & "<br>"
		End If
		
		'Response.Write CellNo
		'Response.End 
		
		If Len(CellNo) > 0 Then
			%>
			<script language="javascript">
				var msgStatus, replyText 
				msgStatus = SendSms(<%=CellNo%>)
						                                                   
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
					window.location.href='KenGenSMS.asp'
					}
			</script>
			<%
		End If
	Next
	
	%>
	<script language="javascript">
		alert('<%=i%> SMS Text Message(s) sent successfully.')
		window.location.href='KenGenSMS.asp'
	</script>
	<%
	Response.End
End if 
%>

</BODY>
</HTML>