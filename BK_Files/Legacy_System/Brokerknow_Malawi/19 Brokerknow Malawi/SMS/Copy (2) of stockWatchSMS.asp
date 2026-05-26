<!--#include file="../libroutines.asp"-->
<HTML>
<HEAD>
<TITLE> SMS </TITLE>
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css">
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">

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
		window.location.replace("stockWatchSMS.asp?gen=1&stockWatch="+stockWatch+"&Contracts="+Contracts+"&Debits="+Debits);
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
	<form method="post" action="stockWatchSMS.asp" name="FrmSMS">
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
	stockWatch = trim(Request.QueryString("stockWatch"))
	If stockWatch="" Then stockWatch=0
		
	Contracts = trim(Request.QueryString("Contracts"))
	If Contracts ="" Then Contracts=0
	
	Debits = trim(Request.QueryString("Debits"))
	If Debits ="" Then Debits =0
	
	''create the server objects here 
	Set rs=server.createobject("Adodb.recordset")
	Set rs1=server.createobject("Adodb.recordset")

	Set Conn = GetActiveConnection("KBroker")	

	' first get the paths where the files are to be stored
	'sqlStr = "select * from PathConfigurations order by DocumentID"
	'Rs.Open sqlStr, Conn, 0, 1
	
	'While Not Rs.EOF
	'		If Rs("DocumentID") = 1 Then strPath = Rs("Path")
	'		If Rs("DocumentID") = 2 Then strPathStockwatch = Rs("Path")
	'		If Rs("DocumentID") = 3 Then strPathDebtors = Rs("Path")
	'		If Rs("DocumentID") = 4 Then strPathContracts = Rs("Path")
	'	Rs.MoveNext
	'Wend
	'Rs.Close
	
	'Set Document Paths
	strPath = server.MapPath(".") & "\bin"
	strPathStockwatch = server.MapPath(".") & "\bin"
	strPathDebtors = server.MapPath(".") & "\bin"
	strPathContracts = server.MapPath(".") & "\bin"
	
	'generate all urs stuff here, first check the files one wants to generate
	If stockWatch = 1 Then
		'get all the list of clients with mobile numbers
		sqlStr= "SELECT * FROM StockWatchList ORDER BY Client_DPA_, StockWatch_DPA_"

		sqlStr ="SELECT TOP 100 PERCENT dbo.StockWatch.Client_DPA_, COUNT(dbo.StockWatch.Client_DPA_) AS StockCount, 	dbo.StockWatchList.Client_DPA_ AS ClientCode,  " & _
			" dbo.StockWatchList.SecurityCode, dbo.StockWatchList.Price, dbo.StockWatchList.ClientCellTel " & _
			" FROM dbo.StockWatch INNER JOIN " & _
			" dbo.StockWatchList ON dbo.StockWatch.Client_DPA_ = dbo.StockWatchList.Client_DPA_ " & _
			" WHERE (dbo.StockWatch.Deleted <> 1) " & _
			" GROUP BY dbo.StockWatch.Client_DPA_, dbo.StockWatchList.Client_DPA_, dbo.StockWatchList.SecurityCode, dbo.StockWatchList.Price,  " & _
			" dbo.StockWatchList.ClientCellTel " & _
			" ORDER BY dbo.StockWatch.Client_DPA_"

		sqlStr ="SELECT DISTINCT TOP 100 PERCENT Client_DPA_,ClientCellTel FROM StockWatchList order by client_DPA_"
		
		rs.open sqlStr, conn,0,1
					
		Dim client
		Dim cellNo
		Dim SmsStr
		Dim PricenSec
		Dim first
		Dim smsStr3
		show =1
		
		If rs.EOF Or rs.BOF Then
			%>
			<SCRIPT LANGUAGE="JavaScript">
				alert("No Stock Watches To SMS");
				window.location.href='StockWatchSMS.asp'
			</SCRIPT>
			<%
			'Response.End
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
					
			While Not rs.EOF 
					'create a string for the sms here
					clientNo = rs("Client_DPA_")
													
					sqlstr1 ="select * from StockWatchList where client_DPA_ =" & rs("Client_DPA_")

					rs1.open sqlstr1, conn, 0, 1
									
					i=0
					While Not rs1.EOF 
							'get the securities here
							If  trim(secStr) = "" Then
								secStr = rs1("SecurityCode") & " " & rs1("Price")	
							Else
								'check If the characters are greater than 150
								If len(secStr &" "& rs1("SecurityCode")	& " " & rs1("Price"))>150 Then
									secStr1 = secStr1 &" "& rs1("SecurityCode")	& " " & rs1("Price")
								Else
									secStr = secStr &" "& rs1("SecurityCode")	& " " & rs1("Price")	
								End If
							End If
						rs1.movenext
					Wend
					rs1.close

					cellNo = replace(rs("ClientCellTel"), " ","")
					cellNo = replace(rs("ClientCellTel"), "-","")
					
					If Len(cellNo) > 0 Then
						'Remove 0 at position 1
						'If mid(cellNo,1,1)=0 Then cellNo = mid(cellNo,2,10) Else cellno = mid(cellNo,1,9)

						cellNo = "CSV:" & Replace(cellNo," ","")
						smsStr =  cellNo &"|"& secStr & vbCrLf
					End If
								
					If secStr1<>"" Then smsStr_1 =  cellNo &"|"& secStr1 & vbCrLf
				
					smsStr1 = smsStr1 & smsStr & smsStr_1
					secStr=""
					smsStr_1=""
					secStr1=""
				rs.movenext
			Wend

			' add this text to the csv file
			HeaderText ="api-id:1362866" & vbCrLf  
			HeaderText =HeaderText  & "user:jthiore" & vbCrLf 
			HeaderText =HeaderText  & "password:jthiore1" & vbCrLf 
			HeaderText =HeaderText  & "text: #field1#" & vbCrLf
			HeaderText =HeaderText  & "Delimiter:|" & vbCrLf
			HeaderText =HeaderText  & "max_credits:3" & vbCrLf
			HeaderText =HeaderText  & "escalate:3" & vbCrLf
			HeaderText =HeaderText  & "from:AAKS" & vbCrLf

			thePath= strPathStockwatch & "\sw_"& replace(date(),"/","") & ".txt"
				
			Set fso = server.createobject("Scripting.FileSystemObject")
			Set txtFile = fso.createTextFile(thePath)
				 
			txtFile.WriteLine(HeaderText & "" & smsStr1)
			txtFile.close
				 
			Set txtFile=Nothing
			Set fso= Nothing
			rs.close

			Set rs=Nothing
			conn.close

			Response.Write "<b>Stock Watch List</b> <br>"		
			Response.Write replace(HeaderText & "" & smsStr1,vbCrLf,"<br>")
		End If
	End If

	'Check whether the debit has been selected
	If Contracts = 1 Then	
		'get all the list of clients with mobile numbers
		Set rs=server.createobject("Adodb.recordset")
		Set rs1=server.createobject("Adodb.recordset")

		Set Conn = GetActiveConnection("KBroker")
		sqlStr= "SELECT * FROM smsContract WHERE (ClientCellTel IS Not NULL) AND (ClientCellTel <> N'') and (updateOnContract=1)  ORDER BY Client_DPA_"
		sqlStr= "SELECT * FROM smsContract WHERE (ClientCellTel IS Not NULL) AND (ClientCellTel <> N'')  ORDER BY Client_DPA_"

		rs.open sqlStr, conn,0,1
			
		show =1
		If rs.EOF Or rs.BOF Then
			%>
			<SCRIPT LANGUAGE="JavaScript">
				alert("No Contracts To SMS");
				window.history.back(0);
			</SCRIPT>
			<%
			'Response.End
		Else
			rs.movefirst
			cNo=0
			
			While Not rs.EOF 
					'create a string for the sms here
					' format the celltell number
					cellNo = replace(rs("ClientCellTel")," ","")
					If mid(cellNo,1,1)=0 Then cellNo = mid(cellNo,2,10) Else cellno = mid(cellNo,1,9)
					
					If Not isnumeric(cellNo)  Or trim(cellNo)="" Then
						cellNo = cellNo & " *"
						cNo=1
					Else
						cNo=0
						' add the prefix to the cell no
						cellNo = "CSV:254" & cellNo		
					End If
					
					'If it is a purchase contract Then
					If cint(cnNo)= cint(0) Then
						If rs("OrderType_DPA_")=1 Then
							smsStr =smsStr & cellNo &"|"& rs("LotQty") & " " & rs("SecurityCode")& " SHARES PURCHASED ON " & formatdate(rs("lotTdate")) & " @ KES " & rs("LotPrice")&" PER SHARE. SETTLEMENT OF KES " & rs("contractAmount") & " DUE " & formatdate(rs("contractSettlementDate"))& vbCrLf
						Else
							smsStr =smsStr & cellNo  &"|"& rs("LotQty") & " " & rs("SecurityCode")& " SHARES SOLD ON "& formatdate(rs("lotTdate")) &" @ KES " & rs("LotPrice")& " PER SHARE " & vbCrLf
						End If
					End If
				rs.movenext
			Wend

			'Create a text file here where u will save ur data
			HeaderText ="api-id:1362866" & vbCrLf  
			HeaderText =HeaderText  & "user:jthiore" & vbCrLf 
			HeaderText =HeaderText  & "password:jthiore1" & vbCrLf 
			HeaderText =HeaderText  & "text: #field1#" & vbCrLf
			HeaderText =HeaderText  & "Delimiter:|" & vbCrLf
			HeaderText =HeaderText  & "max_credits:3" & vbCrLf
			HeaderText =HeaderText  & "escalate:3" & vbCrLf
			HeaderText =HeaderText  & "from:AAKS" & vbCrLf
			
			thePath= strPathDebtors & "\ca_"& replace(date(),"/","") & ".txt"
			
			Set fso = server.createobject("Scripting.FileSystemObject")
			Set txtFile = fso.createTextFile(thePath)
			
			txtFile.WriteLine(HeaderText & "" & smsStr)
			txtFile.close
			
			Set txtFile=Nothing
			Set fso= Nothing
			
			rs.close
			conn.close
			
			Response.Write "<br>&nbsp;<br><b>Contracts List</b><br>"
			Response.Write replace(HeaderText & "" & smsStr,vbCrLf,"<br>")
		End If
	End If

	'Check for debits
	If Debits =1 Then
		'get all the list of clients with mobile numbers
		Set rs=server.createobject("Adodb.recordset")
		Set rs1=server.createobject("Adodb.recordset")

		Set Conn = GetActiveConnection("KBroker")
		sqlStr= "SELECT * FROM smsDebtors WHERE (ClientCellTel IS Not NULL) AND (ClientCellTel <> N'') ORDER BY Client_DPA_"

		'Set the debt date to be 5 days b4 leo
		selectedDate =formatdate(dateadd("d",-5,now()))
		sqlStr = " SELECT DISTINCT TOP 100 PERCENT a.Client_DPA_, MAX(a.TransDate) AS LastDate, b.CurrentBal AS Balance, Client.CreditLimit, Client.ClientCellTel " & _
			" FROM (SELECT     * " & _
			" FROM dbo.ClientTransactionList " & _
			" WHERE Transdate <= '"& selectedDate &"') a INNER JOIN " & _
			" Client ON a.Client_DPA_ = Client.Client_DPA_ INNER JOIN " & _
			" (SELECT SUM(ISNULL(dbo.StatementList.Credit - dbo.StatementList.Debit, 0)) + dbo.Client.ClientOpeningBal AS CurrentBal,  " & _
			" dbo.StatementList.Client_DPA_ " & _
			" FROM dbo.StatementList INNER JOIN " & _
			" dbo.Client ON dbo.StatementList.Client_DPA_ = dbo.Client.Client_DPA_ " & _
			" WHERE (dbo.Client.Deleted = 0 AND Transdate <= '"& selectedDate &"') " & _
			" GROUP BY dbo.StatementList.Client_DPA_, dbo.Client.ClientOpeningBal) b ON Client.Client_DPA_ = b.Client_DPA_ " & _
			" WHERE (Client.Deleted = 0) " & _
			" GROUP BY a.Client_DPA_, b.CurrentBal, Client.CreditLimit, Client.ClientCellTel, Client.updateOnDebt " & _
			" HAVING (b.CurrentBal < 0) AND (Client.updateOnDebt = 1) AND (Client.ClientCellTel IS Not NULL) AND (RTRIM(LTRIM(Client.ClientCellTel)) <> '')"

		rs.open sqlStr, conn,0,1
		
		show =1
		If rs.EOF Or rs.BOF Then
			%>
			<SCRIPT LANGUAGE="JavaScript">
				alert("No Debtor Information To SMS");
				window.history.back(0);
			</SCRIPT>
			<%
			'Response.End
		Else
			smsStr=""
			rs.movefirst
			
			While Not rs.EOF 
					'create a string for the sms here
					' format the celltell number
					cellNo = replace(rs("ClientCellTel")," ","")
					If mid(cellNo,1,1)=0 Then cellNo = mid(cellNo,2,10) Else cellno = mid(cellNo,1,9)
					
					If Not isnumeric(cellNo)  Or trim(cellNo)="" Then
						cellNo = cellNo & " *"
					Else
						' add the prefix to the cell no
						cellNo = "CSV:254" & cellNo		
					End If
					smsStr =smsStr & cellNo &"|" & "Please note that you have an outstanding balance of KES " & formatnumber(0-rs("Balance"),2) & " overdue. Please settle."& vbCrLf
				rs.movenext
			Wend

			'path = request.form("path")& "ClientStockWatch "& replace(formatdate(now()),"-","") &" "& replace(formatdatetime(now(),vbshorttime),":","") &".txt"
			'Response.Write path
			'Response.End

			HeaderText ="api-id:1362866" & vbCrLf  
			HeaderText =HeaderText  & "user:jthiore" & vbCrLf 
			HeaderText =HeaderText  & "password:jthiore1" & vbCrLf 
			HeaderText =HeaderText  & "text: #field1#" & vbCrLf
			HeaderText =HeaderText  & "Delimiter:|" & vbCrLf
			HeaderText =HeaderText  & "max_credits:3" & vbCrLf
			HeaderText =HeaderText  & "escalate:3" & vbCrLf
			HeaderText =HeaderText  & "from:AAKS" & vbCrLf

			thePath= strPathDebtors & "\db_"& replace(date(),"/","") & ".txt"
					
			Set fso = server.createobject("Scripting.FileSystemObject")
			Set txtFile = fso.createTextFile(thePath)
					 
			txtFile.WriteLine(HeaderText &""& smsStr)
			txtFile.close
					 
			Set txtFile=Nothing
			Set fso= Nothing
					 
			rs.close
			conn.close
					 
			Response.Write "<br><br><b>Debits List</b><br>"
			Response.Write replace(HeaderText &""& smsStr,vbCrLf,"<br>")
		End If
	End If

	%>
	<BR>
	<input type ="button" name ="back" value="<< Back" onclick ="javascript: window.location.href='stockWatchSMS.asp'">
	<%
End If
%>

</BODY>
</HTML>
