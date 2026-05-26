<!--#include file="../libroutines.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE> SMS </TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
</HEAD>

<BODY>
	
<%
	gen = trim(request.form("gen"))
	if gen ="" then gen=0
	if gen <>1 then
%>	
	<form method=post action="ContractSms.asp">
	<table width="50%">
	<tr>
		<td>&nbsp;<BR></td> <td></td>
	</tr>
	<tr>
		<td colspan="2"> <U><B>Select the SMS files you want to Generate</B></U></td>
	</tr>
	<tr>
		<td><input type="checkbox"  name="stockWatch" value="1" class="borderless">&nbsp;Stock Watch</td><td></td>
	</tr>
	<tr>
		<td><input type="checkbox" checked name="Contracts" value="1" class="borderless">&nbsp;Contracts</td><td></td>
	</tr>
	<tr>
		<td><input type="checkbox"  name="Debits" value="1" class="borderless">&nbsp;Debits</td><td></td>
	</tr>
	<tr>
		<td colspan="2"><input type ="Submit" value="Generate"></td>
	</tr>
	<input type ="hidden" value ="1" name="gen">
	</table>
	</form>

	<%
	else
		stockWatch = trim(request.form("stockWatch"))
		if stockWatch="" then stockWatch=0
		Contracts = trim(request.form("Contracts"))
		if Contracts ="" then Contracts=0
		Debits = trim(request.form("Debits"))
		if Debits ="" then Debits =0
		
			
			''create the server objects here 
			set rs=server.createobject("Adodb.recordset")
			set rs1=server.createobject("Adodb.recordset")

			Set Conn = GetActiveConnection("KBroker")	

			 ' first get the paths where the files are to be stored
			sqlStr = "select * from PathConfigurations order by DocumentID"
				Rs.Open sqlStr, Conn, 0, 1
				While Not Rs.EOF
				   If Rs("DocumentID") = 1 Then strPath = Rs("Path")
				   If Rs("DocumentID") = 2 Then strPathStockwatch = Rs("Path")
				   If Rs("DocumentID") = 3 Then strPathDebtors = Rs("Path")
				   If Rs("DocumentID") = 4 Then strPathContracts = Rs("Path")
					Rs.MoveNext
				Wend
				Rs.Close

			'generate all urs stuff here, first check the files one wants to generate
			if stockWatch =1 then
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

					'response.write  sqlStr


				sqlStr ="SELECT DISTINCT TOP 100 PERCENT Client_DPA_,ClientCellTel FROM StockWatchList order by client_DPA_"
				rs.open sqlStr, conn,0,1
				dim client
				dim cellNo
				dim SmsStr
				dim PricenSec
				dim first
				dim smsStr3
				show =1
				if rs.eof or rs.bof then
					%>
					
					<SCRIPT LANGUAGE="JavaScript">
					<!--
						alert("There are not stock watches to generate");
						window.history.back(0);

					//-->
					</SCRIPT>

				<%
					response.end
				else
				
				'check the last time the stocks were updated 
					
				set equityRs = server.createObject("Adodb.recordset")

				equityRs.open "SELECT MAX(MktDate) AS LastImportDate FROM datastream_Market ", conn, 0,1

				if not rs.eof or not rs.bof then
					lastModifiedDate = formatdate(equityRs("LastImportDate"))
				end if

				equityRs.close
				set equityRs= nothing

				' give an alert on the date the last import happened

				%>
					<SCRIPT LANGUAGE="JavaScript">
					<!--
							alert('The last importation of the price list happend on <%=lastModifiedDate%>');
					//-->
					</SCRIPT>
				
				<%
				rs.movefirst
				
					while not rs.eof 
					'create a string for the sms here
					
					clientNo = rs("Client_DPA_")
					cellNo = replace(rs("ClientCellTel"), " ","")
									
					sqlstr1 ="select * from StockWatchList where client_DPA_ =" & rs("Client_DPA_")

					rs1.open sqlstr1, conn, 0, 1
					
						i=0
						while not rs1.eof 
						'get the securities here
						
						if  trim(secStr) = "" then
							secStr = rs1("SecurityCode") & " " & rs1("Price")	
						else
							'check if the characters are greater than 150
							if len(secStr &" "& rs1("SecurityCode")	& " " & rs1("Price"))>150 then
								
								secStr1 = secStr1 &" "& rs1("SecurityCode")	& " " & rs1("Price")
							else
							secStr = secStr &" "& rs1("SecurityCode")	& " " & rs1("Price")	
							end if
						end if
						
						rs1.movenext
					wend
					 rs1.close

					'make sure the cell fone number are ok

					if trim(cellNo)="" or cellNo=null then
						cellNo="*"
						'smsStr =  cellNo &","& secStr & "<br>"
						'smsStr =  "<font color='red'>" & cellNo &","& secStr & "</font><br>"
						smsStr =""
					else
					if mid(cellNo,1,1)=0 then cellNo = mid(cellNo,2,10) else cellno = mid(cellNo,1,9)
					if not isnumeric(cellNo)  or trim(cellNo)="" then
						cellNo = cellNo & " *"
						'smsStr =  "<font color='red'>" & cellNo &","& secStr & "</font><br>"
						smsStr =""
					else
						' add the prefix to the cell no
						cellNo = "CSV: 254" & cellNo
						smsStr =  cellNo &"|"& secStr & chr(13)
						
					end if
					
					end if
					if secStr1<>"" then smsStr_1 =  cellNo &"|"& secStr1 & chr(13)
						smsStr1 = smsStr1 & smsStr & smsStr_1
						secStr=""
						smsStr_1=""
						secStr1=""
						rs.movenext
					wend


	'Create a text file here where u will save ur data
			'path = request.form("path")& "ClientStockWatch "& replace(formatdate(now()),"-","") &" "& replace(formatdatetime(now(),vbshorttime),":","") &".txt"
			'response.write path
			'response.end


	' add this text to the csv file

			HeaderText ="api-id: 1127740" & chr(13)  
			HeaderText =HeaderText  & "user: african alliance" & chr(13) 
			HeaderText =HeaderText  & "password: alliance1" & chr(13) 
			HeaderText =HeaderText  & "text: #field1#" & chr(13)
			HeaderText =HeaderText  & "Delimiter: |" & chr(13)

			path1= strPathStockwatch & "\sw_"& replace(date(),"/","") & ".txt"
			
			'Response.Write path1
			'Response.end
			 set fso = server.createobject("Scripting.FileSystemObject")
			 set txtFile = fso.createTextFile(path1)
			 txtFile.WriteLine(HeaderText & "" & smsStr1)
			 txtFile.close
			 set txtFile=nothing
			 set fso= nothing
			 rs.close

			 set rs=nothing
			 conn.close

			response.write "<b>Stock Watch List</b> <br>"		
			response.write replace(HeaderText & "" & smsStr1,chr(13),"<br>")
			end if
		end if

		
	 'Check whether the debit has been selected
	  if Contracts = 1 then	
		'get all the list of clients with mobile numbers
				
			set rs=server.createobject("Adodb.recordset")
			set rs1=server.createobject("Adodb.recordset")

			Set Conn = GetActiveConnection("KBroker")
			sqlStr= "SELECT * FROM smsContract WHERE (ClientCellTel IS NOT NULL) AND (ClientCellTel <> N'') and (updateOnContract=1)  ORDER BY Client_DPA_"
		sqlStr= "SELECT * FROM smsContract WHERE (ClientCellTel IS NOT NULL) AND (ClientCellTel <> N'')  ORDER BY Client_DPA_"

			

			rs.open sqlStr, conn,0,1
			
			show =1
			if rs.eof or rs.bof then
				%>
				
				<SCRIPT LANGUAGE="JavaScript">
				<!--
					alert("There no contracts created today for sms");
					window.history.back(0);

				//-->
				</SCRIPT>
				<%
				response.end
			else


			rs.movefirst
			cNo=0
				while not rs.eof 
				'create a string for the sms here
				' format the celltell number
				cellNo = replace(rs("ClientCellTel")," ","")
				if mid(cellNo,1,1)=0 then cellNo = mid(cellNo,2,10) else cellno = mid(cellNo,1,9)
				
				if not isnumeric(cellNo)  or trim(cellNo)="" then
					cellNo = cellNo & " *"
					cNo=1
				else
					cNo=0
					' add the prefix to the cell no
					cellNo = "CSV: 254" & cellNo		
				end if
				
				
				
				'if it is a purchase contract then
				if cint(cnNo)= cint(0) then
				'Response.Write cNo & "<br>"
				 if rs("OrderType_DPA_")=1 then
					smsStr =smsStr & cellNo &"|"& rs("LotQty") & " " & rs("SecurityCode")& " SHARES PURCHASED ON " & formatdate(rs("lotTdate")) & " @ KES " & rs("LotPrice")&" PER SHARE. SETTLEMENT OF KES " & rs("contractAmount") & " DUE " & formatdate(rs("contractSettlementDate"))& chr(13)
				 else

					smsStr =smsStr & cellNo  &"|"& rs("LotQty") & " " & rs("SecurityCode")& " SHARES SOLD ON "& formatdate(rs("lotTdate")) &" @ KES " & rs("LotPrice")& " PER SHARE " & chr(13)
				end if
				end if
					'response.write replace(rs("contractAmount")," ","")
					'response.end
				rs.movenext
				wend


'Create a text file here where u will save ur data
		'path = request.form("path")& "ClientStockWatch "& replace(formatdate(now()),"-","") &" "& replace(formatdatetime(now(),vbshorttime),":","") &".txt"
		'response.write path
		'response.end

		HeaderText ="api-id: 1127740" & chr(13)  
		HeaderText =HeaderText  & "user: african alliance" & chr(13) 
		HeaderText =HeaderText  & "password: alliance1" & chr(13) 
		HeaderText =HeaderText  & "text: #field1#" & chr(13)
		HeaderText =HeaderText  & "Delimiter: |" & chr(13)

		
		path1= strPathDebtors & "\ca_"& replace(date(),"/","") & ".txt"
		 set fso = server.createobject("Scripting.FileSystemObject")
		 set txtFile = fso.createTextFile(path1)
		 txtFile.WriteLine(HeaderText & "" & smsStr)
		 txtFile.close
		 set txtFile=nothing
		 set fso= nothing
		rs.close
		conn.close
	response.write "<br>&nbsp;<br><b>Contracts List</b><br>"
	response.write replace(HeaderText & "" & smsStr,chr(13),"<br>")
	end if

	end if

	'Check for debits

	If Debits =1 then
			
		'get all the list of clients with mobile numbers
				
			set rs=server.createobject("Adodb.recordset")
			set rs1=server.createobject("Adodb.recordset")

			Set Conn = GetActiveConnection("KBroker")
		
		'	rs.open sqlStr, conn,0,1
		
		 sqlStr = "select * from debita"
	     Rs.Open sqlStr, Conn, 0, 1
         Rs.MoveFirst
			show =1

			
			if rs.eof or rs.bof then
				%>
				
				<SCRIPT LANGUAGE="JavaScript">
				<!--
					alert("There are no sms clients with debts");
					window.history.back(0);

				//-->
				</SCRIPT>
				<%
				response.end
			else

			smsStr=""
			rs.movefirst
			
			While Not Rs.EOF
        'create a string for the sms here
        ' format the celltell number
			cellNo = Replace(Rs("ClientCellTel"), " ", "")
			If Mid(cellNo, 1, 1) = 0 Then cellNo = Mid(cellNo, 2, 10) Else cellNo = Mid(cellNo, 1, 9)
			If Not IsNumeric(cellNo) Or Trim(cellNo) = "" Then
				cellNo = cellNo & " *"
			Else
				' add the prefix to the cell no
				cellNo = "CSV: 254" & cellNo
			End If
			'get the the current balance to compare with
			
			Rs1.Open "select Balance from debitb where client_DPA_ =" & Rs("Client_DPA_"), Conn, 0, 1
			If Not Rs1.EOF Or Not Rs1.BOF Then
				If Rs1("Balance") < Rs("Balance") Then Bal = Rs("Balance") Else Bal = Rs1("Balance")
			Else
				Bal = Rs("Balance")
			End If
			Rs1.Close
			Bal = FormatNumber(0 - Bal, 2)
				smsStr = smsStr & cellNo & "|" & "Please note that you have an outstanding balance of KES " & FormatNumber(Bal, 2) & " overdue. Please settle." & Chr(13)
			Rs.MoveNext
			Wend

'Create a text file here where u will save ur data
		'path = request.form("path")& "ClientStockWatch "& replace(formatdate(now()),"-","") &" "& replace(formatdatetime(now(),vbshorttime),":","") &".txt"
		'response.write path
		'response.end

		HeaderText ="api-id: 1127740" & chr(13)  
		HeaderText =HeaderText  & "user: african alliance" & chr(13) 
		HeaderText =HeaderText  & "password: alliance1" & chr(13) 
		HeaderText =HeaderText  & "text: #field1#" & chr(13)
		HeaderText =HeaderText  & "Delimiter: |" & chr(13)

	 
		path1= strPathDebtors & "\db_"& replace(date(),"/","") & ".txt"
		
		 set fso = server.createobject("Scripting.FileSystemObject")
		 set txtFile = fso.createTextFile(path1)
		 txtFile.WriteLine(HeaderText &""& smsStr)
		 txtFile.close
		 set txtFile=nothing
		 set fso= nothing
		 rs.close
		 conn.close
response.write "<br><br><b>Debits List</b><br>"
response.write replace(HeaderText &""& smsStr,chr(13),"<br>")
end if

'response.end

	End if

		%>
		&nbsp;<BR>&nbsp;
		<input type ="button" name ="back" value="<< Back" onclick ="javascript:history.back(1);">
		<%
		end if
		

	%>
</BODY>
</HTML>
