<!--#include file="../libroutines.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE> Generate Stock Watch SMS </TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT>

	<script language="JavaScript">
	 function GetDirectory() {
	  strFile = document.frmStockWatch.path.value;
	  intPos = strFile.lastIndexOf("\\");
	  strDirectory = strFile.substring(0, intPos);
	   document.frmStockWatch.path1.value =strDirectory;
	 // alert(strFile + '\n\n' + strDirectory);
	  return false;
	 }
	</script>
</HEAD>
<CENTER><FONT SIZE="3" COLOR=""><B>STOCK WATCH</B></FONT></CENTER>

<%
'======================= Begin_Alter_Across_Entities =================================
		
		const UDLName = "KBroker"
		const DataSource = "StockWatchList"
		const DataEntity = "StockWatchList"
		const DataEntityPlural = "StockWatchLists"
		const ActionFolder = "Operations"
		const ActionPage = "stockWatchSms.asp"
'======================= End_Alter_Across_Entities =================================
	
		Dim action
		Dim conn 
		Dim sqlStr
		Dim rs
		Dim guidStr 
		Dim guid 
	gen=request.form("gen")
	if trim(gen)<>"" then
		'get all the list of clients with mobile numbers
				
			set rs=server.createobject("Adodb.recordset")
			set rs1=server.createobject("Adodb.recordset")

			Set Conn = GetActiveConnection("KBroker")
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

		path = request.form("path")
		path1 = request.form("path1") 
		path1= path1 & "\sw_"& replace(date(),"/","") & ".txt"
		
		'Response.Write path1
		'Response.end
		 set fso = server.createobject("Scripting.FileSystemObject")
		 set txtFile = fso.createTextFile(path1)
		 txtFile.WriteLine(HeaderText & "" & smsStr1)
		 txtFile.close
		 set txtFile=nothing
		 set fso= nothing
		
response.write replace(HeaderText & "" & smsStr1,chr(13),"<br>")
end if
%>
		<br><br>
		<input type="button" name="back" value=" Back " onclick="javascript:window.history.back(0);">
		<%
		response.end
	end if
%>

<BODY>
<FORM METHOD=POST ACTION="stockWatchSms.asp" name ="frmStockWatch">&nbsp;<BR>&nbsp;<BR>
<input type="file" size ="50"  name="path" value="C:\Knownig\Brokerknow\Imports\SMS\Default.txt"><br>&nbsp;<BR>
<INPUT TYPE="submit" name="cmdGenerate" value= "Generate" onclick="GetDirectory();">
<input type="hidden" name="path1" value="">
<input type="hidden" name="gen" value="1">
</FORM>
</BODY>
</HTML>
