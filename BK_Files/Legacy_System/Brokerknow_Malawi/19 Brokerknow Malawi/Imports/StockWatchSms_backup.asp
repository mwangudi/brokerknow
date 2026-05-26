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
</HEAD>

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
			rs.movefirst
			client =rs("Client_DPA_")
				'create a string for the sms here
				while not rs.eof 
				
				
					cellNo = replace(rs("clientCellTel")," ","")
					PricenSec =PricenSec & rs("SecurityCode")	& " "  & rs("Price") & " "
					if client <> rs("Client_DPA_")  then 
						
						PricenSec =  rs("SecurityCode")	& " "  & rs("Price") & " "
						
						if  trim(smsStr3) <> trim(smsStr2)  then
							
							smsStr3= smsStr2 
								 
							'show is is used to make sure the string is not written on the text file
							if show=0 then
								smsStr4 =smsStr4 + smsStr3 & chr(13)
							end if
							sms_Str3 = sms_Str3  & sms_Str1
							
							
						else
							smsStr3=""							
						end if
						
					else
						smsStr2 = ""& cellNo &"," & PricenSec 
						if not isnumeric(cellNo)  and len(trim(cellNo))<>10 then
							sms_Str1 = sms_Str & "<font color='red'>"& "(client code - "& rs("Client_DPA_")& " ->"& Cellno & ")  "& cellNo &"</font>," & PricenSec  & chr(13)
						else
							show=0
							sms_Str1 = sms_Str & ""& cellNo &"," & PricenSec  & chr(13)
							
						end if
					end if
					
					if rs("stockCount")=1 then
						
						if not isnumeric(cellNo)  and len(trim(cellNo))<>10 then
							sms_Str = sms_Str & "<font color='red'>"& "(client code - "& rs("Client_DPA_")& " ->"& Cellno &")"& cellNo &"</font>," & PricenSec  & chr(13)
						else
							sms_Str = sms_Str & ""& cellNo &"," & PricenSec  & chr(13)
							smsStr= smsStr & ""& cellNo &"," & PricenSec  & chr(13)
						end if
						'response.write "'"& cellNo &"','" & rs("SecurityCode")	& " "  & rs("Price") & "<br> "
					end if
					client =rs("Client_DPA_")
				rs. movenext
				
				wend
				
			end if
		smsStr1 = smsStr & smsStr4 
		sms_Str3 = sms_Str & sms_Str3 
		'Create a text file here where u will save ur data
		'path = request.form("path")& "ClientStockWatch "& replace(formatdate(now()),"-","") &" "& replace(formatdatetime(now(),vbshorttime),":","") &".txt"
		'response.write path
		'response.end
		path = request.form("path")
		 set fso = server.createobject("Scripting.FileSystemObject")
		 set txtFile = fso.createTextFile(path )
		 txtFile.WriteLine(smsStr1)
		 txtFile.close
		 set txtFile=nothing
		 set fso= nothing
		
		response.write replace(sms_Str3,chr(13),"<br>")
		%>
		<br><br>
		<input type="button" name="back" value=" Back " onclick="javascript:window.history.back(0);">
		<%
		response.end
	end if
%>

<BODY>
<FORM METHOD=POST ACTION="stockWatchSms.asp">&nbsp;<BR>&nbsp;<BR>
<input type="text" size ="50"  name="path" value="C:\Knownig\Brokerknow\Imports\SMS\generated.txt"><br>&nbsp;<BR>
<INPUT TYPE="submit" name="cmdGenerate" value= "Generate">
<input type="hidden" name="gen" value="1">
</FORM>
</BODY>
</HTML>
