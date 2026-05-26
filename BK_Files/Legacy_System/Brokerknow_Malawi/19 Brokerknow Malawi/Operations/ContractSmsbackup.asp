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

<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT>

	<script language="JavaScript">
	 function GetDirectory() {
	  strFile = document.frmContractsms.path.value;
	  intPos = strFile.lastIndexOf("\\");
	  strDirectory = strFile.substring(0, intPos);
	   document.frmContractsms.path1.value =strDirectory;
	 // alert(strFile + '\n\n' + strDirectory);
	  return false;
	 }
	</script>
</HEAD>
<CENTER><FONT SIZE="3" COLOR=""><B>CONTRACTS</B></FONT></CENTER>

<%
'======================= Begin_Alter_Across_Entities =================================
		
			const UDLName = "KBroker"
		const DataSource = "smsContract"
		const DataEntity = "smsContract"
		const DataEntityPlural = "smsContract"
		const ActionFolder = "Operations"
		const ActionPage = "clientSmsContractList"
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
			sqlStr= "SELECT * FROM smsContract WHERE (ClientCellTel IS NOT NULL) AND (ClientCellTel <> N'') and (updateOnContract=1)  ORDER BY Client_DPA_"
sqlStr= "SELECT * FROM smsContract WHERE (ClientCellTel IS NOT NULL) AND (ClientCellTel <> N'')  ORDER BY Client_DPA_"

			

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

		path = request.form("path")
		path1 = request.form("path1") 
		path1= path1 & "\ca_"& replace(date(),"/","") & ".txt"
		 set fso = server.createobject("Scripting.FileSystemObject")
		 set txtFile = fso.createTextFile(path1)
		 txtFile.WriteLine(HeaderText & "" & smsStr)
		 txtFile.close
		 set txtFile=nothing
		 set fso= nothing

response.write replace(HeaderText & "" & smsStr,chr(13),"<br>")
end if

'response.end

%>		
<br><br>
		<input type="button" name="back" value=" Back " onclick="javascript:window.history.back(0);">
		<%
		response.end
	end if
%>


<BODY>
<FORM METHOD=POST ACTION="ContractSms.asp" name ="frmContractsms">&nbsp;<BR>&nbsp;<BR>
<input type="file" size ="50"  name="path" value="C:\Knownig\Brokerknow\Imports\SMS\Default.txt"><br>&nbsp;<BR>
<INPUT TYPE="submit" name="cmdGenerate" value= "Generate" onclick="GetDirectory();">
<input type="hidden" name="path1" value="">
<input type="hidden" name="gen" value="1">
</FORM>
</BODY>
</HTML>
