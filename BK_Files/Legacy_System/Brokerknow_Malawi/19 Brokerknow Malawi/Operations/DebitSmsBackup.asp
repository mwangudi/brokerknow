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
	  strFile = document.frmDebitSms.path.value;
	  intPos = strFile.lastIndexOf("\\");
	  strDirectory = strFile.substring(0, intPos);
	   document.frmDebitSms.path1.value =strDirectory;
	 // alert(strFile + '\n\n' + strDirectory);
	  return false;
	 }
	</script>
</HEAD>
<CENTER><FONT SIZE="3" COLOR=""><B>CLIENT DEBITS</B></FONT></CENTER>
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
			sqlStr= "SELECT * FROM smsDebtors WHERE (ClientCellTel IS NOT NULL) AND (ClientCellTel <> N'') ORDER BY Client_DPA_"

			'set the debt date to be 5 days b4 leo

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
			 " HAVING (b.CurrentBal < 0) AND (Client.updateOnDebt = 1) AND (Client.ClientCellTel IS NOT NULL) AND (RTRIM(LTRIM(Client.ClientCellTel)) <> '')"
			

		'	response.write  sqlStr
		'	response.end
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
					alert("There are no sms clients with debts");
					window.history.back(0);

				//-->
				</SCRIPT>
				<%
				response.end
			else


			rs.movefirst
			
				while not rs.eof 
				'create a string for the sms here
				' format the celltell number
				cellNo = replace(rs("ClientCellTel")," ","")
				if mid(cellNo,1,1)=0 then cellNo = mid(cellNo,2,10) else cellno = mid(cellNo,1,9)
				if not isnumeric(cellNo)  or trim(cellNo)="" then
					cellNo = cellNo & " *"
				else
					' add the prefix to the cell no
					cellNo = "CSV: 254" & cellNo		
				end if
					smsStr =smsStr & cellNo &"|" & "Please note that you have an outstanding balance of KES " & formatnumber(0-rs("Balance"),2) & " overdue. Please settle."& chr(13)
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

		
		path1= path1 & "\db_"& replace(date(),"/","") & ".txt"
		
		 set fso = server.createobject("Scripting.FileSystemObject")
		 set txtFile = fso.createTextFile(path1)
		 txtFile.WriteLine(HeaderText &""& smsStr)
		 txtFile.close
		 set txtFile=nothing
		 set fso= nothing
		
response.write replace(HeaderText &""& smsStr,chr(13),"<br>")
end if

'response.end
end if
%>		



<BODY>
<FORM METHOD=POST ACTION="DebitSms.asp" name="frmDebitSms">&nbsp;<BR>&nbsp;<BR>
<input type="file" size ="50"  name="path" value="C:\Knownig\Brokerknow\Imports\SMS\Default.txt"><br>&nbsp;<BR>
<INPUT TYPE="submit" name="cmdGenerate" value= "Generate" onclick="GetDirectory();">
<input type="hidden" name="path1" value="">
<input type="hidden" name="gen" value="1">
</FORM>
</BODY>
</HTML>
