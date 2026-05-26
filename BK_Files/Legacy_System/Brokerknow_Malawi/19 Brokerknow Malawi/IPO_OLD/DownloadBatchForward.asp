<%OPTION EXPLICIT%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Download Batch Forwards</title>
	<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
	<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
	<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
	<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
	<SCRIPT language=VBScript src="../scripts/reports.vbs"></SCRIPT>
	<SCRIPT language=Javascript src="../scripts/reports.js"></SCRIPT>

	<script language="javascript">
	 function isNumeric(val){return(parseFloat(val,10)==(val*1));}

	 function validateNumeric(ctrl){
	  var value = ctrl.value;
	  
	  if (value!=''){
		if (isNumeric(value)){}else{ctrl.value=''}
	  }
	 }
	</script>
</head>

<body Class="Reports">
<!--#include file="../libroutines.asp"-->
<%
function WriteTextFile (FilePath, content) 
 dim objCSVFile , objFSO
 Set objCSVFile = CreateObject("ADODB.Stream")	
 Set objFSO = CreateObject("Scripting.FileSystemObject")
 objCSVFile.Open

 objCSVFile.CharSet = "Ascii"
 objCSVFile.WriteText trim(content)
 
 objCSVFile.SaveToFile FilePath, 2
 
 set objCSVFile = nothing

 WriteTextFile  = true
 
 if err.number then WriteTextFile  = false
 On Error GoTo 0
end function

  function savefile(FPath,text)
	'TO ALLOW THE USER TO SAVE THE DOCUMENT
	Dim adoStream, str
	Dim objFSO,objFile, fname  
	const adTypeText = 2
	const adTypeBinary = 1
    
	Set objFSO = CreateObject("Scripting.FileSystemObject")
    
	If objFSO.FileExists(FPath) Then
		
		set objFile = objFSO.Getfile(FPath)
		fname =  objFile.Name
        
		Response.Clear 'first clear the response, and then set the appropriate headers 
		Response.ContentType = "text/csv"
		Response.AddHeader "Content-Disposition", "inline;attachment; filename=""" & fname & """"
		Response.AddHeader "Content-Length", objFile.Size 'The header Content-Length is set so that the browser can properly display the progress bar.
		
		'adoStream.Charset = "ascii"
		'Response.CharSet = "UTF-8" 

		Set adoStream = CreateObject("ADODB.Stream") 
		adoStream.Type = adTypeBinary 
		adoStream.Open() 
		adoStream.LoadFromFile FPath
		Response.BinaryWrite adoStream.Read() 
		adoStream.Close 

		set objFile = nothing
	else
        %>
		 <script language="javascript">
		  alert("Order File not found. Please contact the system administrator.")
		 </script>
		<%
	end if

	Set adoStream = Nothing 
	Set objFSO = Nothing 
    
 end function

function GenerateContent(CDANo,DNumber,Bcode,FilterSQL,fname)

Dim rst, sqlstr, cnn, ContentStr, rsbatch
Dim AppNo, TotalAmt, TotalQty
Dim BatchHeader, BatchStr, BatchNo, BatchType, BatchID, BatchLine
Dim FileHeader, FileFooter
Dim intRstCount, intRsBatchCount
Dim intCount,count
Dim astrRecs, batchdata
Dim AppSerialNo 
Dim batchlineNo 
Dim ClientName 
Dim Qty 
Dim Amount 
Dim PaymentMethod 
Dim Withdraw 
Dim DownloadID
Dim OfferingID

Dim TotalBatches 'count of batches
Dim TotalAppNo ' Total number of applications
Dim TotalShares ' Total no of shares
Dim TotalAgentAmount ' total amount applied for through the agent i.e Discount securities
Dim TotalChequeAmount ' total amount paid directly to the company

Set cnn = Server.CreateObject("ADODB.Connection")
Set rst = Server.CreateObject("ADODB.Recordset")
Set rsbatch = Server.CreateObject("ADODB.Recordset")
Set cnn = GetActiveConnection("KBroker")

BatchType = 1 'Cheques accounts
DownloadID = DNumber
DNumber = right("000" & DNumber,3)
 
cnn.Commandtimeout = 0

'Get sums of batch values
sqlstr = "SELECT Batch_No as BatchNo, SUM(Alloted_Rights) AS TotalQty,  " & _
		 "     SUM(ISNULL(Alloted_Rights, 0) * ISNULL(Offering_Price, 0)) AS TotalAmt, COUNT(Offering_DPA_)  " & _
		 "     AS AppCount " & _
		 " FROM  ForwardsList " & _
		 " where (NOT (Batch_No IS NULL)) AND " & FilterSQL & _
		 " GROUP BY Batch_No"

set rst = cnn.execute(sqlstr)

intRstCount = rst.recordcount

'Prepare to note records downloaded last
 sqlstr = "UPDATE  Offerings " & _
		  " SET LastDownloaded =0 where Forward  = 1"
 
 cnn.execute(sqlstr)
 
if intRstCount > 0 then
 rst.movefirst
 
 astrRecs = rst.getrows()

'Constract File Header
FileHeader = 1 & vbTab & CDANo & vbTab & "00" & vbTab & DNumber & vbTab & Year(Now()) & "-" & right("00" & Month(Now()),2) & "-" & right("00" & Day(Now()),2) & vbTab & vbcrlf

  TotalBatches = 0
  TotalAppNo = 0
  TotalShares = 0
  TotalAgentAmount = 0
  TotalChequeAmount = 0 'Remains zero coz this is an application through the broker

  for intCount = 0 to intRstCount-1 
   BatchID = trim(astrRecs(0,intCount))
   BatchNo = right(Bcode,2) & "00" & right("0000" & BatchID,4)
   AppNo = trim(astrRecs(3,intCount))
   TotalQty = trim(astrRecs(1,intCount))
   TotalAmt = trim(astrRecs(2,intCount))

    BatchHeader = 2 & vbTab & CDANo & vbTab & BatchNo & vbTab & BatchType & vbTab & AppNo & vbTab & TotalQty & vbTab & FormatNumEx(TotalAmt,2) & vbTab & vbcrlf

	'Generate batch line items
    sqlstr = "SELECT ForwardsList.Batch_No, ForwardsList.Client_DPA_, Client.ClientName, ForwardsList.Offering_DPA_, ForwardsList.BatchSeq, ForwardsList.PAL_No As SerialNo,  " & _
			 "                       ForwardsList.Alloted_Rights as Qty, ForwardsList.Offering_Price As Price, ForwardsList.Offering_Price * ForwardsList.Alloted_Rights AS Amount " & _
			 " FROM         ForwardsList INNER JOIN " & _
			 "                       Client ON ForwardsList.Client_DPA_ = Client.Client_DPA_ " & _
			 " WHERE     (ForwardsList.Batch_No = " & BatchID & ") " & _
			 " ORDER BY ForwardsList.BatchSeq"
   
    set rsbatch = cnn.execute(sqlstr)

	intRsBatchCount = rsbatch.recordcount
    
	if intRsBatchCount > 0 then

	   rsbatch.movefirst
	   batchdata = rsbatch.getrows()
	   BatchLine = ""
	  for count = 0 to intRsBatchCount-1
        OfferingID = trim(batchdata(3,count))
        AppSerialNo = right(trim(batchdata(5,count)),6) 'trim to 6 characters to the right
		batchlineNo = trim(batchdata(4,count))
		ClientName = trim(batchdata(2,count))
		Qty = Cdbl(trim(batchdata(6,count)))
		Amount = FormatNumEx(trim(batchdata(8,count)),2)
		PaymentMethod = BatchType ' Same as Batch Type
		Withdraw = 0'(0 = Normal; 1 = Withdraw Application)

		BatchLine = BatchLine & 3 & vbTab & CDANo & vbTab & BatchNo & vbTab & batchlineNo & vbTab & AppSerialNo & vbTab & ClientName & vbTab & Qty & vbTab & Amount & vbTab & PaymentMethod & vbTab & 0 & vbTab & 0 & vbTab & Withdraw & vbTab  & vbcrlf
	    
	  next

    end if
     BatchStr = BatchStr & BatchHeader & BatchLine

	 'Do the totals here
	 TotalBatches = TotalBatches + 1
	 TotalAppNo = TotalAppNo + AppNo
	 TotalShares = TotalShares + TotalQty
	 TotalAgentAmount = 0
	 TotalChequeAmount = TotalChequeAmount + TotalAmt
	 
	'Mark records in batch as downloaded also update file name of download 
        sqlstr = "UPDATE    Offerings " & _
				 " SET      Downloaded = 1, LastDownloaded =1, BatchFileName='" & fname & "',Download_DPA_ =  " & DownloadID & _
				 " WHERE    (Batch_No = " & BatchID & ") AND Forward  = 1  "

		cnn.execute(sqlstr)

  next
 
  'Generate file footer
  FileFooter = 4 & vbTab & CDANo & vbTab & "00" & vbTab & DNumber & vbTab & Year(Now()) & "-" & right("00" & Month(Now()),2) & "-" & right("00" & Day(Now()),2) & vbTab & TotalBatches & vbTab & TotalAppNo & vbTab & TotalShares & vbTab & FormatNumEx(TotalAgentAmount,2) & vbTab & FormatNumEx(TotalChequeAmount,2) & vbTab

 ContentStr = FileHeader & BatchStr & FileFooter
else
 ContentStr = ""
end if

'Clean up
Set rst = nothing
Set rsbatch = nothing
Set cnn = nothing

GenerateContent = ContentStr
end function

Dim genReport, ToNo, FromNo

genReport = trim(Request.Form("genReport"))
ToNo = trim(Request.Form("ToNo"))
FromNo = trim(Request.Form("FromNo"))

If genReport <> "1" Then%>
<Script Language="JavaScript">
		report_SetBodyClass();
		function validateForm(frm){			
			if (frm.FromNo.value==''){
			    alert("Please specify the from batch number.");
				frm.FromNo.focus();
				return;
			}
			
			frm.target = '_self';			
			frm.submit();
		}
		
	</Script>
	<form method="POST" action="DownloadBatchForward.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport" id="genReport">
		<table>
			
			<tr>
				<td><b>Please specify batch number to download</b></td>
				<td>&nbsp;</td>
			</tr>
			<tr>
				<td colspan="2">From &nbsp;<input type="text" name="FromNo" id="FromNo" size="7" onChange = "javascript: validateNumeric(this);"></td>
			</tr>
			<tr>
				<td colspan="2">To&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" name="ToNo" id="ToNo" size="7" onChange = "javascript: validateNumeric(this);"></td>
			</tr>
			<tr>
				<td colspan="2">&nbsp;</td>
			</tr>
			<tr>
				<td colspan=2><input type="Button" class="submit" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Download... ">&nbsp;&nbsp;</td>
			</tr>
		</table>
		
	</form>
	<%
	Response.End
End If

'Construct batch parameters 
Dim Filterstr

if ToNo <> "" AND FromNo <> "" then
 Filterstr = "Batch_No between " & FromNo & " AND " & ToNo & " "
else
 Filterstr = "Batch_No = " & FromNo & " "
end if

' Get Broker code

Dim BrokerCode, CdaID
Dim rs, Conn
Dim DownloadNo

Set Conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set Conn = GetActiveConnection("KBroker")

sqlstr = "SELECT Broker.BrokerCode " & _
		 " FROM  Broker INNER JOIN " & _
		 " CompanyInfo ON Broker.Broker_DPA_ = CompanyInfo.Broker_DPA_"

set rs = conn.execute(sqlstr)

if not (rs.eof or rs.bof) then
 BrokerCode = trim(rs.fields("BrokerCode"))
else
 %>
 <script language="javascript">
  alert("Important company information was not found. Please contact the system administrator.")
 </script>
<%
response.end
end if

BrokerCode = "B" & BrokerCode
CdaID = BrokerCode & " B" ' CDA ID

'Get Download No

sqlstr = "SELECT ISNULL(MAX(Download_DPA_), 0) + 1 AS Download_DPA_ FROM  Offerings"
set rs = conn.execute(sqlstr)

DownloadNo =  trim(rs.fields("Download_DPA_"))

'Get path where to save file (.csv)
Dim pathToFile, parentDirectory, filename
Dim Success, filecontent, sqlstr

parentDirectory = ".."	

'FILE NAME: XXXXXC99D999.txt
    'XXXXX = CDA ID e.g. B01 B  broker codes
	'C = Literal C i.e. Computer
	'99 = Agent's stand alone Computer Number 01 thru 99 else 00 if on Network 
	'D = Literal D i.e. Download
	'999 = Current download number (This is incremented by 1 for each download)

filename= CdaID & "C00D" & right("000" & DownloadNo,3) & ".txt"

pathToFile = Server.MapPath(parentDirectory) & "\IPOS\" & filename

filecontent = GenerateContent(CdaID,DownloadNo,BrokerCode,Filterstr,filename)

if trim(filecontent) = "" then
 %>
	 <script language="javascript">
	  alert("No batch forwards were found for downloading.")
	  document.location.href = 'DownloadBatchForward.asp';
	 </script>
 <% response.end 
end if

Success = WriteTextFile(pathToFile,filecontent)

if Success then
' Save file to a location
call savefile(pathToFile,filecontent)

else
 %>
 <script language="javascript">
  alert("An error has been encountered while downloading batch applications. Please contact the system administrator.")
 </script>
<%
end if

Set Conn = nothing
Set rs = nothing
 response.end
%>
</body>
</html>