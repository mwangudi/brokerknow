  <html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Contract Schedule</title>
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
	 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
	 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
	<!--CALENDAR -->
	<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
	<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
	<SCRIPT language=VBScript src="../scripts/reports.vbs"></SCRIPT>
	<SCRIPT language=Javascript src="../scripts/reports.js"></SCRIPT>
	
	<style media="print">
	
		@page {
			
			margin-left: 1cm;
			margin-right: 1cm;
			margin-top: 1cm;    
			margin-bottom: 1cm;
			writing-mode: tb-rl;
			height: 90%;
			margin: 10% 0%;			
			
			br.newpage{
				page-break-before:always;
			}
			
			
		}
		 
		
	</style>


</head>

<body Class="Reports">

<Script Language="JavaScript">
	function HideRemindSelectLandscape(){
		try{			
			document.getElementById('landRem').style.display = 'none';
		}	
		catch(e){}	
	}
	function ShowRemindSelectLandscape(){
		try{			
			document.getElementById('landRem').style.display = '';
		}	
		catch(e){}	
	}
	window.onbeforeprint = HideRemindSelectLandscape;
	window.onafterprint = ShowRemindSelectLandscape;
</Script>

<!--#include file="../libroutinesTEST.asp"-->

<%
'if enabled then
const beginLeviesCol = 15

genReport = Request.Form("genReport")
selectedTradeDate = Request.Form("txtDate")

If genReport <> "1" Or Not IsDate(selectedTradeDate) Then%>
	<Script Language="JavaScript">
		report_SetBodyClass();		
		function validateForm(frm){			
			if (frm.txtDate.value==''){
				alert("Select a date");
				frm.txtDate.focus();
				return;
			}
			
			frm.target = '_self';			
			frm.submit();
		}
		
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "txtDate","cmdDate","<%= FormatDate(Date) %>",1);
	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="downloadContractschedule.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">	
		<table>
			<tr>
				<td>Select Date</td>
				<td>
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>
			</tr>
			<tr>
				<td colspan=2><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... ">&nbsp;&nbsp;</td>
			</tr>
		</table>
		
	</form>
	
	<%Set rs = Nothing
	Set Conn = Nothing
	Response.End
End If

%>

<!--<td bgColor="#000000" nowrap>&nbsp;</td>-->
      
<%

   Response.Buffer = True
   server.scripttimeout=100000000
   %>
      <%
		Dim fld
		Dim conn 
		Dim sqlStr
		Dim rs
		Dim i
		Dim dailyTotalsArray()
			
		Set conn = GetActiveConnection("KBroker")
 		sqlStr = "ContractLeviesCrossTabToDate '" & FormatDate(selectedTradeDate) & "'"	
		Set Rs = CreateObject("ADODB.Recordset")		
		Rs.CursorLocation = adUseClient		
		  'response.write sqlStr
          'response.end
		Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
		
		'Rs.Filter = "LotTDate = '" & FormatDate(selectedTradeDate) & "'"
		
		if not isEmpty(rs) Then 'If nothing comes back do not try
        if rs.recordcount >=1 then 'with more than one record
          
          
         ' create the temp file.
          Dim fso, MyFile,sRTF, sFileName,sTitle,sTitle2
          Set fso = CreateObject("Scripting.FileSystemObject")
          
          sfname = "Contracts_" & FormatDate(selectedTradeDate) & replace(replace(time(),":","")," ","")
          sfname = "Contracts_" & FormatDate(selectedTradeDate) & "_" & "y" & year(now) & "m" & month(now) & "d" & day(now) & "h" & hour(now) & "m" & minute(now)& "s" & second(now)
          sfname = sfname & ".csv"
          
          'response.write replace(replace(time(),":","")," ","")
          'response.end
          Set MyFile = fso.CreateTextFile(Server.MapPath(".") & "\" & sfname, True)
          
          
          
          'Stitle ="Traded,Sett Date,Code,Client,Sec,Cust,Agent,Br.,Contr,CDS,Price,Qty,Agen,AKS,CDSC,CMA,Comm,Comp,Compensate,NSE,Stam,Tran,Gross,Net"
		  Stitle ="Traded,Sett Date,Code,Client,Sec,Cust,Agent,Br.,Contr,CSD,Price,Qty,Agen,Comm,Comp,Compensate,Gross,Net"
    
             MyFile.WriteLine(Stitle)
          
		'i = 0
		'fldCount = fldCount + 2 'this is hard coding just to skip some unwanted columns 
		for i = beginLeviesCol to rs.fields.count - 2
				Redim Preserve dailyTotalsArray(i - beginLeviesCol)
				dailyTotalsArray(i - beginLeviesCol) = 0
				'i = i + 1
		next
		
		'add gross, and net amount
		Redim Preserve dailyTotalsArray((i) - beginLeviesCol)
		dailyTotalsArray((i) - beginLeviesCol) = 0
		
		Redim Preserve dailyTotalsArray((i + 1) - beginLeviesCol)
		dailyTotalsArray((i + 1) - beginLeviesCol) = 0
		
		'Redim Preserve dailyTotalsArray((i + 2) - beginLeviesCol)
		'dailyTotalsArray((i + 2) - beginLeviesCol) = 0
		%>
	 	
	   <%
 
 
        'Now Loop for each record we find in the Database
        'Do Until rs.EOF
      Do while not (rs.eof)
		
		sFileNamestr = ""
		sFileNamestr = sFileNamestr & FormatDate(rs.Fields("LotTDate")) & ","
		sFileNamestr = sFileNamestr & FormatDate(rs.Fields("ContractSettlementDate")) & ","
		sFileNamestr = sFileNamestr & (rs("Client_DPA_")) & ","
		sFileNamestr = sFileNamestr & (Replace(rs("ClientName"),",","&")) & ","
		sFileNamestr = sFileNamestr & (rs("SecurityCode")) & ","
		sFileNamestr = sFileNamestr & IIF(cbool(rs.Fields("IsCustodian")) = True, "Y", "N") & ","
		
		if len(left(trim(rs("AgentName")),1)) = 1 then 
			agentStr = left(replace(rs("AgentName")," ",""),7)
		else
			agentStr = ""
		end if
		
		sFileNamestr = sFileNamestr & (agentStr) & ","

		sFileNamestr = sFileNamestr & (rs("BrokerCode")) & ","
		sFileNamestr = sFileNamestr & (rs("ContractNumber")) & ","
		sFileNamestr = sFileNamestr & (rs("LotSlipNo")) & ","
		sFileNamestr = sFileNamestr & (rs.Fields("LotPrice")) & ","
		sFileNamestr = sFileNamestr & (rs.Fields("LotQty")) & ","

		sFileNamestr = sFileNamestr & (rs.Fields("Agent")) & ","
        levyTotals = 0
		'sFileNamestr = sFileNamestr & (rs.Fields("AKS")) & ","
		'levyTotals = levyTotals + (rs.Fields("AKS"))
		'sFileNamestr = sFileNamestr & (rs.Fields("CDSC")) & ","
		'levyTotals = levyTotals + (rs.Fields("CDSC"))
		'sFileNamestr = sFileNamestr & (rs.Fields("CMA")) & ","
		'levyTotals = levyTotals + (rs.Fields("CMA"))
		sFileNamestr = sFileNamestr & (rs.Fields("Commission")) & ","
		levyTotals = levyTotals + (rs.Fields("Commission"))
		'sFileNamestr = sFileNamestr & (rs.Fields("Comp")) & ","
		'levyTotals = levyTotals + (rs.Fields("Comp"))
		sFileNamestr = sFileNamestr & "0" & ","
		levyTotals = levyTotals + 0
		'sFileNamestr = sFileNamestr & (rs.Fields("Compensate")) & ","
		'levyTotals = levyTotals + (rs.Fields("Compensate"))
		'sFileNamestr = sFileNamestr & (rs.Fields("NSE")) & ","
		'levyTotals = levyTotals + (rs.Fields("NSE"))
		'sFileNamestr = sFileNamestr & (rs.Fields("Stamps")) & ","
		'levyTotals = levyTotals + (rs.Fields("Stamps"))
		'sFileNamestr = sFileNamestr & (rs.Fields("Transfer")) & ","
		'levyTotals = levyTotals + (rs.Fields("Transfer"))

		grossAmt = rs.Fields("LotGrossAmount") 'rs.Fields("LotPrice") * rs.Fields("LotQty") 

		If rs.Fields("OrderTypeSale").Value = 0 Then 
			netAmt = grossAmt + levyTotals 
		Else
			netAmt = grossAmt - levyTotals 
		End If
            		     
		sFileNamestr = sFileNamestr & (grossAmt) & ","      
		sFileNamestr = sFileNamestr & (netAmt) & "," 
		
		rs.MoveNext
           
     
 'loop and write the next row of cells
'rs.movenext
    
'Also notice the first cell is empty and just for looks
              sRTF = sFileNamestr 
              
              MyFile.WriteLine(sRTF) 'Write the Text to the open file
              
                          
            
           'rs.MoveNext 'Loop and write the next row or cells
          loop
          
          
       '//nitairudisha
        ' 8 + UBound(dailyTotalsArray)
        
				%>
        
      	
				<%	
				else
				%>
                <script language = 'javascript'>
                		alert ("No contracts found using the specified criteria");
                		window.parent.history.go(-1);          		
                </script>
                
                <%  Set Rs = Nothing
					Set Conn = Nothing
                Response.end

          rs.Close
        Set rs = Nothing
        MyFile.Close
      end if
     end if
     
     Response.Write "<META HTTP-EQUIV=""REFRESH"" Content=""0;URL=" & sfname & """>"
     
     Response.Write "<br><p style=""font-size:10pt;font-family:arial;text-align:center;"">Download complete.</p>"
     Response.End 

        
function savefile(FPath)
	'TO ALLOW THE USER TO SAVE THE DOCUMENT
	Dim adoStream, str
	Dim objFSO,objFile, fname  
	const adTypeText = 2
	const adTypeBinary = 1
    
	Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    'response.write FPath
    'response.end 
    
	If objFSO.FileExists(FPath) Then
		
		set objFile = objFSO.Getfile(FPath)
		fname =  objFile.Name
        
		Response.Clear 'first clear the response, and then set the appropriate headers 
		Response.ContentType = "application/unknown"
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
        Response.Write "<br><p style=""font-size:10pt;font-family:arial;text-align:center;"">Download failed.</p>"
        response.end
	end if

	Set adoStream = Nothing 
	Set objFSO = Nothing 
    
 end function
%>