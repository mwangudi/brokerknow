<%
Set conn = GetActiveConnection("KBroker")
Set Rs = CreateObject("ADODB.Recordset")
Set Rs2 = CreateObject("ADODB.Recordset")

If Len(Request.QueryString("f")) > 0 Then
	Response.ContentType = "application/x-unknown" ' arbitrary 
    'Response.ContentType = "application/octet-stream"
    
    fn = Request.QueryString("f")
    
    FPath = Server.MapPath(".") & "\IPO FILES\" & fn
    
    Response.AddHeader "Content-Disposition", "attachment; filename=""" & fn & """"

    Set adoStream = CreateObject("ADODB.Stream") 
    adoStream.Open() 
    'adoStream.Charset = "UTF-8"
    'Response.CharSet = "UTF-8"
    adoStream.Type = 1 
    adoStream.LoadFromFile(FPath) 
    Response.BinaryWrite adoStream.Read()
    adoStream.Close 
    Set adoStream = Nothing 
	
	Response.Flush
	
	Response.End 
End If

If Request.QueryString("Offering") = "" Then
	SQL = "SELECT Security_DPA_ FROM [SecurityListOfferings] " & _
	" WHERE CAST(FLOOR(CAST(ClosingDate AS FLOAT)) AS DATETIME) >= CAST(FLOOR(CAST(GetDate() AS FLOAT)) AS DATETIME)" & _
	" ORDER BY ClosingDate DESC"
	
	Set Rs = conn.Execute(SQL)
						
	If Not (Rs.EOF Or Rs.BOF) Then
		%>
		<script language="javascript">window.location.href='DownloadBatches.asp?Offering=<%=Rs("Security_DPA_")%>'</script>
		<%
	Else
		Offering = 1
	End If
End If
%>
<head>
<title>Download Batches</title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<script language="JavaScript" src="CALENDAR/calendar.js"></script>
<!--END CALENDAR -->

<style media="print">
	@page {
		@top{font-family: Helvetica, Arial, sans-serif;
			font-size: 150%;
			font-weight: bolder;
			text-align: left;
			content: "<%= FormatDate(Date) %>";			
		}
			
		margin-left: 2cm;
		margin-right: 5cm;
		margin-top: 1cm;    
		margin-bottom: 2cm;
		size: portrait;
			
		br.newpage{
			page-break-before:always;
		}
			
	}
</style>

<script language="javascript">
function filterBatch(myOpt)
	{
	if (myOpt == ''){
		alert("Invalid option selected");
		return false;
		}
		
	window.location.href = 'DownloadBatches.asp?Offering=<%=request.querystring("offering")%>&filter=' + myOpt
	}
	var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "txtDate","cmdDate","<%= FormatDate(Date) %>",1);
	var cal1=new ctlSpiffyCalendarBox("cal1", "frmMain", "txtToDate","cmdToDate","<%= FormatDate(Date) %>",1);
</script>
</head>

<body Class="Reports">
<!--#include file="../libroutines.asp"-->

<%
genReport = Trim(Request.Form("genReport"))
Offering = Request.QueryString("Offering")
OfferingType = Request.QueryString("OfferingType")
Offerings_Date = Request("txtDate")
OfferingsTo_Date = Request("txtToDate")

If (genReport <> "1") Then%>
	<Script Language="JavaScript">
		function validateForm(frm){			
			if (frm.cboOfferings.selectedIndex < 0){
				alert("Please select an IPO Offering");
				frm.cboOfferings.focus();
				return;
			}
			if (document.frmMain.optType(0).checked == false && document.frmMain.optType(1).checked == false  ){
				alert("Please select the type of Download File");
				//frm.cboOfferings.focus();
				return;
			}

			
			
			if (document.frmMain.optType(0).checked == true)
				{ OfferingType = 1; }
			else 
				{ OfferingType = 2; }
				var Offering = document.all.item("cboOfferings").value;
			
			frm.action = 'DownloadBatches.asp?Offering=' + Offering + '&OfferingType=' + OfferingType 
			frm.target = '_self';		
			frm.submit();

		}
	</Script>
	
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	
	<form method="POST" action="DownloadBatches.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">
		<table border="0" cellspacing="2" cellpadding="2" width="80%">
			<!--tr><td width="100%" nowrap colspan="2"><!--b>DOWNLOAD APPLICATIONS</b></td></tr-->
			
			<!--tr><td width="100%" nowrap colspan="2">&nbsp;</td></tr-->
			
			<tr>
				<td width="10%" nowrap>Offering Name</td>
				<td width="90%" nowrap>
					<select name='cboOfferings' id='cboOfferings' onChange="window.location.href='DownloadBatches.asp?Offering='+this.value">
					<option selected ParentSecurity = "" minQty = "" stepQty = "" Ratio = "" OfferType = "" SearchPrice = "" value = ''></option>
					<% 
									
				 SQL = " SELECT Security_DPA_, Security_EIT_, SecurityAddr, SecurityCode, SecurityMktPrice, SecurityName, " & _
				       " OrderSecType_DPA_, Immobilised, Sector_DPA_,  " & _
                       " ImportCode, ClosingDate, Price, Offerings, CanTrade, BankAccount_DPA_, ETSCode, IPOCode, IPOClosed," & _
					   " IPOSMSCharge, StartDate, RequiresExtra,  " & _
                       " BatchSize, DefaultSelection, Ratio, MinimumQuantity, StepQuantity, OfferingType, RequiresHoldings," & _
					   " IsRight, ParentSecurity AS ParentSecurity_DPA_,  " & _
                       " Showing, PAL, NSEName, IPOSMSAccount " & _
                       " FROM Security where offerings = 1"
                         
                 sqlStr = " SELECT * FROM [SecurityListOfferings] " & _
						  " WHERE cast(floor(cast(ClosingDate as float)) as datetime) >= cast(floor(cast(GetDate() as float)) as  datetime)" & _
						  "Order By SecurityName ASC"
					
					'Response.Write ("<tr><td>" & sqlStr & "<tr><td>")
			
					
					
					
					Set Rs = conn.Execute(sqlStr)
					
					If Not (Rs.EOF Or Rs.BOF) Then
						Do Until Rs.EOF
								if trim(rs.Fields("DefaultSelection")) = 1 Then
									%>
									<option selected ParentSecurity = "<%=Rs("ParentSecurity_DPA_")%>" minQty = "<%=Rs("MinimumQty")%>" stepQty = "<%=Rs("StepQty")%>" Ratio = "<%=Rs("Ratio")%>" OfferType = "<%=Rs("OfferType")%>" SearchPrice = "<%=rs.Fields("SecurityMktPrice")%>" value = '<%=rs.Fields("Security_DPA_")%>'><%=rs.Fields("SecurityName")%></option>
									<%
									price = rs("SecurityMktPrice")
									ratio = Rs("Ratio")
									offering = rs.Fields("Security_DPA_")
								else
									%>                   						
									<option ParentSecurity = "<%=Rs("ParentSecurity_DPA_")%>" minQty = "<%=Rs("MinimumQty")%>" stepQty = "<%=Rs("StepQty")%>" Ratio = "<%=Rs("Ratio")%>" OfferType = "<%=Rs("OfferType")%>" SearchPrice = "<%=rs.Fields("SecurityMktPrice")%>" value = '<%=rs.Fields("Security_DPA_")%>'><%=rs.Fields("SecurityName")%></option>
									<%
									price = rs("SecurityMktPrice")
									ratio = Rs("Ratio")
									offering = rs.Fields("Security_DPA_")
								end if 
							Rs.MoveNext
						Loop
					End If
					%>
					</select>
				</td>     
			</tr>
			
			<tr>
				<td width="10%" nowrap>Download File</td>
				<td width="90%" nowrap>
					<input type = 'radio' name ='optType' id = 'optType' value=1 >&nbsp;Daily Applications
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					<input type = 'radio' name ='optType' id = 'optType' value=2 ;>&nbsp;Recalled Applications
				<%
				Select Case Request.QueryString("filter")
					Case "1"
						%>
						<script language='javascript'>document.frmMain.optType(0).checked = true;</script>
						<%
					Case "2"
						%>
						<script language='javascript'>document.frmMain.optType(1).checked = true;</script>
						<%
				End Select	
				%>
				
				</td>
			</tr>
			
			<!--  <tr>
				<td width="10%" nowrap>Batch No</td>
				<td width="90%" nowrap><select name='cboBatch' id='cboBatch'>					
				<%
				if Request.QueryString("filter") <> "" then
				    if Request.QueryString("filter") = 1 then
					    SQL = "SELECT DISTINCT Batch_No, OfferingType FROM Offerings WHERE Offering = " & Offering & " AND Deleted <> 1 AND Forward = 0 AND OfferingType = "& Request.QueryString("filter") &" ORDER BY Batch_No DESC"					
					else
					    SQL = "SELECT DISTINCT Batch_No, OfferingType FROM Offerings WHERE Offering = " & Offering & " AND Deleted <> 1 AND Forward = 1 AND OfferingType = 1 ORDER BY Batch_No DESC"					
					end if
				else
					SQL = "SELECT DISTINCT Batch_No, OfferingType FROM Offerings WHERE Offering = " & Offering & " AND Deleted <> 1 AND Forward = 0 AND OfferingType = 1 ORDER BY Batch_No DESC"					
				end if
				
				Set Rs = conn.Execute(SQL)
				
				If Not (Rs.EOF Or Rs.BOF) Then
					Do Until Rs.EOF
						%>
						<option offeringtype='<%=Rs("OfferingType")%>' value = '<%=Rs("Batch_No")%>'><%=Rs("Batch_No")%></option>
						<%
						Rs.MoveNext
					Loop
				End If
				%>
				</select>
				</td>
			</tr>-->
			<tr>
				<td width="10%" nowrap>Date From:</td>
				<td>
				<SCRIPT LANGUAGE="JavaScript">cal.writeControl();	</SCRIPT>
				</td>
			</tr>
			<tr>
				<td width="10%" nowrap>Date To :</td>
				<td>
				<SCRIPT LANGUAGE="JavaScript">cal1.writeControl();	</SCRIPT>
				</td>
			</tr>
			
			<tr>
				<td width="100%" nowrap colspan=2><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... " id="btnGenerate" name="btnGenerate">&nbsp;&nbsp;</td>
			</tr>
		</table>
	</form>
	<%
	Response.End
End If
%>

<%
Dim n
Dim j
dim Offering_DPAs
Select Case OfferingType
	
	'Daily Applications
	Case "1"  
		Batch_No = Request.QueryString("BatchNo")
		Offering = Request.QueryString("Offering")

			
		SQL = " SELECT * FROM  OfferingList INNER JOIN " & _
              " Client ON OfferingList.Client_DPA_ = Client.Client_DPA_ INNER JOIN " & _
              " Class ON Client.Class_DPA_ = Class.Class_DPA_ " & _
              " WHERE     (OfferingList.Offering = "& Offering &") AND (OfferingList.Batch_No = " & Batch_No & ") AND (OfferingList.OfferingType = 1) " & _
              " ORDER BY OfferingList.BatchSeq"
		
		'Response.Write SQL
		'Response.End

		SQL = "DataOffering  '" & formatdate(Offerings_Date)  & "', '"& formatdate(OfferingsTo_Date) &"'," &   Offering &",1"
		'response.write SQL
		Set Rs = conn.execute(SQL)
		
		If Not (Rs.EOF and Rs.BOF) Then
			strFile = ""
			Delim = ","
			
			yyyy = Year(Date)
			mmm = Month(Date): If Len(mmm) = 1 Then mmm = "0" & mmm
			dd = Day(Date): If Len(dd) = 1 Then dd = "0" & dd
			hh = Hour(Time): If Len(hh) = 1 Then hh = "0" & hh
			mm = Minute(Time): If Len(mm) = 1 Then mm = "0" & mm
			ss = Second(Time): If Len(ss) = 1 Then ss = "0" & ss
			
			'fDate = yyyy & mmm & dd & hh & mm & ss 'Format(Date(),"yyyymmddhhmmss")
			fDate = dd & mmm & yyyy 'Format(Date(),"yyyymmddhhmmss")
			
			'Filename
			nZero = 7 - cdbl(Len(Batch_No))
			Batch_No=rs("Batch_No")
			fBatch_No = Batch_No
			For n = 0 To nZero - 1
				'fBatch_No = "0" & fBatch_No
			Next

			fBatch_No = Replace(fBatch_No,"-","8000")
			
			'Code=Rs("AgentCode")
			Code=""
			'FilePath = Server.MapPath(".") & "\CoopFiles\" & "B" & SESSION("BrokerCode") & fDate & fBatch_No & ".csv"
			FilePath = Server.MapPath(".") & "\IPO FILES\" & Code & fDate & fBatch_No & ".csv"
			
			'Response.Write Batch_No : Response.End
			'Header
		
			
			
			
			'No header required
			'strFile = strFile & "APPH" & Delim & fDate & Delim & fBatch_No & Delim & TotalApplications & Delim & TotalShares & Delim & TotalAmount & Delim & "" & Delim & vbCrLf
			StrFile = "Serial No,Amount,No.Of Shares,FirstNames,Middle Name,Surname /Business Name ,Transaction Date,Payment Mode,Cheque Account Number,Cheque Serial Number ,Internal Reference Number,Batch No,Additional Information" & vbCrLf
			
			TotalAsc = 0
			
			Do Until Rs.EOF
			
					StrLine = ""
					'Response.Write Rs("Offering_DPA_")
					'Response.End
					if Offering_DPAs<>""then
						Offering_DPAs = Offering_DPAs & "," & Rs("Offering_DPA_")
					else
						Offering_DPAs = Rs("Offering_DPA_")
					end if
					'Split Client Names
			        Name = Rs("ClientName")
			        if Name <> "" then
						if instr(1,Name,"&") > 0 then
							ClientNames = split(Name,"&")
							FirstName = ClientNames(0)
							MiddleName = ClientNames(1)
							Surname = ""
						else
				            ClientNames = split(Name," ")
							FirstName = ClientNames(0)
							'MiddleName = ClientNames(1)
							ClientNames(0) = ""
							'ClientNames(1) = ""
							Surname = Join(ClientNames)	
						end if
			        else
			            FirstName = Rs("ClientName")
			            MiddleName = ""			
			            Surname = ""
			        end if
					
					'Body
					
					StrLine = StrLine & trim(Rs("PAL_No")) & Delim & Rs("Amount_Payable") & Delim & Rs("Alloted_Rights") & Delim & FirstName & Delim & MiddleName & Delim & Surname & Delim & Rs("Offerings_Date") 
					
					
					'Payment Type
					Select Case Rs("PaymentType")
						Case 1
							StrLine = StrLine & Delim & "Guarantee" & Delim & "" & Delim & ""
						Case 2
							StrLine = StrLine & Delim & "Cash" & Delim & "" & Delim & ""
						Case 3
							StrLine = StrLine & Delim & "Cheque" & Delim & trim(Rs("PaymentAccountNo")) & Delim & trim(Rs("PaymentRef"))
						
					End Select
					
               if enabled then
					if not isnumeric(left(trim(Rs("ClientCDSNo")),len(trim(Rs("ClientCDSNo")))-2)) then
						%>
							<SCRIPT LANGUAGE="JavaScript">
							<!--
							alert('Invalid CDS number for Client Code: <%=trim(Rs("Client_DPA_"))%>');
							window.history.go(-1);
							//-->
							</SCRIPT>
						<%
						response.end
					end if 
				end if
					
					'StrLine = StrLine & Delim & Rs("Offering_DPA_") & Delim & Rs("AgentCode") & Delim &
					'CLng(left(trim(Rs("ClientCDSNo")),len(trim(Rs("ClientCDSNo")))-2))  & Delim & trim(Rs("Batch_No")) 

					StrLine = StrLine & Delim & Rs("Offering_DPA_") & Delim & trim(Rs("Batch_No")) 

                   ' if enabled then
					'Status
					Select Case Rs("Status")
						Case 1
							StrLine = StrLine & Delim & "New Entry" 
						Case 2
							StrLine = StrLine & Delim & "Recall" 
						Case 3
							StrLine = StrLine & Delim & "Modified Entry" 
						
					End Select
					'end if
					
					StrLine = StrLine & Delim
					
					'Agent Line Check Sum
					LineAsc = 0
					
					For j = 1 To Len(StrLine)
						LineAsc = LineAsc + Asc(Mid(StrLine,j,1))
					Next
					
					'No checksum required
					'strFile = strFile & StrLine & LineAsc & Delim & vbCrLf
					strFile = strFile & StrLine & Delim & vbCrLf
					
					'AgentCode
					'Code = Rs("AgentCode")
					
					TotalAsc = TotalAsc + LineAsc
				Rs.MoveNext
			Loop
			
			
			
			'Footer
			'No Footer Required
			'strFile = strFile & "APPF" & Delim & TotalAsc & Delim & vbCrLf
			
			Set FSO = Server.CreateObject("Scripting.FileSystemObject")

			If FSO.FileExists(FilePath) Then
				FSO.DeleteFile FilePath, True
			End If

			Set fTextStream = FSO.CreateTextFile(FilePath, True, False)
			fTextStream.Write strFile
			fTextStream.Close

			Set fTextStream = Nothing
			Set FSO = Nothing
			%>
			<table border="0" cellspacing="2" cellpadding="2" width="80%">
				<tr><td width="100%"><b>Batch File No. <%=Batch_No%> is ready for Download...</b></td></tr>
				<tr><td width="100%">
				<input type=button style="font-family:arial;font-weight:bold;font-size:12pt;background-color:silver;color:black;border:0;" name="download" Value=" DOWNLOAD " onclick="window.location.href='DownloadBatches.asp?f=<%=Code & fDate & Batch_No%>.csv';"></TD>
				</td></tr>
			</table>
			<%
			sqlstr=  " UPDATE Offerings " & _
					 " SET Downloaded = 1, LastDownLoaded = GETDATE(),Download_Date = Getdate() " & _
					 " WHERE(Offering_DPA_ IN ("& Offering_DPAs &"))"
		'response.write sqlstr:response.end
		conn.Execute(sqlstr)
		
		Else
			%>
			<!--br><br-->
			<table border="0" cellspacing="2" cellpadding="2" width="80%">
				<tr><td width="100%"><b>No Batch File available for Download</b></td></tr>
			</table>
			<%
		End If

		sqlUpdt = " UPDATE offerings set downloaded = 1,LastDownloaded = Getdate(),Download_Date = Getdate() WHERE " & _
		          " Batch_No = " & Batch_No
		
	
	Case "2" 'Agent Forwards File
		Batch_No = Request.QueryString("BatchNo")
		Offering = Request.QueryString("Offering")

		SQL = "SELECT * FROM Forwardlist INNER JOIN " & _
                 " Client ON Forwardlist.Client_DPA_ = Client.Client_DPA_ INNER JOIN " & _
                 " Class ON Client.Class_DPA_ = Class.Class_DPA_ " & _
                 " WHERE (Forwardlist.Offering = "& Offering &") AND (Forwardlist.Batch_No = " & Batch_No & ") AND (Forwardlist.OfferingType = 1) " & _
                 " ORDER BY Forwardlist.BatchSeq"

		SQL = "DataOffering  '" & formatdate(Offerings_Date)  & "', '"& formatdate(OfferingsTo_Date) &"'," &   Offering &",2"
                
        'Response.Write SQL
        'Response.end
     
		Set Rs = conn.execute(SQL)

		If Not (Rs.EOF and Rs.BOF) Then
			strFile = ""
			Delim = ","
			
			yyyy = Year(Date)
			mmm = Month(Date): If Len(mmm) = 1 Then mmm = "0" & mmm
			dd = Day(Date): If Len(dd) = 1 Then dd = "0" & dd
			hh = Hour(Time): If Len(hh) = 1 Then hh = "0" & hh
			mm = Minute(Time): If Len(mm) = 1 Then mm = "0" & mm
			ss = Second(Time): If Len(ss) = 1 Then ss = "0" & ss
			Batch_No=rs("Batch_No")

			'fDate = yyyy & mmm & dd & hh & mm & ss 'Format(Date(),"yyyymmddhhmmss")
			fDate = dd & mmm & yyyy 'Format(Date(),"yyyymmddhhmmss")
			
			'Filename
			nZero = 7 - cdbl(Len(Batch_No))
			
			fBatch_No = Batch_No
			'For n = 0 To nZero - 1
				'fBatch_No = "0" & fBatch_No
			'Next
			
			'AgentCode
			Code = ""
			
			'FilePath = Server.MapPath(".") & "\CoopFiles\" & "B" & SESSION("BrokerCode") & fDate & fBatch_No & ".csv"
			FilePath = Server.MapPath(".") & "\IPO FILES\Recall_" & Code & fDate & fBatch_No & ".csv"
			
			'Header
			
			
			'No Header Required		
			'strFile = strFile & "PYTH" & Delim & fDate & Delim & fBatch_No & Delim & TotalApplications & Delim & TotalShares & Delim & TotalAmount & Delim & "" & Delim & vbCrLf
			StrFile = "Serial No,Amount,No.Of Shares,FirstNames,Middle Name,Surname /Business Name ,Transaction Date,Payment Mode,Cheque Account Number,Cheque Serial Number ,Internal Reference Number,Batch No,Additional Information" & vbCrLf
			
			TotalAsc = 0
			
			Do Until Rs.EOF
					StrLine = ""
					
					'Split Client Names
			        Name = Rs("ClientName")
			       'Response.Write (Name)
				    if Offering_DPAs <> "" then
						Offering_DPAs = Offering_DPAs & "," & Rs("Offering_DPA_")
					else
						Offering_DPAs = Rs("Offering_DPA_")
					end if
			        
			        if Name <> "" then
			            if instr(1,Name,"&") > 0 then
			                ClientNames = split(Name,"&")
			                FirstName = ClientNames(0)
			                MiddleName = ""
			                Surname = ClientNames(1)
			            else			            
			                ClientNames = split(Name," ")
			                FirstName = ClientNames(0)
			                MiddleName = ClientNames(1)
			                ClientNames(0) = ""
			                ClientNames(1) = ""
			                Surname = Join(ClientNames)	
			            end if
			        else
			            FirstName = Rs("ClientName")
			            MiddleName = ""			
			            Surname = ""
			        end if
					
					'Body
					
					StrLine = StrLine & trim(Rs("PAL_No")) & Delim & Rs("Amount_Payable") & Delim & Rs("Alloted_Rights") & Delim & FirstName & Delim & MiddleName & Delim & Surname & Delim & Rs("Offerings_Date") 
					
					'Payment Type
					Select Case Rs("PaymentType")
						Case 1
							StrLine = StrLine & Delim & "Guarantee" & Delim & "" & Delim & ""
						Case 2
							StrLine = StrLine & Delim & "Cash" & Delim & "" & Delim & ""
						Case 3
							StrLine = StrLine & Delim & "Cheque" & Delim & trim(Rs("PaymentAccountNo")) & Delim & trim(Rs("PaymentRef"))
						
					End Select

              '-this section of code has been checked out -no client cdsno...
			  if(Enabled) then
					if not isnumeric(left(trim(Rs("ClientCDSNo")),len(trim(Rs("ClientCDSNo")))-2)) then
						%>
							<SCRIPT LANGUAGE="JavaScript">
							<!--
							alert('Invalid CDS number for Client Code: <%=trim(Rs("Client_DPA_"))%>/n Please rectify that CDS and resume the download');
							window.history.go(-1);
							//-->
							</SCRIPT>
						<%
						response.end
					end if 
			  end if
			  
					'StrLine = StrLine & Delim & Rs("Offering_DPA_") & Delim & Rs("AgentCode") & Delim & _
					'CLng(left(trim(Rs("ClientCDSNo")),len(trim(Rs("ClientCDSNo")))-2)) & Delim & trim(Rs("Batch_No")) 

					StrLine = StrLine & Delim & Rs("Offering_DPA_") & Delim & trim(Rs("Batch_No")) 
					
				   'Status
					Select Case Rs("Status")
						Case 1
							StrLine = StrLine & Delim & "New Entry" 
						Case 2
							StrLine = StrLine & Delim & "Recall" 
						Case 3
							StrLine = StrLine & Delim & "Modified Entry" 
						
					End Select
					
				    StrLine = StrLine & Delim
					
					'Agent Line Check Sum
					LineAsc = 0
					
					For j = 1 To Len(StrLine)
						LineAsc = LineAsc + Asc(Mid(StrLine,j,1))
					Next
					
					'No checksum required
					'strFile = strFile & StrLine & LineAsc & Delim & vbCrLf
					strFile = strFile & StrLine & Delim & vbCrLf
					
					
					
					TotalAsc = TotalAsc + LineAsc
				Rs.MoveNext
			Loop
			
			'Footer
			'No Footer Required
			'strFile = strFile & "PYTF" & Delim & TotalAsc & Delim & vbCrLf
			
			Set FSO = Server.CreateObject("Scripting.FileSystemObject")

			If FSO.FileExists(FilePath) Then
				FSO.DeleteFile FilePath, True
			End If

			Set fTextStream = FSO.CreateTextFile(FilePath, True, False)
			fTextStream.Write strFile
			fTextStream.Close

			Set fTextStream = Nothing
			Set FSO = Nothing
			%>
			<table border="0" cellspacing="2" cellpadding="2" width="80%">
				<tr><td width="100%"><b>Batch File No. <%=Batch_No%> is ready for Download...</b></td></tr>
				<tr><td width="100%">
				<input type=button style="font-family:arial;font-weight:bold;font-size:12pt;background-color:silver;color:black;border:0;" name="download" Value=" DOWNLOAD " onclick="window.location.href='DownloadBatches.asp?f=Recall_<%=Code & fDate & Batch_No%>.csv';"></TD>
				</td></tr>
			</table>
			<%
			sqlstr= " UPDATE Offerings " & _
					" SET  Downloaded = 1, LastDownLoaded = GETDATE(),Download_Date = Getdate() " & _
					" WHERE (Offering_DPA_ IN ("& Offering_DPAs &"))"
			'response.write sqlstr:response.end
			conn.Execute(sqlstr)
		Else
			%>
			<!--br><br-->
			<table border="0" cellspacing="2" cellpadding="2" width="80%">
				<tr><td width="100%"><b>No Batch File available for Download</b></td></tr>
			</table>
			<%
		End If
	End Select
	%>
</body>

