<!--#include file="../libroutines.asp"-->
<%
'======================= Begin_Alter_Across_Entities =================================
		
		const UDLName = "KBroker"
		const DataSource = "OfferingsList"
		const DataEntity = "Certificate"
		const DataEntityPlural = "Certificates"
		const ActionFolder = "Operations"
'======================= End_Alter_Across_Entities =================================
	Set conn = GetActiveConnection(UDLName)

	dim action 
	Dim NextNo
	Dim BatchedItems
	Dim BatchNull

	BatchID = Request("ID")
    
	action = Request.QueryString("action")

	if(action="") then
	action = Request.Form("action")
	end if
	
	SearchSerial = Request.QueryString("SearchSerial")

	BatchedItems = Request.Form("CommitParams")	
	buttonAction = Trim(Ucase(Request.Form("cmdAdd")))
	
	Set MaxRs = Server.CreateObject("ADODB.Recordset")
		MaxRs.CursorLocation = adUseClient 

	if(action ="Save") then
		if instr(1,buttonAction,"SEARCH") > 0 then
		SearchSerial = Request.Form("txtSearch")
		WriteDialogRelocateScript "EditBatchedApplication.asp?SearchSerial=" & SearchSerial & "&ID=" & BatchID
		else

		if(BatchedItems="") then
		%>
		<script language="javascript">
		alert('You must choose applications to Batch');
		</script>
		<%
		Response.end
		end if
		
		itemids=Split(BatchedItems,",")
		
		BatchNull = "NULL"
		
		j=0

		conn.BeginTrans

		conn.Execute("update offerings set Downloaded=0,LastDownloaded=0,Batch_No=" & BatchNull & " where Batch_No=" & BatchID)		
			
		For i=0 to UBound(itemids)			
              
			Ids =Split(itemids(i),"<->")
			
			OrderDetail_DPA_=Ids(1)
			ID=Ids(0)
			
			j=i+1

			conn.Execute("Update Offerings Set Batch_No=" & BatchID & ",BatchSeq=" & j & " Where Offering_DPA_=" & ID)
		Next
	
	conn.CommitTrans
	conn.Close
	Set conn = Nothing
	'WritefraEnabledDialogCloseScript
	WriteDialogRelocateScript "EditBatchedApplication.asp?ID=" & BatchID
	Response.End	
	end if	
end if
	
if(action="editipo") then
	
	ItemID=Request.Form("ItemID")
	ID=Request.Form("ID")
	
	Sname=Request.Form("txtname")
	cbotype = Request.Form("cbotype")		
		
		sqlStr="UPDATE Offerings SET shNAME2 ='" & Sname & "',AppType='" & cbotype & "' Where Offering_DPA_ = " & ItemID

		conn.BeginTrans
			conn.Execute(sqlStr)			
		conn.CommitTrans
	
	
	conn.close
	Set conn=nothing	
%>
	    <SCRIPT LANGUAGE="JAVASCRIPT">
		window.parent.parent.frames['maininfo'].location.reload();
	    </SCRIPT>
	 <%

	WriteDialogRelocateScript "EditBatchedApplication.asp?ID=" & ID
	Response.end

	end if
	%>	
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add <%=DataEntity%></title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
<SCRIPT language="JavaScript">
	var calDDate=new ctlSpiffyCalendarBox("calDDate", "frm<%=DataSource%>", "txtRDate","cmdDDate","<%=FormatDate(rDate)%>",1);
</SCRIPT>

<script language="javascript">
function  CommitCDSTrade(theChk,theItem,OrderDetail_DPA_,BatchSize)
			{
			var thestring;
			var laststring;
			var Selectedids = document.getElementById("CommitParams").value;
			var selectedLength,Batchcount;
			
		if (Selectedids != ''){
              var ctrlSize = Selectedids.split(",");
              
              if (BatchSize > 0){Batchcount = BatchSize-1;}else{Batchcount = 0;}
			  
		    selectedLength= ctrlSize.length ;
                 
               if (selectedLength > Batchcount && theChk.checked ){
			    alert('The maximum batch size (' + BatchSize + ') has been exceeded. The selection will be reversed.');
				if (theChk.checked) {theChk.checked = false;}
			    return;
			   }
			} else {
			 if (BatchSize <= 0){
			    BatchSize = 0;
			    alert('Please set the maximum batch size to proceed. The selection will be reversed.');
				if (theChk.checked) {theChk.checked = false;}
			  }
			
			} 
                
				if (theChk.checked)
				{		
					if (document.frmMain.elements("CommitParams").value =="")
						{
						  document.frmMain.elements("CommitParams").value = theItem + '<->' + OrderDetail_DPA_;
						}
					else
						{
						  document.frmMain.elements("CommitParams").value = document.frmMain.elements("CommitParams").value + ',' + theItem + '<->' + OrderDetail_DPA_;
						}
				}
				else
				{
				       laststring= theItem + '<->' + OrderDetail_DPA_;

					if(document.frmMain.elements("CommitParams").value==laststring)
					{
					document.frmMain.elements("CommitParams").value="";
					}
					else
					{
                                var arraycontent = Selectedids.split(",");

                                 if (arraycontent[0]==laststring){
                                     thestring =  laststring + ','; // first element in array so remove trailing comma
                                  } else {
                                     thestring =  ',' + laststring;//remove leading comma
                                  }
				
					document.frmMain.elements("CommitParams").value = replaceSubstring(Selectedids,thestring,'') ;
                              }
				}
				
			}			
			
			function replaceSubstring(inputString, fromString, toString) {
				   // Goes through the inputString and replaces every occurrence of fromString with toString
				   var temp = inputString;
				   if (fromString == "") {
					  return inputString;
				   }
				   if (toString.indexOf(fromString) == -1) { // If the string being replaced is not a part of the replacement string (normal situation)
					  while (temp.indexOf(fromString) != -1) {
						 var toTheLeft = temp.substring(0, temp.indexOf(fromString));
						 var toTheRight = temp.substring(temp.indexOf(fromString)+fromString.length, temp.length);
						 temp = toTheLeft + toString + toTheRight;
					  }
				   } else { // String being replaced is part of replacement string (like "+" being replaced with "++") - prevent an infinite loop
					  var midStrings = new Array("~", "`", "_", "^", "#");
					  var midStringLen = 1;
					  var midString = "";
					  // Find a string that doesn't exist in the inputString to be used
					  // as an "inbetween" string
					  while (midString == "") {
						 for (var i=0; i < midStrings.length; i++) {
							var tempMidString = "";
							for (var j=0; j < midStringLen; j++) { tempMidString += midStrings[i]; }
							if (fromString.indexOf(tempMidString) == -1) {
							   midString = tempMidString;
							   i = midStrings.length + 1;
							}
						 }
					  } // Keep on going until we build an "inbetween" string that doesn't exist
					  // Now go through and do two replaces - first, replace the "fromString" with the "inbetween" string
					  while (temp.indexOf(fromString) != -1) {
						 var toTheLeft = temp.substring(0, temp.indexOf(fromString));
						 var toTheRight = temp.substring(temp.indexOf(fromString)+fromString.length, temp.length);
						 temp = toTheLeft + midString + toTheRight;
					  }
					  // Next, replace the "inbetween" string with the "toString"
					  while (temp.indexOf(midString) != -1) {
						 var toTheLeft = temp.substring(0, temp.indexOf(midString));
						 var toTheRight = temp.substring(temp.indexOf(midString)+midString.length, temp.length);
						 temp = toTheLeft + toString + toTheRight;
					  }
				   } // Ends the check to see if the string being replaced is part of the replacement string or not
				   return temp; // Send the updated string back to the user
				} // Ends the "replaceSubstring" function

</script>

</head>

<body Class="Dialog">
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<%
if(action="edit") then
	ItemID=Request.QueryString("ItemID")
	ID = Request.QueryString("ID")

	Set conn = GetActiveConnection(UDLName)
   
	sqlStr = "Select * from offerings Where Offering_DPA_=" & ItemID
	
	Set Rs = Server.CreateObject("ADODB.Recordset")
		Rs.CursorLocation = adUseClient 
		Rs.Open  SQLServerFormat(HandleQuote(sqlStr)), Conn.ConnectionString, adOpenKeySet, adLockOptimistic
       

	%>
	<form name = 'frmEdit' Method = 'post' action = 'EditBatchedApplication.asp?action=editipo' id='frmEdit'>
		<table width="50%" align="center">
			<tr>
				<td colspan="2" align="center"><font size=2 face="Tahoma" color=blue > <b> Edit Individual Application </b> </font></td>
			</tr>
			<tr>		
				<td>Second Name</td>
				<td>Application Type</td>
				<tr>  
			</tr>
			<tr>				
				<td><input type = 'text' name ='txtname' id = 'txtname' size="20" value = '<%=Rs("shNAME2")%>'></td>
				<td width="83%">
					<select name="cbotype" id="cbotype" size=1>
						<%
						Select Case(Rs("AppType")) 
							case "N"
							%>
							<option value="C">N</option>
							<option value="N">C</option>							
							<option value="J">J</option>							
							<%
							case "J"
							%>
							<option value="J">J</option>
							<option value="C">N</option>							
							<option value="N">C</option>
							<%
							case "C"
							%>
							<option value="N">C</option>
							<option value="C">N</option>
							<option value="J">J</option>							
							<%
						end select
						%>						
					</select>
			</td>
			</tr>
			<tr>
				<td></td>
				<td align="right"><input type = 'Submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save ">&nbsp;
				<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.location.replace('EditBatchedApplication.asp?ID=<%=ID%>')"></td>
			</tr>
			<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
			<input type = 'hidden' name ='ItemID' id = 'ItemID' value="<%=ItemID%>">
		</table>
	</form>
	<%

	conn.Close
	Set conn = Nothing

	Set Rs = nothing
else
%>

<form name = 'frm<%=DataSource%>' method = 'post' action = 'EditBatchedApplication.asp' id='frmMain' >
<table>
<tr><td colspan="2" align="right">		
			Serial No:&nbsp;&nbsp;<input type = 'text' name ='txtSearch' id = 'txtSearch' size="20" value = ''>&nbsp;&nbsp;<input type = 'Submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Search ">		
	</td></tr>
</table>

<table border="0" cellspacing="1" cellpadding="1"  width="100%">
  <tr>
	<td  valign="top">&nbsp;</td>
	<td  valign="top"><b><font color="#000080">Serial No</font></b></td>
	<td  valign="top"><b><font color="#000080">Offering Name</font></b></td>
	<td  valign="top"><b><font color="#000080">Code</font></b></td>
	<td  valign="top"><b><font color="#000080">Client Name</font></b></td>	
	<td  valign="top" align="right"><b><font color="#000080">Quantity Applied</font></b></td>
	<td  valign="top" align="right"><b><font color="#000080">Price</font></b></td>
	<td  valign="top" align="right"><b><font color="#000080">Payable</font></b></td>	
  </tr>
<%

sqlStr = "SELECT * FROM [" & DataSource & "] where Batch_no=" & BatchID & " order by BatchSeq asc"
		 
		Set Rs = Server.CreateObject("ADODB.Recordset")
		Rs.CursorLocation = adUseClient 
		Rs.Open  SQLServerFormat(HandleQuote(sqlStr)), Conn.ConnectionString, adOpenKeySet, adLockOptimistic
        
		dim bgcolor
		 bgcolor="#FFFFFF"
		
		CheckedArray = ""

        If rs.EOF Or rs.BOF Then
		%>
		<script language="javascript">
		alert('No applications to batch');
		</script>
		<%
		Response.end
		else
			do until Rs.eof
				
				'Populate Checked Orders Array
			if CheckedArray = "" then
			   CheckedArray = trim(rs.Fields("Offering_DPA_")) & "<->" & trim(rs.Fields("Offering_DPA_"))
			 Else 
               CheckedArray  = CheckedArray & "," & trim(rs.Fields("Offering_DPA_")) & "<->" & trim(rs.Fields("Offering_DPA_"))
			 end if
		
				commitTrade = "<input type='checkbox' checked class='BorderLess' value='" & rs.Fields("Offering_DPA_") & "' name='chkCommit' id='chkCommit' onClick = 'javascript: CommitCDSTrade(this, " & rs.Fields("Offering_DPA_") & ", " & rs.Fields("Offering_DPA_") & "," & rs.Fields("BatchSize") & ");'>"
                
				if bgcolor="#F0EFDB" then  bgcolor="#FFFFFF" else bgcolor ="#F0EFDB"
				%>
				<tr bgcolor="<%=bgcolor%>" onMouseover="JavaScript: this.bgColor='#99CCFF'" onMouseout="JavaScript: this.bgColor='<%=bgcolor%>'">
					<td  valign="top"><%=commitTrade%></td>
					<td  valign="top"><a href="EditBatchedApplication.asp?ID=<%=BatchID%>&action=edit&ItemID=<%=Rs("Offering_DPA_")%>"><%=Rs("Pal_no")%></a></td>
					<td  valign="top"><%=Rs("SecurityCode")%></td>
					<td  valign="top"><%=Rs("Client_DPA_")%></td>
					<td  valign="top" nowrap><%=mid(Rs("ClientName"),1,30)%></td>				
					<td  valign="top" align="right"><%=FormatNumEx(Rs("Alloted_Rights"),0)%></td>
					<td  valign="top" align="right"><%=FormatNum(Rs("Offering_Price"))%></td>
					<td  valign="top" align="right"><%=FormatNum(Rs("amount_Payable"))%></td>
			    </tr>
			<%
			Rs.movenext
			loop
		end if

		if(SearchSerial<>"") then
		  sqlStr = "SELECT * FROM [" & DataSource & "] where Batch_no is null AND (PAL_No like '%" & SearchSerial & "%') AND ClosingDate > cast(Floor(cast(GetDate() as Float)) as datetime) order by pal_no asc"
		else
		  sqlStr = "SELECT * FROM [" & DataSource & "] where Batch_no is null AND ClosingDate >= cast(Floor(cast(GetDate() as Float)) as datetime) order by pal_no asc"
		end if	
 
		MaxRs.Open  SQLServerFormat(HandleQuote(sqlStr)), Conn.ConnectionString, adOpenKeySet, adLockOptimistic
        
		'dim bgcolor
		 bgcolor="#FFFFFF"
		
		'CheckedArray = ""

        If MaxRs.EOF Or MaxRs.BOF Then		
		else
			do until MaxRs.eof
				
				commitTrade = "<input type='checkbox' class='BorderLess' name='chkCommit' onClick = 'javascript: CommitCDSTrade(this, " & MaxRs.Fields("Offering_DPA_") & ", " & MaxRs.Fields("Offering_DPA_") & "," & MaxRs.Fields("BatchSize") & ");'>"

				if bgcolor="#F0EFDB" then  bgcolor="#FFFFFF" else bgcolor ="#F0EFDB"
				%>
				<tr bgcolor="<%=bgcolor%>" onMouseover="JavaScript: this.bgColor='#99CCFF'" onMouseout="JavaScript: this.bgColor='<%=bgcolor%>'">
				<td  valign="top"><%=commitTrade%></td>
				<td  valign="top"><a href="EditBatchedApplication.asp?ID=<%=BatchID%>&action=edit&ItemID=<%=MaxRs("Offering_DPA_")%>"><%=MaxRs("Pal_no")%></a></td>
				<td  valign="top"><%=MaxRs("SecurityCode")%></td>
				<td  valign="top"><%=MaxRs("Client_DPA_")%></td>
				<td  valign="top" nowrap><%=mid(MaxRs("ClientName"),1,30)%></td>				
				<td  valign="top" align="right"><%=FormatNumEx(MaxRs("Alloted_Rights"),0)%></td>
				<td  valign="top" align="right"><%=FormatNum(MaxRs("Offering_Price"))%></td>
				<td  valign="top" align="right"><%=FormatNum(MaxRs("amount_Payable"))%></td>
			    </tr>
			<%
			MaxRs.movenext
			loop
		end if

%>
 <tr>
	<td colspan="7">&nbsp;</td>
	<td>
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save ">
	</td>
	<td>
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();">
	</td>
 </tr>
 </table>
 <input type = 'hidden' name ='CommitParams' id = 'CommitParams' value="<%=CheckedArray%>">
 <input type = 'hidden' name ='action' id = 'action' value="Save">&nbsp;
 <input type = 'hidden' name ='ID' id = 'ID' value="<%=BatchID%>">&nbsp;
</form>
<%
end if
%>
</body>

</html>
