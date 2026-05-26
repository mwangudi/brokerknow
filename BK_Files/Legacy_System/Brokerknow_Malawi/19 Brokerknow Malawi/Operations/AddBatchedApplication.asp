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


	action = Request.Form("action")
	SearchSerial = Request.QueryString("SearchSerial")
	BatchedItems = Request.Form("CommitParams")	
	buttonAction = Trim(Ucase(Request.Form("cmdAdd")))
	
	Set MaxRs = Server.CreateObject("ADODB.Recordset")
		MaxRs.CursorLocation = adUseClient 

	if(action ="Save") then
		if instr(1,buttonAction,"SEARCH") > 0 then
		SearchSerial = Request.Form("txtSearch")
		WriteDialogRelocateScript "AddBatchedApplication.asp?SearchSerial=" & SearchSerial
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

		sqlmax="Select Max(isnull(Batch_No,0))+1 as NextBatch from offerings where deleted=0"
		
		j=0
		Set MaxRs = conn.Execute(sqlmax)
		NextNo = MaxRs("NextBatch")
		
		conn.BeginTrans	
		For i=0 to UBound(itemids)			
              
			Ids =Split(itemids(i),"<->")
			
			OrderDetail_DPA_=Ids(1)
			ID=Ids(0)
			j=i+1

			conn.Execute("Update Offerings Set Batch_No=" & NextNo & " ,BatchSeq = " & j & " Where Offering_DPA_=" & ID)
		Next
	
	conn.CommitTrans
	conn.Close
	Set conn = Nothing
	WritefraEnabledDialogCloseScript
	Response.End	
	end if	
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
			var selectedLenght,Batchcount;
			
			if (Selectedids != ''){
              var ctrlSize = Selectedids.split(",");
              
              if (BatchSize > 0){Batchcount = BatchSize-1;}else{Batchcount = 0;}
			  
			  selectedLenght = ctrlSize.length ;
               if (selectedLenght > Batchcount && theChk.checked){
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
					
					laststring=theItem + '<->' + OrderDetail_DPA_;
                     
                      var ctrlcontent = Selectedids.split(",");
                      
                      if (ctrlcontent[0]==laststring){
                       thestring= theItem + '<->' + OrderDetail_DPA_ +  ',';
                      }else {
                          thestring=',' + theItem + '<->' + OrderDetail_DPA_;
                      }
                      
						if(document.frmMain.elements("CommitParams").value==laststring)
						{
							document.frmMain.elements("CommitParams").value="";
						}
						else
						{
							document.frmMain.elements("CommitParams").value = replaceSubstring(document.frmMain.elements("CommitParams").value,thestring,'') ;
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

<form name = 'frm<%=DataSource%>' method = 'post' action = 'AddBatchedApplication.asp' id='frmMain' >
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



if(SearchSerial<>"") then
  sqlStr = "SELECT * FROM [" & DataSource & "] where Batch_no is null AND (PAL_No like '%" & SearchSerial & "%') AND ClosingDate > cast(Floor(cast(GetDate() as Float)) as datetime) order by pal_no asc"
else
  sqlStr = "SELECT * FROM [" & DataSource & "] where Batch_no is null AND ClosingDate >= cast(Floor(cast(GetDate() as Float)) as datetime) order by pal_no asc"
end if	


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
				
				commitTrade = "<input type='checkbox' class='BorderLess' value='" & rs.Fields("Offering_DPA_") & "' name='chkCommit' id='chkCommit' onClick = 'javascript: CommitCDSTrade(this, " & rs.Fields("Offering_DPA_") & ", " & rs.Fields("Offering_DPA_") & "," & rs.Fields("BatchSize") & ");'>"
                
				if bgcolor="#F0EFDB" then  bgcolor="#FFFFFF" else bgcolor ="#F0EFDB"
				%>
				<tr bgcolor="<%=bgcolor%>" onMouseover="JavaScript: this.bgColor='#99CCFF'" onMouseout="JavaScript: this.bgColor='<%=bgcolor%>'">
					<td  valign="top"><%=commitTrade%></td>
					<td  valign="top"><%=Rs("Pal_no")%></td>
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
</form>

</body>

</html>
