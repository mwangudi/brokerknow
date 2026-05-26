<!--#include file="../libroutines.asp"-->
<%
	Set conn = GetActiveConnection("KBroker")

	dim action 
	Dim NextNo
	Dim BatchedItems
	Dim BatchNull
	sqlStr = "SELECT * FROM OfferingsList WHERE Offering_DPA_ = " & Request("ID")
	set rs = conn.execute(sqlStr)

	if not (rs.eof or rs.bof) then		BatchID = rs("Batch_No")		SecurityDPA = rs("Offering")
	else		BatchID = 0
		SecurityDPA = 0	end if        'Response.Write Request("ID")    'Response.End     
	action = Request.QueryString("action")

	if(action="") then
		action = Request.Form("action")
	end if
	
	SearchSerial = Request.QueryString("SearchSerial")

	BatchedItems = Request.Form("CommitParams")	
	buttonAction = Trim(Ucase(Request.Form("cmdAdd")))
	
	Set MaxRs = Server.CreateObject("ADODB.Recordset")
	MaxRs.CursorLocation = adUseClient 

	if (action = "Save") then
		if instr(1,buttonAction,"SEARCH") > 0 then
			SearchSerial = Request.Form("txtSearch")
			WriteDialogRelocateScript "EditBatchedApplication.asp?SearchSerial=" & SearchSerial & "&ID=" & BatchID
		else
			if(BatchedItems="") then
				%>
				<script language="javascript">
					alert('No applications selected for batching');
				</script>
				<%
				Response.end
			end if
		
			itemids=Split(BatchedItems,",")
		
			j=0

			conn.BeginTrans
										'Response.Write "update offerings set Downloaded=0,LastDownloaded=0,Batch_No=Null where Batch_No=" & BatchID & "" & "<BR>"					
					conn.Execute("update offerings set Downloaded=0,LastDownloaded=0,Batch_No=Null where Batch_No=" & BatchID)		
							
					For i=0 to UBound(itemids)			
						      
						Ids =Split(itemids(i),"<->")
							
						OrderDetail_DPA_=Ids(1)
							
						j=i+1
												'Response.Write "Update Offerings Set Batch_No=" & BatchID & ",BatchSeq=" & j & " Where Offering_DPA_=" & Ids(0)						
						conn.Execute("Update Offerings Set Batch_No=" & BatchID & ",BatchSeq=" & j & " Where Offering_DPA_=" & Ids(0))
					Next
	
			conn.CommitTrans
	
	'Response.end
		conn.Close
	Set conn = Nothing
		'WriteDialogRelocateScript "EditBatchedApplication.asp?ID=" & ID	WriteDialogCloseScript	
	Response.End	
	end if	
end if
	
if (action="editipo") then
	
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
<title>Edit Batched Application</title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 

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
<form name = 'frmBatched' method = 'post' action = 'EditBatchedApplication.asp' id='frmMain' >
<br><table>
<tr><td colspan="2" align="right">		
			Serial No:&nbsp;&nbsp;<input type = 'text' name ='txtSearch' id = 'txtSearch' size="20" value = ''>&nbsp;&nbsp;<input type = 'Submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Search ">		
	</td></tr>
</table>
<br><br>
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

		sqlStr = "SELECT * FROM OfferingsList WHERE Batch_No = " & BatchID & " AND Offering = "& SecurityDPA &" ORDER BY BatchSeq ASC"
				
		Set Rs = Server.CreateObject("ADODB.Recordset")
		Rs.CursorLocation = adUseClient 
		Rs.Open  SqlStr, Conn.ConnectionString, adOpenKeySet, adLockOptimistic
        
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

		if(SearchSerial<>"") then
		  sqlStr = "SELECT * FROM OfferingsList where Batch_no is null AND (PAL_No like '%" & SearchSerial & "%') AND ClosingDate > cast(Floor(cast(GetDate() as Float)) as datetime) order by pal_no asc"
		else
		  sqlStr = "SELECT * FROM OfferingsList where Batch_no is null AND ClosingDate >= cast(Floor(cast(GetDate() as Float)) as datetime) order by pal_no asc"
		end if	
 
		MaxRs.Open  SqlStr, Conn.ConnectionString, adOpenKeySet, adLockOptimistic
        
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
				<td  valign="top"><%=MaxRs("Pal_no")%></td>
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
	<td align=right>
		<br><br><input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save ">
	&nbsp;&nbsp;
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();">
	</td>
 </tr>
 </table>
 <input type = 'hidden' name ='CommitParams' id = 'CommitParams' value="<%=CheckedArray%>">
 <input type = 'hidden' name ='action' id = 'action' value="Save">&nbsp;
 <input type = 'hidden' name ='ID' id = 'ID' value="<%=Request("ID")%>">&nbsp;
</form>
</body>

</html>
