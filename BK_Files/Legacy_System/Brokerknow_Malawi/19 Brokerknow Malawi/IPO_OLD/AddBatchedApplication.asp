<!--#include file="../libroutines.asp"-->
<%
	Set conn = GetActiveConnection("KBroker")

	dim action 
	Dim NextNo
	Dim BatchedItems

	action = Request.Form("action")
	SearchSerial = Request.QueryString("SearchSerial")
	BatchedItems = Request.Form("CommitParams")	
	buttonAction = Trim(Ucase(Request.Form("cmdAdd")))
	
	UserID = SESSION("UserID")
	
	Offering = Request("Offering")	
	
	If Offering = "" Then Offering = Request.Form("cboOfferings")
	
	'Response.Write Offering 'Request.Form
	'Response.End 
	
	If Offering <> "" Then
	Set MaxRs = Server.CreateObject("ADODB.Recordset")
	MaxRs.CursorLocation = adUseClient 

	if (action = "Save") then
		if instr(1,buttonAction,"SEARCH") > 0 then
			SearchSerial = Request.Form("txtSearch")
			WriteDialogRelocateScript "AddBatchedApplication.asp?SearchSerial=" & SearchSerial
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
				For i=0 to UBound(itemids)			
				    ''GET BATCH NO
					'------------------------------------------------------------------------------------------------------------------
					SqlStr = "SELECT MAX(ISNULL(Offerings.Batch_No, 0)) AS BatchNo FROM Offerings INNER JOIN Security ON Offerings.Offering = Security.Security_DPA_" & _
						" WHERE (Offerings.ChangedBy = "& UserID &") AND (Offerings.Offering = "& offering &") AND Deleted = 0"
					Set Rst = conn.Execute(SqlStr)
	
					If Not (Rst.EOF Or Rst.BOF) Then
						theBatchNo = 0
						If IsNull(Rst("BatchNo")) Then
							theBatchNo = 0
						Else
							theBatchNo = Rst("BatchNo")
						End If
							
						If theBatchNo = 0 Then
							SqlStr = "SELECT MAX(ISNULL(Offerings.Batch_No, 0)) AS BatchNo FROM Offerings INNER JOIN Security ON Offerings.Offering = Security.Security_DPA_" & _
								" WHERE (Offerings.ChangedBy = "& UserID &") AND (Offerings.Offering = "& offering &") AND Deleted = 0"
							Set Rst2 = conn.Execute(SqlStr)
	
							If Not (Rst2.EOF Or Rst2.BOF) Then
								If IsNull(Rst2("BatchNo")) Then
									SqlStr = "SELECT MAX(ISNULL(Offerings.Batch_No, 0)) AS BatchNo FROM Offerings INNER JOIN Security ON Offerings.Offering = Security.Security_DPA_" & _
									" WHERE (Offerings.Offering = "& offering &") AND Deleted = 0"
									Set Rst3 = conn.Execute(SqlStr)
										
									If Not (Rst3.EOF Or Rst3.BOF) Then
										BatchNo = Rst3("BatchNo") + 1
									Else
										BatchNo = 1
									End If					
								Else
									BatchNo = Rst2("BatchNo") + 1
								End If
							Else
								SqlStr = "SELECT MAX(ISNULL(Offerings.Batch_No, 0)) AS BatchNo FROM Offerings INNER JOIN Security ON Offerings.Offering = Security.Security_DPA_" & _
								" WHERE (Offerings.Offering = "& offering &") AND Deleted = 0"
								Set Rst3 = conn.Execute(SqlStr)
											
								If Not (Rst3.EOF Or Rst3.BOF) Then
									BatchNo = Rst3("BatchNo") + 1
								Else
									BatchNo = 1
								End If	 
							End If
						Else
							SqlStr = "SELECT COUNT(Offerings.Offering_DPA_) AS OfferingCount, Security.BatchSize" & _
							" FROM Offerings INNER JOIN Security ON Offerings.Offering = Security.Security_DPA_" & _
							" WHERE (Offerings.ChangedBy = "& UserID &") AND (Offerings.Offering = "& offering &") AND Offerings.Batch_No = "& theBatchNo &" AND Deleted = 0" & _
							" GROUP BY Security.BatchSize"
							Set Rst2 = conn.Execute(SqlStr)
								
							If Not (Rst2.EOF Or Rst2.BOF) Then
								If Rst2("OfferingCount") = Rst2("BatchSize") Then
									BatchNo = Rst("BatchNo") + 1
								Else
									BatchNo = Rst("BatchNo")
								End If
							Else
								BatchNo = Rst("BatchNo")
							End If
						End If
					Else
						SqlStr = "SELECT MAX(ISNULL(Offerings.Batch_No, 0)) AS BatchNo FROM Offerings INNER JOIN Security ON Offerings.Offering = Security.Security_DPA_" & _
							" WHERE (Offerings.ChangedBy = "& UserID &") AND (Offerings.Offering = "& offering &") AND Deleted = 0"
						Set Rst2 = conn.Execute(SqlStr)
	
						If Not (Rst2.EOF Or Rst2.BOF) Then
							BatchNo = Rst2("BatchNo") + 1
						Else
							SqlStr = "SELECT MAX(ISNULL(Offerings.Batch_No, 0)) AS BatchNo FROM Offerings INNER JOIN Security ON Offerings.Offering = Security.Security_DPA_" & _
							" WHERE (Offerings.Offering = "& offering &") AND Deleted = 0"
							Set Rst3 = conn.Execute(SqlStr)
											
							If Not (Rst3.EOF Or Rst3.BOF) Then
								BatchNo = Rst3("BatchNo") + 1
							Else
								BatchNo = 1
							End If
						End If
					End If
	
					set Rst = Nothing : set Rst2 = Nothing : set Rst3 = Nothing
					'------------------------------------------------------------------------------------------------------------------
			
					'sqlmax="Select Max(isnull(Batch_No,0))+1 as NextBatch from offerings where deleted=0"
					'Set MaxRs = conn.Execute(sqlmax)
					'NextNo = MaxRs("NextBatch")
			
					NextNo = BatchNo
					  
					Ids =Split(itemids(i),"<->")
					
					OrderDetail_DPA_=Ids(1)
					ID=Ids(0)
					j=i+1
					
					SqlStr = "UPDATE Offerings SET Batch_No = " & NextNo & ", BatchSeq = " & j & " WHERE Offering_DPA_=" & ID
					
					'Response.Write sqlstr
					'Response.End 
					
					conn.Execute(SqlStr)
				Next
			conn.CommitTrans
			
			conn.Close
			Set conn = Nothing
	
			WritefraEnabledDialogCloseScript
			Response.End	
		end if	
	end if
	end if
	%>
<html>

<head>
<title>Add Batched Application</title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 

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

<form name = 'frmBatch' method = 'post' action = 'AddBatchedApplication.asp' id='frmMain' >
	<br>
	
	<table align=center border="0" cellspacing="1" cellpadding="1" width="98%">
		<tr>
			<td width="20%" nowrap>Offering Name</td>
			<td width="80%" nowrap>
				<select name = 'cboOfferings' id = 'cboOfferings' size="1" onChange="window.location.href='AddBatchedApplication.asp?Offering='+this.value">
				<% 
				sqlStr = "SELECT * FROM [SecurityListOfferings] " & _
				" WHERE cast(floor(cast(ClosingDate as float)) as datetime) >= cast(floor(cast(GetDate() as float)) as datetime)" & _
				" Order By SecurityName ASC"
				Set rs = conn.Execute(SqlStr)
				
				If Not (rs.EOF Or rs.BOF) Then
					Do Until rs.EOF
							if Offering = "" then
								if trim(rs.Fields("DefaultSelection")) = 1 Then
									%>
									<option selected ParentSecurity = "<%=Rs("ParentSecurity_DPA_")%>" Ratio = "<%=Rs("Ratio")%>" OfferType = "<%=Rs("OfferType_DPA_")%>" SearchPrice = "<%=rs.Fields("SecurityMktPrice")%>" value = '<%=rs.Fields("Security_DPA_")%>'><%=rs.Fields("SecurityName")%></option>
									<%
									price = rs("SecurityMktPrice")
									ratio = Rs("Ratio")
								else
									%>                   						
									<option ParentSecurity = "<%=Rs("ParentSecurity_DPA_")%>" Ratio = "<%=Rs("Ratio")%>" OfferType = "<%=Rs("OfferType_DPA_")%>" SearchPrice = "<%=rs.Fields("SecurityMktPrice")%>" value = '<%=rs.Fields("Security_DPA_")%>'><%=rs.Fields("SecurityName")%></option>
									<%
									'price = rs("SecurityMktPrice")
								end if 
							else
								if trim(rs.Fields("Security_DPA_")) = trim(Offering) Then
									%>
									<option selected ParentSecurity = "<%=Rs("ParentSecurity_DPA_")%>" Ratio = "<%=Rs("Ratio")%>" OfferType = "<%=Rs("OfferType_DPA_")%>" SearchPrice = "<%=rs.Fields("SecurityMktPrice")%>" value = '<%=rs.Fields("Security_DPA_")%>'><%=rs.Fields("SecurityName")%></option>
									<%
									price = rs("SecurityMktPrice")
									ratio = Rs("Ratio")
								else
									%>                   						
									<option ParentSecurity = "<%=Rs("ParentSecurity_DPA_")%>" Ratio = "<%=Rs("Ratio")%>" OfferType = "<%=Rs("OfferType_DPA_")%>" SearchPrice = "<%=rs.Fields("SecurityMktPrice")%>" value = '<%=rs.Fields("Security_DPA_")%>'><%=rs.Fields("SecurityName")%></option>
									<%
									'price = rs("SecurityMktPrice")
								end if 
							end if
						rs.MoveNext
					Loop
				End If
				%>
				</select>
			</td>     
		</tr>
		
		<tr>
			<td width="20%" nowrap>Serial No</td>
			<td width="80%" nowrap>		
			<input type = 'text' name ='txtSearch' id = 'txtSearch' size="36" value = ''>
			&nbsp;&nbsp;
			<input type = 'Submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Search ">		
			</td>
		</tr>
	</table>

	<br><br>

	<%if Offering <> "" then
		if(SearchSerial<>"") then
			sqlStr = "SELECT * FROM OfferingsList WHERE Forward = 0 AND Offering = "& Offering &" AND Batch_no is null AND (PAL_No like '%" & SearchSerial & "%') AND ClosingDate > cast(Floor(cast(GetDate() as Float)) as datetime) order by pal_no asc"
		else
			sqlStr = "SELECT * FROM OfferingsList WHERE Forward = 0 AND Offering = "& Offering &" AND Batch_no is null AND ClosingDate >= cast(Floor(cast(GetDate() as Float)) as datetime) order by pal_no asc"
		end if	

		Set Rs = Server.CreateObject("ADODB.Recordset")
		Rs.CursorLocation = adUseClient 
		Rs.Open  SqlStr, Conn.ConnectionString, adOpenKeySet, adLockOptimistic
        
		dim bgcolor
		bgcolor="#FFFFFF"
		
		CheckedArray = ""

		If rs.EOF Or rs.BOF Then
			%>
				<table align=center border="0" cellspacing="1" cellpadding="1"  width="98%">
				<tr bgcolor="<%=bgcolor%>">
					<td width="100%" align=left>No applications found for batching.</td>
			    </tr>
			    </table>
				<%
			Response.end
		else
			%>
			<table align=center border="0" cellspacing="1" cellpadding="1"  width="98%">
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
			<td align=right>
			<BR><BR><input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save ">
			&nbsp;&nbsp;
			<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();">
			</td>
		</tr>
	</table>
	<%
	end if

	If Offering = "" Then%>
	<script language="javascript">
		window.location.href = 'AddBatchedApplication.asp?Offering='+ document.all.item("cboOfferings").options[document.all.item("cboOfferings").selectedIndex].value;
	</script>
	<%End If%>

<input type = 'hidden' name ='CommitParams' id = 'CommitParams' value="<%=CheckedArray%>">
<input type = 'hidden' name ='action' id = 'action' value="Save">&nbsp;
</form>

</body>
</html>