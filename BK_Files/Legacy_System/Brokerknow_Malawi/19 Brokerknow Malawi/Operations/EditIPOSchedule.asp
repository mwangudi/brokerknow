<!--#include file="../libroutines.asp"-->
<%
	
	'======================= Begin_Alter_Across_Entities =================================
		
		const UDLName = "KBroker"
		const DataSource = "OfferingList"
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
		
		OfferingDPA = Request.Form("chkItem")
		'ChkItem = Cint(Request.Form("chkCommit92"))
		chkType = Cint(Request.Form("chkType"))	
		
		ChkNo=Request.Form("txtchkno")			
		BankCode = Request.Form ("txtbnkcode")
		AppType = Request.Form("cbotype")
		DrwAcc = Request.Form("DrwAcc")
		
		NextNo = Request.Form("NextItem")
		
		sqlmax="Select Max(isnull(ItemNo,0))+1 as NextBatch from tblApplDetails"
		
		Set MaxRs = conn.Execute(sqlmax)
		
		NextNo = MaxRs("NextBatch")
		
		if instr(1,buttonAction,"SEARCH") > 0 then
		SearchSerial = Request.Form("txtSearch")
		WriteDialogRelocateScript "EditIPOSchedule.asp?SearchSerial=" & SearchSerial & "&ID=" & BatchID
		else
		
		'Set conn1 = GetActiveConnection("kcb")

		if(chkType=2) then
			
			sqlStr = "SELECT UserName " & _
			 " FROM         users " & _
			 " WHERE     (UserID = " & Session("UserID") & ")"
		
		Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		CapturedBy = rs("UserName")

			sqlStr="Select * from OfferingList where Offering_DPA_ = " & OfferingDPA
			
			'Response.Write sqlStr
			'Response.End
			
			Set MaxRs = conn.Execute(sqlStr)	
		

		conn.BeginTrans	
			Do while MaxRs.eof=false
			
			sqlStr = "Insert into tblApplDetails (ScdID, AgentID, PayRecID, ItemNo, AppType, NoShares, AmtPayed, ChqNo, BankCode, DrwAcc, shNAME1, CapturedBy,OfferingDPA) VALUES( " & BatchID & " ,'023' " & _
					"," & BatchID & NextNo & " ," & NextNo & ", '" & AppType & "'," & MaxRs("Alloted_Rights") & "," & MaxRs("Offering_Price")*MaxRs("Alloted_Rights") & ",'" & ChkNo & "','" & BankCode & "' " & _
					",'" & DrwAcc & "','" & replace(MaxRs("ClientName"),"'","") & "','" & CapturedBy & "'," & MaxRs("Offering_DPA_") & ")"
					
					'Response.write sqlStr
					'Response.end

					conn.Execute(sqlstr)
					conn.Execute("UPDATE tblschedule set ICount=ICount+1 where ScheduleID = " & BatchID)
			MaxRs.moveNext
			loop
		
		conn.Execute("UPDATE offerings Set Paid=1 where Offering_DPA_=" & OfferingDPA)

		conn.CommitTrans
		'end if
		else
		conn.BeginTrans	
		 sqlStr ="Delete From tblApplDetails where OfferingDPA = " & OfferingDPA
		 conn.Execute(sqlStr)
		 conn.Execute("UPDATE tblschedule set ICount=ICount-1 where ScheduleID = " & BatchID)
		conn.CommitTrans
		
		conn.Execute("UPDATE offerings Set Paid=0 where Offering_DPA_=" & OfferingDPA)
		end if
	
	end if
	
	conn.Close
	'conn1.Close

	Set conn = Nothing
	
	
	WriteDialogRelocateScript "EditIPOSchedule.asp?ID=" & BatchID
	
	Response.end
end if

	if(action="editipo") then
	
	ItemID=Request.Form("ItemID")
	ID=Request.Form("ID")
	
	ChkNo=Request.Form("txtChk")
	Bankcode=Request.Form("txtBank")
	DrwAcc=Request.Form("txtDrawer")
	Sname=Request.Form("txtname")
	cbotype = Request.Form("cbotype")

		If Trim(ChkNo) = "" Then%>
				 <script language = 'vbscript'>
						ShowMessage "Please type the Cheque No"
						
				 </script>
				 <% response.end
		 End If	 
		 
		 'validate Account
		 If Trim(BankCode) = "" Then%>
				 <script language = 'vbscript'>
						ShowMessage "Please Enter the Bank Code"					         		
				 </script>
				 <% response.end
		 End If
		 		
		'Check whether price has been entered
		If Trim(DrwAcc) = "" Then%>
			<script language = 'vbscript'>
					ShowMessage "Please Enter the Drawer Account "
					
			</script>
			<% response.end
		End If
		
		sqlStr="UPDATE tblApplDetails SET ChqNo='" & ChkNo & "',BankCode = '" & Bankcode & "',DrwAcc = '" & DrwAcc & "'" & _
		",shNAME2 ='" & Sname & "',AppType='" & cbotype & "' Where PayRecID=" & ItemID

		'Set conn = GetActiveConnection("kcb")

		conn.BeginTrans
			conn.Execute(sqlStr)			
		conn.CommitTrans
	
	
	conn.close
	Set conn=nothing
	
	'response.redirect "EditIPOSchedule.asp?ID=" & ID
	WriteDialogRelocateScript "EditIPOSchedule.asp?ID=" & ID
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
function  CommitCDSTrade(theChk,theItem,OrderDetail_DPA_)
			{
			var thestring;
			var laststring;
                
				//document.frmMain.elements("delAction").value = "Execute";
				//document.frmMain.elements("OrderDetail_DPA_").value = OrderDetail_DPA_;
				
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
				
				thestring=',' + theItem + '<->' + OrderDetail_DPA_;
				laststring=theItem + '<->' + OrderDetail_DPA_;
				
				//alert(laststring);

					if(document.frmMain.elements("CommitParams").value==laststring)
					{
					document.frmMain.elements("CommitParams").value="";
					}
					else
					{					
					document.frmMain.elements("CommitParams").value = replaceSubstring(document.frmMain.elements("CommitParams").value,thestring,'') ;
					}
				}
				
				//alert(document.frmMain.elements("CommitParams").value);
				//ItemSelected(theItem);
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

<script language='vbscript'>				
	function SaveChanges(offering,thetype)
	Dim thename
	thename = offering
	'ShowMessage offering
			document.frmMain.elements("chkItem").value = thename
			document.frmMain.elements("chkType").value = thetype

			frm<%=DataSource%>.submit
	end function
</script>		

</head>

<body Class="Dialog">
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<%
if(action="edit") then
	ItemID=Request.QueryString("ItemID")
	ID = Request.QueryString("ID")

	'Set conn = GetActiveConnection("kcb")

	sqlStr = "Select * from tblApplDetails Where PayRecID=" & ItemID
	
	Set Rs = Server.CreateObject("ADODB.Recordset")
		Rs.CursorLocation = adUseClient 
		Rs.Open  SQLServerFormat(HandleQuote(sqlStr)), Conn.ConnectionString, adOpenKeySet, adLockOptimistic
       

	%>
	<form name = 'frmEdit' Method = 'post' action = 'EditIPOSchedule.asp?action=editipo' id='frmEdit'>
		<table>
			<tr>
				<td colspan="4" align="center"><font size=2 face="Tahoma" color=blue > <b> Edit Individual Application </b> </font></td>
			</tr>
			<tr>
				<td>Cheque No</td>
				<td>Bank Code</td>
				<td>Drawer Acc</td>
				<td>Second Name</td>
				<td>Application Type</td>
				<tr>  
			</tr>
			<tr>
				<td><input type = 'text' name ='txtChk' id = 'txtChk' size="20" value = '<%=Rs("ChqNo")%>'></td>
				<td><input type = 'text' name ='txtBank' id = 'txtBank' size="20" value = '<%=Rs("BankCode")%>'></td>
				<td><input type = 'text' name ='txtDrawer' id = 'txtDrawer' size="20" value = '<%=Rs("DrwAcc")%>'></td>
				<td><input type = 'text' name ='txtname' id = 'txtname' size="20" value = '<%=Rs("shNAME2")%>'></td>
				<td width="83%">
					<select name="cbotype" id="cbotype" size=1>	
						<option value="C">N</option>
						<option value="J">J</option>
						<option value="N">C</option>
					</select>
			</td>
			</tr>
			<tr>
				<td colspan="3"></td>
				<td align="right"><input type = 'Submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save ">&nbsp;
				<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.location.replace('EditIPOSchedule.asp?ID=<%=ID%>')"></td>
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
<form name = 'frm<%=DataSource%>' method = 'post' action = 'EditIPOSchedule.asp' id='frmMain' >
<table>
<tr><td colspan="2" align="right">		
			Serial No:&nbsp;&nbsp;<input type = 'text' name ='txtSearch' id = 'txtSearch' size="20" value = ''>&nbsp;&nbsp;<input type = 'Submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Search ">		
	</td></tr>
</table>

<table border="0" cellspacing="1" cellpadding="1"  width="100%">
  <tr>
	<td  valign="top">&nbsp;</td>
	<td  valign="top"><b><font color="#000080">Item No</font></b></td>
	<td  valign="top"><b><font color="#000080">Cheque No</font></b></td>
	<td  valign="top"><b><font color="#000080">Bank Code</font></b></td>
	<td  valign="top"><b><font color="#000080">Client Name</font></b></td>	
	<td  valign="top" align="right"><b><font color="#000080">Quantity Applied</font></b></td>	
	<td  valign="top" align="right"><b><font color="#000080">Payable</font></b></td>
	<td  valign="top" align="right"><b><font color="#000080">Edit</font></b></td>
  </tr>
<%

'Set conn1 = GetActiveConnection("kcb")

sqlStr = "SELECT * FROM tblApplDetails where ScdID =" & BatchID & " order by ItemNo Asc"
				
		Set Rs = Server.CreateObject("ADODB.Recordset")
		Rs.CursorLocation = adUseClient 
		Rs.Open  SQLServerFormat(HandleQuote(sqlStr)), Conn.ConnectionString, adOpenKeySet, adLockOptimistic
        
		dim bgcolor
		 bgcolor="#FFFFFF"
		
		CheckedArray = ""
		
		i=0
        If rs.EOF Or rs.BOF Then
		%>
		<script language="javascript">
		alert('No applications to Schedule');
		</script>
		<%
		Response.end
		else
			do until Rs.eof				
				'Populate Checked Orders Array
			if CheckedArray = "" then
			   CheckedArray = trim(rs.Fields("OfferingDPA")) & "<->" & trim(rs.Fields("ScdID"))
			 Else 
               CheckedArray  = CheckedArray & "," & trim(rs.Fields("OfferingDPA")) & "<->" & trim(rs.Fields("ScdID"))
			 end if
		
				commitTrade = "<input type='checkbox' checked class='BorderLess' name='chkCommit" & rs.Fields("OfferingDPA") & "' id='chkCommit" & rs.Fields("OfferingDPA") & "' onClick = 'SaveChanges(" & rs.Fields("OfferingDPA") & ",1);'>"

				if bgcolor="#F0EFDB" then  bgcolor="#FFFFFF" else bgcolor ="#F0EFDB"
				%>
				<tr bgcolor="<%=bgcolor%>" onMouseover="JavaScript: this.bgColor='#99CCFF'" onMouseout="JavaScript: this.bgColor='<%=bgcolor%>'">
				<td  valign="top"><%=commitTrade%></td>
				<td  valign="top"><%=Rs("ItemNo")%></td>
				<td  valign="top"><%=Rs("ChqNo")%></td>
				<td  valign="top"><%=Rs("BankCode")%></td>
				<td  valign="top" nowrap><%=mid(Rs("shNAME1"),1,30)%></td>				
				<td  valign="top" align="right"><%=FormatNumEx(Rs("NoShares"),0)%></td>				
				<td  valign="top" align="right"><%=FormatNum(Rs("AmtPayed"))%></td>
				<td  valign="top"><a href="editIPOSchedule.asp?ID=<%=Rs("ScdID")%>&action=edit&ItemID=<%=Rs("PayRecID")%>"><%=Rs("ItemNo")%></a></td>
			    </tr>
			<%
			ChkNo=Rs("ChqNo")			
			BankCode = Rs("BankCode")
			AppType = Rs("AppType")
			DrwAcc = Rs("DrwAcc")

			i=i+1
			Rs.movenext
			loop
		end if

		if(Cint(i)<25) then
			if(SearchSerial<>"") then
			sqlStr = "SELECT * FROM [" & DataSource & "] where Batch_no is not null AND (PAL_No like '%" & SearchSerial & "%') order by pal_no asc"
			else
			sqlStr = "SELECT * FROM [" & DataSource & "] where Batch_no is not null and paid=0 order by pal_no asc"
			end if	

			'Response.write  sqlStr
			'Response.end

			'Set MaxRs = Server.CreateObject("ADODB.Recordset")
			'Rs.CursorLocation = adUseClient 
			MaxRs.Open  SQLServerFormat(HandleQuote(sqlStr)), Conn.ConnectionString, adOpenKeySet, adLockOptimistic
			
			'dim bgcolor
			 bgcolor="#FFFFFF"
			
			'CheckedArray = ""

			If MaxRs.EOF Or MaxRs.BOF Then		
			else
				do until MaxRs.eof				
						
					commitTrade = "<input type='checkbox' class='BorderLess' name='chkCommit" & MaxRs.Fields("Offering_DPA_") & "' id='chkCommit" & MaxRs.Fields("Offering_DPA_") & "' onClick = 'SaveChanges(" & MaxRs.Fields("Offering_DPA_") & ",2);'>"

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
 <input type = 'hidden' name ='chkItem' id = 'chkItem' value="">&nbsp;
 <input type = 'hidden' name ='chkType' id = 'chkType' value="">&nbsp;
 <input type = 'hidden' name ='NextItem' id = 'NextItem' value="<%=i+1%>">&nbsp;
 <input type = 'hidden' name ='txtchkno' id = 'txtchkno' value="<%=ChkNo%>">&nbsp;
 <input type = 'hidden' name ='txtbnkcode' id = 'txtbnkcode' value="<%=BankCode%>">&nbsp;
 <input type = 'hidden' name ='AppType' id = 'AppType' value="<%=AppType%>">&nbsp;
 <input type = 'hidden' name ='DrwAcc' id = 'DrwAcc' value="<%=DrwAcc%>">&nbsp;
   		
</form>
<%
end if
%>
</body>

</html>
