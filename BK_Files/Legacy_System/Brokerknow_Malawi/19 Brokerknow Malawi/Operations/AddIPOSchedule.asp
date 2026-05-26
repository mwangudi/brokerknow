<!--#include file="../libroutines.asp"-->
<%
	
	'======================= Begin_Alter_Across_Entities =================================
		
		const UDLName = "KBroker"
		const UDLName1 = "kcb"
		const DataSource = "BatchedApplications"
		const DataEntity = "Certificate"
		const DataEntityPlural = "Certificates"
		const ActionFolder = "Operations"
'======================= End_Alter_Across_Entities =================================
	Set conn = GetActiveConnection(UDLName)

	dim action 
	Dim NextNo
	Dim Batchno
	Dim ChkNo
	Dim BankCode
	Dim AppType
	Dim DrwAcc
	Dim CapturedBy
	
	action = Request.Form("action")
	buttonAction = Trim(Ucase(Request.Form("cmdAdd")))
	
	Set MaxRs = Server.CreateObject("ADODB.Recordset")
		MaxRs.CursorLocation = adUseClient 
	
	Set conn = GetActiveConnection("KBroker")
	
	if(action ="Save") then
		
	Batchno=Request.Form("cbobatch")	
	ChkNo=Request.Form("txtchkno")			
	BankCode = Request.Form ("txtbnkcode")
	AppType = Request.Form("cbotype")
	DrwAcc = Request.Form("txtDrwAcc")
	
		'validate Entity
		 If Trim(Batchno) = "" Then%>
				 <script language = 'vbscript'>
						ShowMessage "Please Select the Batch No"
						
				 </script>
				 <% response.end
		 End If
		 
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
		 'validate Bank
		 If Trim(AppType) = "" Then%>
				 <script language = 'vbscript'>
						ShowMessage "Please Specify the App Type"
						
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
		
		sqlStr = "SELECT UserName " & _
			 " FROM         users " & _
			 " WHERE     (UserID = " & Session("UserID") & ")"
		
		Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		CapturedBy = rs("UserName")		
		

		sqlStr ="Select Max(ScdID) + 1 as ScheduleNo from tblApplDetails"
		Set rs = conn.Execute(sqlStr)
		
		ScheduleNo = rs("ScheduleNo")	
		
		if(Isnull(ScheduleNo)) then
		ScheduleNo =1 
		end if
		
			
		sqlStr="Select * from OfferingList where Batch_No=" & Batchno & " and paid=0 order by BatchSeq asc"
		
		'Response.write sqlStr
		'Response.end

		i=1
		Set MaxRs = conn.Execute(sqlStr)
			if Not(MaxRs.eof or MaxRs.bof) then	
				conn.BeginTrans			
					Do while MaxRs.eof=false
						
						if(i>25) then
						sqlstr="INSERT into tblschedule(ScheduleID, SchdStatus, ICount, BatchNo, BDate, BTime, AgentID) VALUES(" & ScheduleNo & ",'C'," & i-1 & ",'" & Batchno & "',cast(floor(cast(GetDate() as float)) as DateTime),'" & Time() & "','023')"
						
						conn.Execute(sqlstr)
						conn.Execute("UPDATE offerings Set Paid=1 where Batch_No=" & Batchno)

						i=1
						ScheduleNo=ScheduleNo + 1
						end if
					
					sqlStr = "Insert into tblApplDetails (ScdID, AgentID, PayRecID, ItemNo, AppType, NoShares, AmtPayed, ChqNo, BankCode, DrwAcc, shNAME1, CapturedBy,OfferingDPA) VALUES( " & ScheduleNo & " ,'023' " & _
					"," & ScheduleNo & i & " ," & i & ", '" & AppType & "'," & MaxRs("Alloted_Rights") & "," & MaxRs("Offering_Price")*MaxRs("Alloted_Rights") & ",'" & ChkNo & "','" & BankCode & "' " & _
					",'" & DrwAcc & "','" & replace(MaxRs("ClientName"),"'","") & "','" & CapturedBy & "'," & MaxRs("Offering_DPA_") & ")"
					
					'Response.write sqlStr
					'Response.end

					conn.Execute(sqlstr)
										
					i=i+1
					MaxRs.MoveNext
					loop
				
				'if(Cint(i)<>25) then
				sqlstr="INSERT into tblschedule(ScheduleID, SchdStatus, ICount, BatchNo, BDate, BTime, AgentID) VALUES(" & ScheduleNo & ",'C'," & i-1 & ",'" & Batchno & "',cast(floor(cast(GetDate() as float)) as DateTime),'" & Time() & "','023')"
				'end if

				conn.Execute(sqlstr)
				conn.Execute("UPDATE offerings Set Paid=1 where Batch_No=" & Batchno)

				conn.CommitTrans
			end if
	conn.Close
	
	Set conn = Nothing
	
	WritefraEnabledDialogCloseScript
	Response.End	
	'end if	
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

</head>

<body Class="Dialog">
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>

<form name = 'frm<%=DataSource%>' method = 'post' action = 'AddIPOSchedule.asp' id='frmMain' >
<table border="0" cellspacing="1" cellpadding="1"  width="100%">
 <tr>
    <td width="15%">Batch No</td>
    <td width="54%"><select name = 'cbobatch' id = 'cbobatch' size="1" >
    	
<%
		'Set conn = GetActiveConnection("KBroker")
		sqlStr = "SELECT DISTINCT Batch_No " & _
		" FROM         Offerings " & _
		" WHERE     (NOT (Batch_No IS NULL)) AND (Paid = 0) AND (Deleted = 0)"
		'" WHERE     (NOT (Batch_No IS NULL)) AND (Deleted = 0)"
 
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF                 
				%>
			   <option value = '<%=rs.Fields("Batch_No")%>'><%=rs.Fields("Batch_No")%></option>
             	<%		
                 rs.MoveNext
                Loop
        End If
%>

    </select></td>    
	</td>
  </tr>
  <tr>
	<td width="20%">Cheque No</td>
	<td width="80%"><input type="text" name="txtchkno" id="txtchkno" size="25" value=''></td>
  </tr>
  <tr>
	<td width="20%">Bank Code</td>
	<td width="80%"><input type="text" name="txtbnkcode" id="txtbnkcode" size="25" value=''></td>
  </tr>
  <tr>
    <td width="17%">Application Type</td>
    <td width="83%">
		<select name="cbotype" id="cbotype" size=1>	
			<option value="C">N</option>
			<option value="J">J</option>
			<option value="N">C</option>
		</select>
	</td>
  </tr>
  <tr>
	<td width="20%">Drawer Account</td>
	<td width="80%"><input type="text" name="txtDrwAcc" id="txtDrwAcc" size="35" value=''></td>
  </tr>
 <tr>	
	<td>
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Generate ">
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
