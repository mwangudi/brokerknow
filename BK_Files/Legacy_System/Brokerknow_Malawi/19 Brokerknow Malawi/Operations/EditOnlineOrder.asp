<!--#include file="../libroutines.asp"-->
<%
	const UDLName = "KBroker"
	const DataSource = "EditOnlineOrder"
	const DataEntity = "OnlineOrder"
	const DataEntityPlural = "OnlineOrders"
	const ActionFolder = "Operations"
	
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
   Dim ID
   Dim rsEdit	
	
	from =Request.QueryString("action")
	
	Set conn = GetActiveConnection("KBroker")

	if(trim(from)="") then
	action = ucase(Request.Form("action"))
	else
	action=ucase(from)
	end if
	
	ID = Request("ID")		
	
		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for editing"
                		
                </script>
                <% response.end
        End If    
	
	select case action	
		case "EXECUTE_HEADER"
			Dim reason
			Dim ApprovalAction			
			Dim orderDate
			Dim ref
			Dim secType	
			

			reason = Request.Form("txtReason")
			'ApprovalAction = Request.Form("txtApproval")
			ref = Request.Form("txtRef")
			Action =Cint(Request.Form("cboAction"))			
			
			ClientEmail=Request.Form("email") 			
				
				select case Action
				case 1
					AcceptOrder(ID)
					Subject = "Online Order Approved"
					Body = "Your Online Order has been approved" & vbCrLf
					Body = Body & "Date : " & Date() & vbCrLf
					Body = Body & "Online Order No : " & OrderID & vbCrLf
					Body = Body & "Approved Order No : " & vbCrLf
				case 2	
					'Action = UCase(Trim(Request.QueryString("appaction")))
					'if Action = "REJECT-INSUFFICIENT FUNDS" then
					'	Reason = "Insufficient funds"
					'elseif Action = "REJECT-PRICE HAS CHANGED" then
					'	Reason = "Price has changed"
					'else
					'	Reason = "User request"
					'end if							

					Subject = "Online Order Rejected"
					Body = "Your Online Order has been rejected." & vbCrLf
					Body = Body & "Date : " & Date() & vbCrLf
					Body = Body & "Online Order No : " & OrderID & vbCrLf
					Body = Body & "Reason : " & Reason & vbCrLf
					Body = Body & "NB: Please do not reply to this email address." & vbCrLf					
					
					Conn.Execute("Update WebtbOrder Set Action=2")
				case else					
					Subject = "Online Order Reviewed"
					Body = "Your Online Order is under review." & vbCrLf
					Body = Body & "Date : " & Date() & vbCrLf
					Body = Body & "Online Order No : " & OrderID & vbCrLf
					Body = Body & "Reason : " & Reason & vbCrLf
					Body = Body & "NB: Please do not reply to this email address." & vbCrLf

					Conn.Execute("Update WebtbOrder Set Action=3")
			end select
			
			if(Isnull(ClietEmail) or ClietEmail="") then
			else
			SendMail ClientEmail, Subject, Body		
			end if
			
           WritefraEnabledDialogCloseScript
           Response.End 
		sub AcceptOrder(ID)
		  
                
			sqlStr = "SELECT * FROM WebtbOrder WHERE (Order_DPA_ = " & ID & ")"
			Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
			
			If (rs.EOF Or rs.BOF) Then%>
				<script language = 'vbscript'>
						ShowMessage "The specified order cannot be found."
						window.history.back			
				</script>
				<% response.end
			End If
				
			set guid = server.createobject("NDUtils.CGUID")
			guidStr = guid.GenerateGUID
					 
			 sqlStr = "INSERT INTO [tbOrder] (OrderDate,OrderHold,OrderRef,Order_DPA_,Order_EIT_,Branch_DPA_,OrderSecType_DPA_,Client_DPA_,OrderType_DPA_,OrderAutoReleaseDate,OrderHoldType_DPA_,OrderCompounded) SELECT " & _
					 "#" & rs.fields("OrderDate") & "#" & " as OrderDate" & _
					 "," & " " & 1 & " " & " as OrderHold" & _
					 "," & "'" & rs.fields("OrderRef") & "'" & " as OrderRef" & _ 
					 "," & " " & "iif(isnull(max([Order_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'tbOrder'),max([Order_DPA_]) + 1)" & " " & " as Order_DPA_" & _
					 "," & "'" & guidStr & "'" & " as Order_EIT_" & _ 
					 "," & " " & Session("Branch_DPA_") & " " & " as Branch_DPA_" & _
					 "," & " " & rs.fields("OrderSecType_DPA_") & " " & " as OrderSecType_DPA_" & _
					 "," & " " & rs.fields("Client_DPA_") & " " & " as Client_DPA_" & _
					 "," & " " & rs.fields("OrderType_DPA_") & " " & " as OrderType_DPA_" & _
					 "," & " NULL " & " as OrderAutoReleaseDate" & _
					 "," & " " & 2 & " " & " as OrderHoldType_DPA_" & _
					 "," & " " & abs(cint(rs.fields("OrderCompounded"))) & " " & " as OrderCompounded" & _
					 " FROM [tbOrder]"		 
			
			 conn.BeginTrans
			 sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
			 conn.Execute sqlStr
						 
				'obtain header key value
				sqlStr = "SELECT [tbOrder.Order_DPA_] FROM [tbOrder] WHERE [tbOrder.Order_EIT_] = " & "'" & guidStr & "'"
							 
				 Dim orderRS
				Set orderRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
				If (orderRS.EOF Or orderRS.BOF) Then%>
						<script language = 'vbscript'>
								ShowMessage "A serious error has been encountered while saving the data. Try saving again"
												
						</script>
						<% response.end
				End If
						 
			'save detail data
			sqlStr = "INSERT INTO [OrdDetail] (OrdDetailCertNo,OrdDetailPrice,OrdDetailQty,OrdDetailValidity" & _
					",OrdDetail_DPA_,Order_DPA_,Security_DPA_) SELECT " & _
					"''" & " as OrdDetailCertNo" & _
					"," & "'" & rs.fields("OrdDetailPrice") & "'" & " as OrdDetailPrice" & _
					"," & " " & CDbl(rs.fields("OrdDetailQty")) & " " & " as OrdDetailQty" & _
					"," & "#" & rs.fields("OrdDetailValidity") & "#" & " as OrdDetailValidity" & _
					"," & " " & "iif(isnull(max([OrdDetail_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'OrdDetail'),max([OrdDetail_DPA_]) + 1)" & " " & " as OrdDetail_DPA_" & _
					"," & " " & orderRS.Fields("Order_DPA_") & " " & " as Order_DPA_" & _
					"," & " " & rs.fields("Security_DPA_") & " " & " as Security_DPA_" & _
					" FROM [OrdDetail]"
						 
			sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
			conn.Execute = sqlStr
						
			sqlStr = "UPDATE WebtbOrder SET Accepted = 1 WHERE(Order_DPA_ = " & RS.Fields("Order_DPA_") & ")"
			
			conn.Execute = sqlStr
			conn.CommitTrans
			conn.Close
				 		
    end sub

	Function GetOrderID(detailID)
		Dim getRs
		Set getConn = GetActiveConnection("KBroker")
		sqlStr = "SELECT OrdDetailList.Order_DPA_ FROM OrdDetailList WHERE " & _
                "  OrdDetailList.OrdDetail_DPA_=" & detailID
                
        set getRs = getConn.Execute(sqlStr)
        If Not (getRs.EOF OR getRs.BOF) Then
			GetOrderID = getRs("Order_DPA_")
		Else
			GetOrderID = detailID
		End If	        
		
   End Function
    
sub SendMail(toRecipient, subject, bodyText)
	cc = ""
	bcc = ""
	
	Set Conn = GetActiveConnection("KBroker")

	SQL = "SELECT Email FROM         Users INNER JOIN " & _
		 "                       UserGroups ON Users.UserID = UserGroups.UserID " & _
		 " WHERE     (NOT (Users.email IS NULL)) AND (UserGroups.GroupID = 1) OR " & _
		 "                       (UserGroups.GroupID = 2) OR " & _
		 "                       (UserGroups.GroupID = 6)"

		Set Rs = Conn.Execute(SQL)
		
		if not (Rs.eof and Rs.bof) then
			Do while Rs.eof=false
			cc=cc & Rs("Email") & ","
			Rs.movenext
			loop
		end if

	SQL = "SELECT * FROM MAilConfigList"
	Set Rs = Conn.Execute(SQL)
	
	If (Rs.EOF Or Rs.BOF) Then%>
		<Script Language="JavaScript">
			alert("The mail configurations have not been set.");
		</Script>
		<%Response.End 
	End If
				
	Const cdoSchema = "http://schemas.microsoft.com/cdo/configuration/" 
	Set objMsg = CreateObject("CDO.Message") 
	
	With objMsg
	
		.Configuration.Fields.Item(cdoSchema & "sendusing") = Rs.Fields("SendUsingMethod").Value 
		.Configuration.Fields.Item(cdoSchema & "smtpserver") = Rs.Fields("SMTPServer").Value 
		.Configuration.Fields.Item(cdoSchema & "smtpserverport") = Rs.Fields("SMTPServerPort").Value 
		.Configuration.Fields.Update 		 
		.Subject = subject
		.Sender = Rs.Fields("SendDisplayName").Value 
		.To = toRecipient
		If cc <> "" Then
			.CC = cc
		End If
		
		If bcc <> "" Then
			.BCC = bcc
		End If
		
		If bodyText <> "" Then
			.TextBody = bodyText
		End If
		
		.HTMLBody = bodyText	' "file://" & pathTo
		
		On Error Resume Next
		.Send 
		
		If Err.number > 0 Then%>
			<Script Language="JavaScript">
				alert("The page could not be sent. An unexpected error occured: <%= Err.description %>")
			</Script>
			<%Response.End 
		End If
	End With
	
	Set Rs = Nothing
	Set Conn = Nothing
	Set objMsg = Nothing
	Set Fso = Nothing%>
		<Script Language="JavaScript">
			alert("The page was sent successfully.")		
		</Script>
	<%	
end sub

	end select
  %>  
  
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit <%=DataEntity%></title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
 
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>

<script language="javascript">
	function AllowedNavigation()
		{
			validNavigate = true;
		}
		
	function DeleteOrderItem()
	{
			window.parent.frames("detail").HandleDeleteAction();
	}

	function  UpdateOrderStatus(action)
	{
	 var actionvalue=action.value;
	 var rejectval=0;
	 
	 if(actionvalue==2 || actionvalue==3)
	 {
	 	document.frmMain.elements["txtReason"].readonly = false;	
	 }
	 else
	 {
		document.frmMain.elements["txtReason"].readonly = true;	
	 }
	}	
	
</script>	
</head>

<body Class="Dialog" SCROLL="NO">

<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<form name = 'frm<%=DataSource%>' id='frmMain' method = 'post' action = 'EditOnlineOrder.asp' >
<%
        Set conn = GetActiveConnection("KBroker")
             
        
        sqlStr = "SELECT     WebtbOrder.Order_DPA_, Client.Client_DPA_, Client.ClientName,Client.ClientEmail,Branch.BranchDescription as BranchName, WebtbOrder.OrderCompounded, WebtbOrder.Accepted,  " & _
				 "                       WebtbOrder.ApprovalAction, WebtbOrder.Reason, WebtbOrder.UserName, WebtbOrder.ActionDate, WebtbOrder.ActionTime, WebtbOrder.OrderDate,  " & _
				 "                       WebtbOrder.OrderRef, OrderSecType.OrderSecTypeDisplayName, OrderType.OrderTypeDescription  as OrderTypeName,OrderType.OrderTypeSale,webtbOrder.Rejected,Action" & _
				 " FROM         WebtbOrder INNER JOIN " & _
				 "                       Client ON WebtbOrder.Client_DPA_ = Client.Client_DPA_ INNER JOIN " & _
				 "                       Branch ON Client.Branch_DPA_ = Branch.Branch_DPA_ INNER JOIN" & _
                  "    OrderSecType ON WebtbOrder.OrderSecType_DPA_ = OrderSecType.OrderSecType_DPA_ INNER JOIN " & _
                  "    OrderType ON WebtbOrder.OrderType_DPA_ = OrderType.OrderType_DPA_" & _
				 " WHERE     (WebtbOrder.Order_DPA_ = " & ID & ")"
					   
	   'Response.write(sqlStr)
	   'Response.end

        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then
			WriteDialogRefuseOpenScript%>
                <script language = 'vbscript'>
                		window.parent.dialogArguments.opener.alert "The selected Order cannot be retrieved for editing, " & Chr(13) & " because it has been released for trading."
                </script>
                <% 			   
                response.end
        End If
%>
<table border="0" width="600">
  <tr>
    <td width="15%">Order No.</td>
    <td width="35%"><input readonly = 'true' class=readonly  STYLE="WIDTH: 100px; text-align: right" type = 'text' name ='txtOrderNo' id = 'txtOrderNo' value = '<%=ID%>' size="20"></td>
    <td width="15%">Security Type</td>
    <td width="35%"><input readonly = 'true' class=readonly STYLE="WIDTH: 100px;" type = 'text' name ='txtOrderSecType' id = "txtOrderSecType" value = '<%=rs.Fields("OrderSecTypeDisplayname")%>' size="20">
    </td>
  </tr>
  <tr>
    <td width="15%">Branch</td>
    <td width="35%"><input readonly = 'true' class=readonly STYLE="WIDTH: 100px;" type = 'text' name ='txtBranch' id = "txtBranch" value = '<%=rs.Fields("BranchName")%>' size="20">
</td>
<td width="15%">Order Type</td>
    <td width="35%"><input readonly = 'true' class=readonly  type = 'text' name ='txtOrderType' id = "txtOrderType" value = '<%=rs.Fields("OrderTypeName")%>' size="20">
    <input type = 'hidden' name ='txtOrderTypeSale' id = "txtOrderTypeSale" value = '<%=rs.Fields("OrderTypeSale")%>' size="20">
 </td>    
  <tr>
    <td width="15%">Client</td>
    <td width="35%"><input readonly = 'true' class=readonly  type = 'text' name ='txtClient' id = "txtClient" value = '<%=rs.Fields("ClientName")%>' size="20">
</td>
  <td width="15%">Date</td>
    <td width="35%"><input readonly = 'true' class=readonly  type = 'text' name ='txtDate' id = "txtDate" value = '<%=FormatDate(rs.Fields("OrderDate"))%>' size="20"></td>  
  </tr>
  <tr>
  <td width="15%">Action</td>
  <td width="35%"><select name="cboAction" onchange="UpdateOrderStatus(this)">	
			<% if rs("Action")=1 then%>		
			<option selected SearchCode = "0" SearchText = "Accept" value = '1'>Accept</option>			
			<option value='2'>Reject</option>
			<option value='3'>Review</option>
			<% end if %>
			
			<% if rs("Action")=2 then%>		
			<option SearchCode = "0" SearchText = "Accept" value = '1'>Accept</option>			
			<option selected value='2'>Reject</option>
			<option value='3'>Review</option>
			<% end if %>
			
			<% if rs("Action")=3 then%>		
			<option SearchCode = "0" SearchText = "Accept" value = '1'>Accept</option>			
			<option value='2'>Reject</option>
			<option selected value='3'>Review</option>
			<% end if %>						
		</select>
      </td>      
  <td width="15%">Reason</td>
    <td width="35%" rowspan="2">
		
		<textarea id="txtReason" name="txtReason" rows='3' cols="25"><%=rs.Fields("Reason")%></textarea>		
    </td>	
  </tr>  	  
   <tr><td width="15%">Ref No</td>
    <td width="35%">
		<input type = 'text' name ='txtRef' id = 'txtRef' value = '<%=rs.Fields("OrderRef")%>' size="20" >
		<input type = 'hidden' name ='txtRefBk' id = 'txtRefBk' value = '<%=rs.Fields("OrderRef")%>' size="20">
    </td>
  
  </tr>
  <tr>
    <td width="100%" colspan=4 align=center>
	<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value="Save" onclick = "AllowedNavigation()">
     <input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Close " OnClick="JavaScript: window.self.close();">
    	<input type = 'hidden' name ='action' id = 'action' value="Execute_Header">
    	<input type = 'hidden' name ='ID' id = 'ID' value="<%= Rs.Fields("Order_DPA_").Value %>">
		<input type = 'hidden' name ='email' id = 'email' value="<%= Rs.Fields("ClientEmail").Value %>">    	    			  	
    </td>
  </tr>
  </table>
  </form>
</body>

</html>

