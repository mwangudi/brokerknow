<!--#include file="../libroutines.asp"-->
<%
	const UDLName = "KBroker"
	const DataSource = "EditOrder"
	const DataEntity = "Order"
	const DataEntityPlural = "Orders"
	const ActionFolder = "Operations"
	
	
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
   Dim ID
   Dim rsEdit
		
	from =Request.QueryString("action")
	
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
			Dim branch
			Dim client
			Dim orderType
			Dim orderDate
			Dim ref
			Dim secType
			Dim compound
			Dim hold
			Dim interbank
			 
			compound = cint(Request.Form("CompoundStatus"))
			interbank = cint(Request.Form("InterBankStatus"))
			hold = cint(Request.Form("HoldStatus"))
			branch = Request.Form("cboBranch")
			client = Request.Form("cboClient")
			orderType = Request.Form("cboOrderType")
			orderDate = Request.Form("txtDate")
			ref = Request.Form("txtRef")
			secType = Request.Form("cboOrderSecType")       
			toCancel = Request.Form("cmdCancel")
			
			If toCancel <> "" Then
				WriteDialogCloseScript
				Response.End
			End If   
				'validate Branch
				'If Trim(branch) = "" Then% >
'                        <script language = 'vbscript'>
'                                ShowMessage "Please specify the Branch"
'
'                        </script>
'                        < % response.end
'                End If
'                'validate Client
'                If Trim(Client) = "" Then% >
'                        <script language = 'vbscript'>
'                                ShowMessage "Please specify the Client"
'
'                        </script>
'                        <% response.end
'                End If
'                'validate Order Type
'                'If Trim(orderType) = "" Then% >
'                        <script language = 'vbscript'>
'                                ShowMessage "Please specify the Order Type"
'
'                        </script>
'                        < % response.end
'                End If
'                'validate Hold
'                If Trim(Hold) = "" Then% >
'                        <script language = 'vbscript'>
'                                ShowMessage "Please specify the Hold status"
'
'                        </script>
'                        < % response.end
'                End If
'                'validate Reference No.
'                'If Trim(ref) = "" Then% >
'                        <script language = 'vbscript'>
'                                'ShowMessage "Please specify the Reference No."
'
'                        </script>
'                        < % 'response.end
'            '   End If
'                'validate size of Reference No.
'                If Len(Ref) > 100 Then% >
'                        <script language = 'vbscript'>
'                        ShowMessage "Reference No. can only be 100 characters in length"
'
'                        </script>
'                        < % response.end
'                End If
'                'validate security type
'                If Trim(secType) = "" Then% >
'                        <script language = 'vbscript'>
'                                ShowMessage "Please specify the security type"
'
'                        </script>
'                        < % response.end
'                End If
'


			    
				Set conn = GetActiveConnection("KBroker")
			    
				'save data
				'sqlStr = "UPDATE [tbOrder] SET OrderDate = " & "#" & FormatDate(orderDate) & "#" & ",OrderHold = " & " " & hold & " " & "" & _
'                        ",OrderRef = " & "'" & ref & "'" & ",Branch_DPA_ = " & " " & branch & " " & ",Client_DPA_ = " & " " & client & " " & "" & _
'                        ",OrderType_DPA_ = " & " " & orderType & " " & _
'                        ",OrderSecType_DPA_ = " & " " & secType & " " & _
'                        ",OrderCompounded = " & " " & compound & " " & _
'                        " WHERE Order_DPA_  = " & ID
	
				Dim clearReleaseInfo
				
				if hold = 1 then
					clearReleaseInfo = " ,OrderReleasedBy = NULL" & _
										" ,OrderDateReleased = NULL"
				else
					clearReleaseInfo = ""
				end if
				
				sqlStr = "UPDATE [tbOrder] SET OrderCompounded = " & " " & compound & " " & _
						",InterBank = " & "" & interbank & " " & _
						" ,OrderHold = " & " " & hold & " " & clearReleaseInfo & _						
						",OrderRef = " & "'" & ref & "'" & _
						" WHERE Order_DPA_  = " & ID                

				    
				conn.BeginTrans
						conn.Execute SQLServerFormat(HandleQuote(sqlStr))
				conn.CommitTrans
				'retrieve the item ID
				sqlStr = "SELECT OrdDetail.OrdDetail_DPA_  FROM OrdDetail WHERE OrdDetail.Order_DPA_=" & ID
				Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
				ID = Rs.Fields("OrdDetail_DPA_")
				conn.Close
				Set conn = Nothing
				
				response.redirect DataSource & "Header.asp?ID=" & ID
		Case "FIRST"
		ID = ID	
		Case Else
			ID = GetOrderID(ID)	
    end select
    
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

	function  UpdateOrderStatus()
	{
		var compoundVal = "0"; //order not to be compounded
		var holdVal = "0"; //order not to be held
		var interbank = "0";
		
		if (document.frmMain.elements("chkCompound").checked)
		{
			compoundVal = "1";//order to be compounded
		}
		document.frmMain.elements("CompoundStatus").value = compoundVal;
		
		if (document.frmMain.elements("chkHold").checked)
		{
			holdVal = "1";//order to be held
		}
		document.frmMain.elements("HoldStatus").value = holdVal;
		
		if (document.frmMain.elements("chkInterBank").checked)
		{
			interbank = "1";//order to be compounded
		}
		document.frmMain.elements("InterBankStatus").value = interbank;
				
		SaveChanges();
	}
	
	function UpdateRef()
	{
		var oldVal = document.frmMain.elements("txtRefBk").value;
		var newVal = document.frmMain.elements("txtRef").value;
		
		if(oldVal !== newVal)
		{
			SaveChanges();
		}
	}
</script>
<script language='vbscript'>				
	function SaveChanges()
			frm<%=DataSource%>Header.submit
	end function
</script>		
</head>

<body Class="Dialog" SCROLL="NO">

<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<form name = 'frm<%=DataSource%>Header' id='frmMain' method = 'post' action = '<%=DataSource%>Header.asp' >
<%
        Set conn = GetActiveConnection("KBroker")
             
        
        sqlStr = "SELECT tbOrder.OrderCompounded,tbOrder.InterBank, tbOrder.OrderDate, tbOrder.OrderHold, tbOrder.OrderRef, tbOrder.Order_DPA_, tbOrder.OrderAutoReleaseDate,  " & _
                 "     BranchList.BranchName, ClientList.ClientName, OrderSecTypeList.OrderSecTypeDisplayName, OrderTypeList.OrderTypeName,  " & _
                  "    OrderTypeList.OrderTypeSale, OrderHoldTypeList.OrderHoldTypeName " & _
				"		FROM OrderHoldTypeList RIGHT OUTER JOIN " & _
                "      OrderTypeList INNER JOIN " & _
                "      OrderSecTypeList INNER JOIN " & _
                "      ClientList INNER JOIN " & _
                "      BranchList INNER JOIN " & _
                "      tbOrder ON BranchList.Branch_DPA_ = tbOrder.Branch_DPA_ ON ClientList.Client_DPA_ = tbOrder.Client_DPA_ ON  " & _
                "      OrderSecTypeList.OrderSecType_DPA_ = tbOrder.OrderSecType_DPA_ ON OrderTypeList.OrderType_DPA_ = tbOrder.OrderType_DPA_ ON  " & _
                "      OrderHoldTypeList.OrderHoldType_DPA_ = tbOrder.OrderHoldType_DPA_ " & _
				"	WHERE   tbOrder.Order_DPA_ = " & ID 
       
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
<table border="0" width="100%">
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
    <td width="15%">Hold</td>
    <td width="35%">
    <%if cbool(Rs.Fields("OrderHold")) then%>
		<input type=checkbox disabled  checked value='True' name='chkHold' id='chkHold' onClick = 'UpdateOrderStatus();'> 
	<%else%>
		<input type=checkbox   value='False' name='chkHold' id='chkHold' onClick = 'UpdateOrderStatus();'> 
	<%end if%>
</td>
  </tr>
  <tr>
    <td width="15%">Client</td>
    <td width="35%"><input readonly = 'true' class=readonly  type = 'text' name ='txtClient' id = "txtClient" value = '<%=rs.Fields("ClientName")%>' size="20">
</td>
  <td width="15%">Date</td>
    <td width="35%"><input readonly = 'true' class=readonly  type = 'text' name ='txtDate' id = "txtDate" value = '<%=FormatDate(rs.Fields("OrderDate"))%>' size="20"></td>  
  </tr>
  <tr>
    <td width="15%">Order Type</td>
    <td width="35%"><input readonly = 'true' class=readonly  type = 'text' name ='txtOrderType' id = "txtOrderType" value = '<%=rs.Fields("OrderTypeName")%>' size="20">
    <input type = 'hidden' name ='txtOrderTypeSale' id = "txtOrderTypeSale" value = '<%=rs.Fields("OrderTypeSale")%>' size="20">
 </td>
  <td width="15%">Compound&nbsp;</td>
    <td width="35%">
    <%if cbool(Rs.Fields("OrderCompounded")) then%>
		<input type=checkbox  checked value='True' name='chkCompound' id='chkCompound' onClick = 'UpdateOrderStatus();'> 
	<%else%>
		<input type=checkbox   value='False' name='chkCompound' id='chkCompound' onClick = 'UpdateOrderStatus();'> 
	<%end if%>
    </td>  
  </tr>
  <tr>
    <td width="15%">Ref No</td>
    <td width="35%">
		<input type = 'text' name ='txtRef' id = 'txtRef' value = '<%=rs.Fields("OrderRef")%>' size="20" onblur='UpdateRef()'>
		<input type = 'hidden' name ='txtRefBk' id = 'txtRefBk' value = '<%=rs.Fields("OrderRef")%>' size="20">
    </td>
    <td width="15%">Inter Bank&nbsp;</td>
    <td width="35%">
    <%if cbool(Rs.Fields("InterBank")) then%>
		<input type=checkbox  checked value='True' name='chkInterBank' id='chkInterBank' onClick = 'UpdateOrderStatus();'> 
	<%else%>
		<input type=checkbox   value='False' name='chkInterBank' id='chkInterBank' onClick = 'UpdateOrderStatus();'> 
	<%end if%>
    </td>
  </tr>
  <tr>
    <td width="100%" colspan=4 align=center>
    <input type = 'button' Class=Buttons name ='cmdDelete' id = 'cmdDelete' value=" Delete " onClick='DeleteOrderItem();'>
    <input type = 'button' Class=Buttons name ='cmdPrint' id = 'cmdPrint' value=" Print  " onClick="javascript: window.parent.location.href='OrderForm.asp?order_id=<%=ID%>'">
     <input type = 'submit' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Close ">
    	<input type = 'hidden' name ='action' id = 'action' value="Execute_Header">
    	<input type = 'hidden' name ='ID' id = 'ID' value="<%= Rs.Fields("Order_DPA_").Value %>">
    	<input type = 'hidden' name ='CompoundStatus' id = 'CompoundStatus' value='<%= cint(Rs.Fields("OrderCompounded").Value) %>'>
    	<input type = 'hidden' name ='InterBankStatus' id = 'InterBankStatus' value='<%= cint(Rs.Fields("InterBank").Value) %>'>
    	<input type = 'hidden' name ='HoldStatus' id = 'HoldStatus' value='<%= cint(Rs.Fields("OrderHold").Value) %>'>
    </td>
  </tr>
  </table>
  </form>
</body>

</html>
