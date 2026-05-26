<!--#include file="../libroutines.asp"-->
<%
	const UDLName = "KBroker"
	const DataSource = "EditLot"
	const DataEntity = "Lot"
	const DataEntityPlural = "Lots"
	const ActionFolder = "Operations"
	
	
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
   Dim ID
   Dim rsEdit
	
	action = ucase(Request.Form("action"))
    
          
		If action = "EXECUTE_HEADER" Then
			toCancel = Request.Form("cmdCancel")
			If toCancel <> "" Then
				WriteDialogCloseScript
				Response.End
			End If  
		End If

    Dim IDHolder
	Dim IDArray
	
	IDHolder = Request("ID")

	If (Trim(IDHolder) = "") or (IDHolder = "0") Then%>
			<script language = 'vbscript'>
                	ShowMessage "Please select an Order item for Lot allocation"
                	window.self.close
			</script>
			<%response.end
	End If
	
	IDArray = split(IDHolder,"<->")
	ID = IDArray(lbound(IDArray))
	
   	sqlStr = "SELECT * FROM OrdDetailList WHERE OrdDetailList.OrdDetail_DPA_= " & ID
   	
   	Set conn = GetActiveConnection("KBroker")
   	set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
   	If rs.EOF Or rs.BOF Then%>
            <script language = 'vbscript'>
                	window.self.ShowMessage "The selected Order item cannot be retrieved for lot allocation"
                	
            </script>
            <% response.end
    End If
    
    If Request.QueryString("action") = "save" Then
   		ContractSettlementDate = Request.QueryString("sDate")      

   		IDHolder = Request.QueryString("ID")
   		IDArray = split(IDHolder,"<->")
		ID = IDArray(lbound(IDArray))
		ItemID = IDArray(ubound(IDArray))
	
		Set conn = GetActiveConnection("KBroker")
			
		sqlstr = "SELECT * FROM Lot INNER JOIN Contract ON Lot.Contract_DPA_ = Contract.Contract_DPA_" & _
		" WHERE (Lot.OrdDetail_DPA_ = "& ID &")"
		Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		If (rs.EOF Or rs.BOF) Then
			%>
			<script language = 'vbscript'>
			 		ShowMessage "Save not possible."
			</script>
			<%
			response.end
		End If
			
		contractID = rs("Contract_DPA_")
					
		conn.BeginTrans
		SaveItem contractID, ContractSettlementDate
		conn.CommitTrans
		conn.Close
		Set conn = Nothing
		%>
		<script language="javascript">
		 window.location.href = 'EditLotHeader.asp?ID=<%=ID%>'
		</script>
		<%
		Response.End
	End If
    
	function SaveItem(Contract_DPA_,ContractSettlementDate)
		UserId=Session("UserID")
		
		sqlStr = "Update Contracts Set ModifiedBy=" & UserId & ", DateModified=GetDate(), ContractSettlementDate=#" & FormatDate(ContractSettlementDate) & "# WHERE Contract_DPA_ = " & Contract_DPA_
		'Response.Write sqlstr
		'Response.End 		
		conn.Execute SQLServerFormat(HandleQuote(sqlStr))
	end function    	
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
<%
	sqlstr = "SELECT Contract.ContractSettlementDate FROM Lot INNER JOIN Contract ON Lot.Contract_DPA_ = Contract.Contract_DPA_" & _
	" WHERE (Lot.OrdDetail_DPA_ = "& ID &")"

   	Set conn = GetActiveConnection("KBroker")
   	set rst = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
   	
   	If Not rst.EOF Or rst.BOF Then
		'set default settlement date
		DefaultDate = FormatDate(rst.Fields("ContractSettlementDate"))
	Else
		'set default settlement date
		DefaultDate = FormatDate(date)
	End If
%>
<SCRIPT language="JavaScript">
	var calDate = new ctlSpiffyCalendarBox("calDate", "frm<%=DataSource%>Header", "txtSettleDate","cmdSettleDate","<%=DefaultDate%>",1);
</SCRIPT>

<script language="JavaScript">

function DeleteOrderItem()
{
		window.parent.frames("detail").HandleDeleteAction();
}

function SaveOrderItem()
{
	sDate = window.document.all.item("txtSettleDate").value
	window.location.href = 'EditLotHeader.asp?action=save&ID=<%=IDHolder%>&sDate='+sDate
}
</script>		
</head>

<body Class="Dialog">

<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<form name = 'frm<%=DataSource%>Header' id='frmMain' method = 'post' action = '<%=DataSource%>Header.asp' >
<table border="0" width="100%">
  <tr>
    <td width="17%" height="25">Order No</td>
    <td width="83%" height="25"><input readonly = 'true' class=readonly  type = 'text'  STYLE="width: 100px; text-align: right"  name ='txtOrderNo' id = 'txtOrderNo' size="10" value = '<%=rs.Fields("Order_DPA_")%>'></td>
  </tr>
  <tr>
    <td width="17%" height="25">Item No</td>
    <td width="83%" height="25"><input readonly = 'true' STYLE="width: 100px; text-align: right" class=readonly  type = 'text' name ='txtItemNo' id = 'txtItemNo' size="10" value = '<%=rs.Fields("OrdDetail_DPA_")%>'></td>
  </tr>
  <tr>
    <td width="17%" height="25">Order Type</td>
    <td width="83%" height="25"><input readonly = 'true' class=readonly  type = 'text' STYLE="width: 150px" name ='txtOrderType' id = 'txtOrderType' size="20" value = '<%=rs.Fields("OrdDetailType")%>'></td>
  </tr>
  <tr>
    <td width="17%" height="25">Instrument</td>
    <td width="83%" height="25"><input readonly = 'true' class=readonly  type = 'text' STYLE="width: 150px" name ='txtInstrument' id = 'txtInstrument' size="20" value = '<%=rs.Fields("OrdDetailSecType")%>'></td>
  </tr>
   <tr>
    <td width="17%" height="25">Client</td>
    <td width="83%" height="25">
<input readonly = 'true' class=readonly  type = 'text' name ='txtClient' id = 'txtClient' STYLE="width: 300px" size="35" value = '<%=rs.Fields("OrdDetailClient")%>'></td>
  </tr>

  <tr>
    <td width="17%" height="25">Security</td>
    <td width="83%" height="25"><input readonly = 'true' class=readonly STYLE="width: 300px"  type = 'text' name ='txtSecurity' id = 'txtSecurity' size="35" value = '<%=rs.Fields("OrdDetailSecurity")%>'></td>
  </tr>
  <tr>
    <td width="17%" height="25">Balance<%=itemid%></td>
    <td width="83%" height="25"><input readonly = 'true' STYLE="width: 150px; text-align: right" class=readonly  type = 'text' name ='txtBalance' id = 'txtBalance' size="20" value = '<%= FormatNum(rs.Fields("BalanceQty")) %>'></td>
  </tr>
 
 <!--<tr>
    <td nowrap width="17%" height="25">Settlement Date</td>
    <td nowrap width="83%" height="25"><SCRIPT language="JavaScript">calDate.writeControl();</SCRIPT></td>
  </tr>-->
  
  <tr>
    <td align="right" colspan=2 nowrap width="17%" height="25">
		<!--<input type = 'button' class=buttons name ='cmdSave' id = 'cmdSave' value=" Save " onClick='SaveOrderItem();'>
		&nbsp;&nbsp;-->
		<input type = 'button' style="display:none" class=buttons name ='cmdDelete' id = 'cmdDelete' value=" Delete " onClick='DeleteOrderItem();'>
		&nbsp;&nbsp;
		<input type = 'submit' class=buttons name ='cmdCancel' id = 'cmdCancel' value=" Close ">
    </td>
  </tr>
  
 	<input type = 'hidden' name ='action' id = 'action' value="Execute_Header">
	<input type = 'hidden' name ='ID' id = 'ID' value="<%=IDHolder%>">
  </table>
  
  </form>
</body>

</html>
