<!--#include file="../libroutines.asp"-->
<%
	
	'======================= Begin_Alter_Across_Entities =================================
		
		const UDLName = "KBroker"
		const DataSource = "AddDelivery"
		const DataEntity = "Delivery"
		const DataEntityPlural = "Deliveries"
		const ActionFolder = "Operations"
'======================= End_Alter_Across_Entities =================================
	Dim action
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim ID
	
	action = ucase(Request.Form("action"))
	
	ID = Request("ID")
	If (Trim(ID) = "") or (ID = "0") Then%>
			<script language = 'vbscript'>
                	ShowMessage "Please select a Contract for delivery"
                	window.self.close
			</script>
			<%response.end
	End If
	
	if action = "EXECUTE" then
		Dim transferNo
		Dim dDate
       
       transferNo = Request.Form("txtTransferNo")
       dDate = Request.Form("txtDDate")
      
       
        'validate Transfer No
        If Trim(transferNo) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Transfer No"
                		
                </script>
                <% response.end
        End If
        'validate size of Transfer No
        If Len(transferNo) > 20 Then%>
                <script language = 'vbscript'>
                ShowMessage "Transfer No can only be 20 characters in length"
                
                </script>
                <% response.end
        End If
      
        'save data
        sqlStr = "UPDATE [Contract] SET ContractTransferNo = '" & transferNo &  "'," & _
                " ContractDeliveryDate = #" & FormatDate(dDate) & "#, " & _
                " ContractDelivered = 1 WHERE Contract_DPA_ = " & ID
        Set conn = GetActiveConnection("KBroker")
        
        conn.BeginTrans
                conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        conn.CommitTrans
        conn.Close
        Set conn = Nothing
        WritefraEnabledDialogCloseScript
        Response.End
   	end If
	
	'fetch contract
	sqlStr = "SELECT * FROM ContractList WHERE Contract_DPA_=" & ID
	Set conn = GetActiveConnection("KBroker")
   	set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
   	If rs.EOF Or rs.BOF Then%>
            <script language = 'vbscript'>
                	window.self.ShowMessage "The selected Contract cannot be retrieved for delivery"
                	
            </script>
            <% response.end
    End If 
    
    'get date to display
    if isnull(rs.Fields("ContractDeliveryDate")) then
			DDate = Date
    else
			DDate = rs.Fields("ContractDeliveryDate")
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
	var calDDate=new ctlSpiffyCalendarBox("calDDate", "frm<%=DataSource%>", "txtDDate","cmdDDate","<%=FormatDate(DDate)%>",1);
</SCRIPT>
</head>

<body Class="Dialog">
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>

<form name = 'frm<%=DataSource%>' method = 'post' action = '<%=DataSource%>.asp' id='frmMain' >
<table border="0" cellspacing="1" cellpadding="1" height="218">
  <tr>
    <td nowrap height="13"> Contract</td>
    <td nowrap height="13"><input type = 'text' name ='txtContract' id = 'txtContract' size="20"  STYLE="WIDTH: 80PX"  value = '<%=rs.Fields("ContractNumber")%>'   readonly = 'true' class=readonly >
    </td>
  </tr>
  <tr>
    <td nowrap height="15"> Broker</td>
    <td nowrap height="15"><input type = 'text' name ='txtBroker' id = 'txtBroker' size="20"  STYLE="WIDTH: 80PX"  value = '<%=rs.Fields("BrokerCode")%>'   readonly = 'true' class=readonly ></td>
  </tr>
  <tr>
    <td nowrap height="15"> Client</td>
    <td nowrap height="15"><input type = 'text' name ='txtClient' id = 'txtClient' size="20"  STYLE="WIDTH: 200PX"  value = '<%=rs.Fields("OrdDetailClient")%>'   readonly = 'true' class=readonly ></td>
  </tr>
  <tr>
    <td nowrap height="15"> Security</td>
    <td nowrap height="15"><input type = 'text' name ='txtSecurity' id = "txtSecurity" size="20"  STYLE="WIDTH: 200PX" value = '<%=rs.Fields("OrdDetailSecurity")%>'   readonly = 'true' class=readonly ></td>
  </tr>
  <tr>
    <td nowrap height="15"> Transaction Type</td>
    <td nowrap height="15"><input type = 'text' name ='txtTransType' id = 'txtTransType' size="20" STYLE="WIDTH: 80PX"  value = '<%=rs.Fields("OrdDetailType")%>'   readonly = 'true' class=readonly ></td>
  </tr>
  <tr>
    <td nowrap height="25"> CDS Ref</td>
    <td nowrap height="25"><input type = 'text' name ='txtSlipNo' id = 'txtSlipNo' size="20" STYLE="WIDTH: 100PX"   value = '<%=rs.Fields("LotSlipNo")%>'   readonly = 'true' class=readonly ></td>
  </tr>
  <tr>
    <td nowrap height="24"> Quantity</td>
    <td nowrap height="24"><input type = 'text' name ='txtQty' id = "txtQty" size="20"  STYLE="WIDTH: 80PX; TEXT-ALIGN: RIGHT"   value = '<%= FormatNumCommasOnly(rs.Fields("LotQty")) %>'   readonly = 'true' class=readonly ></td>
  </tr>
  <tr>
    <td nowrap height="25"> Price</td>
    <td nowrap height="25"><input type = 'text' name ='txtPrice' id = "txtPrice" size="20" STYLE="WIDTH: 80PX; TEXT-ALIGN: RIGHT"   value = '<%= FormatNum(rs.Fields("LotPrice")) %>'   readonly = 'true' class=readonly ></td>
  </tr>
  <tr>
    <td nowrap height="25"> Delivery Date</td>
    <td nowrap height="25"><SCRIPT language="JavaScript">calDDate.writeControl();</SCRIPT></td>
  </tr>
  <tr>
    <td nowrap height="25"> Transfer No</td>
    <td nowrap height="25"><input type = 'text' name ='txtTransferNo' id = "txtTransferNo" size="20" value = '<%=rs.Fields("ContractTransferNo")%>'></td>
  </tr>
</table>
<table border=0 cellspacing=0 cellpadding=0 align=bottom width=100%>  
  <tr>
    <td align=right>
    <BR>
    <BR>
    <BR>
    <input type = 'submit' class=buttons name ='cmdAdd' id = 'cmdAdd' value=" Save ">
    &nbsp;&nbsp;&nbsp;
    <input type = 'button' class=buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.close();">
    <input type = 'hidden' name ='action' id = 'action' value="Execute">&nbsp;
    <input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
     </td>
  </tr>
 </table>
</form>

</body>

</html>
