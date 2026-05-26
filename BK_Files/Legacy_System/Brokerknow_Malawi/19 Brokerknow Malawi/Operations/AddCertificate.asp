<!--#include file="../libroutines.asp"-->
<%
	
	'======================= Begin_Alter_Across_Entities =================================
		
		const UDLName = "KBroker"
		const DataSource = "AddCertificate"
		const DataEntity = "Certificate"
		const DataEntityPlural = "Certificates"
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
                	ShowMessage "Please select a Contract"
                	window.self.close
			</script>
			<%
			
			response.end
	End If
	
	if action = "EXECUTE" then
		Dim certificate
		Dim rDate
       
       certificate = Request.Form("txtCertificate")
       rDate = Request.Form("txtRDate")
      
       
        'validate Certificate
        If Trim(certificate) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Certificate"
                		
                </script>
                <% 
				ReloadPage(ID)
				response.end
        End If
        'validate size of Certificate
        If Len(certificate) > 20 Then%>
                <script language = 'vbscript'>
                ShowMessage "Certificate can only be 20 characters in length"
                
                </script>
                <% 
				
				ReloadPage(ID)
				response.end
        End If
      
        'save data
        sqlStr = "UPDATE [Contract] SET ContractNCertificate = '" & certificate &  "'," & _
                " ContractNCDate = #" & FormatDate(rDate) & "#, " & _
                " ContractNCDelivered = 1 WHERE Contract_DPA_ = " & ID
        Set conn = GetActiveConnection("KBroker")
        
        conn.BeginTrans
                conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        conn.CommitTrans
        conn.Close
        Set conn = Nothing

		%>
		<SCRIPT LANGUAGE="JAVASCRIPT">
			window.opener.location= window.opener.location;
		</script>
		<%
        WritefraEnabledDialogCloseScript2
        Response.End
   	end If
	
	'fetch contract
	sqlStr = "SELECT * FROM ContractList WHERE Contract_DPA_=" & ID
	Set conn = GetActiveConnection("KBroker")
   	set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
   	If rs.EOF Or rs.BOF Then%>
            <script language = 'vbscript'>
                	window.self.ShowMessage "The selected Contract cannot be retrieved"
                	
            </script>
            <% 
			ReloadPage(ID)
			response.end
    End If 
    
    'get date to display
   ' if isnull(rs.Fields("ContractNCDate")) then
			'rDate = Date
   ' else
			rDate = rs.Fields("settlementdate")
   ' end if
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


	function forceSubmit()
	{
		setOpener();
		var targetPage = window.dialogArguments.opener.document.all.item("frmMain").elements("EditPage").value		
		document.frmMain.method='post';
		document.frmMain.target='_self';
		document.frmMain.submit();
	}
	
	function setOpener()
	{
		window.self.opener = window.dialogArguments.opener;
	}
</SCRIPT>
</head>

<body Class="Dialog"  onLoad="javascript: setOpener()">
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>

<form name = 'frm<%=DataSource%>' method = 'post' action = '<%=DataSource%>.asp' id='frmMain' >
<table border="0" cellspacing="1" cellpadding="1" height="218">
  <tr>
    <td nowrap height="13"> Contract</td>
    <td nowrap height="13"><input type = 'text' name ='txtContract' id = 'txtContract' size="20" STYLE="WIDTH: 80PX"  value = '<%=rs.Fields("ContractNumber")%>'   readonly = 'true' class=readonly >
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
    <td nowrap height="15"><input type = 'text' name ='txtSecurity' id = "txtSecurity" size="20" STYLE="WIDTH: 200PX"    value = '<%=rs.Fields("OrdDetailSecurity")%>'   readonly = 'true' class=readonly ></td>
  </tr>
  <tr>
    <td nowrap height="25"> CDS Ref</td>
    <td nowrap height="25"><input type = 'text' name ='txtSlipNo' id = 'txtSlipNo' size="20" STYLE="WIDTH: 100PX"   value = '<%=rs.Fields("LotSlipNo")%>'   readonly = 'true' class=readonly ></td>
  </tr>
  <tr>
    <td nowrap height="24"> Quantity</td>
    <td nowrap height="24"><input type = 'text' name ='txtQty' id = "txtQty" size="20" STYLE="WIDTH: 80PX; TEXT-ALIGN: RIGHT"  value = '<%= FormatNumCommasOnly(rs.Fields("LotQty")) %>'   readonly = 'true' class=readonly ></td>
  </tr>
  <tr>
    <td nowrap height="25"> Price</td>
    <td nowrap height="25"><input type = 'text' name ='txtPrice' id = "txtPrice" size="20" STYLE="WIDTH: 80PX; TEXT-ALIGN: RIGHT"   value = '<%= FormatNum(rs.Fields("LotPrice")) %>'   readonly = 'true' class=readonly ></td>
  </tr>
  <tr>
    <td nowrap height="25"> Date Recieved</td>
    <td nowrap height="25"><SCRIPT language="JavaScript">calDDate.writeControl();</SCRIPT></td>
  </tr>
  <tr>
    <td nowrap height="25"> New Certificate</td>
    <td nowrap height="25"><input type = 'text' name ='txtCertificate' id = "txtCertificate" size="20" value = '<%=rs.Fields("ContractNCertificate")%>'></td>
  </tr>
</table>
<table border=0 cellspacing=0 cellpadding=0 align=bottom width=100%>  
  <tr>
    <td align=right>
    <BR>
    <BR>
    <BR>
    <input type = 'button' class=buttons name ='cmdAdd' id = 'cmdAdd' value=" Save "  onClick= "javascript: forceSubmit();">
    &nbsp;&nbsp;&nbsp;
    <input type = 'button' class=buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.close();">
    <input type = 'hidden' name ='action' id = 'action' value="Execute">&nbsp;
    <input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
	<input type = 'hidden' name ='buttonAction' id = 'action' value="Save">
     </td>
  </tr>
 </table>
</form>

</body>

</html>
