<!--#include virtual="libroutines.asp"-->
<%
	
	const UDLName = "KBroker"
	const DataSource = "DB_Affirmation"
	const DataEntity = "SettlementDate"
	const DataEntityPlural = "SettlementDates"
	const ActionFolder = "Operations"
	
	Dim UserId
	Dim action
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim rsEdit
	Dim ID
	
	action = ucase(Request.Form("action"))
	ID = Request("ID")
	UserId=Session("UserID")
                
		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for editing"	
                </script>
                <% response.end
        End If
        
	if  action = "EXECUTE" then
                
					Dim RecordID
					Dim SettlementDate 
					Dim ContractNo
			        
					SettlementDate = Trim(Request.Form("txtDate"))
					ContractNo = Trim(Request.Form("ContractNo"))
		            
		            'Validate
		             If Trim(ContractNo) = "" Then%>
                      <script language = 'vbscript'>
                		ShowMessage "This record cannot be saved! Contract Number Missing"	
                      </script>
                      <% 
                      WritefraEnabledDialogCloseScript
                      response.end
                      
                     End If		 
		            		        
					'save data
					Set conn = GetActiveConnection("KBroker")
					
					conn.BeginTrans
					
						   sqlStr = "UPDATE contract SET " & _
									" ContractSettlementDate = #" & SettlementDate & "#" & _ 
									" WHERE Contract_DPA_= " & ID
							sqlStr = SQLServerFormat(HandleQuote(sqlStr))
							conn.Execute sqlStr
					
					conn.CommitTrans		
					WritefraEnabledDialogCloseScript2
				    conn.Close
					Set conn = Nothing
					
					Response.End
	
   end if
   
	Set conn = GetActiveConnection("KBroker")
	sqlStr = "SELECT * FROM DB_Affirmation2 WHERE Contract_DPA_=" & ID
			        
	Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
    If rs.EOF Or rs.BOF Then%>
	   <script language = 'vbscript'>
	    window.self.ShowMessage "The selected <%=DataEntity%> cannot be retrieved for editing"             		
	   </script>
		<% response.end
	 End If

%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>EDIT <%=DataEntity%></title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/accountList.js"></SCRIPT>
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<script language="JavaScript" src="CALENDAR/calendar.js"></script>
<!--END CALENDAR -->

<script language='javascript'>
	function forceSubmit()
	{
		//setOpener();
		//var targetPage = window.dialogArguments.opener.document.all.item("frmMain").elements("EditPage").value;
				
		document.frm<%=DataSource%>.method='post';
		document.frm<%=DataSource%>.target='_self';
		document.frm<%=DataSource%>.submit();	
		
	}
	
	function setOpener()
	{

		window.self.opener = window.dialogArguments.opener;
				//alert(window.dialogArguments.opener.location);
	}
</script>
</head>

<body Class="Dialog" onLoad="javascript: setOpener()">
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<SCRIPT language="JavaScript">
	var cal=new ctlSpiffyCalendarBox("cal", "frm<%=DataSource%>", "txtDate", "cmdDate","<%=FormatDate(rs.Fields("SettlementDate")) %>",1);
</SCRIPT>

<form name = 'frm<%=DataSource%>' method = 'post' action = 'EditSettlementDate.asp' id = "frmMain">
<table border="0" width="100%" cellspacing="1" cellpadding="1">
  <tr>
    <td width="40%">Settlement Date</td>
    <td width="60%"><SCRIPT language="JavaScript">cal.writeControl();</SCRIPT></td>
 
  </tr>
  <tr>
    <td width="40%">Trade Date</td>
    <td width="60%"><input type ="text" Name ="TradeDate" value="<%=FormatDate(rs.Fields("LotTDate"))%>" readonly class=readonly style= "width:256" size="20"></td>
 
  </tr>
  <tr>
    <td width="40%">Client</td>
    <td width="60%"><input type ="text" Name ="Client" value="<%=rs.Fields("Account")%>" readonly class=readonly style= "width:256" size="20"></td>
 
  </tr>
  <tr>
    <td width="40%">Order No</td>
    <td width="60%"><input type ="text" Name ="OrderNo" value="<%=rs.Fields("Order_DPA_")%>" readonly class=readonly style= "width:256" size="20"></td>
 
  </tr>
  <tr>
    <td width="40%">Contract</td>
    <td width="60%"><input type ="text" Name ="Contract" value="<%=rs.Fields("ContractNumber")%>" readonly class=readonly style= "width:256" size="20"></td>
 
  </tr>
  <tr>
    <td width="40%">Security</td>
    <td width="60%"><input type ="text" Name ="TradeDate" value="<%=rs.Fields("SecurityCode")%>" readonly class=readonly style= "width:256" size="20"></td>
 
  </tr>
  <tr>
    <td width="40%">Slip No.</td>
    <td width="60%"><input type ="text" Name ="SlipNo" value="<%=rs.Fields("LotSlipNo")%>" readonly class=readonly style= "width:256" size="20"></td>
 
  </tr>
  <tr>
    <td width="40%">Price</td>
    <td width="60%"><input type ="text" Name ="Price" value="<%=FormatNumber(rs.Fields("LotPrice"),4)%>" readonly class=readonly style= "width:256" size="20"></td>
 
  </tr>
  <tr>
    <td width="40%">Quantity</td>
    <td width="60%"><input type ="text" Name ="Quantity" value="<%=rs.Fields("LotQty")%>" readonly class=readonly style= "width:256" size="20"></td>
 
  </tr>
  <tr>
    <td width="40%">Gross</td>
    <td width="60%"><input type ="text" Name ="Gross" value="<%=formatnumber(rs.Fields("SettlementGrossAmount"),2)%>" readonly class=readonly style= "width:256" size="20"></td>
 
  </tr>
  <%
  Commission=rs.Fields("Comm") + rs.Fields("Levytotal")
  %>
  <tr>
    <td width="40%">Commission</td>
    <td width="60%"><input type ="text" Name ="Commission" value="<%=Formatnumber(Commission,2)%>" readonly class=readonly style= "width:256" size="20"></td>
 
  </tr>
  <tr>
    <td width="40%">Net Amount</td>
    <td width="60%"><input type ="text" Name ="NetAmount" value="<%=Formatnumber(rs.Fields("SettlementAmount"),2)%>" readonly class=readonly style= "width:256" size="20"></td>
 
  </tr>
  <tr>
	  <td width="100%" colspan=3 align="center" valign=absBottom>
		<BR><BR>
		<input type = 'button' Class=Buttons name ='cmdEdit' id = 'cmdEdit' value=" Edit " onclick = "forceSubmit();">
    	<input type = 'button' Class=Buttons name ='cmdCancel' id = "cmdCancel" value="Cancel" onclick = "JavaScript: window.self.close()">
		<input type = 'hidden' name ='action' id = 'action' value="EXECUTE">
		<input type = 'hidden' name ='ID' id = 'ID' value="<%= Rs.Fields("Contract_DPA_").Value %>">
		<input type = 'hidden' name ='ContractNo' id = 'ContractNo' value="<%= Rs.Fields("ContractNumber").Value %>">
		<input type = 'hidden' name ='buttonAction' id = 'action' value="Save">
	</td>
  </tr>
</table>
</form>
</body>
</html>