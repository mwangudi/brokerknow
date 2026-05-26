<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Contract Delivery</title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>

 <script language='vbscript'>
 function ItemSelected(itemID)
 		frmContractDelivery.elements("ID").value = itemID
 		frmContractDelivery.submit
 end function
 </script>
</head>

<body><!--#include file="../libroutines.asp"-->

<CENTER>
	<DIV class="ListNugget" id="AdvSearchHead" style="WIDTH: 640px" name="AdvSearchHead">
		<TABLE class="ListNuggetHeader" cellPadding="0" cellSpacing="0" width="100%" name="AdvSearchtestHeader"> 
			<TR>
			<TD class="ListNuggetTitleCellWhite"
					onselectstart="window.event.cancelBubble=true; return false;"   
					onclick="PartWrapperToggle('AdvSearchHead');">
					<A class=ListNuggetTitle onclick="return PartWrapperToggle('AdvSearchHead');"  
					 href="javascript:PartWrapperToggle('AdvSearchHead');">Contract Delivery
					</A>
				</TD>
			 
				<TD class=ListNuggetButtonCellWhite onclick="PartWrapperToggle('AdvSearchHead');">
				<DIV class=ListNuggetButton>
					<IMG class=ListNuggetUpButton id=AdvSearchUp height=17 alt="Hide options" src="../images/blue-chevron_up.gif" width=17 align=right border=0 name=AdvSearchHeadUp>
					<IMG class=ListNuggetDownButton id=AdvSearchDown height=17 alt=Options src="../images/gray-chevron_down.gif" width=17 align=right border=0 name=AdvSearchHeadDown>
				</DIV>
			</TD>
			</TR>
		</TABLE>
<%
	Dim conn 
   Dim sqlStr
   Dim rs
	
	Set conn = GetActiveConnection("KBroker")
    
        
	action = ucase(Request.Form("action"))
	if action = "EXECUTE" then
		  
       ID = Request.Form("ID")

		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		MsgBox "No contract specified for delivery"
                		
                </script>
                <%response.end
        End If

       
        'save data
        sqlStr = "UPDATE [Contract] SET ContractDelDate = " & "#" & date & "#" & ",ContractDelivered = Yes" & _
                " WHERE Contract_DPA_  = " & ID
        conn.BeginTrans
                conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        conn.CommitTrans
       
   	end If
%>
<DIV class="ListNuggetBody" id="AdvSearchHeadBody" name="AdvSearchHeadBody" style="WIDTH: 640px">
<table class="srch_bg" style="MARGIN-TOP: 0px" cellPadding="1" width=100% cellSpacing="0" border="0">  
<tr><td>
<form name = 'frmContractDelivery' method = 'post' action = 'ContractDelivery.asp' >
<table border="0" width="100%">
 <tr>
    <td colspan = '2'>
    <table border="0" width="100%">
    <tr>
    	<td width="24%"><b><font color="#000080"></font></b></td>
      <td width="24%"><b><font color="#000080">Contract No</font></b></td>
      <td width="16%"><b><font color="#000080">Client</font></b></td>
      <td width="12%"><b><font color="#000080">Security</font></b></td>
     </tr>
 <%

	Set conn = GetActiveConnection("KBroker")
 	sqlStr = "SELECT * FROM PendingDeliveryList"			
	Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
     If Not(rs.EOF Or rs.BOF) Then
        'rs.MoveFirst
        Do Until rs.EOF%>
        		<tr>
        			<td><a href = 'vbScript:ItemSelected(<%=rs.Fields("Contract_DPA_")%>)'>Deliver</a></td>
        			<td width="24%"><b><font color="#000080"><%=rs.Fields("ContractCNumber")%></font></b></td>
					<td width="24%"><b><font color="#000080"><%=rs.Fields("Client")%></font></b></td>
					<td width="16%"><b><font color="#000080"><%=rs.Fields("SecurityName")%></font></b></td>
				</tr>
             <%rs.MoveNext
        Loop
     else%>
                <script language = 'vbscript'>
                		MsgBox "No pending deliveries found"
                		
                </script>
                <% response.end
     
   End if
 %>
  </table>
    </td>
    <input type = 'hidden' name ='ID' id = 'ID' >
    <input type = 'hidden' name ='action' id = 'action' value="Execute">

  </tr>
  </table>
</form>


</td>
</tr>
</table>
</div>
</div>

</body>

</html>
