<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>View Order</title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>

<script language="javascript">
function printDoc() {
	document.all.item("printLink").style.display = "none"
	print()
	document.all.item("printLink").style.display = ""
}
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
					 href="javascript:PartWrapperToggle('AdvSearchHead');">View Order Information
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
		
<DIV class="ListNuggetBody" id="AdvSearchHeadBody" name="AdvSearchHeadBody" style="WIDTH: 640px">
<table class="srch_bg" style="MARGIN-TOP: 0px" cellPadding="1" width=100% cellSpacing="0" border="0">  
<tr><td>



<%
	
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
   Dim ID
	
	ID = Request("ID")

		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		MsgBox "No record specified for viewing"
                		
                </script>
                <% response.end
        End If
%>
<form name = 'frmViewOrder' method = 'post' action = 'ViewOrderList.asp' >
<table border="0" width="628">
<tr name="printLink" id="printLink">
    <td width="101"><input type = 'submit' Class=Buttons name ='cmdAddDetail' id = 'cmdAddDetail' value="OK"></td>
    <td width="513"><a href="javascript:printDoc()"><img src="../images/printLink.gif" border="0" alt=PRINT><font color="blue">PRINT</font></a>
    </td>
  </tr>
  <%
        Set conn = GetActiveConnection("KBroker")
        
        
        sqlStr = "SELECT OrderDate,OrderHold,OrderRef,Order_DPA_,BranchName,ClientName,OrderTypeName,OwnerName FROM [OwnerList] INNER JOIN ([OrderTypeList] INNER JOIN ([ClientList] INNER JOIN ([BranchList] INNER JOIN [Order] ON BranchList.Branch_DPA_ = Order.Branch_DPA_) ON ClientList.Client_DPA_ = Order.Client_DPA_) ON OrderTypeList.OrderType_DPA_ = Order.OrderType_DPA_) ON OwnerList.Owner_DPA_ = Order.Owner_DPA_ WHERE Order_DPA_  = " & ID        
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		MsgBox "The selected Order cannot be retrieved for editing"
                		
                </script>
                <% response.end
        End If%>
  <tr>
    <td width="101"><b>Branch</b></td>
    <td width="513"><%=rs.Fields("BranchName")%></td>
  </tr>
  <tr>
    <td width="101"><b>Client</b></td>
    <td width="513"><%=rs.Fields("ClientName")%></td>
  </tr>
  <tr>
    <td width="101"><b>Order Type</b></td>
    <td width="513"><%=rs.Fields("OrderTypeName")%></td>
  </tr>
  <tr>
    <td width="101"><b>Owner</b></td>
    <td width="513"><%=rs.Fields("OwnerName")%></td>
  </tr>
  <tr>
  <%
		Dim hold
        
        if rs.Fields("OrderHold") then
        		hold = "Yes"
        else
        		hold = "No"
        end if%>
    <td width="101"><b>Hold</b></td>
    <td width="513"><%=hold%></td>
  </tr>
  <tr>
    <td width="101"><b>Date</b></td>
    <td width="513"><%=rs.Fields("OrderDate")%></td>
  </tr>
  <tr>
    <td width="101"><b>Ref No</b></td>
    <td width="513"><%=rs.Fields("OrderRef")%></td>
  </tr>
  <tr>
    <td colspan = '2' width="620">
    <table border="0" width="100%">
    <tr>
      <td width="32%"><b><font color="#000080">Security</font></b></td>
      <td width="16%"><b><font color="#000080">Quantity</font></b></td>
      <td width="12%"><b><font color="#000080">Price</font></b></td>
      <td width="24%"><b><font color="#000080">Certificate</font></b> <b><font color="#000080">No</font></b>.</td>
      <td width="16%"><b><font color="#000080">Validity</font></b></td>
    </tr>    
 <%
 	sqlStr = "SELECT OrdDetailCertNo,OrdDetailPrice,OrdDetailQty,OrdDetailValidity" & _
                ",OrdDetail_DPA_,SecurityName,Security_DPA_ FROM [SecurityList] INNER JOIN [OrdDetail] ON SecurityList.Security_DPA_ = OrdDetail.Security_DPA_ WHERE Order_DPA_  = " & ID
       
    Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
    If Not(rs.EOF Or rs.BOF) Then
        rs.MoveFirst
        Do Until rs.EOF%>
        		<tr>
                       <td><%=rs.Fields("SecurityName")%></td>
                        <td><%=rs.Fields("OrdDetailQty")%></td>
                        <td><%=rs.Fields("OrdDetailPrice")%></td>
                        <td><%=rs.Fields("OrdDetailCertNo")%></td>
                        <td><%=rs.Fields("OrdDetailValidity")%></td>
                        
             </tr>
             <%rs.MoveNext
        Loop
   End if
 %>
  </table>
    </td>
  </tr>
  <tr>
    <td width="101"></td>
    <td width="513" align="left">
    </td>
  </tr>
  <tr>
    <td width="101"></td>
    <td width="513" align="left">
    </td>
  </tr>
  <tr>
    <td width="101"></td>
    <td width="513" align="left">
      <p align="center">___________________________&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      _______________________
    </td>
  </tr>
  <tr>
    <td width="101"></td>
  </center>
    <td width="513" align="left">
      <p align="center"><b>Client Signature&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      Date</b></p>
    </td>
  </tr>


<CENTER>
  
</table>
</form>


</td>
</tr>
</table>
</div>
</div>

</body>

</html>
