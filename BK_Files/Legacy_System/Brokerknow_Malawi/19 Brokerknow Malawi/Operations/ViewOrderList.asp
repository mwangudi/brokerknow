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
 
 <script language='vbscript'>
 function ItemSelected(itemID)
 		frmViewOrderList.elements("ID").value = itemID
 		frmViewOrderList.submit
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
	Dim conn 
   Dim sqlStr
   Dim rs
	
		 'sqlStr = "SELECT OrderDate,OrderHold,OrderRef,Order_DPA_,BranchName,ClientName,OrderTypeName,OwnerName FROM [OwnerList] INNER JOIN ([OrderTypeList] INNER JOIN ([ClientList] INNER JOIN ([BranchList] INNER JOIN [Order] ON BranchList.Branch_DPA_ = Order.Branch_DPA_) ON ClientList.Client_DPA_ = Order.Client_DPA_) ON OrderTypeList.OrderType_DPA_ = Order.OrderType_DPA_) ON OwnerList.Owner_DPA_ = Order.Owner_DPA_ "
		 sqlStr = "SELECT * FROM [OrderList]"
		 Set conn = GetActiveConnection("KBroker")
        
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then
                %><p>No Orders found</p><%
                response.end
        End If
        
        rs.MoveFirst
        
%>
<form name = 'frmViewOrderList' method = 'post' action = 'ViewOrder.asp' >
<table border="0" width="100%">
<tr>
	<td width="34%"><b><font color="#000080">Client</font></b></td>
	<td width="34%"><b><font color="#000080">Date</font></b></td>
	<td width="34%"><b><font color="#000080">Reference No.</font></b></td>
	<td width="34%"><b><font color="#000080">Type</font></b></td>
  </tr>
<%		Do Until rs.EOF%>
                <tr>
                        <td width="34%"><%=rs.Fields("ClientName")%></td>
                        <td width="34%"><%=rs.Fields("OrderDate")%></td>
                        <td width="34%"><b><font color="#000080"><a href = 'vbScript:ItemSelected(<%=rs.Fields("Order_DPA_")%>)'><%=rs.Fields("OrderRef")%></a></font></b></td>
                        <td width="34%"><%=rs.Fields("OrderTypeName")%></td>
                </tr>
                <%rs.MoveNext
        Loop
        conn.Close
        Set conn = Nothing%>
</table>
    		<input type = 'hidden' name ='ID' id = 'ID' >
</form>

</td>
</tr>
</table>
</div>
</div>

</body>

</html>
