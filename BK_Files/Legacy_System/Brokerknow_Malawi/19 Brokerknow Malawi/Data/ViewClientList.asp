<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>View Clients</title>
  <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
 <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css"> 
 <SCRIPT language=Javascript src="scripts/common.js"></SCRIPT>

 <script language='vbscript'>
 function ItemSelected(itemID)
 		frmViewClientList.elements("ID").value = itemID
 		frmViewClientList.submit
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
					 href="javascript:PartWrapperToggle('AdvSearchHead');">View Client Information
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
	
		 'sqlStr = "SELECT ClientAddr,ClientBDate,ClientCellTel,ClientContact,ClientEmail" & _
                '",ClientFax,ClientGeneric1,ClientGeneric2,ClientGeneric3" & _
       '         ",ClientHomeTel,ClientIDPass,ClientName,ClientOfficeTel,ClientPhoto" & _
        '        ",ClientVIP,Client_DPA_,AgentName,BranchName,ClassClass,CommissionRate,GenderGender" & _
         '       ",OwnerName,ResidencyName FROM [ResidencyList] INNER JOIN ([OwnerList] INNER JOIN ([GenderList] " & _
          '      "INNER JOIN ([CommissionList] INNER JOIN ([ClassList] INNER JOIN ([BranchList] INNER JOIN ([AgentList] " & _
           '     "INNER JOIN [Client] ON AgentList.Agent_DPA_ = Client.Agent_DPA_) ON BranchList.Branch_DPA_ = Client.Branch_DPA_) " & _
            '    "ON ClassList.Class_DPA_ = Client.Class_DPA_) ON CommissionList.Commission_DPA_ = Client.Commission_DPA_) ON " & _
             '   "GenderList.Gender_DPA_ = Client.Gender_DPA_) ON OwnerList.Owner_DPA_ = Client.Owner_DPA_) ON " & _
              '  "ResidencyList.Residency_DPA_ = Client.Residency_DPA_"
         sqlStr = "SELECT * FROM [ClientList]"
         Set conn = GetActiveConnection("KBroker")
        
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then
                %><p>No Clients found</p><%
                response.end
        End If
        
        rs.MoveFirst
        
%>



<form name = 'frmViewClientList' method = 'post' action = 'ViewClient.asp' >
<table border="0" width="100%">
<tr>
	<td width="34%"><b><font color="#000080">Name</font></b></td>
	<td width="34%"><b><font color="#000080">Owner</font></b></td>
	<td width="34%"><b><font color="#000080">Contact</font></b></td>
	<td width="34%"><b><font color="#000080">Office Phone</font></b></td>
	<td width="34%"><b><font color="#000080">Cell Phone</font></b></td>
	<td width="34%"><b><font color="#000080">Email</font></b></td>
	<td width="34%"><b><font color="#000080">Address</font></b></td>
	<td width="34%"><b><font color="#000080">High Net Worth</font></b></td>
  </tr>
<%		Do Until rs.EOF%>
                <tr>
                        <td width="34%"><b><font color="#000080"><a href = 'vbScript:ItemSelected(<%=rs.Fields("Client_DPA_")%>)'><%=rs.Fields("ClientName")%></a></font></b></td>
                        <td width="34%"><%=rs.Fields("OwnerName")%>&nbsp</td>
                        <td width="34%"><%=rs.Fields("ClientContact")%>&nbsp</td>
                        <td width="34%"><%=rs.Fields("ClientOfficeTel")%>&nbsp</td>
                        <td width="34%"><%=rs.Fields("ClientCellTel")%>&nbsp</td>
                        <td width="34%"><%=rs.Fields("ClientEmail")%>&nbsp</td>
                        <td width="34%"><%=rs.Fields("ClientAddr")%>&nbsp</td>
                        <%if(rs.Fields("ClientVIP")) then%>
							<td width="34%">Yes&nbsp</td>
						<%else%>
							<td width="34%">No&nbsp</td>
						<%end if%>
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
</center>

</body>

</html>










