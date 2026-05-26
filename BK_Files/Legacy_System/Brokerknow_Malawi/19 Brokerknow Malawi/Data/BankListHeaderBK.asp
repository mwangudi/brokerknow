<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit Bank</title>
  <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <script language='vbscript'>
 function ItemSelected(itemID)
 		frmEditBankList.elements("ID").value = itemID
 		frmEditBankList.submit
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
					 href="javascript:PartWrapperToggle('AdvSearchHead');">Edit
                    Bank</A>
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
	
		 sqlStr = "SELECT * FROM [BankList]"
         Set conn = GetActiveConnection("KBroker")
        
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then
                %><p>No Banks found</p><%
                response.end
        End If
        
        rs.MoveFirst
        
%>
<form name = 'frmEditBankList' method = 'post' action = 'EditBank.asp' >
<table border="1" width="100%">
<tr>
	<td width="34%"><b><font color="#000080">Name</font></b></td>
  </tr>
<%		Do Until rs.EOF%>
                <tr>
                        <td width="34%"><b><font color="#000080"><a href = 'vbScript:ItemSelected(<%=rs.Fields("Bank_DPA_")%>)'><%=rs.Fields("BankName")%></a></font></b></td>
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

