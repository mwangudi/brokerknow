<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>View Client</title>

  <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
 <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css"> 
 <SCRIPT language=Javascript src="scripts/common.js"></SCRIPT>
 
</head>

<body><!--#include file="../libroutines.asp"-->

<%
	
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
   Dim ID
	
	ID = Request.Form("ID")

		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		MsgBox "No record specified for viewing"
                		window.history.back
                </script>
                <% response.end
        End If
%>

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

<form name = 'frmViewClient' method = 'post' action = 'ViewClientList.asp' >
<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <%
        Set conn = GetActiveConnection("KBroker")
        
        
        sqlStr = "SELECT ClientAddr,ClientBDate,ClientCellTel,ClientContact,ClientEmail" & _
                ",ClientFax,ClientGeneric1,ClientGeneric2,ClientGeneric3" & _
                ",ClientHomeTel,ClientIDPass,ClientName,ClientOfficeTel,ClientPhoto" & _
                ",ClientVIP,Client_DPA_,AgentName,BranchName,ClassClass,CommissionRate,GenderGender" & _
                ",OwnerName,ResidencyName FROM [ResidencyList] INNER JOIN ([OwnerList] INNER JOIN ([GenderList] INNER JOIN ([CommissionList] INNER JOIN ([ClassList] INNER JOIN ([BranchList] INNER JOIN ([AgentList] INNER JOIN [Client] ON AgentList.Agent_DPA_ = Client.Agent_DPA_) ON BranchList.Branch_DPA_ = Client.Branch_DPA_) ON ClassList.Class_DPA_ = Client.Class_DPA_) ON CommissionList.Commission_DPA_ = Client.Commission_DPA_) ON GenderList.Gender_DPA_ = Client.Gender_DPA_) ON OwnerList.Owner_DPA_ = Client.Owner_DPA_) ON ResidencyList.Residency_DPA_ = Client.Residency_DPA_ WHERE Client_DPA_  = " & ID        
        
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		MsgBox "The selected Client cannot be retrieved for viewing"
                		window.history.back
                </script>
                <% response.end
        End If%>
	<tr>
    <td width="17%"><b>Branch</b></td>
    <td width="83%">
        <%=rs.Fields("BranchName")%>
        </td>
  </tr>
   <tr>
    <td width="17%"><b>Agent</b></td>
    <td width="83%"><%=rs.Fields("AgentName")%></td>
  </tr>

 <tr>
    <td width="17%"><b>Class</b></td>
    <td width="83%"><%=rs.Fields("ClassClass")%></td>
  </tr>

 <tr>
    <td width="17%"><b>Owner</b></td>
    <td width="83%"><%=rs.Fields("OwnerName")%></td>
  </tr>

  <tr>
    <td width="17%"><b>Commission</b></td>
    <td width="83%"><%=rs.Fields("CommissionRate")%></td>
  </tr>
  <tr>
    <td width="17%"><b>Gender</b></td>
    <td width="83%"><%=rs.Fields("GenderGender")%></td>
  </tr>
  <tr>
    <td width="17%"><b>Residency</b></td>
    <td width="83%"><%=rs.Fields("ResidencyName")%></td>
  </tr>

  <tr>
    <td width="17%"><b>Address</b></td>
    <td width="83%"><%=rs.Fields("ClientAddr")%></td>
  </tr>
  <tr>
    <td width="17%"><b>Date of Birth</b></td>
    <td width="83%"><%=rs.Fields("ClientBDate")%></td>
  </tr>
  <tr>
    <td width="17%"><b>Name</b></td>
    <td width="83%"><%=rs.Fields("ClientName")%></td>
  </tr>
  <tr>
    <td width="17%"><b>Contact Name</b></td>
    <td width="83%"><%=rs.Fields("ClientContact")%></td>
  </tr>
  <tr>
    <td width="17%"><b>Cell Phone</b></td>
    <td width="83%"><%=rs.Fields("ClientCellTel")%></td>
  </tr>
  <tr>
    <td width="17%"><b>Email</b></td>
    <td width="83%"><%=rs.Fields("ClientEmail")%></td>
  </tr>
  <tr>
    <td width="17%"><b>Fax</b></td>
    <td width="83%"><%=rs.Fields("ClientFax")%></td>
  </tr>
  <tr>
    <td width="17%"><b>Home Phone</b></td>
    <td width="83%"><%=rs.Fields("ClientHomeTel")%></td>
  </tr>
  <tr>
    <td width="17%"><b>ID/Passport</b></td>
    <td width="83%"><%=rs.Fields("ClientIDPass")%></td>
  </tr>
  <tr>
    <td width="17%"><b>Office Phone</b></td>
    <td width="83%"><%=rs.Fields("ClientOfficeTel")%></td>
  </tr>
  <tr>
    <td width="17%"><b>Photo</b></td>
    <td width="83%"><%=rs.Fields("ClientPhoto")%></td>
  </tr>
  <tr>
  <%
		Dim vip
        
        if rs.Fields("ClientVIP") then
        		vip = "Yes"
        else
        		vip = "No"
        end if%>
    <td width="17%"><b>High Net Worth</b></td>
    <td width="83%"> <%=vip%></td>
  </tr>
  <tr>
    <td width="17%"><b>Generic 1</b></td>
    <td width="83%"><%=rs.Fields("ClientGeneric1")%></td>
  </tr>
  <tr>
    <td width="17%"><b>Generic 2</b></td>
    <td width="83%"><%=rs.Fields("ClientGeneric2")%></td>
  </tr>
  <tr>
    <td width="17%"><b>Generic 3</b></td>
    <td width="83%"><%=rs.Fields("ClientGeneric3")%></td>
  </tr>
  <tr>
    <td width="17%"><input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value="OK"></td>
    <td width="83%">
    	
    </td>
  </tr>
</table>
</form>



</td>
</tr>
</table>
</div>
</center>
</body>

</html>
















