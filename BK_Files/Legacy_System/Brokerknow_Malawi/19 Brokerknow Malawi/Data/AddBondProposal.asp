<!--#include file="../libroutines.asp"-->
<%
	const UDLName = "KBroker"
	const DataSource = "EditOrder"
	const DataEntity = "Order"
	const DataEntityPlural = "Orders"
	const ActionFolder = "Operations"
	Dim ID
	
	ID = Request.QueryString("ID")
	action=Request.QueryString("action")
		
	If Trim(ID) = "" Then%>
            <script language = 'vbscript'>
                	ShowMessage "No record specified for editing"
                	
            </script>
            <% response.end
    End If	%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit <%=DataEntity%></title>
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css">
</head>

<body Class="Dialog" marginLeft=0 marginwidth=0 margintop=0 marginheight=0 SCROLLING=NO>
	<table cellspacing=0 cellpadding=0 style="left: 0;top: 0"><tr><td>
			<IFRAME FRAMEBORDER=0 marginwidth="0" marginheight="0" NAME="header" SCROLLING=yes SRC="AddBondProposal2.asp" height="430" width="600"></IFRAME>
	 </td></tr> 
</tr>
</table>
</body>

</html>