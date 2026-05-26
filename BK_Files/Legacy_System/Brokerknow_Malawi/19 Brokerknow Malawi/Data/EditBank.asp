<!--#include file="../libroutines.asp"-->
<%
	const UDLName = "KBroker"
		const DataSource = "BankList"
		const DataEntity = "Bank"
		const DataEntityPlural = "Banks"
		const ActionFolder = "Data"
	
	Dim ID
	
	ID = Request("ID")

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
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit <%=DataEntity%></title>
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css">
</head>


<body Class="Dialog" leftMargin=0 marginwidth=0 topMargin=0 marginheight=0 scroll=No>
	<table cellspacing=0 cellpadding=0 style="left: 0;top: 0"><tr><td>
			<IFRAME FRAMEBORDER=0 marginwidth="0" marginheight="0" NAME="header" SCROLLING=NO SRC="<%=DataSource%>Header.asp?ID=<%=ID%>" width="420px" height="110px"></IFRAME>
	 </td></tr>	 
	 <tr><td>
			<IFRAME FRAMEBORDER=0 marginwidth="0" marginheight="0" NAME="detail" SCROLLING=yes SRC="<%=DataSource%>Item.asp?ID=<%=ID%>" width="420px" height="170px"></IFRAME>
</td>
</tr>
</table>
</body>

</html>
