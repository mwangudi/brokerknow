<!--#include file="../libroutines.asp"-->
<%
		const UDLName = "KBroker"
		const DataSource = "StockWatchList"
		const DataEntity = "StockWatchList"
		const DataEntityPlural = "StockWatchLists"
		const ActionFolder = "Operations"
		
	
	Dim ID
	Dim idHeld
	
	ID = Request("ID")
	ID1= split(ID,"-")
	itemId= ID1(0)
	clientID =ID1(1)
	If Trim(clientID) = "" Then%>
            <script language = 'vbscript'>
                	ShowMessage "No record specified for editing"
                	
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
<title>Edit <%=DataEntity%></title>
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css">
</head>

<body Class="Dialog" marginLeft=0 marginwidth=0 margintop=0 marginheight=0>
	<table cellspacing=0 cellpadding=0 style="left: 0;top: 0">
	 <tr><td>
			<IFRAME FRAMEBORDER=0 marginwidth="0" marginheight="0" NAME="detail" SCROLLING=auto SRC="EditStockWatchItem.asp?ID=<%=clientID%>" width="600px" height="810px"></IFRAME>
</td>
</tr>
</table>
</body>

</html>