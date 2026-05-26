<!--#include file="../libroutines.asp"-->
<%
	const UDLName = "KBroker"
	const DataSource = "EditLot"
	const DataEntity = "Lot"
	const DataEntityPlural = "Lots"
	const ActionFolder = "Operations"
	
	Dim IDHolder
	Dim IDArray
	Dim ItemID
	Dim ID
	
	IDHolder = Request("ID")
	If (Trim(IDHolder) = "") or (IDHolder = "0") Then%>
			<script language = 'vbscript'>
                	ShowMessage "Please select an Order item for Lot allocation"
                	window.self.close
			</script>
			<%response.end
	End If
	
	IDArray = split(IDHolder,"<->")
	ID = IDArray(lbound(IDArray))
	ItemID = IDArray(ubound(IDArray))
	
	if (ItemID = 0) then
			Response.redirect "AddLot.asp?ID=" & IDHolder
			Response.end
	end if
	%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit <%=DataEntity%></title>
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css">
<script type="text/javascript" language="javascript">
<!--
function refreshIframe(){
    var iframeid = "detail";
    var h = parent.window.getElementById(iframeid).src;
    parent.window.getElementById(iframeid).src = h;
}
//-->
</script>
<script language="JavaScript">
			<!--
			function calcHeight()
			{
			  var the_height= document.getElementById('detail').contentWindow.document.body.scrollHeight;
			   
				the_height= the_height + 5000;
			  document.getElementById('detail').height = the_height;
			  
			}


			//-->
		</script>
</head>

<body Class="Dialog" marginLeft=0 marginwidth=0 margintop=0 marginheight=0>
<table cellspacing=0 cellpadding=0 style="left: 0;top: 0">
	<tr><td>
		<IFRAME FRAMEBORDER=0 marginwidth="0" marginheight="0" NAME="header" SCROLLING=NO SRC="<%=DataSource%>Header.asp?ID=<%=IDHolder%>" width="770px" height="250px"></IFRAME>
	</td></tr>
		 
	<tr><td>
		<IFRAME FRAMEBORDER=0 marginwidth="10" marginheight="0" NAME="detail" id="detail" SCROLLING=no SRC="<%=DataSource%>Item.asp?ID=<%=IDHolder%>" width="870px" height="1000px" onload=" calcHeight();"></IFRAME>
	</td></tr>
</table>
</body>

</html>
