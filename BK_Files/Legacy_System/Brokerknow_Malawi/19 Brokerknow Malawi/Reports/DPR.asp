
<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Daily Processing Reconciliation</title>
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
	 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
	 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
	 <SCRIPT language=VBScript src="../scripts/reports.vbs"></SCRIPT>
	 <SCRIPT language=Javascript src="../scripts/reports.js"></SCRIPT>
	

<style media="print">
	
		@page {
			margin-left: 2cm;
			margin-right: 5cm;
			margin-top: 1cm;    
			margin-bottom: 2cm;
			writing-mode: tb-rl;
			height: 80%;
			margin: 10% 0%;						
			br.newpage{
				page-break-before:always;
			}		
		}		 
		
	</style>

</head>
<Script Language="JavaScript">
	
	function GoBack() {
		window.location.replace("DPR_List.asp");
	}
	
	
</Script>
<!--#include file="../libroutinesTEST.asp"-->
<body >
<%

pageToDisplay = trim(request("page"))
pageTitle = trim(request("pageTitle"))
'response.write pageToDisplay :response.end


%>

<table cellspacing=0 cellpadding=0 style="left: 0;top: 0" width ="100%">
	<tr>
		<td nowrap>
			<b><font face="Arial Narrow" size="4">
			<!--<%=pageTitle%>-->
			</font></b>
		</td>
		<td nowrap align=center>
			<INPUT type=Button  value="<< Go Back" name="GoBack" ID="GoBack" OnClick="JavaScript: GoBack();" >
		</td>
		<!--<td nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font>--></td>
	</tr>	
       <tr>
		  <td COLSPAN=3><font face="Arial" size="2">&nbsp;</font></td>
	</tr>
	<tr><td COLSPAN=3>
		<IFRAME FRAMEBORDER=0 marginwidth="10" marginheight="0" width="870" height="640" NAME="detail" id="detail" SCROLLING=yes SRC="<%=pageToDisplay%>" ></IFRAME>
	</td></tr>
</table>

</body>
</html>