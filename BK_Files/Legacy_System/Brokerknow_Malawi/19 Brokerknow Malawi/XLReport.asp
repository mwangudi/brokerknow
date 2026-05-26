<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2//EN">

<html dir=ltr><head><title>BrokerKnow Report Printing Framework ©2004</title>

<META NAME="ROBOTS" CONTENT="NOINDEX">
<META HTTP-EQUIV="Content-Type" content="text/html; charset=Windows-1252">
<LINK REL="stylesheet" TYPE="text/css" HREF="style/default.css">
</head>

<%
	xlFile = Request.QueryString("xlFile")
	
	If xlFile = "" Then%>
		<Script Language="JavaScript">
			window.parent.self.close();
		</Script>
		<%
		Response.End
	End If
%>


<!--frameset rows="*"-->



<frameset rows="*" FRAMEBORDER="0" FRAMESPACING="0">
	<frameset rows="60, *">
        	<frame src="XLReportHeader.asp" name="contentsR" scrolling="NO"  FRAMEBORDER="0" FRAMESPACING="0">
        	<frame src="<%= xlFile %>" name="maininfoR" scrolling=auto marginheight="0" marginwidth="0" FRAMEBORDER="No" FRAMESPACING="0">        	        	      		        		
     </frameset>	
</frameset>
	
	


    <noframes>


<body bgcolor="#FFFFFF" text="#000000">

</body>
    </noframes>
</frameset>


<frameset>


<noframes>


<body bgcolor="#FFFFFF" text="#000000"><font face="Verdana,Arial,Helvetica">

</noframes>

</frameset>
</frameset>
</html>

