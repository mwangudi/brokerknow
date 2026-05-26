<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN"><html>

<head><TITLE>BrokerKnow Header Bar</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=unicode">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0" >
<META HTTP-EQUIV="Expires" CONTENT="0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<LINK href="STYLE/default.css" type=TEXT/CSS rel=STYLESHEET> 
<LINK href="STYLE/webparts.css" type=TEXT/CSS rel=STYLESHEET>
<SCRIPT language=Javascript src="scripts/common.js"></SCRIPT>
<base target="maininfo">
</head>
<BODY>

<!--#include file="libroutines.asp"-->

<%

Dim Conn

Dim RS
Dim mnuCaption

UserID = Session("UserID")

%>

<table border="0" cellpadding="0" cellspacing="0" valign="top" width="100%">
  <TR>
	<TD align="left" valign=top noWrap width="60">
		<IMG height=59 alt="" src="images/helpeasy.jpg" width=59 align=textTop border=0>&nbsp;		
	</TD>
	<TD align="left" valign=center width="124">
      &nbsp;<IMG height=24 alt="" src="images/logo.jpg" width=109 border=0>		
	</TD>
	
	<TD align="left" valign=bottom nowrap width=200>
		<FONT face=Arial color=brown size=2><b ID="DataDescription">&nbsp;
			
		</b></FONT>
	</TD>
	
	<TD align="right" valign=top nowrap align="right" >
		<%
		 'the path should be relative to the data folder
		companyLogo = Session("companyLogo")
		companyName = Session("CompanyName")
		
		companyLogo = ""
		If 	companyLogo <> "" Then
			theImagePath = companyLogo

			'check existence of actual image file
			Dim Fso		 		
			Set Fso = Server.CreateObject("Scripting.FileSystemObject")
			If Fso.FileExists(Request.ServerVariables("APPL_PHYSICAL_PATH") & theImagePath) Then
				Response.Write "<IMG height=45 alt=""" & companyName & """ src=""" & theImagePath & """ border=0>"		
			End If
			Set Fso = Nothing
		End If%>
		<table>			
		<%Response.Write  "<tr><td cellpadding=""0""><font face=Arial size=2>Licensed to: </font></td><td cellpadding=""0""><b><font color='#000080'>" & companyName & "</font></b></font></td></tr>"
		Response.Write  "<tr><td cellpadding=""0""><font face=Arial size=2>Branch: </font></td><td cellpadding=""0""><b><font color='#000080'>" & Session("Branch") & "</font></b></font></td></tr>"
		Response.Write  "<tr><td cellpadding=""0""><font face=Arial size=2>User: </font></td><td cellpadding=""0""><b><font color='#000080'>" & Session("CurrentUser") & "</font></b></font></td></tr>"
		%>
		</table>
	</TD>
	
  </TR>
</table>

</BODY>

</html>
