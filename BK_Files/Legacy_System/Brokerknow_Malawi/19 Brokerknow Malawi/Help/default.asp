<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2//EN">

<% @Language = "VBScript" %>

<html dir=ltr><head><title>Documentation</title>

<META NAME="ROBOTS" CONTENT="NOINDEX">
<META HTTP-EQUIV="Content-Type" content="text/html; charset=Windows-1252">
</head>


<% Set OBJbrowser = Server.CreateObject("MSWC.BrowserType")
	BrsType = Request.ServerVariables("HTTP_User-Agent")
	MachType=Request.ServerVariables("HTTP_UA-CPU")
	If InStr(BrsType, "MSIE") Then
		If InStr(BrsType, "Windows") Then
			File="contents.asp" 
			Size="30"
			Scroll="Auto"
		Else
			File="coflat.htm"
			Size="34"
			Scroll="Yes"
		End If
		If MachType="Alpha" Then
			File="contalph.asp"
			Size="30"
			Scroll="Auto" 
		End If
	Else

		File="coflat.htm"
		Size="34"
		Scroll="Yes"
	End If
%>

 
<%
        If Request.QueryString("jumpurl") <> "" Then
                strMainUrl = Request.QueryString("jumpurl")
        Else
                strMainUrl = "iiwltop.htm" 
        End If
 %>


<!--frameset cols="275,*"-->


<frameset rows="<% =Size%>,*" FRAMEBORDER="0" FRAMESPACING="0">
	<frame src="navbar.asp" name="NavBar" scrolling="No" noresize marginheight="0" marginwidth="0" framespacing="0" frameborder="No">
	<frameset cols="284,*">
        	<frame src=<% =File%> name="contents"  scrolling=<% =Scroll%> FRAMEBORDER="0" FRAMESPACING="0">
        	<frame src=<% =strMainUrl%> name="main" FRAMEBORDER="0" FRAMESPACING="0">
	</frameset>
</frameset>


<noframes>


<body bgcolor="#FFFFFF" text="#000000"><font face="Verdana,Arial,Helvetica">

</noframes>

</body>
</html>