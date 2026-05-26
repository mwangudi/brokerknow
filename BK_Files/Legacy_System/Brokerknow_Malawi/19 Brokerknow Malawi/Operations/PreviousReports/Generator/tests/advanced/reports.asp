<%

If Session("UserName")="" Then
	Response.Redirect "login.asp"
End If

Dim objWiz 
Set objWiz = Server.CreateObject("AspWebSolution.ReportWizard2") 

With objWiz
	.LincenseKey = "ABCD-EFGH-IJKL-MNOP"	
	If CBool(Session("IsAdmin")) Then
		.IsAdmin = True
	Else
		.CanAdd = CBool(Session("CanAdd"))
		.CanEdit = CBool(Session("CanEdit"))
		.CanDelete = CBool(Session("CanDelete"))
	End If
	
	If Session("Departments")<>"" Then
		.UseReports = Session("Departments")
	End If
	
	If CBool(Session("UseImages")) Then
		.LocaleFile = "locale/locale_img.xml"
	Else
		.LocaleFolder = "locale"
		.AutoDetectLocale = True
	End If
	
	.ImagesUrl = "images"
	.Title = "ACME Reporting Solution"	
	.HideProcedures = True
	.ReportsFile = Server.MapPath("../data/sample.xml")
	'.Width = "90%"
End With

'Make sure you initialize before adding any HTML text
objWiz.Init("PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & Server.MapPath("../data/sample.mdb") & ";") 
    
%>

<html>
	<head>
		<title>Advanced Reporting</title>
	</head>
	<body scroll="auto">
		<div align=right style="font-size:8pt;">
			<a href="default.asp"><font color=gray face=tahoma>Logout (<%=Session("UserName")%>)</font></a>
		</div>
		<%
			'put this where you want your reports 
			objWiz.DisplayReports 
		%>
	</body>
</html>

<%    
Set objWiz = Nothing
%>
