<%

Dim objWiz 
Set objWiz = Server.CreateObject("AspWebSolution.ReportWizard2") 
    
'Make sure you initialize before adding any HTML text
objWiz.Init("PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & Server.MapPath("../data/sample.mdb") & ";") 
%>

<html>
	<head>
		<title>Simple Repoprts</title>
	</head>
	<body scroll="auto">
		<%
			'put this where you want your reports 
			objWiz.DisplayReports 
		%>
	</body>
</html>

<%    
Set objWiz = Nothing
%>
