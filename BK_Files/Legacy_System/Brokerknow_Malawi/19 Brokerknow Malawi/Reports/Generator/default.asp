<!--#include file="../../libroutines.asp"-->

<%

Session("UserName") = Session("UserID")
Session("IsAdmin") = True
Session("CanAdd") = True 'False
Session("CanEdit") = True 'False
Session("CanDelete") = True 'False

Session("UseImages") = True

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
		.AutoDetectLocale = False
	End If
	
	.ImagesUrl = "images"
	.Title = "BrokerKnow Reporting Sub-system"	
	.HideProcedures = True
	
	.ReportsFile = Server.MapPath("data/gen_reports.xml")
	
	'.HideTables = True
	'.HideViews = True
	
End With

Set Conn = GetActiveConnection("KBroker")

'Make sure you initialize before adding any HTML text
objWiz.Init  Conn.ConnectionString

Set Conn = Nothing
    
%>

<html>
	<head>
		<title>BrokerKnow Advanced Reporting Sub-system</title>		
		<SCRIPT language=Javascript src="../../scripts/common.js"></SCRIPT> 
		<SCRIPT language=Javascript src="../../scripts/fhsupport.js"></SCRIPT>
	
	</head>
	<body scroll="auto">
		
		<%
			'put this where you want your reports 
			objWiz.DisplayReports 
		%>
	</body>
</html>

<Script Language="JavaScript">
	window.onload = doMyReportAction;
	
	function doMyReportAction(){
	
	try{
		var doc = document.getElementsByTagName('input').item('__FARGS');
		var nextDoc = document.getElementsByTagName('input').item('__SQL');
		  
		//alert(nextDoc.outerHTML) 
		var currBodyHTML = document.body.innerHTML;
		
		var beginCut = currBodyHTML.indexOf(doc.outerHTML);
		var endCut = currBodyHTML.indexOf(nextDoc.outerHTML);
		var resultBodyHTML1 = currBodyHTML.substr(0, beginCut);
		var resultBodyHTML2 = currBodyHTML.substr(endCut);
		
		var resultBodyHTML = resultBodyHTML1 + resultBodyHTML2;
		
		//document.body.innerHTML = resultBodyHTML; 
		
		//alert(resultBodyHTML)
		//document.body.insertAdjacentHTML(resultBodyHTML, "before");	
	}
	
	catch(e){}
	}

</Script>

<%    
Set objWiz = Nothing
%>
