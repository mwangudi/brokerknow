<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE>Save File</TITLE>
</Head>
<Body>
<%
 
'save images here
				'  Variables
				'  *********
					Dim mySmartUpload
					Dim intCount
					filetext=Request.QueryString("filetext")					
				'  Object creation
					Set mySmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload")

				'  Upload
					mySmartUpload.Upload()						
				 'intCount = mySmartUpload.Save("Picture")%>
					<Script Language="JavaScript">	
						//window.alert("<%= intCount & " files uploaded." %>");						
					</Script>
					<%	
				Set mySmartUpload = Nothing
				Response.Redirect "upload.asp?filetext=" & filetext

%>
</body>
</html>