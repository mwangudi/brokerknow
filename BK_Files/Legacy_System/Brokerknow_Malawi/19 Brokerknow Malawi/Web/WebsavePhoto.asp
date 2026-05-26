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
					intCount = mySmartUpload.Save("../Data/Picture")
					'Response.write(filetext)
					'Response.end

					%>
					<Script Language="JavaScript">	
						//window.alert("<%= intCount & " files uploaded." %>");						
					</Script>
					<%	
				Set mySmartUpload = Nothing
				Response.Redirect "Webupload.asp?filetext=" & filetext

%>
</body>
</html>