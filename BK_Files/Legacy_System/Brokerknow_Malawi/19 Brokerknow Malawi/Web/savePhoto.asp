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
					
					'Response.write(filetext)
					'Response.end
										
					intCount = mySmartUpload.Save("Photos")					
					
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