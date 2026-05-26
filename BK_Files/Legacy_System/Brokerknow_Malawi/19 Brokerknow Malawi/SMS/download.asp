<% 
    Response.ContentType = "application/x-unknown" ' arbitrary 
    'Response.ContentType = "application/octet-stream"
    
    fn = Request.QueryString("f")
    
    FPath = server.MapPath(".") & "\bin\" & fn
    
    'Response.AddHeader "Content-Disposition", "attachment; filename=" & fn & ""
	
	Response.AddHeader "Content-Disposition", "attachment; filename=""" & fn & """"

    Set adoStream = CreateObject("ADODB.Stream") 
    adoStream.Open() 
    adoStream.Type = 1 
    adoStream.LoadFromFile(FPath) 
    Response.BinaryWrite adoStream.Read() 
    adoStream.Close 
    Set adoStream = Nothing 
	
	'Response.WriteFile(fn)
	'Response.BinaryWrite fn
	
	'Response.Flush 	
    Response.End 
    
%>
