<%@ Language=VBScript %>
<%
  Dim oConn,oRS,strConn,sSQLServer
  
  ' Build the connection string. Replace <username> and <strong password> with
  ' the username and password of an account that has permissions on the database.
  sSQLServer = "webserver"
  strConn = "Provider=SQLOLEDB.1;Persist Security Info=False;" & _
            "User ID=sa;Password=;Initial Catalog=pubs;Data Source=" & sSQLServer
  ' Set our return content type.
  Response.ContentType = "text/xml"
  
  ' Create a connection.
  set oConn = Server.CreateObject("ADODB.Connection")
  ' Open the connection.
  oConn.Open strConn
  ' Execute the SQL statement.
  set oRS = oConn.Execute(Request.QueryString("SQL"))
  ' Save the recordset in the Response object.
  oRS.Save Response,1
%>
