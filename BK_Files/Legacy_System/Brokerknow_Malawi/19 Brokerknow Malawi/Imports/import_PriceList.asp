<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE>Price List File Upload</TITLE>
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT>
</head>
	<body>
	<!--#include virtual="libroutines.asp"-->
	<!--#include file="flUploader.asp"-->	

	<%
	Set Rs = Server.CreateObject("ADODB.Recordset")			
	Set Rs1 = Server.CreateObject("ADODB.Recordset")			
	Set Conn = Server.CreateObject("ADODB.Connection")
	Set Conn = GetActiveConnection("KBroker")

	' load object
	Dim load
	Set load = new Loader
				
	' calling initialize method
	load.initialize
				
	' File name
	Dim fileName, import_UDLPath
	fileName = LCase(load.getFileName("sourcefile"))			
	
	' Path where file will be uploaded
	Dim pathToFile, CurrentDirectory
	CurrentDirectory = "."	
		
	pathToFile = Server.MapPath(CurrentDirectory) & "\" & filename
		
	Dim fso
	Dim msgExists
	Dim SqlStr2
		
	Set fso = server.CreateObject("Scripting.FileSystemObject") 
		
	
	If (fso.FileExists(pathToFile) = True) Or (fso.FolderExists(pathToFile) = True) Then
		msgExists = "Could not upload file. A file or folder with that name already exists"
		fso.DeleteFile pathToFile
		Set fso = Nothing
		%>
		<SCRIPT LANGUAGE="JavaScript">
		<!--
			alert("<%=msgExists%>");
			window.history.back(0);
		//-->
		</SCRIPT>
		<%
		response.end
	Else		
		'validate file
		'	check for proper broker
		'Dim filNameOnly
			
		'filNameOnly = 
		' Uploading file data
		Dim fileUploaded
		fileUploaded = load.saveToFile ("sourcefile", pathToFile)
	End If

	'response.write pathToFile
	'response.end
		
	' destroy load object
	Set load = Nothing	
	
	' if the file is uploaded then
	If (fileUploaded = True) Then
		File = pathToFile '"C:\Folder\File.xls"
		
		''DATE
		''----------------------------------------------------------------------------------------------
		Dim xlApp
		Dim xlBook
		Dim xlSheet
		Dim theDate

		Set xlApp = Server.CreateObject("Excel.Application")
		
		xlApp.Visible = False
		
		Set xlBook = xlApp.Workbooks.Open(File)
		Set xlSheet = xlBook.Worksheets(1)

		theDate = xlSheet.Range("H1").Value

		xlBook.Close False
		xlApp.Quit

		Set xlBook = Nothing
		Set xlApp = Nothing
		
		'Response.Write theDate
		'Response.End 
		
		If IsNull(theDate) Or Len(theDate) = 0 Then
			theDate = Date()
		End If
		''----------------------------------------------------------------------------------------------
		
		Set objConn = Server.CreateObject("ADODB.Connection")
		objConn.Open "DBQ=" & File & ";" & _ 
		"DRIVER={Microsoft Excel Driver (*.xls)};" 
			
		Set objRS = Server.CreateObject("ADODB.Recordset")
		objRS.ActiveConnection = objConn
		objRS.CursorType = 3 'Static cursor.
		'objRS.Source = "Select * from [Sheet1$A1:C28]"
		objRS.Source = "Select * from [Price List$B11:G100]"
		objRS.Open
		%>
		<form name="frmPrices" action ="CommitPriceList.asp">
		<%
		Response.Write("<TABLE border='0' cellspacing=0 cellpadding=2 width='80%' align=center style=""border: 0 solid gray;"">")
			objRS.MoveFirst
			'Delete the previous entered records

			conn.execute("delete from _ImportPriceList")
			'open the recordset rs for inserts
			rs.open "Select * from _ImportPriceList where 1=1",conn,3,2
			conn.BeginTrans
			
			'Response.Write("<TR><TD style=""border: 1 solid gray;""><b>Security</b></TD><TD style=""border: 1 solid gray;""><b>Price</b></TD></TR>")
			
		    do until objRS.EOF
				Response.Write("<TR>")
				
				For X = 0 To objRS.Fields.Count - 1
					if X = 1 or X=5 then
						'Response.write("<TD>" & pc(objRS.Fields.Item(X).Value)) & "  sdfsd"
						if ucase(trim(mid(objRS.Fields.Item(X).Value,1,23))) = "FIXED INCOME SECURITIES" then
							exit do 
							exit for
						end if
					
						if trim(objRS.Fields.Item(0).Value)<>"" and X=5 then
							rs.addnew
							rs1.open "select security_DPA_ from security where LTrim(Rtrim(NSEName)) like '%" & ucase(trim(mid(objRS.Fields.Item(1).Value,1,15))) &"%'", conn, 0,1
							
							if not rs1.eof or not rs1.bof then
								secDPA = rs1("security_DPA_")
							else 
								secDPA =null
							end if
					
							rs1.close
							rs("secName")=trim(objRS.Fields.Item(1).Value)
							rs("price")=trim(objRS.Fields.Item(5).Value)
							rs("SecKnow_DPA_")=secDPA
							rs("importdate")=FormatDate(theDate)
							rs.update
						end if
						Response.write("<TD style=""border: 1 solid gray;"">&nbsp;" & objRS.Fields.Item(X).Value & "</TD>")
					else
						'Response.write("<TD style=""border: 1 solid gray;"">&nbsp;</TD>")
					end if
				Next
				
				objRS.MoveNext
				Response.Write("</TR>" & vbcrlf)
			loop
     
			set rs1= nothing
			rs.close
			set rs=nothing
			conn.CommitTrans       
 			objRS.close
			set objRS=nothing
			objConn.close
			set objConn=nothing
	end if
		
	%>
		<tr><td style="border: 1 solid gray;" colspan="2">
			<TABLE border=0 cellspacing=0 cellpadding=2 width=100% align=center style="border: 0 solid gray;">
				<%
				'Response.Write("<TR><TD><b>NSE Daily Price List ["& FormatDate(theDate) &"]</b></TD></TR>")
				On Error Resume Next
				Set objConn = Server.CreateObject("ADODB.Connection")

				objConn.Open "DBQ=" & File & ";" & "DRIVER={Microsoft Excel Driver (*.xls)};" 

				''BONDS
				Set objRS = Server.CreateObject("ADODB.Recordset")

				objRS.ActiveConnection = objConn
				objRS.CursorType = 3'Static cursor
				objRS.Source = "Select * from [Price List$H176:I178]"
				objRS.Open

				Response.Write("<TR><TD>Turnover in Bonds Today</TD></TR>")

				do until objRS.EOF
						Response.Write "<TR>"
						Response.write "<TD><b>" & FormatNum(objRS.Fields.Item(0).Value) & "</b><TD>"
						Response.write "<input type=hidden name=hidBond id=hidBond value="& objRS.Fields.Item(0).Value &">"
						Response.Write "</TR>" & vbCrLf 
					objRS.MoveNext
				loop
					     
				objRS.close
				set objRS=nothing

				''EQUITY
				Set objRS = Server.CreateObject("ADODB.Recordset")

				objRS.ActiveConnection = objConn
				objRS.CursorType = 3'Static cursor
				objRS.Source = "Select * from [Price List$H182:I184]"
				objRS.Open

				Response.Write("<TR><TD>Equity Turnover Today</TD></TR>")

				do until objRS.EOF
						Response.Write "<TR>"
						Response.write "<TD><b>" & FormatNum(objRS.Fields.Item(0).Value) & "</b><TD>"
						Response.write "<input type=hidden name=hidEquity id=hidEquity value="& objRS.Fields.Item(0).Value &">"
						Response.Write "</TR>" & vbCrLf 
					objRS.MoveNext
				loop
					     
				objRS.close
				set objRS=nothing

				objConn.close
				set objConn=nothing
				%>
			</table>
		<td><tr>
		
		<tr><td colspan="4">
		<input type="hidden" value="1" name="commit">
		<input type="hidden" value="<%=theDate%>" name="theDate">
		<input type="submit" name="submit" value="Commit Prices">
		<td><tr>
		</table>
		
	</form>
	<%
	
	 Response.Write("</TABLE>")
	
	'response.write "Heh, Heh"
	'response.end
	%>
	</body>
</html>


