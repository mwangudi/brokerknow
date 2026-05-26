<!--#include file="../libroutines.asp"-->
<%
	const UDLName = "KBroker"
	const DataSource = "EditJournal"
	const DataEntity = "Journal"
	const DataEntityPlural = "Journals"
	const ActionFolder = "Operations"
	
	Dim ID
	Dim idHeld
	
	ID = Request("ID")
	idHeld = Ucase(Trim(Request("IDHeld")))
	%>
            <script language = 'vbscript'>
                	'ShowMessage "<%=ID%>"
                	
            </script>
            <%

	If Trim(ID) = "" Then%>
            <script language = 'vbscript'>
                	ShowMessage "No record specified for editing"
                	
            </script>
            <% response.end
    End If	
    
    if idHeld <> "JOURNAL" then
			ID = GetOrderID(ID)
    end if
    
    Function GetOrderID(detailID)
		Dim getRs
		Set getConn = GetActiveConnection("KBroker")
		sqlStr = "SELECT Journal_DPA_ FROM " & DataEntity & "FullList WHERE " & _
                "  JournalEntry_DPA_ =" & detailID
                
        set getRs = getConn.Execute(sqlStr)
        If Not (getRs.EOF OR getRs.BOF) Then
			GetOrderID = getRs("Journal_DPA_")
		Else
			GetOrderID = 0
		End If	        
		
  End Function     
    
    %>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit <%=DataEntity%></title>
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css">
</head>

<body Class="Dialog" marginLeft=0 marginwidth=0 margintop=0 marginheight=0>
	<table cellspacing=0 cellpadding=0 style="left: 0;top: 0"><tr><td>
			<IFRAME FRAMEBORDER=0 marginwidth="0" marginheight="0" NAME="header" SCROLLING=no SRC="<%=DataSource%>Header.asp?ID=<%=ID%>" width="790px" height="150px"></IFRAME>
	 </td></tr>
	 
	 <tr><td>
			<IFRAME FRAMEBORDER=0 marginwidth="0" marginheight="0" NAME="detail" SCROLLING=no SRC="<%=DataSource%>Item2.asp?ID=<%=ID%>" width="790px" height="410px"></IFRAME>
</td>
</tr>
</table>
</body>

</html>